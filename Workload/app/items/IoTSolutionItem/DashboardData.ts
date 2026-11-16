// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { kqlString } from "./HubModeling.ts";
import { quoteKqlIdentifier, type SampleStream } from "./ModelAuthoring.ts";

export interface DashboardField {
  key: string;
  isNumeric?: boolean;
  included: boolean;
  component?: string;
  leaf?: string;
}
export interface DashboardTelemetryField extends DashboardField { isNumeric: boolean }
export const DASHBOARD_NUMERIC_TYPES = '("int","long","real","decimal","double")';
const visibleKey = (key: string) => `isnotempty(trim(" \\t\\r\\n", ${key})) and ${key} !in ("$metadata", "$version", "__t")`;

function dashboardIdentity(stream: SampleStream): string[] {
  return stream === "telemetry" ? [
    "| extend deviceId = tostring(user_headers.cloudEvents_subject)",
  ] : [
    "| extend deviceId = tostring(user_headers.cloudEvents_deviceid), moduleId = tostring(user_headers.cloudEvents_moduleid)",
    '| extend deviceId = iff(isempty(moduleId), deviceId, strcat(deviceId, "/", moduleId))',
  ];
}

function dashboardSource(table: string, stream: SampleStream) {
  const raw = table === "dashboard_raw" ? "dashboard_raw_1" : "dashboard_raw";
  const events = table === "dashboard_events" ? "dashboard_events_1" : "dashboard_events";
  const expectedKind = kqlString(stream === "telemetry" ? "measurement" : "lifecycle");
  const declarations = [
    `let ${raw} = (`,
    quoteKqlIdentifier(table),
    "| where ingestion_time() >= ago(30d)",
    "| project data, user_headers",
    "| extend data = parse_json(data)",
    '| extend dashboard_error = case(',
    'gettype(data) != "dictionary", "Raw dashboard data must contain one JSON object per row.",',
    '(gettype(data.schema_id) != "string" or data.schema_id != "com.do.telemetry.observation.v1" or ' +
      'gettype(data.schema_version) != "string" or data.schema_version != "1.0" or ' +
      'gettype(data.contract_version) != "string" or data.contract_version != "1.0"), ' +
      '"Unsupported DO observation schema or version; expected DOObservationV1 schema/contract 1.0.",',
    `tostring(data.event_kind) != ${expectedKind}, "The DO event_kind does not match the selected dashboard stream.",`,
    '(gettype(data.event_time) !in ("string", "datetime") or isnull(todatetime(data.event_time))), "The DO observation has an invalid event_time.",',
    '(gettype(data.payload) != "dictionary" or gettype(data.payload.value) != "dictionary"), ' +
      '"The DO observation must contain an object at payload.value.",',
    `gettype(user_headers.${stream === "telemetry" ? "cloudEvents_subject" : "cloudEvents_deviceid"}) != "string" or ` +
      `isempty(tostring(user_headers.${stream === "telemetry" ? "cloudEvents_subject" : "cloudEvents_deviceid"})), ` +
      `"The selected ${stream === "telemetry" ? "telemetry" : "properties"} source must provide CloudEvents ${stream === "telemetry" ? "subject" : "device"} identity.",`,
    ...(stream === "telemetry" ? [
      'gettype(user_headers.cloudEvents_iothubdtsubject) !in ("null", "string"), "Telemetry component metadata must be a string or null.",',
    ] : [
      'gettype(user_headers.cloudEvents_moduleid) !in ("null", "string"), "Properties module metadata must be a string or null.",',
    ]),
    '"")',
    ");",
    `let ${events} = ${raw}`,
    "| where isempty(dashboard_error)",
    "| extend event_time = todatetime(data.event_time), data = data.payload.value",
    ...dashboardIdentity(stream),
    `| extend component = ${stream === "telemetry" ? "tostring(user_headers.cloudEvents_iothubdtsubject)" : '""'};`,
  ].join("\n");
  return { declarations, raw, events };
}

/** Short runtime KQL for the inspected DO/CloudEvents contract; diagnostics stay in setup. */
export function dashboardRawQuery(table: string, stream: SampleStream, component?: string): string {
  return [
    quoteKqlIdentifier(table),
    "| where ingestion_time() between (['_startTime'] .. ['_endTime'])",
    ...dashboardIdentity(stream),
    "| where deviceId == _deviceId",
    ...(stream === "telemetry" && component !== undefined
      ? [`| where tostring(user_headers.cloudEvents_iothubdtsubject) == ${kqlString(component)}`] : []),
    "| extend data = parse_json(data)",
    "| extend event_time = todatetime(data.event_time)",
  ].join("\n");
}

export function dashboardFieldQuery(table: string, stream: SampleStream): string {
  const source = dashboardSource(table, stream);
  const fields = stream === "telemetry" ? [
    "| extend j = data",
    "| extend keys = bag_keys(j)",
    "| mv-expand k = keys to typeof(string)",
    `| where ${visibleKey("k")}`,
    "| extend v = j[k]",
    "| extend valueType = gettype(v)",
    '| extend Key = iff(component == "", k, strcat(component, "_", k))',
    "| summarize arg_max(event_time, valueType) by Key, Component = component, Leaf = k",
    `| extend IsNumericType = valueType in ${DASHBOARD_NUMERIC_TYPES}`,
    '| project Key, IsNumericType, Component, Leaf, Error = ""',
  ] : [
    "| extend rp = data.properties.reported",
    "| extend keys = bag_keys(rp)",
    "| mv-expand key = keys to typeof(string)",
    `| where ${visibleKey("key")} and key !startswith "$" and key != "iothub-enqueuedtime" and key != "iothub-connection-device-id"`,
    "| extend val = rp[key]",
    '| extend isComponent = (gettype(val) == "dictionary" and tostring(val["__t"]) == "c")',
    "| mv-expand subkey = iff(isComponent, bag_keys(val), pack_array(key)) to typeof(string)",
    `| where ${visibleKey("subkey")} and subkey !startswith "$"`,
    '| extend Component = iff(isComponent, key, ""), Leaf = subkey',
    '| extend Key = iff(isComponent, strcat(key, "_", subkey), key)',
    "| distinct Key, Component, Leaf",
    '| project Key, Component, Leaf, Error = ""',
  ];
  return [
    source.declarations,
    "union (",
    source.events,
    ...fields,
    "), (",
    `${source.raw} | where isnotempty(dashboard_error)`,
    "| distinct Error = dashboard_error",
    `| project Key = "", ${stream === "telemetry" ? "IsNumericType = false, " : ""}Component = "", Leaf = "", Error`,
    ")",
    "| order by Key asc",
  ].join("\n");
}

export function readDashboardFields(rows: Record<string, unknown>[], stream: "telemetry"): DashboardTelemetryField[];
export function readDashboardFields(rows: Record<string, unknown>[], stream: "twin"): DashboardField[];
export function readDashboardFields(rows: Record<string, unknown>[], stream: SampleStream): DashboardField[] {
  const errors = [...new Set(rows.filter(row => typeof row.Error === "string" && row.Error).map(row => row.Error))];
  if (errors.length) throw new Error(`${stream === "telemetry" ? "Telemetry" : "Properties"}: ${errors.join(" ")}`);
  const seen = new Set<string>();
  const fields = rows.map(row => {
    if (typeof row.Key !== "string" || !row.Key.trim() || typeof row.Component !== "string" ||
        typeof row.Leaf !== "string" || !row.Leaf.trim() || row.Error !== "" ||
        row.Key !== (row.Component ? `${row.Component}_${row.Leaf}` : row.Leaf)) {
      throw new Error("The database returned unsupported dashboard field metadata.");
    }
    const identity = JSON.stringify([row.Component, row.Leaf]);
    if (seen.has(identity)) throw new Error("The database returned duplicate dashboard field identities.");
    seen.add(identity);
    if (stream === "telemetry" && row.IsNumericType !== true && row.IsNumericType !== false && row.IsNumericType !== 1 && row.IsNumericType !== 0) {
      throw new Error("The database returned an invalid dashboard numeric classification.");
    }
    return { key: row.Key, included: true, ...(stream === "telemetry" ? {
      isNumeric: row.IsNumericType === true || row.IsNumericType === 1,
    } : {}), component: row.Component || undefined, leaf: row.Leaf };
  });
  const counts = new Map<string, number>();
  for (const field of fields) counts.set(field.key, (counts.get(field.key) || 0) + 1);
  const used = new Set(fields.filter(field => counts.get(field.key) === 1).map(field => field.key));
  for (const field of fields) if (counts.get(field.key)! > 1) {
    const identity = JSON.stringify([field.component || "", field.leaf]);
    let key = field.component ? `${JSON.stringify(field.component)}[${JSON.stringify(field.leaf)}]` : `[${JSON.stringify(field.leaf)}]`;
    while (used.has(key)) key += ` ${identity}`;
    field.key = key; used.add(key);
  }
  return fields;
}
