// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { MODEL_SEMANTIC_INFO_PREFIX, parseIotMeta, quoteKqlIdentifier, type AuthoringSource } from "./ModelAuthoring.ts";
import {
  coverageReady, scopedDoProjection, fieldValidation, kqlString, modelValidation, orderedModeledFields, requireSingleObservationRows, verifiedEmptyCoverage,
  type CoverageSnapshot, type HubModel, type ModelField, type ModelingInput,
} from "./HubModeling.ts";
import { buildDataAgentContext, type DataAgentContext } from "./DataAgentContext.ts";
import { columnDocstring, publishedFieldFacts } from "./FieldSemantics.ts";
import { docstringJson, emitJsonDocstring } from "./Docstrings.ts";

export const SYSTEM_EVENT_TIME_DOCSTRING = columnDocstring({ name: "event_time", description: "Event timestamp" });

export interface SemanticDocumentation {
  table: string;
  columns: Record<string, string>;
}
export interface OutputColumn { name: string; type: string }
export type CreationCommand = {
  label: string;
  csl: string;
  target: string;
} & (
  | { kind: "table"; columns: OutputColumn[] }
  | { kind: "view"; source: string; query: string; columns: OutputColumn[] }
  | { kind: "policy"; source: string; query: string }
  | { kind: "tableDoc" | "viewDoc"; docstring: string }
  | { kind: "columnDocs" }
);
export interface ModeledDataPlan {
  names: { modeled: string; twinChanges: string; twinState: string };
  commands: CreationCommand[];
  documentation: SemanticDocumentation;
  dataAgent: DataAgentContext;
  prerequisites: string[];
}
export interface ModelingDocstringContext { source: AuthoringSource; names: { twinState: string } }
function modelObjectDocstring(model: HubModel, description: string, controls: string[]): string {
  const metadata = model.declaredModelMetadata || (model.id ? { [model.kind === "wot" ? "id" : "@id"]: model.id } : {});
  const identity = metadata[model.kind === "wot" ? "id" : "@id"];
  if (identity !== (model.id || undefined)) {
    throw new Error("The declared model semantic identity no longer matches the selected root. Upload the model again.");
  }
  parseIotMeta(MODEL_SEMANTIC_INFO_PREFIX + docstringJson(metadata));
  const docstring = emitJsonDocstring(MODEL_SEMANTIC_INFO_PREFIX, metadata, {
    before: [description],
    after: controls,
    requiredKeys: ["id", "@id", "version"],
  });
  parseIotMeta(docstring);
  return docstring;
}
export function semanticDocumentation(model: HubModel, fields: ModelField[], context?: ModelingDocstringContext): SemanticDocumentation {
  const selected = orderedModeledFields(fields);
  const columns: Record<string, string> = Object.create(null);
  columns.deviceId = columnDocstring({ name: "deviceId", description: "IoT Hub device Id" });
  columns.event_time = SYSTEM_EVENT_TIME_DOCSTRING;
  for (const field of selected) {
    columns[field.column] = columnDocstring(publishedFieldFacts(field));
  }
  return {
    table: modelObjectDocstring(model, "Modeled device telemetry enriched with property LKVs", [
      "@iot.meta tableRole=modeled-data",
      ...(context ? [`@iot.meta source=${modelingSource(context.source, context.source.telemetryTable)}; ${modelingSource(context.source, context.names.twinState)}`] : []),
    ]),
    columns,
  };
}
export function modelingSource(source: AuthoringSource, table: string): string {
  return `${source.queryServiceUri.replace(/^https?:\/\//, "").replace(/\/$/, "")}/${source.databaseName}/${table}`;
}
export function helperObjectDocstrings(model: HubModel, source: AuthoringSource, normalizedTable: string) {
  return {
    normalized: modelObjectDocstring(model, "Normalized device property changes in Entity-Attribute-Value (EAV) format", [
      "@iot.meta tableRole=properties-eav",
      `@iot.meta source=${modelingSource(source, source.twinTable)}`,
    ]),
    lkv: modelObjectDocstring(model, "Property last known values", [
      "@iot.meta viewRole=properties-lkv",
      `@iot.meta source=${modelingSource(source, normalizedTable)}`,
    ]),
  };
}

const pathExpression = (root: string, path: string[]): string =>
  root + path.map(key => `[${kqlString(key)}]`).join("");
function telemetryValue(field: ModelField): string {
  return `iff(${field.component ? `component == ${kqlString(field.component)}` : "isempty(component)"}, ` +
    `${pathExpression("payload", field.path)}, dynamic(null))`;
}
function cast(type: string, expression: string): string {
  const casts: Record<string, string> = {
    string: "tostring", real: "todouble", int: "toint", long: "tolong",
    datetime: "todatetime", bool: "tobool", dynamic: "todynamic",
  };
  if (!casts[type]) throw new Error(`Unsupported output type: ${type}.`);
  return `${casts[type]}(${expression})`;
}
function updatePolicy(target: string, source: string, query: string): CreationCommand {
  return { kind: "policy", target, source, query, label: `Set update policy on ${target}`,
    csl: `.alter table ${quoteKqlIdentifier(target)} policy update ${kqlString(JSON.stringify([{
      IsEnabled: true, Source: source, Query: query, IsTransactional: false, PropagateIngestionProperties: false,
    }]))}` };
}
function emptyTable(target: string, columns: OutputColumn[]): CreationCommand {
  return { kind: "table", target, columns, label: `Create ${target}`,
    csl: `.set ${quoteKqlIdentifier(target)} <|\ndatatable(` +
      columns.map(column => `${quoteKqlIdentifier(column.name)}:${column.type}`).join(", ") + ")[]" };
}

export function buildModeledDataPlan(
  input: ModelingInput, snapshot: CoverageSnapshot, waiting: boolean,
  names: ModeledDataPlan["names"]
): ModeledDataPlan {
  const modelErrors = modelValidation(input.model);
  if (modelErrors.length) throw new Error(modelErrors.join(" "));
  if (snapshot) for (const stream of ["telemetry", "twin"] as const) {
    requireSingleObservationRows(snapshot[stream], stream);
    if (snapshot.sourceProbe) requireSingleObservationRows(snapshot.sourceProbe[stream], stream);
  }
  if (!coverageReady(input, snapshot, waiting)) throw new Error("Verify the current processing scope before creating modeled data.");
  if (!snapshot.telemetry.payloads.length && !verifiedEmptyCoverage(input, snapshot)) {
    throw new Error("Verify nonempty, valid raw-source samples in both streams before creating modeled data with no matching telemetry.");
  }
  const errors = fieldValidation(input.fields);
  if (errors.length) throw new Error(errors.join(" "));
  const allNames = [names.modeled, names.twinChanges, names.twinState];
  if (allNames.some(name => !/^[a-zA-Z_]\w*$/.test(name)) || new Set(allNames).size !== allNames.length ||
      allNames.some(name => name === input.source.telemetryTable || name === input.source.twinTable)) {
    throw new Error("Use distinct output names starting with a letter or underscore, followed by letters, digits or underscores. Do not reuse a raw table name.");
  }
  const selected = orderedModeledFields(input.fields);
  const twins = selected.filter(field => field.role === "twin");
  const occupiedNames = new Set([...allNames, input.source.telemetryTable, input.source.twinTable]);
  let lkvAlias = "lkv_bag";
  for (let suffix = 1; occupiedNames.has(lkvAlias); suffix++) lkvAlias = `lkv_bag_${suffix}`;
  const commands: CreationCommand[] = [];
  const hasComponents = selected.some(field => field.component);
  commands.push(emptyTable(names.twinChanges, [
    { name: "deviceId", type: "string" }, { name: "propertyName", type: "string" },
    { name: "propertyValue", type: "dynamic" }, { name: "event_time", type: "datetime" },
  ]));
  const twinQuery = [
    quoteKqlIdentifier(input.source.twinTable), scopedDoProjection(input.scope, "twin"),
    "| where isnotempty(deviceId) and isnotnull(event_time)",
    "| extend rp = payload",
    "| extend keys = bag_keys(rp)",
    "| mv-expand key = keys to typeof(string)",
    '| where key !startswith "$" and key != "iothub-enqueuedtime" and key != "iothub-connection-device-id"',
    ...(hasComponents ? [
      '| extend val = rp[key]',
      '| extend isComponent = (gettype(val) == "dictionary" and tostring(val["__t"]) == "c")',
      "| mv-expand subkey = iff(isComponent, bag_keys(val), pack_array(key)) to typeof(string)",
      '| where subkey != "__t"',
      '| project deviceId, propertyName = iff(isComponent, strcat(key, "_", subkey), key), propertyValue = iff(isComponent, val[subkey], val), event_time',
    ] : ["| project deviceId, propertyName = key, propertyValue = rp[key], event_time"]),
  ].join("\n");
  const twinPolicy = updatePolicy(names.twinChanges, input.source.twinTable, twinQuery);
  const stateQuery = `${quoteKqlIdentifier(names.twinChanges)}\n| where ingestion_time() > ago(30d)\n` +
    "| summarize arg_max(event_time, propertyValue) by deviceId, propertyName";
  commands.push({ kind: "view", target: names.twinState, source: names.twinChanges, query: stateQuery,
    columns: [{ name: "deviceId", type: "string" }, { name: "propertyName", type: "string" },
      { name: "event_time", type: "datetime" }, { name: "propertyValue", type: "dynamic" }],
    label: `Create ${names.twinState}`,
    csl: `.create materialized-view with (backfill=false) ${quoteKqlIdentifier(names.twinState)} on table ${quoteKqlIdentifier(names.twinChanges)} {\n` +
      `${stateQuery}\n}` });
  commands.push(emptyTable(names.modeled, [
    { name: "deviceId", type: "string" }, { name: "event_time", type: "datetime" },
    ...selected.map(field => ({ name: field.column, type: field.kustoType })),
  ]));
  const projected = selected.map(field => `${quoteKqlIdentifier(field.column)}=${cast(field.kustoType,
    field.role === "telemetry" ? telemetryValue(field)
      : `__state[${kqlString(field.component ? `${field.component}_${field.path[0]}` : field.path[0])}]`)}`);
  const query = [
    ...(twins.length ? [`let ${lkvAlias} = ${quoteKqlIdentifier(names.twinState)} | summarize __state=make_bag(bag_pack(propertyName, propertyValue)) by deviceId;`] : []),
    quoteKqlIdentifier(input.source.telemetryTable), scopedDoProjection(input.scope, "telemetry"),
    "| where isnotempty(deviceId) and isnotnull(event_time)",
    ...(twins.length ? [`| lookup kind=leftouter ${lkvAlias} on deviceId`] : []),
    `| project deviceId, event_time, ${projected.join(", ")}`,
  ].join("\n");
  const documentation = semanticDocumentation(input.model, input.fields, { source: input.source, names });
  const helperDocs = helperObjectDocstrings(input.model, input.source, names.twinChanges);
  if (commands.some(command => command.kind === "table" && command.target === names.twinChanges)) {
    const docstring = helperDocs.normalized;
    commands.push({ kind: "tableDoc", target: names.twinChanges, label: `Set docstring on ${names.twinChanges}`, docstring,
      csl: `.alter table ${quoteKqlIdentifier(names.twinChanges)} docstring ${kqlString(docstring)}` });
  }
  if (commands.some(command => command.kind === "view" && command.target === names.twinState)) {
    const docstring = helperDocs.lkv;
    commands.push({ kind: "viewDoc", target: names.twinState, label: `Set docstring on ${names.twinState}`, docstring,
      csl: `.alter materialized-view ${quoteKqlIdentifier(names.twinState)} docstring ${kqlString(docstring)}` });
  }
  commands.push({ kind: "tableDoc", target: names.modeled, label: "Set ModeledData meaning",
    docstring: documentation.table,
    csl: `.alter table ${quoteKqlIdentifier(names.modeled)} docstring ${kqlString(documentation.table)}` });
  commands.push({ kind: "columnDocs", target: names.modeled, label: "Set native column meanings",
    csl: `.alter table ${quoteKqlIdentifier(names.modeled)} column-docstrings (` +
      Object.entries(documentation.columns).map(([column, doc]) => `${quoteKqlIdentifier(column)}:${kqlString(doc)}`).join(", ") + ")" });
  commands.push(twinPolicy);
  commands.push(updatePolicy(names.modeled, input.source.telemetryTable, query));
  const prerequisites = [
    "The creation operation must verify unused output names, permissions, ordered schemas and ownership before installing policies.",
    "Native device/module bindings come from transport metadata, not ADR references or payload identifiers.",
    "New incoming telemetry is processed; raw history is not rewritten.",
    "Property LKV uses the latest whole received value in the V1 30-day ingestion window. Omitted properties retain their LKV; complex values are not recursively merged.",
    ...snapshot.nativeIssues,
    ...(twins.length ? [
      "Device IDs are assumed unique across Hubs. The same native device/module identity correlates telemetry and last-known twin state.",
      "Verify source streaming-ingestion policy supports lookup/join enrichment; do not change raw ingestion policy implicitly.",
      "Validate materialized-view/update-policy permissions and schema compatibility before enabling the nontransactional policies.",
      "Twin state is last-known state available at processing time, not an event-time/as-of reconstruction.",
    ] : []),
  ];
  return { names, commands, documentation, prerequisites,
    dataAgent: buildDataAgentContext({
      database: input.source.databaseName, table: names.modeled, docstring: documentation.table,
      columns: [
        { name: "deviceId", type: "string", docstring: documentation.columns.deviceId },
        { name: "event_time", type: "datetime", docstring: documentation.columns.event_time },
        ...selected.map(field => ({ name: field.column, type: field.kustoType, docstring: documentation.columns[field.column] })),
      ],
    }) };
}

export function modeledDataScript(plan: ModeledDataPlan): string {
  return [...plan.prerequisites.map(prerequisite => `// ${prerequisite}`), "",
    ...plan.commands.flatMap(command => [`// ${command.label}`, command.csl, ""])].join("\n");
}
