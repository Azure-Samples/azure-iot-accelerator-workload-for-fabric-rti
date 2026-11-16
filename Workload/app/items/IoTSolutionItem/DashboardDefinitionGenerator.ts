// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Generates a Fabric Real-Time Dashboard definition JSON from inspected
 * telemetry and property fields. The output follows the Kusto dashboard
 * schema (version 88) and can be posted to the Fabric REST API.
 */

import { dashboardRawQuery, type DashboardField } from "./DashboardData.ts";
import { kqlString } from "./HubModeling.ts";
import { quoteKqlIdentifier } from "./ModelAuthoring.ts";
export type { DashboardField } from "./DashboardData.ts";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface DashboardGenerationParams {
  title: string;
  telemetryTable: string;
  propertiesTable: string;
  telemetryFields: DashboardField[];
  propertyFields: DashboardField[];
  queryServiceUri: string;
  databaseId: string;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function uuid(): string {
  return crypto.randomUUID();
}

/** Color scheme values accepted by the dashboard schema: color--(0) through color--(28) */
const CHART_COLORS = [
  "color--(0)",
  "color--(1)",
  "color--(2)",
  "color--(3)",
  "color--(4)",
  "color--(5)",
  "color--(6)",
  "color--(7)",
  "color--(8)",
  "color--(9)",
  "color--(10)",
  "color--(11)",
  "color--(12)",
  "color--(13)",
  "color--(14)",
  "color--(15)",
  "color--(16)",
  "color--(17)",
  "color--(18)",
  "color--(19)",
  "color--(20)",
  "color--(21)",
  "color--(22)",
  "color--(23)",
  "color--(24)",
  "color--(25)",
  "color--(26)",
  "color--(27)",
  "color--(28)",
];

// ---------------------------------------------------------------------------
// Query builders
// ---------------------------------------------------------------------------

function buildTimeSeriesTelemetryQuery(table: string, field: DashboardField): string {
  const leaf = field.leaf ?? field.key;
  const component = field.component ?? "";
  return [
    dashboardRawQuery(table, "telemetry", component),
    "| where event_time between (['_startTime'] .. ['_endTime'])",
    `| extend Value = todecimal(data.payload.value[${kqlString(leaf)}])`,
    "| where isnotempty(Value)",
    "| project event_time, Value",
    "| order by event_time asc",
  ].join("\n");
}

function buildNonTimeSeriesTelemetryQuery(table: string, field: DashboardField): string {
  const leaf = field.leaf ?? field.key;
  const component = field.component ?? "";
  return [
    dashboardRawQuery(table, "telemetry", component),
    "| where event_time between (['_startTime'] .. ['_endTime'])",
    `| extend Value = data.payload.value[${kqlString(leaf)}]`,
    "| where isnotempty(Value)",
    "| project event_time, Value",
    "| order by event_time desc",
    "| limit 10",
  ].join("\n");
}

function buildPropertyQuery(table: string, field: DashboardField): string {
  const leaf = field.leaf ?? field.key;
  const component = field.component ?? "";
  // Component reported properties are nested under the component key (with a __t="c" marker);
  // root properties sit directly on properties.reported.
  const componentExpr = `rp[${kqlString(component)}]`;
  const valueExpr = component ? `${componentExpr}[${kqlString(leaf)}]` : `rp[${kqlString(leaf)}]`;
  return [
    dashboardRawQuery(table, "twin"),
    "| extend rp = data.payload.value.properties.reported",
    ...(component ? [`| where tostring(${componentExpr}["__t"]) == "c"`] : []),
    `| extend Value = ${valueExpr}`,
    ...(!component ? ['| where tostring(Value["__t"]) != "c"'] : []),
    "| where isnotempty(Value)",
    "| project event_time, Value",
    "| order by event_time desc",
    "| limit 1",
  ].join("\n");
}

function buildDeviceListQuery(table: string): string {
  return [
    quoteKqlIdentifier(table),
    "| where ingestion_time() between (['_startTime'] .. ['_endTime'])",
    "| distinct deviceId = tostring(user_headers.cloudEvents_subject)",
  ].join("\n");
}

// ---------------------------------------------------------------------------
// Tile visual options templates
// ---------------------------------------------------------------------------

const LINE_CHART_OPTIONS: Record<string, unknown> = {
  multipleYAxes: {
    base: {
      id: "-1", label: "", columns: [] as string[],
      yAxisMaximumValue: null as null, yAxisMinimumValue: null as null,
      yAxisScale: "linear", horizontalLines: [] as unknown[],
    },
    additional: [] as unknown[],
    showMultiplePanels: false,
  },
  seriesColors: {},
  hideLegend: false,
  legendLocation: "bottom",
  xColumnTitle: "",
  xColumn: "event_time",
  yColumns: ["Value"],
  seriesColumns: null as null,
  xAxisScale: "linear",
  verticalLine: "",
  crossFilterDisabled: false,
  drillthroughDisabled: false,
  forceAxisTicks: false,
  crossFilter: [] as unknown[],
  drillthrough: [] as unknown[],
};

const TABLE_OPTIONS: Record<string, unknown> = {
  table__enableRenderLinks: true,
  colorRulesDisabled: true,
  crossFilterDisabled: false,
  drillthroughDisabled: false,
  crossFilter: [] as unknown[],
  drillthrough: [] as unknown[],
  table__renderLinks: [] as unknown[],
  colorRules: [] as unknown[],
};

const CARD_OPTIONS: Record<string, unknown> = {
  multiStat__textSize: "auto",
  multiStat__valueColumn: "Value",
  colorRulesDisabled: false,
  colorRules: [] as unknown[],
};

// ---------------------------------------------------------------------------
// Main generator
// ---------------------------------------------------------------------------

export function generateDashboardDefinition(params: DashboardGenerationParams): object {
  const {
    title,
    telemetryTable,
    propertiesTable,
    telemetryFields,
    propertyFields,
    queryServiceUri,
    databaseId,
  } = params;

  const included = {
    timeSeries: telemetryFields.filter((f) => f.included && f.isNumeric),
    nonTimeSeries: telemetryFields.filter((f) => f.included && !f.isNumeric),
    properties: propertyFields.filter((f) => f.included),
  };

  const dashboardId = uuid();
  const pageId = uuid();
  const dataSourceId = uuid();

  // Build data source
  const dataSource = {
    kind: "kusto-trident",
    clusterUri: queryServiceUri,
    name: "IoT",
    id: dataSourceId,
    databaseArtifactId: databaseId,
    database: databaseId,
    workspace: "00000000-0000-0000-0000-000000000000",
    scopeId: "kusto-trident",
  };

  // Device list query (for parameter dropdown)
  const deviceListQueryId = uuid();
  const deviceListQuery = {
    dataSource: { kind: "inline", dataSourceId },
    text: buildDeviceListQuery(telemetryTable),
    id: deviceListQueryId,
    usedVariables: ["_endTime", "_startTime"],
  };

  // Parameters
  const timeRangeParam = {
    kind: "duration",
    id: uuid(),
    displayName: "Time range",
    description: "",
    beginVariableName: "_startTime",
    endVariableName: "_endTime",
    defaultValue: { kind: "dynamic", count: 1, unit: "hours" },
    showOnPages: { kind: "all" },
  };

  const deviceIdParam = {
    kind: "string",
    id: uuid(),
    displayName: "Device ID",
    description: "",
    variableName: "_deviceId",
    selectionType: "scalar",
    includeAllOption: false,
    defaultValue: { kind: "no-selection" },
    dataSource: {
      kind: "query",
      columns: { value: "deviceId" },
      queryRef: { kind: "query", queryId: deviceListQueryId },
    },
    showOnPages: { kind: "all" },
  };

  const tiles: object[] = [{
    id: uuid(),
    title: "",
    hideTitle: true,
    visualType: "designTile",
    pageId,
    layout: { x: 0, y: 0, width: 18, height: 3 },
    visualOptions: {
      designTile__kind: "textBox",
      textBox__html: '<div style="font-size: 24px;">Device dashboard</div><div>(Choose a device ID above to get started)</div>',
    },
  }];
  const queries: object[] = [deviceListQuery];
  let currentY = 3;

  function addSectionDivider(label: string): void {
    tiles.push({
      id: uuid(),
      title: "",
      hideTitle: true,
      visualType: "designTile",
      pageId,
      layout: { x: 0, y: currentY, width: 24, height: 1 },
      visualOptions: { designTile__kind: "divider", divider__label: label },
    });
    currentY += 1;
  }

  // --- Time-series telemetry (line charts, 2 per row, 9×7) ---
  if (included.timeSeries.length > 0) addSectionDivider("Time-series telemetry");
  for (let i = 0; i < included.timeSeries.length; i++) {
    const field = included.timeSeries[i];
    const queryId = uuid();
    const col = i % 2; // 0 or 1
    const x = col * 9;
    if (col === 0 && i > 0) currentY += 7;

    queries.push({
      dataSource: { kind: "inline", dataSourceId },
      text: buildTimeSeriesTelemetryQuery(telemetryTable, field),
      id: queryId,
      usedVariables: ["_deviceId", "_endTime", "_startTime"],
    });

    const color = CHART_COLORS[i % CHART_COLORS.length];
    tiles.push({
      id: uuid(),
      title: field.key,
      visualType: "line",
      pageId,
      layout: { x, y: currentY, width: 9, height: 7 },
      queryRef: { kind: "query", queryId },
      visualOptions: {
        ...LINE_CHART_OPTIONS,
        seriesColors: { Value: color },
      },
    });
  }
  if (included.timeSeries.length > 0) currentY += 7;

  // --- Non-time-series telemetry (tables, 2 per row, 9×7) ---
  if (included.nonTimeSeries.length > 0) addSectionDivider("Non-time-series telemetry (last 10 values)");
  for (let i = 0; i < included.nonTimeSeries.length; i++) {
    const field = included.nonTimeSeries[i];
    const queryId = uuid();
    const col = i % 2;
    const x = col * 9;
    if (col === 0 && i > 0) currentY += 7;

    queries.push({
      dataSource: { kind: "inline", dataSourceId },
      text: buildNonTimeSeriesTelemetryQuery(telemetryTable, field),
      id: queryId,
      usedVariables: ["_deviceId", "_endTime", "_startTime"],
    });

    tiles.push({
      id: uuid(),
      title: field.key,
      visualType: "table",
      pageId,
      layout: { x, y: currentY, width: 9, height: 7 },
      queryRef: { kind: "query", queryId },
      visualOptions: { ...TABLE_OPTIONS },
    });
  }
  if (included.nonTimeSeries.length > 0) currentY += 7;

  // --- Properties (cards, 3 per row, 6×3) ---
  if (included.properties.length > 0) addSectionDivider("Device properties (last known value)");
  for (let i = 0; i < included.properties.length; i++) {
    const field = included.properties[i];
    const queryId = uuid();
    const col = i % 3;
    const x = col * 6;
    if (col === 0 && i > 0) currentY += 3;

    queries.push({
      dataSource: { kind: "inline", dataSourceId },
      text: buildPropertyQuery(propertiesTable, field),
      id: queryId,
      usedVariables: ["_deviceId", "_endTime", "_startTime"],
    });

    tiles.push({
      id: uuid(),
      title: field.key,
      visualType: "card",
      pageId,
      layout: { x, y: currentY, width: 6, height: 3 },
      queryRef: { kind: "query", queryId },
      visualOptions: { ...CARD_OPTIONS },
    });
  }

  return {
    $schema: "https://dataexplorer.azure.com/static/d/schema/88/dashboard.json",
    id: dashboardId,
    eTag: `"${uuid()}"`,
    title,
    schema_version: 88,
    tiles,
    baseQueries: [],
    parameters: [timeRangeParam, deviceIdParam],
    dataSources: [dataSource],
    pages: [{ name: "Page 1", id: pageId }],
    queries,
    embeddedApps: [],
  };
}

/**
 * Base64-encode a string (handles Unicode).
 */
export function toBase64(str: string): string {
  const bytes = new TextEncoder().encode(str);
  let binary = "";
  for (let i = 0; i < bytes.length; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}
