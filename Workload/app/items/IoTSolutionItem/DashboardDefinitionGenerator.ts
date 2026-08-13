/**
 * Generates a Fabric Real-Time Dashboard definition JSON from inspected
 * telemetry and property fields. The output follows the Kusto dashboard
 * schema (version 69) and can be posted to the Fabric REST API.
 */

import { RAW_EVENT_NORMALIZATION_KQL } from "./RawEventKql";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface DashboardField {
  key: string;
  isNumeric?: boolean;
  included: boolean;
  /**
   * Component name this field belongs to (undefined/empty for root-interface fields).
   * Component telemetry arrives as separate messages tagged with IoTSubject=<component>, and
   * component reported properties are nested under the component key. `key` is the flattened
   * display/column name (`<component>_<leaf>`); `component` + `leaf` recover the raw wire shape.
   */
  component?: string;
  /** Raw payload key for this field (the unprefixed leaf name). Defaults to `key` when absent. */
  leaf?: string;
}

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

// Component telemetry arrives as separate messages tagged with IoTSubject=<component>; root
// telemetry has an empty IoTSubject. Guarding on the subject ensures a field reads only its own
// source and prevents a root field from matching a component message (or vice versa) when they
// share a leaf name. Returns "" for root fields, which matches the empty IoTSubject.
function subjectFilter(component: string): string {
  return `| where tostring(headers.IoTSubject) == '${component}'`;
}

function buildTimeSeriesTelemetryQuery(table: string, field: DashboardField): string {
  const leaf = field.leaf ?? field.key;
  const component = field.component ?? "";
  return [
    table,
    "| where ingestion_time() >= ago(30d)",
    RAW_EVENT_NORMALIZATION_KQL,
    "| extend deviceId = tostring(headers.IoTConnectionDeviceId)",
    "| where isnotempty(deviceId)",
    "| where deviceId == _deviceId",
    subjectFilter(component),
    "| extend enqueuedTime = todatetime(headers.IoTEnqueueTime)",
    "| where isnotempty(enqueuedTime)",
    "| where enqueuedTime between (['_startTime'] .. ['_endTime'])",
    "| extend telemetry = data",
    `| extend Value = todecimal(telemetry['${leaf}'])`,
    "| where isnotempty(Value)",
    "| project enqueuedTime, Value",
    "| order by enqueuedTime asc",
  ].join("\n");
}

function buildNonTimeSeriesTelemetryQuery(table: string, field: DashboardField): string {
  const leaf = field.leaf ?? field.key;
  const component = field.component ?? "";
  return [
    table,
    "| where ingestion_time() >= ago(30d)",
    RAW_EVENT_NORMALIZATION_KQL,
    "| extend deviceId = tostring(headers.IoTConnectionDeviceId)",
    "| where isnotempty(deviceId)",
    "| where deviceId == _deviceId",
    subjectFilter(component),
    "| extend enqueuedTime = todatetime(headers.IoTEnqueueTime)",
    "| where isnotempty(enqueuedTime)",
    "| where enqueuedTime between (['_startTime'] .. ['_endTime'])",
    "| extend telemetry = data",
    `| extend Value = telemetry['${leaf}']`,
    "| where isnotempty(Value)",
    "| project Time=enqueuedTime, Value",
    "| order by Time desc",
    "| limit 10",
  ].join("\n");
}

function buildPropertyQuery(table: string, field: DashboardField): string {
  const leaf = field.leaf ?? field.key;
  const component = field.component ?? "";
  // Component reported properties are nested under the component key (with a __t="c" marker);
  // root properties sit directly on properties.reported.
  const valueExpr = component ? `rp['${component}']['${leaf}']` : `rp['${leaf}']`;
  return [
    table,
    "| where ingestion_time() >= ago(30d)",
    RAW_EVENT_NORMALIZATION_KQL,
    "| extend deviceId = tostring(headers.IoTConnectionDeviceId)",
    "| where isnotempty(deviceId)",
    "| where deviceId == _deviceId",
    "| extend enqueuedTime = todatetime(headers.IoTEnqueueTime)",
    "| where isnotempty(enqueuedTime)",
    "| extend rp = data.properties.reported",
    `| extend Value = ${valueExpr}`,
    "| where isnotempty(Value)",
    "| project Time=enqueuedTime, Value",
    "| order by Time desc",
    "| limit 1",
  ].join("\n");
}

function buildDeviceListQuery(table: string): string {
  return [
    table,
    "| where ingestion_time() >= ago(30d)",
    RAW_EVENT_NORMALIZATION_KQL,
    "| extend deviceId = tostring(headers.IoTConnectionDeviceId)",
    "| where isnotempty(deviceId)",
    "| distinct deviceId",
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
  xColumn: "enqueuedTime",
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
    usedVariables: [] as string[],
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

  const tiles: object[] = [];
  const queries: object[] = [deviceListQuery];
  let currentY = 0;

  // --- Time-series telemetry (line charts, 2 per row, 9×7) ---
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

  // --- Properties (cards, 6 per row, 3×3) ---
  for (let i = 0; i < included.properties.length; i++) {
    const field = included.properties[i];
    const queryId = uuid();
    const col = i % 6;
    const x = col * 3;
    if (col === 0 && i > 0) currentY += 3;

    queries.push({
      dataSource: { kind: "inline", dataSourceId },
      text: buildPropertyQuery(propertiesTable, field),
      id: queryId,
      usedVariables: ["_deviceId"],
    });

    tiles.push({
      id: uuid(),
      title: field.key,
      visualType: "card",
      pageId,
      layout: { x, y: currentY, width: 3, height: 3 },
      queryRef: { kind: "query", queryId },
      visualOptions: { ...CARD_OPTIONS },
    });
  }

  return {
    $schema: "https://pbiadx.powerbi.com/static/d/schema/69/dashboard.json",
    id: dashboardId,
    eTag: `"${uuid()}"`,
    title,
    schema_version: 69,
    tiles,
    baseQueries: [],
    parameters: [timeRangeParam, deviceIdParam],
    dataSources: [dataSource],
    pages: [{ name: "Page 1", id: pageId }],
    queries,
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
