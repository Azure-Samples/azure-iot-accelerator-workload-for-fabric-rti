// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { declaredModelName, isJsonObject, modelMeaningDocstring, modelResourceNameBase, parseIotMeta, quoteKqlIdentifier, readDocstringTextLine as meaningLine, readModelIdentityMetadata, type JsonObject, type JsonValue } from "./ModelAuthoring.ts";
import { readColumnSemantics } from "./FieldSemantics.ts";
import { parseDocstringJson } from "./Docstrings.ts";
import { buildDataAgentQueryGuidelines, buildIndependentLatestValuesQuery, buildLatestValueQuery } from "./dataAgentQueryGuidance.ts";

export interface DocumentedColumn { name: string; type: string; docstring?: string }
export interface DocumentedSource {
  database: string;
  table: string;
  docstring: string;
  columns: DocumentedColumn[];
}
export interface DataAgentContext {
  agentDescription: string;
  agentInstructions: string;
  sourceDescription: string;
  sourceInstructions: string;
  examples: { question: string; query: string }[];
}

export function requireEventTimeColumn(columns: { name: string; type: string }[]): "event_time" {
  const candidates = columns.filter(column => column.name === "event_time");
  if (candidates.length !== 1 || candidates[0].type !== "datetime") {
    throw new Error("The inspected source must contain exactly one event_time:datetime column. Existing resources were not renamed or migrated.");
  }
  return "event_time";
}

/** Adapt V1 table-level column facts once; native column documentation takes precedence. */
export function normalizeDocumentedSource(source: DocumentedSource): DocumentedSource {
  const legacy = new Map<string, { displayName?: string; description?: string; kind?: string; unit?: string; component?: string }>();
  const lineRegex = /^- (\S+?)(?: — "([^"]*)")? \(([^)]*)\)(?::\s*(.*))?$/;
  for (const line of source.docstring.split(/\r?\n/)) {
    const match = lineRegex.exec(line.trim());
    if (!match) continue;
    const [, name, displayName, attrs, description] = match;
    const attributes = attrs.split(",").map(value => value.trim());
    legacy.set(name, {
      displayName, description,
      kind: attributes.find(value => value === "telemetry" || value === "property"),
      unit: attributes.find(value => value.startsWith("unit:"))?.slice("unit:".length).trim(),
      component: attributes.find(value => value.startsWith("component:"))?.slice("component:".length).trim(),
    });
  }
  const propertyUpdates = parseIotMeta(source.docstring).rowKind === "property-update";
  return { ...source, columns: source.columns.map(column => {
    const old = legacy.get(column.name);
    if (column.docstring || !old) return { ...column };
    const docstring = [old.description === "" ? 'Description: ""' : old.description,
      old.displayName !== undefined ? `Title: ${JSON.stringify(old.displayName)}` : "",
      old.kind === "telemetry" ? "Telemetry." : old.kind === "property"
        ? propertyUpdates ? "Reported property value." : "Twin property (last-known value)." : "",
      old.unit ? `Unit: ${old.unit}.` : "",
      old.component ? `Component: ${JSON.stringify(old.component)}` : "",
    ].filter(Boolean).join("\n");
    return { ...column, ...(docstring ? { docstring } : {}) };
  }) };
}

export function documentedColumns(schema: unknown): DocumentedColumn[] {
  if (!isJsonObject(schema) || !Array.isArray(schema.OrderedColumns) || !schema.OrderedColumns.length) {
    throw new Error("The selected table schema does not contain inspected columns.");
  }
  return schema.OrderedColumns.map(column => {
    if (!isJsonObject(column) || typeof column.Name !== "string" ||
        (typeof column.CslType !== "string" && typeof column.Type !== "string")) {
      throw new Error("The selected table returned an unsupported column schema.");
    }
    return { name: column.Name, type: String(column.CslType || column.Type),
      ...(typeof column.DocString === "string" && column.DocString ? { docstring: column.DocString } : {}) };
  });
}

/** Native exports and V1 nest tables in Tables; the published 1.0 schema omits this runtime grouping kind. */
export function buildDataAgentDataSource(source: DocumentedSource, workspaceId: string, databaseId: string, tableId: string) {
  const { context, columns } = renderDocumentedSource(source);
  const systemTypes: Record<string, string> = { string: "System.String", bool: "System.SByte",
    datetime: "System.DateTime", int: "System.Int32", long: "System.Int64", real: "System.Double",
    double: "System.Double", decimal: "System.Decimal", dynamic: "System.Object", guid: "System.Guid", timespan: "System.TimeSpan" };
  return {
    $schema: "https://developer.microsoft.com/json-schemas/fabric/item/dataAgent/definition/dataSource/1.0.0/schema.json",
    artifactId: databaseId, workspaceId, displayName: source.database, type: "kusto",
    dataSourceInstructions: context.sourceInstructions, userDescription: context.sourceDescription, metadata: {},
    elements: [{
      id: crypto.randomUUID(), is_selected: false, display_name: "Tables", type: "table_grouping",
      description: null as string | null,
      children: [{
        id: tableId, is_selected: true, display_name: source.table, type: "kusto.table",
        description: context.sourceDescription,
        children: source.columns.map((column, index) => ({ id: column.name, is_selected: true,
          display_name: column.name, type: "kusto.column", data_type: systemTypes[column.type] || "System.Object",
          description: columns[index].description || null, children: [] as never[] })),
      }],
    }],
  };
}

const AGENT_PURPOSE = "You are an IoT data assistant. You help users explore and analyze IoT device data stored in a KQL (Kusto Query Language) database.";
const FACT_LABELS: Record<string, string> = {
  type: "WoT type", schema: "DTDL schema", "@type": "Semantic @type", semanticTypes: "Semantic types",
  comment: "Comment", readOnly: "Read-only", writeOnly: "Write-only", writable: "Writable",
  observable: "Observable", format: "Format", const: "Constant", default: "Default", enum: "Enum", codes: "Codes",
  minimum: "Minimum", exclusiveMinimum: "Exclusive minimum", maximum: "Maximum", exclusiveMaximum: "Exclusive maximum",
  multipleOf: "Multiple of", minLength: "Minimum length", maxLength: "Maximum length", pattern: "Pattern",
  contentEncoding: "Content encoding", contentMediaType: "Content media type", items: "Items", minItems: "Minimum items",
  maxItems: "Maximum items", properties: "Properties", required: "Required members", oneOf: "Schema alternatives",
  data: "Event data schema", valueSchema: "Value schema", enumValues: "Enum values", fields: "Fields",
  mapKey: "Map key", mapValue: "Map value", elementSchema: "Element schema", enumValue: "Enum value",
};

function inline(text: string): string {
  return text === "" ? '""' : text.replace(/\r/g, "\\r").replace(/\n/g, "\\n");
}

function code(text: string): string {
  const fence = "`".repeat(1 + Math.max(0, ...[...text.matchAll(/`+/g)].map(match => match[0].length)));
  const padding = text.startsWith("`") || text.endsWith("`") ? " " : "";
  return `${fence}${padding}${inline(text)}${padding}${fence}`;
}

function localized(value: JsonValue | undefined, dtdl: boolean): string | undefined {
  if (typeof value === "string") return value;
  return isJsonObject(value) ? (dtdl ? ["en", ...Object.keys(value)] : ["en-US", "en"])
    .map(key => value[key]).find((text): text is string => typeof text === "string") : undefined;
}

export function defaultDataAgentName(docstring: string, table: string, suffix: string): string {
  return `DataAgent_${modelResourceNameBase(docstring, table)}_${suffix}`;
}

export function dataAgentExamplesMarkdown(examples: DataAgentContext["examples"]): string {
  return examples.map(example => `Q: ${example.question}\n\n\`\`\`kql\n${example.query}\n\`\`\`\n\n`).join("");
}

function readableDtdlSchema(value: JsonValue): JsonValue {
  if (!isJsonObject(value)) return value;
  return Object.fromEntries(Object.entries(value).filter(([key]) => key !== "@id").map(([key, child]) => [
    key, ["schema", "valueSchema", "elementSchema", "mapKey", "mapValue"].includes(key) ? readableDtdlSchema(child)
      : ["fields", "enumValues"].includes(key) && Array.isArray(child) ? child.map(readableDtdlSchema) : child,
  ]));
}

function describeColumn(column: DocumentedColumn, model: ReturnType<typeof readModelIdentityMetadata>) {
  const doc = column.docstring || "", semantics = readColumnSemantics(doc, model);
  let facts: JsonObject = {};
  if (semantics?.version === 1) {
    const metadata = semantics.field.metadata;
    facts = { ...(metadata.kind === "wot" && isJsonObject(metadata.schema) ? metadata.schema : { schema: metadata.schema }),
      ...metadata.annotations };
  } else if (semantics) facts = semantics.field.facts;
  if ("type" in facts && "schema" in facts || "@type" in facts && "semanticTypes" in facts) {
    throw new Error("Ambiguous documented semantic type aliases.");
  }
  const plain = !semantics || semantics.version === 1;
  const telemetry = plain && /(?:^|\n)Telemetry\.(?:\n|$)/.test(doc);
  const property = plain && /(?:^|\n)(?:Twin property \(last-known value\)|Reported property value)\.(?:\n|$)/.test(doc);
  if (telemetry && property) throw new Error("Ambiguous documented column roles.");
  const role = facts.role === "telemetry" || telemetry ? "telemetry"
    : facts.role === "twin property (last-known whole received value)" || property ? "property" : undefined;
  const prose = plain ? doc.split(/\r?\n/).filter(line => !line.startsWith("@iot.meta") &&
    !/^(?:Telemetry\.|Twin property \(last-known value\)\.|Reported property value\.|Title:|Description:|Unit:|Component:)/.test(line)).join("\n").trim() : "";
  const dtdl = semantics?.version === 1 ? semantics.field.metadata.kind === "dtdl" : "schema" in facts || "displayName" in facts;
  const title = localized(dtdl ? facts.displayName : facts.titles, dtdl) ?? localized(facts.title, dtdl) ?? (plain ? meaningLine(doc, "Title") : undefined);
  const description = localized(facts.descriptions, dtdl) ?? localized(facts.description, dtdl) ??
    (plain ? meaningLine(doc, "Description") : undefined) ?? (prose || undefined);
  const unit = typeof facts.unit === "string" ? facts.unit : plain ? /^Unit: (.*)\.$/m.exec(doc)?.[1] : undefined;
  const component = typeof facts.component === "string" ? facts.component : plain ? meaningLine(doc, "Component") : undefined;
  const main = [title, description].filter((value): value is string => value !== undefined).map(inline).join(" \u2014 ");
  const qualifiers = [unit !== undefined ? `unit: ${inline(unit)}` : "", component !== undefined ? `component: ${inline(component)}` : ""].filter(Boolean);
  const first = `${main}${qualifiers.length ? `${main ? " " : ""}(${qualifiers.join(", ")})` : ""}`;
  const details = Object.entries(FACT_LABELS).filter(([key]) => Object.prototype.hasOwnProperty.call(facts, key))
    .map(([key, label]) => `${label}: ${JSON.stringify(key === "schema" && dtdl ? readableDtdlSchema(facts[key]) : facts[key])}`);
  for (const [key, label] of [["title", "Declared title"], ["titles", "Titles"], ["displayName", "Display names"],
    ["description", "Descriptions"], ["descriptions", "Descriptions"]] as const) {
    if (isJsonObject(facts[key])) details.push(`${label}: ${JSON.stringify(facts[key])}`);
  }
  if (typeof facts.title === "string" && facts.title !== title) details.push(`Declared title: ${JSON.stringify(facts.title)}`);
  if (typeof facts.description === "string" && facts.description !== description) details.push(`Declared description: ${JSON.stringify(facts.description)}`);
  const bounded = semantics !== undefined && semantics.version !== 1 && semantics.bounded;
  if (bounded) details.push("Metadata: incomplete (truncated=true).");
  return { ...column, role, component, bounded,
    line: `- ${code(column.name)} (${column.type})${first ? `: ${first}` : ""}${details.length ? "\n\n" + details.map(detail => `  - ${detail}`).join("\n") : ""}`,
    description: [first, details.map(detail => `- ${detail}`).join("\n")].filter(Boolean).join("\n\n"),
  };
}

function renderDocumentedSource(original: DocumentedSource) {
  const source = normalizeDocumentedSource(original);
  if (!source.table || !source.columns.length) throw new Error("Select and inspect a documented source before generating Data Agent context.");
  const table = quoteKqlIdentifier(source.table);
  const metadata = parseIotMeta(source.docstring);
  const modelIdentity = readModelIdentityMetadata(source.docstring);
  const modelInfo = metadata.modelSemanticInfo ? parseDocstringJson(metadata.modelSemanticInfo, "model") : {};
  const hasIdentity = source.columns.some(column => column.name === "deviceId" && column.type === "string");
  requireEventTimeColumn(source.columns);
  const columns = source.columns.map(column => describeColumn(column, modelIdentity));
  const fields = columns.filter(column => !["deviceId", "event_time"].includes(column.name));
  const telemetry = fields.filter(column => column.role === "telemetry"), properties = fields.filter(column => column.role === "property");
  const numeric = telemetry.find(column => ["int", "long", "real", "decimal", "double"].includes(column.type));
  const latest = numeric ?? fields[0];
  const examples: DataAgentContext["examples"] = [];
  if (hasIdentity) examples.push(
    { question: "What devices are in the system?", query: `${table} | distinct deviceId` },
    { question: "Show me the latest modeled records for a device (replace YOUR_DEVICE_ID)",
      query: `${table} | where deviceId == "YOUR_DEVICE_ID" | top 10 by event_time desc` },
  );
  if (hasIdentity && numeric) {
    const column = quoteKqlIdentifier(numeric.name);
    examples.push({
      question: `What is the average ${numeric.name} over the last hour?`,
      query: `${table} | where event_time >= ago(1h) and isnotnull(${column}) | summarize avg(${column}) by bin(event_time, 5m) | render timechart`,
    });
  }
  if (hasIdentity && latest) examples.push({
    question: `Show the latest ${latest.type === "string" ? "nonempty" : "non-null"} ${latest.name} for each device`,
    query: buildLatestValueQuery(source.table, latest.name, latest.type),
  });
  if (hasIdentity && properties.length) examples.push({
    question: `Show the latest available ${properties[0].type === "string" ? "nonempty" : "non-null"} reported ${properties[0].name} for each device`,
    query: buildLatestValueQuery(source.table, properties[0].name, properties[0].type),
  });
  if (hasIdentity) examples.push({
    question: metadata.rowKind === "property-update" ? "How many reported-property updates per device in the last 24 hours?"
      : metadata.tableRole === "modeled-data" || metadata.rowKind === "telemetry-event"
        ? "How many readings per device in the last 24 hours?" : "How many modeled records per device in the last 24 hours?",
    query: `${table} | where event_time >= ago(24h) | summarize Count=count() by deviceId | order by Count desc`,
  });
  const componentTelemetry = telemetry.find(column => !!column.component);
  if (hasIdentity && componentTelemetry) examples.push({
    question: `Show the latest ${componentTelemetry.type === "string" ? "nonempty" : "non-null"} ${componentTelemetry.name} telemetry from component ${componentTelemetry.component} for each device`,
    query: buildLatestValueQuery(source.table, componentTelemetry.name, componentTelemetry.type),
  });
  const firstMeasurement = telemetry.find(column => !column.component) ?? telemetry[0];
  const secondMeasurement = telemetry.find(column => column !== firstMeasurement && !!column.component)
    ?? telemetry.find(column => column !== firstMeasurement);
  if (hasIdentity && firstMeasurement && secondMeasurement) examples.push({
    question: `Show the latest available ${firstMeasurement.name} and ${secondMeasurement.name} for each device, with each measurement's own timestamp`,
    query: buildIndependentLatestValuesQuery(source.table, firstMeasurement, secondMeasurement),
  });
  const meaning = modelMeaningDocstring(source.docstring);
  const sourceMeaning = meaning.split(/\r?\n/).filter(line => !line.startsWith("@iot.meta")).join("\n").trim();
  const modelName = declaredModelName(source.docstring), modelDescription = meaningLine(meaning, "description");
  const modelLines = [
    ...(modelIdentity.id ? [`- Model ID: ${inline(modelIdentity.id)}`] : []),
    ...(modelName !== undefined ? [`- Model name: ${inline(modelName)}`] : []),
    ...(modelIdentity.documentVersion !== undefined ? [`- Model version: ${inline(modelIdentity.documentVersion)}`] : []),
    ...(modelDescription !== undefined ? [`- Description: ${inline(modelDescription)}`] : []),
  ];
  for (const [key, label] of [["title", "Declared title"], ["titles", "Model name translations"], ["displayName", "Model name translations"],
    ["description", "Description translations"], ["descriptions", "Description translations"]] as const) {
    if (isJsonObject(modelInfo[key])) modelLines.push(`- ${label}: ${JSON.stringify(modelInfo[key])}`);
  }
  if (typeof modelInfo.title === "string" && modelInfo.title !== modelName) modelLines.push(`- Declared model title: ${JSON.stringify(modelInfo.title)}`);
  if (typeof modelInfo.description === "string" && modelInfo.description !== modelDescription) modelLines.push(`- Declared description: ${JSON.stringify(modelInfo.description)}`);
  const bounded = modelInfo.truncated === true || /^truncated: true$/m.test(meaning) || columns.some(column => column.bounded);
  const rowDescription = metadata.rowKind === "property-update"
    ? "Each row represents a reported property update from an IoT device. Only properties included in that update are populated; other property columns can be null."
    : metadata.tableRole === "modeled-data" || metadata.rowKind === "telemetry-event"
      ? `Each row represents a telemetry event from an IoT device${properties.length ? ", enriched with the device's latest selected reported properties" : ""}.`
      : "Rows contain the inspected modeled columns below; no row kind is documented.";
  const section = (heading: string, values: typeof columns) => values.length ? `### ${heading}\n\n${values.map(column => column.line).join("\n\n")}\n\n` : "";
  const datasetDescription = sourceMeaning.split(/\r?\n/).filter(line => !/^(?:title:|description:|truncated:|- )/.test(line)).join("\n").trim();
  const sourceInstructions = `## Data Model\n\nThe main table is ${code(source.table)} in the ${code(source.database)} database. ${rowDescription}\n\n` +
    (datasetDescription && datasetDescription !== "Modeled device telemetry enriched with property LKVs"
      ? `Dataset description: ${inline(datasetDescription)}\n\n` : "") +
    (modelLines.length ? `### Device Model\n${modelLines.join("\n")}\n\n` : "") +
    section("Key Columns", columns.filter(column => ["deviceId", "event_time"].includes(column.name)).map(column =>
      column.name === "event_time" ? { ...column, line: column.line + "\n\n  - Primary time axis for all time-based queries." } : column)) +
    section("Telemetry Columns (real-time measurements)", telemetry) +
    section(metadata.rowKind === "property-update" ? "Property Columns (values included in each reported-property update)"
      : "Property Columns (device-reported metadata, enriched via last-known-value lookup)", properties) +
    section("Additional Columns (role not documented)", fields.filter(column => !column.role)) +
    buildDataAgentQueryGuidelines(metadata.rowKind, {
      enrichedProperties: metadata.rowKind !== "property-update" && properties.length > 0,
      dynamic: columns.some(column => column.type === "dynamic"), bounded,
    });
  const context: DataAgentContext = {
    agentDescription: "Explore IoT operations using the configured modeled-data sources.",
    agentInstructions: AGENT_PURPOSE,
    sourceDescription: sourceMeaning.split("\n")[0] || source.table,
    sourceInstructions, examples,
  };
  return { context, columns };
}

/** Render only inspected native documentation, never hidden authoring models or inferred field meanings. */
export function buildDataAgentContext(source: DocumentedSource): DataAgentContext {
  return renderDocumentedSource(source).context;
}
