// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { bindNativeMetadata, type NativeBinding } from "./HubTransport.ts";
import { docstringJson, parseDocstringJson } from "./Docstrings.ts";

export type JsonValue = null | boolean | number | string | JsonValue[] | JsonObject;
export interface JsonObject { [key: string]: JsonValue }

export interface ObservedSchema {
  $schema?: string;
  type?: string | string[];
  properties?: Record<string, ObservedSchema>;
  items?: ObservedSchema;
}

export type SampleStream = "telemetry" | "twin";
export const SAMPLE_DEPTHS = [100, 500, 1000] as const;
const MAX_ROW_CHARACTERS = 2 * 1024 * 1024;

export interface RawSampleRow {
  data: unknown;
  headers: unknown;
  user_headers: unknown;
}

export interface PayloadSample {
  value: JsonObject;
  reported?: JsonObject;
  component?: string;
  eventTime?: string;
  memberTypeIRI?: string;
  headers?: JsonObject;
  userHeaders?: JsonObject;
  native?: NativeBinding;
}

export interface StreamSample {
  rawRows: number;
  observations: number;
  skippedWithoutReportedProperties: number;
  payloads: PayloadSample[];
  batchedRows?: number;
}

export interface AuthoringSource {
  queryServiceUri: string;
  databaseName: string;
  telemetryTable: string;
  twinTable: string;
}

export interface AuthoringSample {
  limit: number;
  telemetry: StreamSample;
  twin: StreamSample;
  schema: ObservedSchema;
}

export const MODEL_SEMANTIC_INFO_PREFIX = "@iot.meta modelSemanticInfo=";

function modelSemanticInfo(value: string): JsonObject {
  const info = parseDocstringJson(value, "model");
  if (Object.keys(info).some(key => !["id", "@id", "title", "titles", "displayName", "description", "descriptions", "version", "truncated"].includes(key))) {
    throw new Error("Unsupported model semantic metadata attribute.");
  }
  if ("id" in info && "@id" in info || "displayName" in info && ("title" in info || "titles" in info) ||
      "@id" in info && ("version" in info || "title" in info || "titles" in info || "descriptions" in info) ||
      "id" in info && "displayName" in info) throw new Error("Ambiguous native model semantic metadata aliases.");
  for (const key of ["id", "@id"]) if (info[key] !== undefined) {
    const id = info[key];
    if (typeof id !== "string" || /[\u0000-\u0020\u007f]/.test(id)) throw new Error("Malformed canonical model semantic identity.");
    try { new URL(id); } catch { throw new Error("Malformed canonical model semantic identity."); }
  }
  for (const key of ["title", "displayName", "description"]) {
    const item = info[key];
    if (item !== undefined && typeof item !== "string" &&
        !(isJsonObject(item) && Object.values(item).every(text => typeof text === "string"))) throw new Error(`Malformed model semantic ${key}.`);
  }
  for (const key of ["titles", "descriptions"]) {
    const item = info[key];
    if (item !== undefined && (!isJsonObject(item) || !Object.values(item).every(text => typeof text === "string"))) {
      throw new Error(`Malformed model semantic ${key}.`);
    }
  }
  if (info.version !== undefined && typeof info.version !== "string" ||
      info.truncated !== undefined && info.truncated !== true) throw new Error("Malformed model semantic version or truncation disclosure.");
  return info;
}

export function parseIotMeta(docstring: string): Record<string, string> {
  const metadata: Record<string, string> = {};
  for (const line of docstring.split(/\r?\n/)) {
    if (!line.startsWith("@iot.meta")) continue;
    const match = /^@iot\.meta\s+([A-Za-z][\w]*)\s*(?:=|:)\s*(\S(?:.*\S)?)$/.exec(line);
    if (!match) throw new Error("Malformed line-anchored IoT metadata tag.");
    if (Object.prototype.hasOwnProperty.call(metadata, match[1])) throw new Error(`Duplicate IoT metadata tag: ${match[1]}.`);
    if (match[1].startsWith("modelSemanticInfo") && !line.startsWith(MODEL_SEMANTIC_INFO_PREFIX)) {
      throw new Error("Malformed model semantic metadata tag; use @iot.meta modelSemanticInfo=.");
    }
    metadata[match[1]] = match[2];
  }
  if (metadata.modelSemanticInfo !== undefined) {
    modelSemanticInfo(metadata.modelSemanticInfo);
    if (["modelId", "deviceModel", "modelVersion"].some(key => Object.prototype.hasOwnProperty.call(metadata, key)) ||
        docstring.split(/\r?\n/).some(line => /^(?:title|description):/.test(line))) {
      throw new Error("Ambiguous legacy and JSON model semantic metadata aliases.");
    }
  }
  return metadata;
}
export function readModelIdentityMetadata(docstring: string): { id?: string; documentVersion?: string } {
  const metadata = parseIotMeta(docstring);
  if (metadata.modelSemanticInfo !== undefined) {
    const info = modelSemanticInfo(metadata.modelSemanticInfo), id = info.id ?? info["@id"];
    const version = info.version ?? (typeof info["@id"] === "string" ? /;(\d+)$/.exec(info["@id"])?.[1] : undefined);
    return { ...(typeof id === "string" ? { id } : {}), ...(typeof version === "string" ? { documentVersion: version } : {}) };
  }
  const modern = metadata.tableRole === "modeled-data" ? metadata.deviceModel : undefined;
  if (modern !== undefined && metadata.modelId !== undefined) throw new Error("Duplicate model identity metadata aliases.");
  const candidate = metadata.modelId ?? modern;
  let id: string | undefined;
  if (candidate !== undefined) {
    try {
      if (/\s/.test(candidate)) throw new Error("Invalid model identity.");
      new URL(candidate); id = candidate;
    } catch {
      if (metadata.modelId !== undefined || docstring.startsWith("Modeled device telemetry enriched with property LKVs")) {
        throw new Error("Malformed canonical model identity metadata.");
      }
      // V1 deviceModel tags may name a model rather than identify it by absolute IRI.
    }
  }
  return { ...(id !== undefined ? { id } : {}),
    ...(metadata.modelVersion !== undefined ? { documentVersion: metadata.modelVersion } : {}) };
}

/** Compatibility view of the existing dataset-meaning lines, not an instruction renderer. */
export function modelMeaningDocstring(docstring: string): string {
  const metadata = parseIotMeta(docstring);
  if (metadata.modelSemanticInfo === undefined) return docstring;
  const info = modelSemanticInfo(metadata.modelSemanticInfo);
  const localized = (value: JsonValue | undefined, preferred: readonly string[], fallback = false): string | undefined =>
    typeof value === "string" ? value : isJsonObject(value)
      ? [...preferred, ...(fallback ? Object.keys(value) : [])].map(key => value[key]).find((text): text is string => typeof text === "string")
      : undefined;
  const dtdl = "displayName" in info || "@id" in info;
  const title = dtdl ? localized(info.displayName, ["en"], true)
    : localized(info.titles, ["en-US", "en"]) ?? localized(info.title, ["en-US", "en"]);
  const description = dtdl ? localized(info.description, ["en"], true)
    : localized(info.descriptions, ["en-US", "en"]) ?? localized(info.description, ["en-US", "en"]);
  return docstring.split(/\r?\n/).map(line => line.startsWith(MODEL_SEMANTIC_INFO_PREFIX)
    ? [...(title !== undefined ? [`title: ${docstringJson(title)}`] : []),
      ...(description !== undefined ? [`description: ${docstringJson(description)}`] : []),
      ...(info.truncated === true ? ["truncated: true"] : [])].join("\n") : line).join("\n");
}

export function readDocstringTextLine(docstring: string, key: string): string | undefined {
  const line = docstring.split(/\r?\n/).find(line => line.startsWith(`${key}: `));
  if (!line) return undefined;
  const text = line.slice(key.length + 2);
  if (!text.startsWith('"')) return text;
  const value: unknown = JSON.parse(text);
  if (typeof value !== "string") throw new Error(`Malformed documented ${key}.`);
  return value;
}

export function declaredModelName(docstring: string): string | undefined {
  readModelIdentityMetadata(docstring);
  return readDocstringTextLine(modelMeaningDocstring(docstring), "title");
}

export function modelResourceNameBase(docstring: string, table: string): string {
  const name = declaredModelName(docstring);
  return (name?.trim() ? name : table).replace(/[^a-zA-Z0-9_]/g, "_");
}

export function filterTablesByRole(tables: string[], docstrings: Record<string, string>, role: string): string[] {
  return tables.filter(table => !!docstrings[table] && parseIotMeta(docstrings[table]).tableRole === role);
}

export function isJsonObject(value: unknown): value is JsonObject {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

export function canonicalJson(value: unknown): string {
  const ordered = (item: unknown): unknown => {
    if (Array.isArray(item)) return item.map(ordered);
    if (isJsonObject(item)) return Object.fromEntries(Object.keys(item).sort().map(key => [key, ordered(item[key])]));
    return item;
  };
  const result = JSON.stringify(ordered(value));
  if (result === undefined) throw new Error("A modeling definition must be JSON serializable.");
  return result;
}

function parseJson(value: unknown, location: string): unknown {
  if (typeof value !== "string") throw new Error(`${location} must be a JSON string in the data column.`);
  if (value.length > MAX_ROW_CHARACTERS) {
    throw new Error(`${location} exceeds the 2 MB sample-row limit. Select a smaller payload sample.`);
  }
  try {
    return JSON.parse(value);
  } catch {
    throw new Error(`${location} contains invalid JSON. Correct the sample data and try again.`);
  }
}

function reportedPayload(value: JsonObject, depth = 0): JsonObject {
  if (depth > 48) throw new Error("A reported-property payload exceeds the supported nesting depth (48).");
  return Object.fromEntries(
    Object.entries(value)
      .filter(([key]) => !key.startsWith("$") && key !== "__t")
      .map(([key, item]) => [
        key,
        isJsonObject(item) && item.__t === "c" ? reportedPayload(item, depth + 1) : item,
      ])
  );
}

function metadataObject(value: unknown, location: string): JsonObject {
  let decoded = value;
  if (typeof value === "string") {
    try { decoded = JSON.parse(value); } catch {
      throw new Error(`${location} metadata contains invalid JSON.`);
    }
  }
  if (decoded === null || decoded === undefined) return {};
  if (!isJsonObject(decoded)) throw new Error(`${location} metadata must be an object or null.`);
  return decoded;
}

/** Decode the verified DOObservationV1 envelope, not the incumbent CloudEvents body. */
export function decodeDoSamples(rows: RawSampleRow[], stream: SampleStream, limit: number): StreamSample {
  validateSampleDepth(limit);
  const result: StreamSample = {
    rawRows: rows.length, observations: 0, skippedWithoutReportedProperties: 0, payloads: [],
  };
  for (const [index, row] of rows.entries()) {
    if (result.observations >= limit) break;
    const location = `${stream === "telemetry" ? "Telemetry" : "Twin-change"} row ${index + 1}`;
    const decoded = parseJson(row.data, location);
    if (Array.isArray(decoded)) result.batchedRows = (result.batchedRows || 0) + 1;
    // Accept an observation array or one intact observation. Do not assume Eventstream flattens arrays.
    const observations = Array.isArray(decoded) ? decoded : [decoded];
    if (observations.length === 0) throw new Error(`${location} has an empty observation array.`);
    for (const observation of observations) {
      if (result.observations >= limit) break;
      if (!isJsonObject(observation) ||
          observation.schema_id !== "com.do.telemetry.observation.v1" ||
          observation.schema_version !== "1.0" || observation.contract_version !== "1.0") {
        throw new Error(
          `${location} is not a supported DOObservationV1 observation (contract/schema version 1.0). ` +
          "Matching raw columns alone do not establish compatibility. Run Ingest device data from your IoT Hub to receive new DOObservationV1 events in compatible raw tables, then select those tables and choose Validate again."
        );
      }
      if (observation.event_kind !== (stream === "telemetry" ? "measurement" : "lifecycle") ||
          typeof observation.event_time !== "string" || !Number.isFinite(Date.parse(observation.event_time))) {
        throw new Error(`${location} has an unexpected event kind or invalid event_time. Check the selected stream.`);
      }
      if (!isJsonObject(observation.payload) ||
          !Object.prototype.hasOwnProperty.call(observation.payload, "value")) {
        throw new Error(`${location} does not contain payload.value.`);
      }
      result.observations++;
      const value = observation.payload.value;
      if (!isJsonObject(value)) {
        throw new Error(
          `${location} has a non-object payload.value. It is valid observation data but unsuitable for object-schema export.`
        );
      }
      let payload = value;
      let reported: JsonObject | undefined;
      if (stream === "twin") {
        const properties = value.properties;
        if (!isJsonObject(properties) || !isJsonObject(properties.reported)) {
          result.skippedWithoutReportedProperties++;
          continue;
        }
        reported = properties.reported;
        payload = reportedPayload(reported);
      }
      const userHeaders = metadataObject(row.user_headers, location);
      const headers = metadataObject(row.headers, location);
      const native = bindNativeMetadata(headers, userHeaders, stream);
      result.payloads.push({ value: payload, ...(reported ? { reported } : {}), ...(native.component ? { component: native.component } : {}),
        eventTime: observation.event_time,
        ...(typeof observation.payload.memberTypeIRI === "string" ? { memberTypeIRI: observation.payload.memberTypeIRI } : {}),
        headers, userHeaders, native });
    }
  }
  return result;
}

function inferValueSchema(values: JsonValue[], depth: number): ObservedSchema {
  if (depth > 48) throw new Error("A payload exceeds the supported schema nesting depth (48).");
  if (values.length === 0) return {};
  const types = new Set(values.map(value => {
    if (value === null) return "null";
    if (Array.isArray(value)) return "array";
    if (typeof value === "number") {
      if (!Number.isFinite(value)) throw new Error("A payload contains a non-finite number.");
      return Number.isInteger(value) ? "integer" : "number";
    }
    if (["string", "boolean", "object"].includes(typeof value)) return typeof value;
    throw new Error("A payload contains a non-JSON value.");
  }));
  if (types.has("number")) types.delete("integer");
  const orderedTypes = [...types].sort();
  const schema: ObservedSchema = { type: orderedTypes.length === 1 ? orderedTypes[0] : orderedTypes };
  const objects = values.filter(isJsonObject);
  if (objects.length > 0) {
    const keys = [...new Set(objects.flatMap(Object.keys))].sort();
    schema.properties = Object.fromEntries(keys.map(key => [
      key,
      inferValueSchema(
        objects.filter(object => Object.prototype.hasOwnProperty.call(object, key)).map(object => object[key]),
        depth + 1
      ),
    ]));
  }
  const arrays = values.filter((value): value is JsonValue[] => Array.isArray(value));
  if (arrays.length > 0) schema.items = inferValueSchema(arrays.flat(), depth + 1);
  return schema;
}

export function inferPayloadSchema(payloads: JsonObject[]): ObservedSchema {
  if (payloads.length === 0) {
    throw new Error("No object payload samples are available. Add sample telemetry or reported twin changes, or select other tables.");
  }
  if (!payloads.every(isJsonObject)) throw new Error("Schema export requires decoded JSON payload objects.");
  return {
    $schema: "https://json-schema.org/draft/2020-12/schema",
    ...inferValueSchema(payloads, 0),
  };
}

/** The WoT preparation helper consumes root fields only; source samples retain their full component provenance. */
export function inferRootPreparationSchema(telemetry: StreamSample, twin: StreamSample): ObservedSchema {
  const rootTelemetry = telemetry.payloads.flatMap(sample => {
    const componentIssue = sample.native?.issues.find(issue => issue.startsWith("Telemetry component metadata "));
    if (componentIssue) throw new Error(componentIssue);
    if (sample.component !== undefined && typeof sample.component !== "string") {
      throw new Error("Telemetry component metadata is not a string.");
    }
    return sample.component ? [] : [sample.value];
  });
  const rootReported = twin.payloads.map(sample => {
    if (!isJsonObject(sample.reported)) {
      throw new Error("The reported-property sample is missing its original reported payload. Choose Validate again before preparing a model.");
    }
    const root = Object.fromEntries(Object.entries(sample.reported).filter(([, value]) => !isJsonObject(value) || value.__t !== "c"));
    return reportedPayload(root);
  });
  const payloads = [...rootTelemetry, ...rootReported];
  if ((telemetry.payloads.length || twin.payloads.length) && !payloads.some(value => Object.keys(value).length > 0)) {
    throw new Error("The sampled records contain no root-level fields for model preparation.");
  }
  return inferPayloadSchema(payloads);
}

export function validateSampleDepth(limit: number): void {
  if (!(SAMPLE_DEPTHS as readonly number[]).includes(limit)) {
    throw new Error("Choose a sample depth of 100, 500, or 1,000 observations per stream.");
  }
}

export function quoteKqlIdentifier(name: string): string {
  if (!name.trim() || /[\x00-\x1f]/.test(name)) throw new Error("Select a valid KQL table name.");
  return `['${name.replace(/\\/g, "\\\\").replace(/'/g, "\\'")}']`;
}

export interface KqlTableResult {
  Columns: { ColumnName: string; ColumnType?: string }[];
  Rows: unknown[][];
}

export function parseKqlTableResult(value: unknown): KqlTableResult {
  if (!isJsonObject(value) || value.error ||
      (value.Exceptions !== undefined && value.Exceptions !== null &&
        (!Array.isArray(value.Exceptions) || value.Exceptions.length > 0))) {
    throw new Error("The database did not complete the query. Check access and query limits, then try again.");
  }
  const tables = value.Tables;
  if (!Array.isArray(tables)) throw new Error("The database returned an unexpected query result.");
  const statusTables = new Set(tables.filter(item => isJsonObject(item) && item.TableName === "QueryStatus"));
  let primaryOrdinal: number | undefined;
  for (const directory of tables.filter(item => isJsonObject(item) && Array.isArray(item.Columns) &&
    item.Columns.some(column => isJsonObject(column) && column.ColumnName === "Ordinal") &&
    item.Columns.some(column => isJsonObject(column) && column.ColumnName === "Kind"))) {
    if (!isJsonObject(directory) || !Array.isArray(directory.Columns) || !Array.isArray(directory.Rows)) {
      throw new Error("The database returned an invalid result directory.");
    }
    const names = directory.Columns.map(column => isJsonObject(column) ? column.ColumnName : undefined);
    const ordinal = names.indexOf("Ordinal"), kind = names.indexOf("Kind");
    for (const row of directory.Rows) {
      if (!Array.isArray(row) || row.length !== names.length || !Number.isInteger(row[ordinal]) ||
          Number(row[ordinal]) < 0 || Number(row[ordinal]) >= tables.length || typeof row[kind] !== "string") {
        throw new Error("The database returned an invalid result directory.");
      }
      const index = Number(row[ordinal]);
      if (row[kind] === "QueryStatus") statusTables.add(tables[index]);
      if (row[kind] === "QueryResult" && primaryOrdinal === undefined) primaryOrdinal = index;
    }
  }
  // v1 can name every table Table_N and identify QueryStatus only in its result directory.
  for (const candidate of tables) {
    if (isJsonObject(candidate) && candidate.TableName !== "PrimaryResult" && Array.isArray(candidate.Columns) &&
        candidate.Columns.some(column => isJsonObject(column) && column.ColumnName === "Severity") &&
        candidate.Columns.some(column => isJsonObject(column) && column.ColumnName === "StatusCode")) {
      statusTables.add(candidate);
    }
  }
  for (const status of statusTables) {
    if (!isJsonObject(status) || !Array.isArray(status.Columns) || !Array.isArray(status.Rows) || !status.Rows.length) {
      throw new Error("The database returned an invalid query status. Try the read again.");
    }
    const names = status.Columns.map(column => isJsonObject(column) ? column.ColumnName : undefined);
    const severity = names.indexOf("Severity"), code = names.indexOf("StatusCode");
    if (names.some(name => typeof name !== "string") || new Set(names).size !== names.length || severity < 0 || code < 0) {
      throw new Error("The database returned an invalid query status. Try the read again.");
    }
    for (const row of status.Rows) {
      if (!Array.isArray(row) || row.length !== names.length ||
          !Number.isInteger(row[severity]) || !Number.isInteger(row[code]) || Number(row[severity]) < 0) {
        throw new Error("The database returned an invalid query status. Try the read again.");
      }
      if (Number(row[severity]) <= 2 || row[code] !== 0) {
        throw new Error("The database did not complete the query. Partial results were discarded. Check query limits and try again.");
      }
    }
  }
  const table = primaryOrdinal !== undefined ? tables[primaryOrdinal] : tables.find(item => isJsonObject(item) &&
    Array.isArray(item.Columns) && Array.isArray(item.Rows) &&
    !statusTables.has(item));
  if (!isJsonObject(table) || !Array.isArray(table.Columns) || !Array.isArray(table.Rows)) {
    throw new Error("The database returned an unexpected table shape. Refresh the source and try again.");
  }
  const columns = table.Columns;
  const rows = table.Rows;
  if (!columns.every(column => isJsonObject(column) && typeof column.ColumnName === "string") ||
      !rows.every(row => Array.isArray(row) && row.length === columns.length)) {
    throw new Error("The database returned an unexpected table shape. Refresh the source and try again.");
  }
  return {
    Columns: columns.map(column => {
      if (!isJsonObject(column)) throw new Error("Invalid query column.");
      return { ColumnName: String(column.ColumnName), ColumnType: String(column.ColumnType || "") };
    }),
    Rows: rows.map(row => {
      if (!Array.isArray(row)) throw new Error("Invalid query row.");
      return row;
    }),
  };
}

interface NativeKqlError {
  code?: string;
  message?: string;
  activityId?: string;
  note?: string;
}
const MAX_ERROR_BYTES = 8192;

function safeNativeMessage(value: unknown, token: string): string | undefined {
  if (typeof value !== "string") return undefined;
  const cleaned = value.split(token).join("[redacted]")
    .replace(/\b(?:Authorization|Cookie|Set-Cookie|X-MS-[\w-]+)["']?\s*[:=][^\r\n]*/gi, "[redacted header]")
    .replace(/\bBearer\s+\S+/gi, "Bearer [redacted]")
    .replace(/\b((?:access_token|refresh_token|client_secret|password|token|api[_-]?key)["']?\s*[:=]\s*)(?:"[^"]*"|'[^']*'|[^,\s;]+)/gi, "$1[redacted]")
    .replace(/[\u0000-\u001f\u007f]/g, " ").replace(/\s+/g, " ").trim();
  return cleaned ? cleaned.slice(0, 600) + (cleaned.length > 600 ? " [truncated]" : "") : undefined;
}

async function readNativeKqlError(response: Response, token: string): Promise<NativeKqlError> {
  if (!response.body) return { note: "The native error body was empty." };
  const reader = response.body.getReader();
  const bytes = new Uint8Array(MAX_ERROR_BYTES);
  let length = 0;
  try {
    while (length < MAX_ERROR_BYTES) {
      const chunk = await reader.read();
      if (chunk.done) break;
      const count = Math.min(chunk.value.length, bytes.length - length);
      bytes.set(chunk.value.subarray(0, count), length); length += count;
      if (length === MAX_ERROR_BYTES) { await reader.cancel(); break; }
    }
  } finally { reader.releaseLock(); }
  const text = new TextDecoder().decode(bytes.subarray(0, Math.min(length, MAX_ERROR_BYTES))).trim();
  let code: unknown, message: unknown;
  try {
    const payload: unknown = JSON.parse(text);
    if (isJsonObject(payload)) {
      const error = isJsonObject(payload.error) ? payload.error : payload;
      code = error.code;
      message = typeof error["@message"] === "string" ? error["@message"] : error.message;
    }
  } catch {
    const native = /^([A-Za-z][\w.-]{0,127}):\s*([^\r\n]+)/.exec(text);
    if (native && !/authorization|cookie|token|secret|password/i.test(native[1])) {
      code = native[1]; message = native[2];
    }
  }
  const safeCode = typeof code === "string" && /^[A-Za-z][\w.-]{0,127}$/.test(code) && !code.includes(token)
    ? code : undefined;
  const safeMessage = safeNativeMessage(message, token);
  const activityId = /\bActivityId["']?\s*[:=]\s*["']?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\b/i.exec(text)?.[1];
  return { code: safeCode, message: safeMessage, activityId,
    note: length === MAX_ERROR_BYTES ? "Native error detail was capped at 8 KiB."
      : !safeCode && !safeMessage ? "No recognized native error code/message was available." : undefined };
}

export class KqlRequestError extends Error {
  readonly status: number;
  readonly code?: string;
  readonly nativeMessage?: string;
  readonly activityId?: string;

  constructor(status: number, native: NativeKqlError) {
    const base = status === 401 || status === 403
      ? `Access to the selected database was denied (${status}). Check your permissions and sign in again.`
      : `The database request failed (${status}).`;
    super([base, native.code || native.message ? `Native error: ${[native.code, native.message].filter(Boolean).join(": ")}` : "",
      native.activityId ? `ActivityId: ${native.activityId}.` : "", native.note || ""].filter(Boolean).join(" "));
    this.name = "KqlRequestError";
    this.status = status; this.code = native.code; this.nativeMessage = native.message; this.activityId = native.activityId;
  }
}

export async function executeKqlRequest(
  source: Pick<AuthoringSource, "queryServiceUri" | "databaseName">,
  token: string,
  csl: string,
  kind: "query" | "mgmt" = "query",
  signal?: AbortSignal,
  fetcher: typeof fetch = fetch
): Promise<KqlTableResult> {
  const uri = new URL(source.queryServiceUri);
  if (uri.protocol !== "https:" || uri.username || uri.password || uri.search || uri.hash ||
      (uri.pathname !== "/" && uri.pathname !== "")) {
    throw new Error("The Eventhouse returned an unsupported query-service address. Refresh the source.");
  }
  if (!source.databaseName.trim() || !token) throw new Error("Select a database and sign in before reading samples.");
  const response = await fetcher(`${uri.origin}/v1/rest/${kind}`, {
    method: "POST",
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", Accept: "application/json" },
    body: JSON.stringify({ db: source.databaseName, csl }),
    signal,
  });
  if (!response.ok) {
    let native: NativeKqlError;
    try {
      signal?.throwIfAborted();
      native = await readNativeKqlError(response, token);
      signal?.throwIfAborted();
    } catch (error: unknown) {
      if (signal?.aborted || (error instanceof Error && error.name === "AbortError")) throw error;
      native = { note: "The native error detail could not be read." };
    }
    throw new KqlRequestError(response.status, native);
  }
  return parseKqlTableResult(await response.json());
}

export const executeAuthoringRead = executeKqlRequest;

export function decodeRawSampleResult(result: KqlTableResult, stream: SampleStream, limit: number): StreamSample {
  const columns = ["data", "headers", "user_headers"];
  const expectedTypes = ["string", "dynamic", "dynamic"];
  if (result.Columns.length !== 3 || !result.Columns.every((column, index) =>
    column.ColumnName === columns[index] &&
      ((index === 0 && ["string", "dynamic", "object"].includes(column.ColumnType?.toLowerCase() || "")) ||
        (column.ColumnType?.toLowerCase() === "object" ? "dynamic" : column.ColumnType?.toLowerCase()) === expectedTypes[index]))) {
    throw new Error(`The ${stream} table must have data:string or dynamic, headers:dynamic, and user_headers:dynamic columns.`);
  }
  if (result.Rows.length > limit) throw new Error("The database exceeded the requested sample limit.");
  const dynamicData = ["dynamic", "object"].includes(result.Columns[0].ColumnType?.toLowerCase() || "");
  return decodeDoSamples(result.Rows.map(row => ({
    data: dynamicData && typeof row[0] !== "string" ? JSON.stringify(row[0]) : row[0],
    headers: row[1], user_headers: row[2],
  })), stream, limit);
}

export async function readPayloadStream(
  source: AuthoringSource, token: string, limit: number, stream: SampleStream,
  signal?: AbortSignal, fetcher: typeof fetch = fetch
): Promise<StreamSample> {
  validateSampleDepth(limit);
  if (source.telemetryTable === source.twinTable) throw new Error("Select separate telemetry and twin-change tables.");
  const table = stream === "telemetry" ? source.telemetryTable : source.twinTable;
  const query = `set truncationmaxsize=8388608;\n${quoteKqlIdentifier(table)}\n` +
    `| top ${limit} by ingestion_time() desc\n| project data, headers, user_headers`;
  const result = await executeAuthoringRead(source, token, query, "query", signal, fetcher);
  try {
    return decodeRawSampleResult(result, stream, limit);
  } catch (error: unknown) {
    const detail = error instanceof Error ? error.message : "The device payload could not be decoded.";
    throw new Error(`We couldn't read ${table} as device data. Check how this data is ingested or choose another table. ${detail}`);
  }
}
export async function readPayloadSamples(
  source: AuthoringSource,
  token: string,
  limit: number,
  signal?: AbortSignal,
  fetcher: typeof fetch = fetch
): Promise<{ telemetry: StreamSample; twin: StreamSample }> {
  validateSampleDepth(limit);
  if (source.telemetryTable === source.twinTable) throw new Error("Select separate telemetry and twin-change tables.");
  const [telemetry, twin] = await Promise.all([
    readPayloadStream(source, token, limit, "telemetry", signal, fetcher),
    readPayloadStream(source, token, limit, "twin", signal, fetcher),
  ]);
  return { telemetry, twin };
}

export async function readAuthoringSamples(
  source: AuthoringSource, token: string, limit: number, signal?: AbortSignal, fetcher: typeof fetch = fetch
): Promise<AuthoringSample> {
  const { telemetry, twin } = await readPayloadSamples(source, token, limit, signal, fetcher);
  return { limit, telemetry, twin,
    schema: inferRootPreparationSchema(telemetry, twin),
  };
}
