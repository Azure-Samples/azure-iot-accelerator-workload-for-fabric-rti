// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import {
  canonicalJson, decodeRawSampleResult, executeAuthoringRead, isJsonObject, quoteKqlIdentifier, validateSampleDepth,
  type AuthoringSource, type JsonObject, type JsonValue, type KqlTableResult, type PayloadSample, type SampleStream, type StreamSample,
} from "./ModelAuthoring.ts";
import { requireThingModelId, selectThingModelRoot, type ThingModelResult } from "./ThingModelParser.ts";
import type { DtdlParseResult } from "./DtdlModelParser.ts";
import { bindNativeMetadata, nativeProjection } from "./HubTransport.ts";
import { dtdlFieldMetadata, wotFieldMetadata, type DeclaredFieldMetadata } from "./FieldSemantics.ts";

export type FieldRole = "telemetry" | "twin";
export interface ModelField {
  id: string;
  path: string[];
  component?: string;
  column: string;
  title: string;
  description?: string;
  unit?: string;
  semanticTypes: string[];
  codes: { value: JsonValue; label?: string; description?: string }[];
  schema: JsonValue;
  kustoType: string;
  included: boolean;
  role?: FieldRole;
  roleSource?: "model" | "hint" | "manual";
  declaredMetadata?: DeclaredFieldMetadata;
}
export interface HubModel {
  kind: "wot" | "dtdl";
  title: string;
  semanticTitle?: string;
  id?: string;
  documentVersion?: string;
  description?: string;
  document: JsonValue;
  declaredModelMetadata?: JsonObject;
  fields: ModelField[];
  ignored?: ThingModelResult["ignored"];
}
export function modelValidation(model: HubModel | undefined): string[] {
  if (!model) return ["Upload a device model to continue"];
  if (model.kind !== "wot") return [];
  try {
    if (requireThingModelId(model.document) !== model.id) {
      return ["The model ID has changed. Upload the file again."];
    }
    return [];
  } catch (error: unknown) {
    return [error instanceof Error ? error.message : "Choose a Thing Model with a valid root id."];
  }
}
export interface ModelReference {
  id: string;
}
export interface ProcessingScope {
  mode: "unselected" | "data" | "reference";
  // The page supplies an explicit choice; omitted preserves existing programmatic predicate semantics.
  additional?: "unselected" | "none" | "specified";
  operator: "and" | "or";
  telemetry: string;
  twin: string;
  reference?: ModelReference;
}
export const RAW_SCOPE_CONTRACT = "decoded-single-observation-raw-row-v2";
export const COVERAGE_SAMPLING_CONTRACT = {
  rawRowLimit: 10000, observationLimit: 100, order: "ingestion-time-desc",
} as const;
export interface ModelingInput {
  source: AuthoringSource;
  model: HubModel;
  fields: ModelField[];
  scope: ProcessingScope;
  limit: number;
}
export interface FieldEvidence {
  fieldId: string;
  telemetry: number;
  twin: number;
  state: "telemetry" | "twin" | "both" | "missing";
}
export interface RoleEvidenceProfile {
  bindings: {
    component?: string;
    path: string[];
    telemetry: number;
    twin: number;
    twinWithComponents: number;
  }[];
}
export interface CoverageSnapshot {
  fingerprint: string;
  kind: "processing" | "preview";
  telemetry: StreamSample;
  twin: StreamSample;
  fields: FieldEvidence[];
  nativeIssues: string[];
  checkedAt: string;
  previewPredicates?: { telemetry: string; twin: string };
  sourceProbe?: { telemetry: StreamSample; twin: StreamSample };
}

export const kqlString = (value: string): string => JSON.stringify(value);
export const pointer = (path: string[]): string => "/" + path.map(part => part.replace(/~/g, "~0").replace(/\//g, "~1")).join("/");
const text = (value: JsonObject, singular: string, plural: string): string | undefined => {
  const translated = value[plural];
  if (isJsonObject(translated)) {
    if (typeof translated["en-US"] === "string") return translated["en-US"];
    if (typeof translated.en === "string") return translated.en;
  }
  return typeof value[singular] === "string" ? value[singular] : undefined;
};
function types(value: JsonValue | undefined): string[] {
  return (Array.isArray(value) ? value : [value]).filter((item): item is string => typeof item === "string");
}
function outputName(path: string[]): string {
  const name = path.join("_").replace(/[^a-zA-Z0-9_]/g, "_");
  const normalized = /^[a-zA-Z_]/.test(name) ? name : `field_${name}`;
  return ["deviceId", "event_time", "sourceHubId"].includes(normalized) ? `${normalized}_value` : normalized;
}
function schemaType(schema: JsonObject): string {
  const mapping: Record<string, string> = { number: "real", integer: "long", boolean: "bool", string: "string" };
  return typeof schema.type === "string" ? mapping[schema.type] || "dynamic" : "dynamic";
}

export function fromThingModel(result: ThingModelResult): HubModel {
  const id = requireThingModelId(result.document);
  if (result.id !== id) throw new Error("The model ID has changed. Upload the file again.");
  if (result.blockers.length || !result.effectiveDocument) {
    throw new Error("Resolve the required Thing Model dependencies before selecting fields.");
  }
  const fields: ModelField[] = [];
  const add = (schema: JsonObject, path: string[], modelPath: string[], prefix: string[], affordance = schema) => {
    const codes: ModelField["codes"] = Array.isArray(schema.enum) ? schema.enum.map(value => ({ value })) : [];
    fields.push({
      id: pointer(modelPath), path, column: outputName([...prefix, ...path]),
      title: text(affordance, "title", "titles") || text(schema, "title", "titles") || path[path.length - 1],
      description: text(affordance, "description", "descriptions") || text(schema, "description", "descriptions"),
      unit: typeof schema.unit === "string" ? schema.unit : undefined,
      semanticTypes: [...new Set([...types(affordance["@type"]), ...types(schema["@type"])])], codes, schema, kustoType: schemaType(schema), included: true,
      declaredMetadata: wotFieldMetadata(schema, affordance, modelPath[0] === "properties"),
    });
  };
  const document = result.effectiveDocument;
  const root = selectThingModelRoot(result.document);
  const declaredModelMetadata: JsonObject = Object.fromEntries(["id", "title", "titles", "description", "descriptions"]
    .filter(key => Object.prototype.hasOwnProperty.call(root, key)).map(key => [key, root[key]]));
  if (isJsonObject(root.version) && typeof root.version.model === "string") declaredModelMetadata.version = root.version.model;
  if (isJsonObject(document.properties)) {
    for (const [name, schema] of Object.entries(document.properties)) {
      if (isJsonObject(schema)) add(schema, [name], ["properties", name], []);
    }
  }
  if (isJsonObject(document.events)) {
    for (const [name, event] of Object.entries(document.events)) {
      if (isJsonObject(event)) {
        add(isJsonObject(event.data) ? event.data : {}, [name], ["events", name, "data"], [], event);
      }
    }
  }
  return { kind: "wot", title: result.title, semanticTitle: text(document, "title", "titles"),
    id: result.id, documentVersion: result.documentVersion,
    description: result.description, document: result.document, declaredModelMetadata, fields,
    ...(result.ignored?.length ? { ignored: result.ignored } : {}) };
}

export function fromDtdlModel(result: DtdlParseResult, document: JsonValue): HubModel {
  const fields = [...result.telemetries, ...result.properties].map(capability => {
    const schema: JsonValue = capability.declaredSchema === undefined ? capability.schema
      : JSON.parse(JSON.stringify(capability.declaredSchema));
    const role: FieldRole = capability.kind === "Telemetry" ? "telemetry" : "twin";
    return {
      id: pointer(capability.component ? ["components", capability.component, capability.kind, capability.leafName]
        : ["root", capability.kind, capability.leafName]),
      path: [capability.leafName], component: capability.component,
      column: capability.name, title: capability.displayName, description: capability.description,
      unit: capability.unit, semanticTypes: capability.semanticTypes || [], codes: capability.enumValues || [],
      schema, kustoType: capability.kustoType, included: capability.included, role, roleSource: "model" as const,
      declaredMetadata: dtdlFieldMetadata(schema, {
        ...(capability.unit !== undefined ? { unit: capability.unit } : {}),
        ...(capability.semanticTypes?.length ? { semanticTypes: capability.semanticTypes } : {}),
        ...(capability.enumValues?.length ? { enumValues: capability.enumValues } : {}),
        ...(capability.declaredBy !== undefined ? { declaredBy: capability.declaredBy } : {}),
      }, capability.declaredAnnotations, capability.declaredSchema !== undefined),
    };
  });
  return { kind: "dtdl", title: result.displayName,
    semanticTitle: result.displayName && result.displayName !== result.modelId ? result.displayName : undefined, id: result.modelId,
    documentVersion: /;(\d+)$/.exec(result.modelId)?.[1], description: result.description, document,
    ...(result.declaredModelMetadata ? { declaredModelMetadata: result.declaredModelMetadata } : {}), fields };
}

export function fieldValidation(fields: ModelField[], { includeRoleErrors = true }: { includeRoleErrors?: boolean } = {}): string[] {
  const selected = fields.filter(field => field.included);
  const errors: string[] = selected.length ? [] : ["Select at least one field to continue"];
  const names = new Set(["deviceId", "event_time", "sourceHubId"]);
  const ids = new Set<string>();
  for (const field of selected) {
    if (field.path.length !== 1 || !field.path[0]) errors.push(`Upload the model again: ${field.title} must bind one declared whole value, not nested leaf fields.`);
    if (!field.role && includeRoleErrors) errors.push(`Select a role for ${field.title}`);
    if (!/^[a-zA-Z_]\w*$/.test(field.column)) errors.push(`The model produces an unsupported column name "${field.column}". Select another field or upload a model with distinct letter/underscore-based names.`);
    if (names.has(field.column)) errors.push(`Two selected fields map to "${field.column}". Select one of them, or upload a model with distinct names. Paths: ${selected.filter(item => item.column === field.column).map(item => item.id).join(", ")}.`);
    if (ids.has(field.id)) errors.push(`Selected fields share the model identity "${field.id}". Upload a model with distinct root/component paths.`);
    ids.add(field.id);
    names.add(field.column);
  }
  return errors;
}

export function orderedModeledFields(fields: ModelField[]): ModelField[] {
  const selected = fields.filter(field => field.included);
  return [...selected.filter(field => field.role === "telemetry"), ...selected.filter(field => field.role === "twin")];
}

function atPath(value: JsonObject, path: string[]): { found: boolean; value?: JsonValue } {
  let current: JsonValue = value;
  for (const key of path) {
    if (!isJsonObject(current) || !Object.prototype.hasOwnProperty.call(current, key)) return { found: false };
    current = current[key];
  }
  return { found: true, value: current };
}
export function fieldValue(sample: PayloadSample, field: Pick<ModelField, "component" | "path">, stream: SampleStream,
  componentEnvelopes = !!field.component): { found: boolean; value?: JsonValue } {
  if (stream === "twin") {
    const reported = sample.reported || sample.value;
    const envelope = field.component ? reported[field.component] : undefined;
    const value = reported[field.path[0]];
    if (field.component && (!isJsonObject(envelope) || envelope.__t !== "c")) return { found: false };
    if (!field.component && componentEnvelopes && isJsonObject(value) && value.__t === "c") return { found: false };
    return atPath(reported, field.component ? [field.component, ...field.path] : field.path);
  }
  if (sample.native?.profile === "unknown" || sample.native?.issues.some(issue => issue.includes("component"))) return { found: false };
  if (field.component) return sample.component === field.component ? atPath(sample.value, field.path) : { found: false };
  if (sample.component) return { found: false };
  return atPath(sample.value, field.path);
}
function evidenceCounts(fieldId: string, telemetry: number, twin: number): FieldEvidence {
  return { fieldId, telemetry, twin, state: telemetry && twin ? "both" : telemetry ? "telemetry" : twin ? "twin" : "missing" };
}
export function prepareRoleEvidence(telemetry: StreamSample, twin: StreamSample): RoleEvidenceProfile {
  const bindings = new Map<string, RoleEvidenceProfile["bindings"][number]>();
  for (const stream of ["telemetry", "twin"] as const) {
    for (const sample of (stream === "telemetry" ? telemetry : twin).payloads) {
      const value = stream === "twin" ? sample.reported || sample.value : sample.value;
      const candidates: Pick<ModelField, "component" | "path">[] = Object.keys(value).map(name => ({
        path: [name], ...(stream === "telemetry" && sample.component ? { component: sample.component } : {}),
      }));
      if (stream === "twin") for (const [component, envelope] of Object.entries(value)) {
        if (isJsonObject(envelope) && envelope.__t === "c") {
          candidates.push(...Object.keys(envelope).map(name => ({ component, path: [name] })));
        }
      }
      for (const candidate of candidates) {
        const key = JSON.stringify([candidate.component || "", candidate.path]);
        let binding = bindings.get(key);
        if (!binding) {
          binding = { ...candidate, telemetry: 0, twin: 0, twinWithComponents: 0 };
          bindings.set(key, binding);
        }
        if (stream === "telemetry") binding.telemetry += Number(fieldValue(sample, candidate, stream).found);
        else {
          binding.twin += Number(fieldValue(sample, candidate, stream, false).found);
          binding.twinWithComponents += Number(fieldValue(sample, candidate, stream, true).found);
        }
      }
    }
  }
  return { bindings: [...bindings.values()] };
}
export function matchRoleEvidence(fields: ModelField[], profile: RoleEvidenceProfile): FieldEvidence[] {
  const componentEnvelopes = fields.some(field => field.included && field.component);
  const bindings = new Map(profile.bindings.map(binding => [JSON.stringify([binding.component || "", binding.path]), binding]));
  return fields.map(field => {
    const binding = bindings.get(JSON.stringify([field.component || "", field.path]));
    return evidenceCounts(field.id, binding?.telemetry || 0,
      (componentEnvelopes ? binding?.twinWithComponents : binding?.twin) || 0);
  });
}
export function inspectFieldEvidence(fields: ModelField[], telemetry: StreamSample, twin: StreamSample): FieldEvidence[] {
  const componentEnvelopes = fields.some(field => field.included && field.component);
  return fields.map(field => {
    const t = telemetry.payloads.filter(sample => fieldValue(sample, field, "telemetry", componentEnvelopes).found).length;
    const p = twin.payloads.filter(sample => fieldValue(sample, field, "twin", componentEnvelopes).found).length;
    return evidenceCounts(field.id, t, p);
  });
}
export function sampleExtent(sample: StreamSample): string | undefined {
  const times = sample.payloads.map(payload => Date.parse(payload.eventTime || "")).filter(Number.isFinite);
  return times.length ? `${new Date(Math.min(...times)).toISOString()} to ${new Date(Math.max(...times)).toISOString()}` : undefined;
}
export function omittedObservedFields(fields: ModelField[], samples: { telemetry: StreamSample; twin: StreamSample }): { telemetry: string[]; twin: string[] } {
  const omitted = (stream: SampleStream): string[] => {
    const paths: string[][] = [];
    const hasComponents = fields.some(field => field.included && field.component);
    for (const sample of samples[stream].payloads) {
      for (const [key, value] of Object.entries(sample.value)) {
        const reported = sample.reported || sample.value;
        if (stream === "twin" && hasComponents && isJsonObject(reported[key]) && reported[key].__t === "c" && isJsonObject(value)) {
          paths.push(...Object.keys(value).map(member => [key, member]));
        } else paths.push(stream === "telemetry" && sample.component ? [sample.component, key] : [key]);
      }
    }
    const selected = fields.filter(field => field.included && field.role === stream)
      .map(field => field.component ? [field.component, ...field.path] : field.path);
    return [...new Set(paths.filter(path => !selected.some(binding =>
      binding.length <= path.length && binding.every((part, index) => part === path[index]))).map(pointer))].sort();
  };
  return { telemetry: omitted("telemetry"), twin: omitted("twin") };
}
export function applyRoleHints(fields: ModelField[], evidence: FieldEvidence[]): ModelField[] {
  return fields.map(field => {
    if (field.roleSource === "manual" || field.roleSource === "model") return field;
    const hint = evidence.find(item => item.fieldId === field.id)?.state;
    const role = hint === "telemetry" || hint === "twin" ? hint : undefined;
    return { ...field, role, roleSource: role ? "hint" : undefined };
  });
}
export function clearAutomaticHints(fields: ModelField[]): ModelField[] {
  return fields.map(field => field.roleSource === "hint" ? { ...field, role: undefined, roleSource: undefined } : field);
}
export function setFieldRole(field: ModelField, role: FieldRole | undefined): ModelField {
  if (field.roleSource === "model" && field.role !== role) throw new Error(`The role for ${field.title} is declared by your DTDL model. Upload the model again to change it.`);
  return { ...field, role, roleSource: "manual" };
}

export function assignableReference(model: HubModel): ModelReference {
  if (model.kind !== "wot") throw new Error("Model assignment is available only for WoT Thing Models.");
  const errors = modelValidation(model);
  if (errors.length) throw new Error(errors.join(" "));
  return { id: requireThingModelId(model.document) };
}

/** Only predicates, not query statements. Quoted punctuation remains valid. */
export function validatePredicate(predicate: string): void {
  let quote = "", depth = 0;
  for (let index = 0; index < predicate.length; index++) {
    const char = predicate[index];
    if (quote) {
      if (char === "\\") { index++; continue; }
      if (char === quote) {
        if (predicate[index + 1] === quote) index++;
        else quote = "";
      }
      continue;
    }
    if (char === "'" || char === '"') quote = char;
    else if (char === "(") depth++;
    else if (char === ")" && --depth < 0) throw new Error("The scope predicate has an unmatched parenthesis.");
    else if (char === ";" || char === "|" || (char === "/" && ["/", "*"].includes(predicate[index + 1]))) {
      throw new Error("Enter a predicate only, without query statements, pipes or comments.");
    }
  }
  if (quote || depth !== 0) throw new Error("Close the scope predicate's quotes and parentheses.");
}
export function effectiveScope(scope: ProcessingScope, stream: SampleStream, preview = false): string {
  if (scope.additional === "unselected" && !preview) throw new Error("Choose whether to specify additional conditions before verifying model coverage.");
  const extra = scope.additional === "none" && !preview ? "" : scope[stream].trim();
  validatePredicate(extra);
  if (scope.mode === "unselected" && !preview) throw new Error("Choose a model reference filter before verifying coverage.");
  if (scope.mode !== "reference" || preview) return extra || "true";
  if (!scope.reference?.id) throw new Error("Use a model with a stable ID, or choose device/data filters.");
  const type = `tostring(data["payload"]["memberTypeIRI"]) == ${kqlString(scope.reference.id)}`;
  return extra ? `(${type}) ${scope.operator} (${extra})` : `(${type})`;
}
export function modelingFingerprint(input: ModelingInput): string {
  return canonicalJson({ architecture: "v1-declared-whole-values", timeColumn: "event_time", predicateContract: RAW_SCOPE_CONTRACT, source: input.source, model: input.model, fields: input.fields,
    scope: input.scope, limit: input.limit, coverageSampling: COVERAGE_SAMPLING_CONTRACT });
}
export function coverageReady(input: ModelingInput, snapshot: CoverageSnapshot | undefined, waiting: boolean): boolean {
  return input.scope.mode !== "unselected" && input.scope.additional !== "unselected" && modelValidation(input.model).length === 0 && !!snapshot && snapshot.kind === "processing" && snapshot.fingerprint === modelingFingerprint(input) &&
    !snapshot.telemetry.batchedRows && !snapshot.twin.batchedRows &&
    !snapshot.sourceProbe?.telemetry.batchedRows && !snapshot.sourceProbe?.twin.batchedRows &&
    fieldValidation(input.fields).length === 0 && (snapshot.telemetry.payloads.length > 0 || waiting);
}
export function verifiedEmptyCoverage(input: ModelingInput, snapshot: CoverageSnapshot | undefined): boolean {
  return !!snapshot && snapshot.telemetry.payloads.length === 0 && !!snapshot.sourceProbe?.telemetry.payloads.length &&
    !!snapshot.sourceProbe.twin.payloads.length &&
    nativeBindingIssues(snapshot.sourceProbe.telemetry, snapshot.sourceProbe.twin).length === 0 &&
    coverageReady(input, snapshot, true);
}
export function scopeSummary(scope: ProcessingScope, stream: SampleStream): string {
  if (scope.mode === "data" && (scope.additional === "none" || !scope[stream].trim()) && scope.additional !== "unselected") {
    return stream === "telemetry" ? "All incoming telemetry in the selected table" : "All incoming twin changes in the selected table";
  }
  return effectiveScope(scope, stream);
}

export function bindingStatus(sample: StreamSample, stream: SampleStream): { identity: string; component: string } {
  if (!sample.payloads.length) return { identity: "Not verified - no matching sample", component: "Not verified - no matching sample" };
  const bindings = sample.payloads.map(item => bindNativeMetadata(item.headers || {}, item.userHeaders || {}, stream));
  return {
    identity: bindings.every(binding => binding.identity && !binding.issues.length) ? "Matched in this sample" : "Needs review",
    component: stream === "twin" ? "Reported-property paths"
      : bindings.some(binding => binding.profile === "unknown" || binding.issues.length) ? "Needs review"
      : bindings.some(binding => binding.component) ? "Component metadata recognized in this sample" : "Root traffic only in this sample",
  };
}

export function nativeIdentity(sample: PayloadSample, stream: SampleStream): string | undefined {
  return bindNativeMetadata(sample.headers || {}, sample.userHeaders || {}, stream).identity;
}
export function processingBindingSamples(snapshot: Pick<CoverageSnapshot, "telemetry" | "twin" | "sourceProbe">) {
  return !snapshot.telemetry.payloads.length && snapshot.sourceProbe
    ? { ...snapshot.sourceProbe, source: "source" as const }
    : { telemetry: snapshot.telemetry, twin: snapshot.twin, source: "processing" as const };
}
export function nativeBindingIssues(telemetry: StreamSample, twin: StreamSample, enrichTwins = true): string[] {
  if (!enrichTwins) twin = { rawRows: 0, observations: 0, skippedWithoutReportedProperties: 0, payloads: [] };
  const issues: string[] = [];
  if (telemetry.payloads.some(sample => !nativeIdentity(sample, "telemetry"))) {
    issues.push("Telemetry device identity: not verified in the transport metadata.");
  }
  if (twin.payloads.some(sample => !nativeIdentity(sample, "twin"))) {
    issues.push("Twin device identity: not verified in the transport metadata.");
  }
  const twinIds = new Set(twin.payloads.map(sample => nativeIdentity(sample, "twin")));
  if (enrichTwins && telemetry.payloads.some(sample => {
    const identity = nativeIdentity(sample, "telemetry");
    return identity?.includes("/") && !twinIds.has(identity);
  })) issues.push("Module telemetry identities have not been matched to native device/module twin identities in this sample.");
  if (!telemetry.payloads.length) issues.push("Native telemetry identity is unverified because the sample is empty.");
  for (const stream of ["telemetry", "twin"] as const) {
    for (const sample of (stream === "telemetry" ? telemetry : twin).payloads) {
      for (const issue of bindNativeMetadata(sample.headers || {}, sample.userHeaders || {}, stream).issues) {
        issues.push(`${stream === "telemetry" ? "Telemetry" : "Twin changes"}: ${issue}`);
      }
    }
  }
  return [...new Set(issues)];
}

function observationProjection(stream: SampleStream): string[] {
  return [
    "| extend __raw = parse_json(data)",
    '| mv-expand __observation = iff(gettype(__raw) == "array", __raw, pack_array(__raw))',
    ...observationContractProjection(stream),
  ];
}
function observationContractProjection(stream: SampleStream): string[] {
  return [
    '| where tostring(__observation.schema_id) == "com.do.telemetry.observation.v1" and tostring(__observation.schema_version) == "1.0" and tostring(__observation.contract_version) == "1.0"',
    `| where tostring(__observation.event_kind) == ${kqlString(stream === "telemetry" ? "measurement" : "lifecycle")}`,
  ];
}
function normalizedProjection(stream: SampleStream): string[] {
  return [
    "| extend data = __observation.payload.value, event_time = todatetime(__observation.event_time), memberTypeIRI = tostring(__observation.payload.memberTypeIRI)",
    "| extend user_headers = parse_json(tostring(user_headers)), headers = parse_json(tostring(headers))",
    ...nativeProjection(stream),
    `| extend payload = ${stream === "telemetry" ? "data" : "data.properties.reported"}`,
    '| where gettype(payload) == "dictionary"',
  ];
}
export function doProjection(stream: SampleStream): string {
  return [...observationProjection(stream), ...normalizedProjection(stream)].join("\n");
}
export function scopedDoProjection(scope: ProcessingScope, stream: SampleStream, preview = false): string {
  return ["| extend data = parse_json(tostring(data))",
    `| where ${effectiveScope(scope, stream, preview)}`,
    "| extend __observation = data",
    ...observationContractProjection(stream),
    ...normalizedProjection(stream)].join("\n");
}
export function requireSingleObservationRows(sample: StreamSample, stream: SampleStream): void {
  if (sample.batchedRows) throw new Error(
    `${stream === "telemetry" ? "Telemetry" : "Reported properties"} raw source contains top-level observation batches. ` +
    "Model scoping requires one observation per raw row. Array-valued fields inside an observation remain supported."
  );
}

function coverageRawInput(input: ModelingInput, stream: SampleStream): string {
  if (input.source.telemetryTable === input.source.twinTable) throw new Error("Select separate telemetry and twin-change tables.");
  const table = stream === "telemetry" ? input.source.telemetryTable : input.source.twinTable;
  // Unknown ingestion clocks must be detected, not silently ranked behind timestamped rows.
  return `set truncationmaxsize=8388608;\n${quoteKqlIdentifier(table)}\n` +
    "| extend __coverage_ingestion_time = ingestion_time()\n" +
    `| top ${COVERAGE_SAMPLING_CONTRACT.rawRowLimit} by __coverage_ingestion_time desc nulls first`;
}
function coverageLimit(input: ModelingInput): number {
  return Math.min(input.limit, COVERAGE_SAMPLING_CONTRACT.observationLimit);
}
function decodeCoverageResult(result: KqlTableResult, stream: SampleStream, limit: number): StreamSample {
  if (result.Columns.length !== 4 || result.Columns[3].ColumnName !== "__coverage_ingestion_time" ||
      result.Columns[3].ColumnType?.toLowerCase() !== "datetime" ||
      result.Rows.some(row => typeof row[3] !== "string" || !Number.isFinite(Date.parse(row[3])))) {
    throw new Error("Recent coverage could not be verified because ingestion timestamps are unavailable or invalid in the raw table. Check its ingestion-time policy and incoming data.");
  }
  return decodeRawSampleResult({
    Columns: result.Columns.slice(0, 3), Rows: result.Rows.map(row => row.slice(0, 3)),
  }, stream, limit);
}
async function readCoverageProbe(input: ModelingInput, token: string, stream: SampleStream,
  signal?: AbortSignal, fetcher: typeof fetch = fetch): Promise<StreamSample> {
  const limit = coverageLimit(input);
  const query = `${coverageRawInput(input, stream)}\n| top ${limit} by __coverage_ingestion_time desc nulls first\n` +
    "| project data, headers, user_headers, __coverage_ingestion_time";
  return decodeCoverageResult(await executeAuthoringRead(input.source, token, query, "query", signal, fetcher), stream, limit);
}
async function readScopedStream(input: ModelingInput, token: string, stream: SampleStream, kind: "processing" | "preview",
  previewScope?: { telemetry: string; twin: string }, signal?: AbortSignal, fetcher: typeof fetch = fetch): Promise<StreamSample> {
  const scope = kind === "preview"
    ? { ...input.scope, telemetry: previewScope?.telemetry || "", twin: previewScope?.twin || "" } : input.scope;
  const limit = coverageLimit(input);
  const query = `${coverageRawInput(input, stream)}\n${scopedDoProjection(scope, stream, kind === "preview")}\n` +
    `| top ${limit} by __coverage_ingestion_time desc nulls first\n` +
    "| project data = tostring(__observation), headers, user_headers, __coverage_ingestion_time";
  return decodeCoverageResult(await executeAuthoringRead(input.source, token, query, "query", signal, fetcher), stream, limit);
}
export async function inspectStreamCoverage(input: ModelingInput, token: string, stream: SampleStream,
  signal?: AbortSignal, fetcher: typeof fetch = fetch): Promise<{ sample: StreamSample; probe: StreamSample; checkedAt: string }> {
  const errors = modelValidation(input.model);
  if (errors.length) throw new Error(errors.join(" "));
  validateSampleDepth(input.limit);
  if (fieldValidation(input.fields).length) throw new Error("Select fields and confirm their roles before checking coverage.");
  effectiveScope(input.scope, stream);
  const probe = await readCoverageProbe(input, token, stream, signal, fetcher);
  requireSingleObservationRows(probe, stream);
  if (signal?.aborted) throw new DOMException("Coverage inspection was cancelled.", "AbortError");
  const sample = await readScopedStream(input, token, stream, "processing", undefined, signal, fetcher);
  return { sample, probe, checkedAt: new Date().toISOString() };
}
export async function inspectScopedCoverage(
  input: ModelingInput, token: string, kind: "processing" | "preview", previewScope?: { telemetry: string; twin: string },
  signal?: AbortSignal, fetcher: typeof fetch = fetch
): Promise<CoverageSnapshot> {
  const errors = modelValidation(input.model);
  if (errors.length) throw new Error(errors.join(" "));
  validateSampleDepth(input.limit);
  if (fieldValidation(input.fields).length) throw new Error("Select fields and confirm their roles before checking coverage.");
  const validationScope = kind === "preview"
    ? { ...input.scope, telemetry: previewScope?.telemetry || "", twin: previewScope?.twin || "" } : input.scope;
  for (const stream of ["telemetry", "twin"] as const) effectiveScope(validationScope, stream, kind === "preview");
  // Probe the actual envelope before filtering, so incompatible data isn't reported as zero matches.
  const [telemetryProbe, twinProbe] = await Promise.all([
    readCoverageProbe(input, token, "telemetry", signal, fetcher),
    readCoverageProbe(input, token, "twin", signal, fetcher),
  ]);
  const sourceProbe = { telemetry: telemetryProbe, twin: twinProbe };
  requireSingleObservationRows(sourceProbe.telemetry, "telemetry");
  requireSingleObservationRows(sourceProbe.twin, "twin");
  if (signal?.aborted) throw new DOMException("Coverage inspection was cancelled.", "AbortError");
  const [telemetry, twin] = await Promise.all([
    readScopedStream(input, token, "telemetry", kind, previewScope, signal, fetcher),
    readScopedStream(input, token, "twin", kind, previewScope, signal, fetcher),
  ]);
  return { fingerprint: modelingFingerprint(input), kind, telemetry, twin, sourceProbe,
    ...(kind === "preview" ? { previewPredicates: { telemetry: previewScope?.telemetry || "", twin: previewScope?.twin || "" } } : {}),
    fields: inspectFieldEvidence(input.fields, telemetry, twin), nativeIssues: nativeBindingIssues(telemetry, twin),
    checkedAt: new Date().toISOString() };
}
