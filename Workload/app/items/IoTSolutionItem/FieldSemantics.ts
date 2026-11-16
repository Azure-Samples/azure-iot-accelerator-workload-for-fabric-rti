// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { isJsonObject, type JsonObject, type JsonValue } from "./ModelAuthoring.ts";
import type { ModelField } from "./HubModeling.ts";
import { DOCSTRING_LIMIT, emitJsonDocstring, isDocstringValue, parseDocstringJson } from "./Docstrings.ts";

export interface DeclaredFieldMetadata {
  kind: "wot" | "dtdl";
  schema: JsonValue;
  annotations: JsonObject;
  declaration?: JsonObject;
  schemaDeclared?: boolean;
}

const DATA_SCHEMA_KEYS = [
  "@type", "title", "titles", "description", "descriptions", "const", "default", "unit", "oneOf", "enum",
  "readOnly", "writeOnly", "format", "type", "minimum", "exclusiveMinimum", "maximum", "exclusiveMaximum",
  "multipleOf", "minLength", "maxLength", "pattern", "contentEncoding", "contentMediaType", "items",
  "minItems", "maxItems", "properties", "required",
] as const;
const ANNOTATION_KEYS = ["@type", "title", "titles", "description", "descriptions"] as const;
export const COLUMN_SEMANTICS_PREFIX = "@iot.meta columnSemantics=";
export const FIELD_SEMANTIC_INFO_PREFIX = "@iot.meta fieldSemanticInfo=";
// Legacy read support only.
export const FLAT_COLUMN_SEMANTICS = "@iot.meta columnSemanticsFormat=facts-v2";
const FACT_KEYS: readonly string[] = [...DATA_SCHEMA_KEYS, "@id", "schema", "valueSchema", "enumValues", "fields",
  "mapKey", "mapValue", "elementSchema", "writable", "observable", "comment"];

export interface ColumnSemanticsContext { kind?: "wot" | "dtdl"; id?: string; documentVersion?: string }
export interface FlatColumnSemantics {
  version: 2;
  field: { name: string; facts: JsonObject };
  bounded: boolean;
  model?: ColumnSemanticsContext;
}
export interface LegacyColumnSemantics {
  version: 1;
  model: { kind: "wot" | "dtdl"; id?: string; documentVersion?: string };
  field: { name: string; metadata: DeclaredFieldMetadata };
}
export interface JsonColumnSemantics {
  version: 3;
  field: { name: string; facts: JsonObject };
  bounded: boolean;
  model?: ColumnSemanticsContext;
}

export function publishedFieldFacts(field: ModelField): JsonObject {
  const metadata = field.declaredMetadata;
  const identity: JsonObject = { name: field.path.join("."), ...(field.component ? { component: field.component } : {}) };
  let facts: JsonObject;
  if (metadata?.kind === "wot") {
    facts = field.id.startsWith("/events/")
      ? { ...identity, ...metadata.annotations, data: metadata.schema }
      : { ...identity, ...(isJsonObject(metadata.schema) ? metadata.schema : {}), ...metadata.annotations };
  } else if (metadata?.kind === "dtdl") {
    const declaration = metadata.declaration || {
      ...(field.description !== undefined ? { description: field.description } : {}),
      ...(field.unit !== undefined ? { unit: field.unit } : {}),
      ...(field.semanticTypes.length ? { "@type": field.semanticTypes } : {}),
    };
    if (declaration.name !== undefined && declaration.name !== identity.name) throw new Error("The DTDL field declaration no longer matches its identity.");
    facts = { ...identity, ...declaration, ...(metadata.schemaDeclared === false ? {} : { schema: metadata.schema }) };
  } else {
    facts = { ...identity, ...(field.description !== undefined ? { description: field.description } : {}),
      ...(field.unit !== undefined ? { unit: field.unit } : {}),
      ...(field.semanticTypes.length ? { "@type": field.semanticTypes } : {}),
      ...(field.codes.length ? { codes: field.codes.map(code => ({ ...code })) } : {}) };
  }
  if (field.role) facts.role = field.role === "twin" ? "twin property (last-known whole received value)" : "telemetry";
  return facts;
}

export function columnDocstring(facts: JsonObject): string {
  validateSemanticInfo(facts);
  const result = emitJsonDocstring(FIELD_SEMANTIC_INFO_PREFIX, facts);
  if (readColumnSemantics(result)?.version !== 3) throw new Error("The published field identity could not fit safely in the docstring.");
  return result;
}

function dataSchema(value: JsonValue): JsonValue {
  if (Array.isArray(value)) return value.map(dataSchema);
  if (!isJsonObject(value)) return value;
  return Object.fromEntries(DATA_SCHEMA_KEYS.filter(key => Object.prototype.hasOwnProperty.call(value, key)).map(key => {
    const declared = value[key];
    if (key === "properties" && isJsonObject(declared)) {
      return [key, Object.fromEntries(Object.entries(declared).map(([name, schema]) => [name, dataSchema(schema)]))];
    }
    return [key, key === "items" || key === "oneOf" ? dataSchema(declared) : declared];
  }));
}

export function wotFieldMetadata(schema: JsonObject, affordance: JsonObject, property: boolean): DeclaredFieldMetadata {
  const keys: readonly string[] = property ? [...ANNOTATION_KEYS, "observable"] : ANNOTATION_KEYS;
  return { kind: "wot", schema: dataSchema(schema),
    annotations: Object.fromEntries(keys.filter(key => Object.prototype.hasOwnProperty.call(affordance, key))
      .map(key => [key, affordance[key]])) };
}

export function dtdlFieldMetadata(schema: JsonValue, annotations: JsonObject, declaration?: JsonObject, schemaDeclared = true): DeclaredFieldMetadata {
  const project = (value: JsonValue): JsonValue => {
    if (Array.isArray(value)) return value.map(project);
    if (!isJsonObject(value)) return value;
    const keys = ["@id", "@type", "displayName", "description", "comment", "unit", "valueSchema", "enumValues",
      "fields", "mapKey", "mapValue", "elementSchema", "schema", "name", "enumValue"];
    return Object.fromEntries(keys.filter(key => Object.prototype.hasOwnProperty.call(value, key)).map(key => [
      key, ["enumValues", "fields", "mapKey", "mapValue", "elementSchema", "schema"].includes(key) ? project(value[key]) : value[key],
    ]));
  };
  return { kind: "dtdl", schema: project(schema), annotations,
    ...(!schemaDeclared ? { schemaDeclared: false } : {}),
    ...(declaration ? { declaration: Object.fromEntries(["@id", "@type", "name", "displayName", "description", "comment", "unit", "writable"]
      .filter(key => Object.prototype.hasOwnProperty.call(declaration, key)).map(key => [key, declaration[key]])) } : {}) };
}

export function declaredMetadataDetails(metadata: DeclaredFieldMetadata): { name: string; value: JsonValue }[] {
  const values: JsonObject = metadata.kind === "wot" && isJsonObject(metadata.schema) ? metadata.schema : { schema: metadata.schema };
  const details = Object.entries(values).filter(([name]) => !["title", "description", "displayName"].includes(name))
    .map(([name, value]) => ({ name, value }));
  for (const [name, value] of Object.entries(metadata.annotations)) {
    if (["title", "description", "displayName"].includes(name)) continue;
    const previous = details.find(detail => detail.name === name);
    if (!previous) details.push({ name, value });
    else if (JSON.stringify(previous.value) !== JSON.stringify(value)) details.push({ name: `affordance.${name}`, value });
  }
  return details;
}

export function displayMetadataValue(value: JsonValue): string {
  const text = JSON.stringify(value);
  return text.length > 180 ? `${text.slice(0, 177)}...` : text;
}

export function columnSemanticAnnotation(model: { kind: string; id?: string; documentVersion?: string },
  name: string, metadata: DeclaredFieldMetadata): string {
  return COLUMN_SEMANTICS_PREFIX + JSON.stringify({ version: 1,
    model: { kind: model.kind, ...(model.id ? { id: model.id } : {}),
      ...(model.documentVersion ? { documentVersion: model.documentVersion } : {}) },
    field: { name, metadata } });
}

function validateSemanticInfo(facts: JsonObject): void {
  if (typeof facts.name !== "string" || !facts.name || facts.component !== undefined && typeof facts.component !== "string" ||
      facts.role !== undefined && (typeof facts.role !== "string" || !["telemetry", "twin property (last-known whole received value)"].includes(facts.role)) ||
      facts.truncated !== undefined && facts.truncated !== true) throw new Error("Malformed column semantic identity or role.");
  if ("type" in facts && "schema" in facts || "@type" in facts && "semanticTypes" in facts || "Declared type" in facts) {
    throw new Error("Ambiguous column semantic type aliases.");
  }
  const keys = new Set([...FACT_KEYS, "name", "displayName", "component", "role", "truncated", "data", "enumValue", "semanticTypes", "codes"]);
  const visit = (value: JsonObject, depth: number) => {
    if (depth > 48 || !isDocstringValue(value)) throw new Error("Malformed or excessively nested column semantic values.");
    if ("type" in value && "schema" in value || "@type" in value && "semanticTypes" in value || "Declared type" in value) {
      throw new Error("Ambiguous column semantic type aliases.");
    }
    for (const key of Object.keys(value)) if (!keys.has(key)) throw new Error(`Unsupported column semantic attribute: ${key}.`);
    for (const key of ["name", "unit", "@id", "comment"]) {
      if (value[key] !== undefined && typeof value[key] !== "string") throw new Error(`Malformed column semantic ${key}.`);
    }
    for (const key of ["title", "description", "displayName"]) {
      const item = value[key];
      if (item !== undefined && typeof item !== "string" &&
          !(isJsonObject(item) && Object.values(item).every(text => typeof text === "string"))) throw new Error(`Malformed column semantic ${key}.`);
    }
    for (const key of ["titles", "descriptions"]) {
      const item = value[key];
      if (item !== undefined && (!isJsonObject(item) || !Object.values(item).every(text => typeof text === "string"))) throw new Error(`Malformed column semantic ${key}.`);
    }
    for (const key of ["type", "@type", "semanticTypes"]) {
      const item = value[key];
      if (item !== undefined && typeof item !== "string" &&
          !(Array.isArray(item) && item.every(text => typeof text === "string"))) throw new Error(`Malformed column semantic ${key}.`);
    }
    for (const key of ["readOnly", "writeOnly", "observable", "writable"]) {
      if (value[key] !== undefined && typeof value[key] !== "boolean") throw new Error(`Malformed column semantic ${key}.`);
    }
    for (const key of ["minimum", "maximum", "exclusiveMinimum", "exclusiveMaximum", "multipleOf", "minLength", "maxLength", "minItems", "maxItems"]) {
      if (value[key] !== undefined && typeof value[key] !== "number") throw new Error(`Malformed column semantic ${key}.`);
    }
    const schema = (item: JsonValue, booleanSchema = false) => {
      if (isJsonObject(item)) visit(item, depth + 1);
      else if (typeof item !== "string" && !(booleanSchema && typeof item === "boolean")) throw new Error("Malformed declared column schema.");
    };
    for (const key of ["schema", "valueSchema", "elementSchema"]) if (value[key] !== undefined) schema(value[key]);
    for (const key of ["data", "mapKey", "mapValue"]) if (value[key] !== undefined) schema(value[key], true);
    if (Array.isArray(value.items)) {
      for (const item of value.items) {
        if (!isJsonObject(item)) throw new Error("Malformed declared column schema.");
        schema(item);
      }
    } else if (value.items !== undefined) schema(value.items, true);
    if (value.properties !== undefined) {
      if (!isJsonObject(value.properties)) throw new Error("Malformed declared property collection.");
      Object.values(value.properties).forEach(item => schema(item, true));
    }
    if (value.oneOf !== undefined) {
      if (!Array.isArray(value.oneOf)) throw new Error("Malformed declared schema alternatives.");
      value.oneOf.forEach(item => schema(item, true));
    }
    if (value.enum !== undefined && !Array.isArray(value.enum) ||
        value.required !== undefined && (!Array.isArray(value.required) || !value.required.every(item => typeof item === "string"))) {
      throw new Error("Malformed declared schema constraint.");
    }
    for (const key of ["fields", "enumValues"]) if (value[key] !== undefined) {
      const items = value[key];
      if (!Array.isArray(items)) throw new Error("Malformed declared schema member collection.");
      const names = new Set<string>();
      for (const item of items) {
        if (!isJsonObject(item) || typeof item.name !== "string" || names.has(item.name)) throw new Error("Malformed or duplicate declared member identity.");
        if (key === "fields" && !Object.prototype.hasOwnProperty.call(item, "schema") ||
            key === "enumValues" && typeof item.enumValue !== "string" && typeof item.enumValue !== "number") {
          throw new Error("Malformed declared schema member.");
        }
        names.add(item.name); visit(item, depth + 1);
      }
    }
  };
  visit(facts, 0);
}

export function readColumnSemantics(docstring: string, context?: ColumnSemanticsContext): LegacyColumnSemantics | FlatColumnSemantics | JsonColumnSemantics | undefined {
  const infoLines = docstring.split(/\r?\n/).filter(line => line.startsWith("@iot.meta fieldSemanticInfo"));
  if (infoLines.length) {
    if (infoLines.length !== 1 || docstring.split(/\r?\n/).filter(Boolean).length !== 1 ||
        !infoLines[0].startsWith(FIELD_SEMANTIC_INFO_PREFIX)) throw new Error("Malformed or ambiguous column semantic metadata formats.");
    if (docstring.length > DOCSTRING_LIMIT) throw new Error("Column semantic metadata exceeds the 1000-character stored limit.");
    const facts = parseDocstringJson(infoLines[0].slice(FIELD_SEMANTIC_INFO_PREFIX.length));
    validateSemanticInfo(facts);
    const bounded = facts.truncated === true;
    delete facts.truncated;
    return { version: 3, field: { name: String(facts.name), facts }, bounded, ...(context ? { model: context } : {}) };
  }
  const lines = docstring.split(/\r?\n/).filter(line => line.startsWith("@iot.meta columnSemantics"));
  const firstLine = docstring.split(/\r?\n/).find(line => line.length > 0);
  if (firstLine?.startsWith('name: "') && lines.some(line => line.startsWith(COLUMN_SEMANTICS_PREFIX))) {
    throw new Error("Ambiguous plain and legacy column semantic formats.");
  }
  const plainFacts = !lines.length && docstring.split(/\r?\n/).some(line => line.startsWith("name:"));
  if (!lines.length && !plainFacts) return undefined;
  if (plainFacts || lines.length === 1 && lines[0] === FLAT_COLUMN_SEMANTICS) {
    const facts: JsonObject = Object.create(null);
    const stack: { indent: number; value: JsonObject; collection: boolean }[] = [{ indent: -1, value: facts, collection: false }];
    for (const line of docstring.split(/\r?\n/)) {
      if (line === FLAT_COLUMN_SEMANTICS || !line) continue;
      const indent = /^\t*/.exec(line)![0].length, content = line.slice(indent);
      if (indent > 48) throw new Error("Column semantic hierarchy exceeds supported depth.");
      while (stack.length > 1 && stack[stack.length - 1].indent >= indent) stack.pop();
      const parent = stack[stack.length - 1];
      if (indent !== parent.indent + 1) throw new Error("Malformed column semantic indentation.");
      const block = /^(properties|fields|items):$/.exec(content);
      const child = /^("(?:\\.|[^"\\])*"):$/.exec(content);
      if (block || child) {
        let key = block?.[1];
        if (child) {
          if (!parent.collection) throw new Error("Unexpected column semantic child field.");
          const decoded: unknown = JSON.parse(child[1]);
          if (typeof decoded !== "string") throw new Error("Malformed column semantic child identity.");
          key = decoded;
        } else if (parent.collection) throw new Error("Missing quoted child field identity.");
        if (!key || Object.prototype.hasOwnProperty.call(parent.value, key)) throw new Error("Duplicate column semantic block.");
        const node: JsonObject = Object.create(null);
        Object.defineProperty(parent.value, key, { value: node, enumerable: true });
        stack.push({ indent, value: node, collection: !!block && block[1] !== "items" });
        continue;
      }
      const match = /^([^:\r\n]+): (.+)$/.exec(content);
      const key = match?.[1] === "Declared type" ? "type" : match?.[1];
      if (!match || !key || !["name", "description", "unit", "role", "semanticTypes", "codes", "truncated",
        ...FACT_KEYS].includes(key) || parent.collection || Object.prototype.hasOwnProperty.call(parent.value, key)) {
        throw new Error("Malformed or duplicate flat column semantic fact.");
      }
      let value: unknown;
      try { value = JSON.parse(match[2]); } catch { throw new Error("Malformed flat column semantic JSON value."); }
      if (!isDocstringValue(value)) throw new Error("Malformed flat column semantic value.");
      parent.value[key] = value;
    }
    if (typeof facts.name !== "string" || facts.truncated !== undefined && facts.truncated !== true ||
        facts.role !== undefined && !["telemetry", "twin property (last-known whole received value)"].includes(String(facts.role))) {
      throw new Error("Malformed flat column semantic identity or role.");
    }
    const bounded = facts.truncated === true;
    delete facts.truncated;
    const validateFacts = (values: JsonObject) => {
      for (const key of ["name", "description", "unit"]) {
        if (values[key] !== undefined && typeof values[key] !== "string") throw new Error(`Malformed column semantic ${key}.`);
      }
      for (const key of ["readOnly", "writeOnly", "observable", "writable"]) {
        if (values[key] !== undefined && typeof values[key] !== "boolean") throw new Error(`Malformed column semantic ${key}.`);
      }
      for (const key of ["minimum", "maximum", "exclusiveMinimum", "exclusiveMaximum", "multipleOf", "minLength", "maxLength", "minItems", "maxItems"]) {
        if (values[key] !== undefined && typeof values[key] !== "number") throw new Error(`Malformed column semantic ${key}.`);
      }
      if (values.type !== undefined && typeof values.type !== "string" &&
          !(Array.isArray(values.type) && values.type.every(item => typeof item === "string"))) {
        throw new Error("Malformed declared column semantic type.");
      }
      for (const key of ["properties", "fields"]) if (values[key] !== undefined) {
        if (!isJsonObject(values[key])) throw new Error("Malformed declared child field collection.");
        for (const child of Object.values(values[key])) {
          if (!isJsonObject(child)) throw new Error("Malformed declared child field metadata.");
          validateFacts(child);
        }
      }
      if (Array.isArray(values.items)) validateSemanticInfo({ name: facts.name, items: values.items });
      else if (isJsonObject(values.items)) validateFacts(values.items);
    };
    validateFacts(facts);
    return { version: 2, field: { name: facts.name, facts }, bounded, ...(context ? { model: context } : {}) };
  }
  if (lines.length !== 1 || !lines[0].startsWith(COLUMN_SEMANTICS_PREFIX)) {
    throw new Error("Malformed column semantic metadata annotation.");
  }
  let value: unknown;
  try { value = JSON.parse(lines[0].slice(COLUMN_SEMANTICS_PREFIX.length)); }
  catch { throw new Error("Malformed column semantic metadata JSON."); }
  if (!isJsonObject(value) || typeof value.version !== "number") throw new Error("Malformed column semantic metadata envelope.");
  if (value.version !== 1) throw new Error(`Unsupported column semantic metadata version: ${value.version}.`);
  if (!isJsonObject(value.model) || !["wot", "dtdl"].includes(String(value.model.kind)) ||
      !isJsonObject(value.field) || typeof value.field.name !== "string" ||
      !isJsonObject(value.field.metadata) || value.field.metadata.kind !== value.model.kind ||
      !Object.prototype.hasOwnProperty.call(value.field.metadata, "schema") ||
      !isJsonObject(value.field.metadata.annotations)) throw new Error("Malformed column semantic metadata envelope.");
  const kind = value.model.kind;
  if (kind !== "wot" && kind !== "dtdl" || value.model.id !== undefined && typeof value.model.id !== "string" ||
      value.model.documentVersion !== undefined && typeof value.model.documentVersion !== "string") {
    throw new Error("Malformed column semantic model identity.");
  }
  return { version: 1, model: { kind, ...(typeof value.model.id === "string" ? { id: value.model.id } : {}),
    ...(typeof value.model.documentVersion === "string" ? { documentVersion: value.model.documentVersion } : {}) },
    field: { name: value.field.name, metadata: { kind, schema: value.field.metadata.schema, annotations: value.field.metadata.annotations } } };
}
