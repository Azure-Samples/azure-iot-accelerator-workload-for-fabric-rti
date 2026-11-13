// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Ajv from "ajv";
import addFormats from "ajv-formats";
import tmSchema from "./schemas/wot-tm.schema.json" with { type: "json" };
import { isJsonObject, type JsonObject, type JsonValue } from "./ModelAuthoring.ts";

export interface ThingModelField {
  name: string;
  kind: "Property" | "Event data";
  title: string;
  declaredType?: string;
}
export interface IgnoredThingModelPart {
  kind: "Property" | "Event data" | "Submodel";
  name: string;
  reason: "null type" | "oneOf" | "submodel";
}

export interface ThingModelResult {
  title: string;
  id?: string;
  documentVersion?: string;
  fields: ThingModelField[];
  blockers: string[];
  description?: string;
  document: JsonValue;
  effectiveDocument?: JsonObject;
  ignored?: IgnoredThingModelPart[];
}

export interface ThingModelResolutionStats {
  workUnits: number;
  expandedNodes: number;
  expandedBytes: number;
  cacheHits: number;
}

export const THING_MODEL_RESOLUTION_LIMITS = Object.freeze({
  workUnits: 500_000,
  expandedNodes: 100_000,
  expandedBytes: 2 * 1024 * 1024,
});

// Unmodified W3C WoT 1.1 REC1.1 schema; provenance/license are beside the schema.
// It contains the later prefixItems keyword in a draft-07 schema, requiring non-strict compilation.
const ajv = new Ajv({ allErrors: true, strict: false });
addFormats(ajv);
const validateStructure = ajv.compile(tmSchema);
export const TM_NAMESPACE = "https://www.w3.org/2022/wot/tm#";
const isSubmodel = (value: unknown) => isRecord(value) &&
  (value.rel === "tm:submodel" || value.rel === `${TM_NAMESPACE}submodel`);
function ignoredSchemaReason(value: unknown): "null type" | "oneOf" | undefined {
  if (!isRecord(value)) return undefined;
  if (value.type === "null" || Object.prototype.hasOwnProperty.call(value, "type") && value.type === null) return "null type";
  return Array.isArray(value.oneOf) ? "oneOf" : undefined;
}
function withoutSubmodels(document: JsonObject, ignored: IgnoredThingModelPart[]): JsonObject {
  const result = { ...document };
  if (Array.isArray(result.links)) result.links = result.links.filter((link, index) => {
    if (!isSubmodel(link)) return true;
    const name = isJsonObject(link) && typeof link.instanceName === "string" && link.instanceName.trim() &&
      !/^[a-z][a-z\d+.-]*:/i.test(link.instanceName) ? link.instanceName : `Submodel ${index + 1}`;
    ignored.push({ kind: "Submodel", name, reason: "submodel" });
    return false;
  });
  for (const key of ["tm:submodel", `${TM_NAMESPACE}submodel`]) {
    if (!Object.prototype.hasOwnProperty.call(result, key)) continue;
    ignored.push({ kind: "Submodel", name: "tm:submodel", reason: "submodel" });
    delete result[key];
  }
  return result;
}
function modelableProjection(document: JsonObject, ignored: IgnoredThingModelPart[]): JsonObject {
  const result = withoutSubmodels(document, ignored);
  for (const [section, kind] of [["properties", "Property"], ["events", "Event data"]] as const) {
    const entries = result[section];
    if (!isJsonObject(entries)) continue;
    result[section] = Object.fromEntries(Object.entries(entries).filter(([name, value]) => {
      const schema = section === "events" && isJsonObject(value) ? value.data : value;
      const reason = ignoredSchemaReason(schema);
      if (!reason) return true;
      ignored.push({ kind, name, reason });
      return false;
    }));
  }
  return result;
}

export function requireThingModelId(document: unknown): string {
  const root = Array.isArray(document) ? selectThingModelRoot(document) : document;
  const id = isRecord(root) ? root.id : undefined;
  if (typeof id !== "string" || !id.trim()) {
    throw new Error("This Thing Model doesn't have an ID. Add an ID, like urn:example:thermostat, then upload the file again.");
  }
  try {
    if (/[\u0000-\u0020\u007f]/.test(id) || /\{\{.*\}\}/.test(id)) throw new Error("Invalid model ID.");
    new URL(id);
  } catch {
    throw new Error("The model ID must be an absolute IRI, like urn:example:thermostat. Fix the ID, then upload the file again.");
  }
  return id;
}

/** A JSON array is upload packaging; each member remains a separate normative TM. */
export function selectThingModelRoot(document: unknown): JsonObject {
  if (isJsonObject(document)) {
    if ("@graph" in document) throw new Error("JSON-LD @graph bundles aren't supported. Use one Thing Model or a JSON array of Thing Models.");
    return document;
  }
  if (!Array.isArray(document) || !document.length || !document.every(isJsonObject)) {
    throw new Error("Upload one WoT Thing Model or a nonempty JSON array of Thing Models.");
  }
  const members = document as JsonObject[];
  const ids = new Map<string, JsonObject>();
  for (const member of members) {
    if ("@graph" in member) throw new Error("JSON-LD @graph bundles aren't supported. Use one Thing Model or a JSON array of Thing Models.");
    const types = Array.isArray(member["@type"]) ? member["@type"] : [member["@type"]];
    if (!types.includes("tm:ThingModel") && !types.includes(`${TM_NAMESPACE}ThingModel`)) {
      throw new Error("Every bundle member must declare its own WoT Thing Model type and absolute id.");
    }
    const id = requireThingModelId(member);
    if (ids.has(id)) throw new Error(`Duplicate model ID in uploaded documents: ${id}.`);
    ids.set(id, member);
  }
  const referenced = new Set<JsonObject>();
  const dependencies = new Map<JsonObject, Set<JsonObject>>();
  const visit = (value: JsonValue, from: JsonObject, depth: number) => {
    if (depth > 48) throw new Error("The Thing Model exceeds the supported nesting depth (48).");
    if (Array.isArray(value)) { value.forEach(item => visit(item, from, depth + 1)); return; }
    if (!isJsonObject(value)) return;
    const reference = typeof value["tm:ref"] === "string" ? value["tm:ref"]
      : (value.rel === "tm:extends" || isSubmodel(value)) && typeof value.href === "string" ? value.href : undefined;
    if (reference) {
      const target = ids.get(reference) || ids.get(reference.split("#")[0]);
      if (target && target !== from) {
        referenced.add(target);
        const targets = dependencies.get(from) || new Set<JsonObject>();
        targets.add(target); dependencies.set(from, targets);
      }
    }
    Object.values(value).forEach(item => visit(item, from, depth + 1));
  };
  members.forEach(member => visit(member, member, 0));
  const roots = members.filter(member => !referenced.has(member));
  if (roots.length !== 1) throw new Error(roots.length
    ? "The Thing Model bundle has multiple roots. Upload one model with its required definitions, not unrelated root models."
    : "The Thing Model bundle has no unambiguous root. Check for cyclic model references.");
  const reachable = new Set(roots), pending = [...roots];
  for (let index = 0; index < pending.length; index++) for (const dependency of dependencies.get(pending[index]) || []) {
    if (!reachable.has(dependency)) { reachable.add(dependency); pending.push(dependency); }
  }
  if (reachable.size !== members.length) throw new Error("The Thing Model bundle contains unrelated or cyclic definitions. Include only the root and its required models.");
  return roots[0];
}

/** RFC 7396: null removes a member; arrays and nonobjects replace rather than merge. */
export function mergeModelPatch(target: JsonValue, patch: JsonValue): JsonValue {
  if (!isJsonObject(patch)) return patch;
  const result: JsonObject = Object.fromEntries(Object.entries(isJsonObject(target) ? target : {}));
  for (const [key, value] of Object.entries(patch)) {
    if (value === null) delete result[key];
    else Object.defineProperty(result, key, { value: mergeModelPatch(result[key], value), enumerable: true, writable: true, configurable: true });
  }
  return result;
}

function inheritModelDefinitions(base: JsonObject, definitions: JsonObject, visit: () => void): JsonObject {
  visit();
  return Object.fromEntries([...new Set([...Object.keys(base), ...Object.keys(definitions)])].map(key => {
    visit();
    if (!Object.prototype.hasOwnProperty.call(definitions, key)) return [key, base[key]];
    const inherited = base[key], declared = definitions[key];
    return [key, isJsonObject(inherited) && isJsonObject(declared)
      ? inheritModelDefinitions(inherited, declared, visit) : declared];
  }));
}

function importValidationShape(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(importValidationShape);
  if (!isRecord(value)) return value;
  // Imported definitions are merge patches, not standalone DataSchemas. Validate their result after resolution.
  if (Object.prototype.hasOwnProperty.call(value, "tm:ref")) return { "tm:ref": value["tm:ref"] };
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, importValidationShape(item)]));
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function localized(value: Record<string, unknown>, singular: string, plural: string): string | undefined {
  const translations = value[plural];
  if (isRecord(translations)) {
    for (const language of ["en-US", "en"]) {
      if (typeof translations[language] === "string") return translations[language];
    }
  }
  if (typeof value[singular] === "string") return value[singular];
  return undefined;
}

function prefixIssues(document: Record<string, unknown>): string[] {
  const issues: string[] = [];
  for (const context of Array.isArray(document["@context"]) ? document["@context"] : []) {
    if (!isRecord(context) || !("tm" in context)) continue;
    const prefix = context.tm;
    const validProtection = !isRecord(prefix) || !Object.prototype.hasOwnProperty.call(prefix, "@protected") ||
      typeof prefix["@protected"] === "boolean";
    const equivalent = prefix === TM_NAMESPACE || (isRecord(prefix) &&
      prefix["@id"] === TM_NAMESPACE && prefix["@prefix"] === true &&
      !("@reverse" in prefix) && !("@context" in prefix) && validProtection);
    if (!equivalent) issues.push(validProtection
      ? `The tm prefix must use the WoT 1.1 namespace ${TM_NAMESPACE}. An expanded prefix definition needs @prefix: true.`
      : "The tm prefix definition's @protected must be true or false, or omitted.");
  }
  return issues;
}

/** This experience requires a root ID; validation never fetches it or assigns Hub roles. */
export function parseThingModel(document: unknown): ThingModelResult {
  if (!isRecord(document)) throw new Error("Upload a WoT Thing Model JSON object.");
  const types = Array.isArray(document["@type"]) ? document["@type"] : [document["@type"]];
  const absoluteType = `${TM_NAMESPACE}ThingModel`;
  if (!types.includes("tm:ThingModel") && !types.includes(absoluteType)) {
    throw new Error(
      'This is not a WoT Thing Model. Upload a document declaring @type "tm:ThingModel"; a Thing Description or JSON Schema is not a Thing Model.'
    );
  }
  const id = requireThingModelId(document);
  const original = JSON.parse(JSON.stringify(document)) as JsonObject;
  const ignored: IgnoredThingModelPart[] = [];
  const projected = modelableProjection(original, ignored);
  const blockers = new Set<string>();
  const visit = (value: unknown, depth: number) => {
    if (depth > 48) throw new Error("The Thing Model exceeds the supported nesting depth (48).");
    if (typeof value === "string" && /\{\{.*\}\}/.test(value)) {
      blockers.add("Template placeholders need resolution before effective fields can be reviewed.");
    }
    if (Array.isArray(value)) {
      value.forEach(item => visit(item, depth + 1));
    } else if (isRecord(value)) {
      if (Object.prototype.hasOwnProperty.call(value, "tm:ref")) {
        blockers.add("tm:ref imports must be supplied in the uploaded model bundle for local resolution.");
      }
      if (Object.prototype.hasOwnProperty.call(value, "$ref")) {
        blockers.add("JSON Schema $ref dependencies are unsupported by local Thing Model validation.");
      }
      if (value.rel === "tm:extends") {
        blockers.add(`${value.rel} dependencies must be supplied in the uploaded model bundle for local resolution.`);
      }
      // These are not the WoT 1.1 link-based inheritance syntax. Do not silently ignore them.
      if (Object.prototype.hasOwnProperty.call(value, "tm:extends")) {
        blockers.add("Declare WoT inheritance/composition through links; this dependency syntax is unsupported.");
      }
      Object.values(value).forEach(item => visit(item, depth + 1));
    }
  };
  visit(projected, 0);
  const validationDocument = {
    ...importValidationShape(projected) as Record<string, unknown>,
    "@type": types.map(type => type === absoluteType ? "tm:ThingModel" : type),
  };
  if (!validateStructure(validationDocument)) {
    const details = (validateStructure.errors || [])
      .map(error => `${error.instancePath || "/"} ${error.message}`).join("; ");
    throw new Error(`This Thing Model doesn't follow the WoT 1.1 format: ${details}`);
  }
  prefixIssues(document).forEach(issue => blockers.add(issue));

  const fields: ThingModelField[] = [];
  if (blockers.size === 0) {
    for (const [name, value] of Object.entries(isRecord(projected.properties) ? projected.properties : {})) {
      if (!isRecord(value)) continue;
      fields.push({
        name, kind: "Property", title: localized(value, "title", "titles") || name,
        ...(typeof value.type === "string" ? { declaredType: value.type } : {}),
      });
    }
    for (const [name, value] of Object.entries(isRecord(projected.events) ? projected.events : {})) {
      if (!isRecord(value)) continue;
      const data = isRecord(value.data) ? value.data : {};
      fields.push({
        name, kind: "Event data", title: localized(value, "title", "titles") || name,
        ...(typeof data.type === "string" ? { declaredType: data.type } : {}),
      });
    }
  }
  return {
    title: localized(document, "title", "titles") || id || "Untitled Thing Model",
    ...(id ? { id } : {}),
    ...(isRecord(document.version) && typeof document.version.model === "string"
      ? { documentVersion: document.version.model } : {}),
    fields,
    ignored,
    blockers: [...blockers],
    ...(localized(document, "description", "descriptions") ? { description: localized(document, "description", "descriptions") } : {}),
    document: original,
    ...(blockers.size === 0 ? { effectiveDocument: projected } : {}),
  };
}

export interface UploadedThingModel {
  fileName: string;
  document: JsonObject;
}

/** Resolve only explicitly supplied local documents. There is deliberately no network resolver. */
export function resolveThingModel(
  document: JsonValue,
  dependencies: UploadedThingModel[] = [],
  fileName = "model.tm.json",
  reportResolution?: (stats: Readonly<ThingModelResolutionStats>) => void
): ThingModelResult {
  const root = selectThingModelRoot(document);
  const original = { ...parseThingModel(root), document: JSON.parse(JSON.stringify(document)) as JsonValue };
  const bundled: UploadedThingModel[] = Array.isArray(document) ? document.flatMap((member, index) =>
    isJsonObject(member) && member !== root ? [{ fileName: `${fileName} [${index + 1}]`, document: member }] : []) : [];
  const entries = [{ fileName, document: root }, ...bundled, ...dependencies];
  const ids = new Map<string, UploadedThingModel>();
  const files = new Map<string, UploadedThingModel>();
  for (const entry of entries) {
    parseThingModel(entry.document);
    const issues = prefixIssues(entry.document);
    if (issues.length) return { ...original, fields: [], effectiveDocument: undefined, blockers: issues };
    if (files.has(entry.fileName)) throw new Error(`Duplicate uploaded filename: ${entry.fileName}.`);
    files.set(entry.fileName, entry);
    const id = entry.document.id;
    if (typeof id === "string" && id) {
      if (ids.has(id)) throw new Error(`Duplicate model ID in uploaded documents: ${id}.`);
      ids.set(id, entry);
    }
  }
  const active = new Set<JsonObject>();
  const resolution = { workUnits: 0, expandedNodes: 0, expandedBytes: 0, cacheHits: 0 };
  const visit = () => {
    if (resolution.workUnits === THING_MODEL_RESOLUTION_LIMITS.workUnits) {
      throw new Error("Thing Model dependency resolution exceeds the supported work limit. Simplify the model or reduce repeated imports.");
    }
    resolution.workUnits++;
  };
  const frames: { depth: number; maximum: number }[] = [];
  const checkDepth = (depth: number) => {
    if (depth > 48) throw new Error("Model dependency resolution exceeds the supported nesting depth (48).");
    const frame = frames[frames.length - 1];
    if (frame) frame.maximum = Math.max(frame.maximum, depth);
  };
  const sizes = new WeakMap<object, { nodes: number; bytes: number }>();
  const stringSizes = new Map<string, number>();
  const encoder = new TextEncoder();
  const literalBytes = (value: JsonValue) => {
    const cached = typeof value === "string" ? stringSizes.get(value) : undefined;
    if (cached !== undefined) return cached;
    const bytes = encoder.encode(JSON.stringify(value)).length;
    if (typeof value === "string") stringSizes.set(value, bytes);
    return bytes;
  };
  const sizeLimit = (nodes: number, bytes: number) => {
    if (nodes > THING_MODEL_RESOLUTION_LIMITS.expandedNodes) {
      throw new Error(`The expanded Thing Model exceeds the supported node limit (${THING_MODEL_RESOLUTION_LIMITS.expandedNodes}). Reduce repeated imports or split the model.`);
    }
    if (bytes > THING_MODEL_RESOLUTION_LIMITS.expandedBytes) {
      throw new Error("The expanded Thing Model exceeds the supported size (2 MB). Reduce repeated imports or split the model.");
    }
    resolution.expandedNodes = Math.max(resolution.expandedNodes, nodes);
    resolution.expandedBytes = Math.max(resolution.expandedBytes, bytes);
  };
  // Count every occurrence in serialized output, including shared cached subtrees, without expanding them again.
  const measure = (value: JsonValue): { nodes: number; bytes: number } => {
    const object = value !== null && typeof value === "object";
    const cached = object ? sizes.get(value) : undefined;
    if (cached) return cached;
    visit();
    let nodes = 1, bytes = object ? 2 : literalBytes(value), count = 0;
    sizeLimit(nodes, bytes);
    if (object) for (const [key, child] of Object.entries(value)) {
      visit();
      const size = measure(child);
      nodes += size.nodes;
      bytes += size.bytes + (count++ ? 1 : 0) + (Array.isArray(value) ? 0 : literalBytes(key) + 1);
      sizeLimit(nodes, bytes);
    }
    const size = { nodes, bytes };
    if (object) sizes.set(value, size);
    return size;
  };
  const checked = <T extends JsonValue>(value: T): T => { measure(value); return value; };
  type CachedResolution = { value: JsonObject; depth: number };
  const caches = new Map<UploadedThingModel, [WeakMap<JsonObject, CachedResolution>, WeakMap<JsonObject, CachedResolution>]>();
  const candidateSchemas = new Set<JsonValue>();
  for (const entry of entries) {
    if (isJsonObject(entry.document.properties)) Object.values(entry.document.properties).forEach(schema => candidateSchemas.add(schema));
    if (isJsonObject(entry.document.events)) for (const event of Object.values(entry.document.events)) {
      if (isJsonObject(event) && isJsonObject(event.data)) candidateSchemas.add(event.data);
    }
  }
  const ignoredSubmodels: IgnoredThingModelPart[] = [];
  const locate = (reference: string, from: UploadedThingModel, kind: "tm:ref" | "tm:extends"): { entry: UploadedThingModel; value: JsonObject } => {
    visit();
    const exact = ids.get(reference);
    if (exact) return { entry: exact, value: exact.document };
    const hash = reference.indexOf("#");
    const location = hash < 0 ? reference : reference.slice(0, hash);
    const fragment = hash < 0 ? "" : reference.slice(hash + 1);
    let entry: UploadedThingModel | undefined = location ? ids.get(location) : from;
    if (!entry && !/^[a-z][a-z\d+.-]*:/i.test(location)) {
      const local = new URL(location, `file:///uploaded/${encodeURIComponent(from.fileName)}`);
      if (local.search) throw new Error(`A query-bearing model reference needs an exact document binding: ${location}.`);
      if (local.pathname.startsWith("/uploaded/")) {
        entry = files.get(decodeURIComponent(local.pathname.slice("/uploaded/".length)));
      }
    }
    if (!entry) throw new Error((kind === "tm:ref"
      ? "This model imports definitions with tm:ref that aren't included in this file. External model resolution isn't supported yet. Include the definitions in this file instead."
      : "This model extends another model with tm:extends that isn't included in this file. External model resolution isn't supported yet. Include those definitions in this file instead.") +
      ` Required model file missing: ${location}.`);
    let value: JsonValue = entry.document;
    if (fragment) {
      const pointer = decodeURIComponent(fragment);
      if (!pointer.startsWith("/") || /~(?![01])/.test(pointer)) throw new Error(`Invalid model fragment: ${reference}.`);
      for (const segment of pointer.slice(1).split("/").map(part => part.replace(/~1/g, "/").replace(/~0/g, "~"))) {
        visit();
        if (!isJsonObject(value) && !Array.isArray(value)) throw new Error(`Required model fragment was not found: ${reference}.`);
        if (Array.isArray(value)) {
          if (!/^(0|[1-9]\d*)$/.test(segment) || Number(segment) >= value.length) {
            throw new Error(`Required model fragment was not found: ${reference}.`);
          }
          value = value[Number(segment)];
        } else {
          if (!Object.prototype.hasOwnProperty.call(value, segment)) throw new Error(`Required model fragment was not found: ${reference}.`);
          value = value[segment];
        }
      }
    }
    if (!isJsonObject(value)) throw new Error(`Model reference must resolve to an object: ${reference}.`);
    return { entry, value };
  };
  const applyImportPatch = (target: JsonValue, patch: JsonValue, entry: UploadedThingModel, depth: number): JsonValue => {
    checkDepth(depth);
    visit();
    if (!isJsonObject(patch)) return resolve(patch, entry, depth + 1);
    // A nested import is its own boundary; its resolved null facts are not a second removal patch.
    if (typeof patch["tm:ref"] === "string") return resolve(patch, entry, depth + 1);
    const result: JsonObject = Object.fromEntries(Object.entries(isJsonObject(target) ? target : {}));
    for (const [key, item] of Object.entries(patch)) {
      visit();
      if (item === null) delete result[key];
      else Object.defineProperty(result, key, {
        value: applyImportPatch(result[key], item, entry, depth + 1),
        enumerable: true, writable: true, configurable: true,
      });
    }
    return checked(result);
  };
  const resolve = (value: JsonValue, entry: UploadedThingModel, depth: number, candidate = candidateSchemas.has(value)): JsonValue => {
    checkDepth(depth);
    visit();
    if (Array.isArray(value)) return checked(value.map(item => resolve(item, entry, depth + 1)));
    if (!isJsonObject(value)) return checked(value);
    // Keep an unsupported candidate intact for final diagnostics; do not resolve its unused branches.
    if (candidate && ignoredSchemaReason(value)) {
      measure(value);
      return checked(JSON.parse(JSON.stringify(value)) as JsonObject);
    }
    if (active.has(value)) throw new Error("Cyclic Thing Model inheritance or import.");
    let modes = caches.get(entry);
    if (!modes) { modes = [new WeakMap(), new WeakMap()]; caches.set(entry, modes); }
    const cache = modes[candidate ? 1 : 0], cached = cache.get(value);
    if (cached) {
      checkDepth(depth + cached.depth);
      resolution.cacheHits++;
      return cached.value;
    }
    const frame = { depth, maximum: depth };
    frames.push(frame);
    active.add(value);
    try {
      let inherited: JsonObject = {};
      if (typeof value["tm:ref"] === "string") {
        const target = locate(value["tm:ref"], entry, "tm:ref");
        const resolved = resolve(target.value, target.entry, depth + 1, candidate);
        if (!isJsonObject(resolved)) throw new Error("Imported definition is not an object.");
        inherited = resolved;
      }
      const links = Array.isArray(value.links) ? value.links : [];
      const extensions = links.filter(link => isJsonObject(link) && link.rel === "tm:extends");
      if (extensions.length > 1) throw new Error("This model extends more than one parent, which is not supported here. Upload a single model that combines those definitions.");
      for (const link of extensions) {
        if (!isJsonObject(link) || typeof link.href !== "string") throw new Error("A tm:extends link requires an href.");
        const target = locate(link.href, entry, "tm:extends");
        const resolved = resolve(target.value, target.entry, depth + 1);
        if (!isJsonObject(resolved)) throw new Error("Inherited Thing Model is not an object.");
        inherited = Object.fromEntries(Object.entries(resolved).filter(([key]) => key !== "id" && key !== "version"));
      }
      const declared = Object.fromEntries(Object.entries(value).filter(([key]) => key !== "tm:ref"));
      const own = value === entry.document ? withoutSubmodels(declared, ignoredSubmodels) : declared;
      if (Array.isArray(own.links)) own.links = own.links.filter(link => !isJsonObject(link) || link.rel !== "tm:extends");
      let result: JsonValue;
      if (typeof value["tm:ref"] === "string") result = applyImportPatch(inherited, own, entry, depth + 1);
      else {
        const resolvedOwn = Object.fromEntries(Object.entries(own).map(([key, item]) => {
          visit();
          return [key, resolve(item, entry, depth + 1)];
        }));
        result = extensions.length ? inheritModelDefinitions(inherited, resolvedOwn, visit) : resolvedOwn;
      }
      if (!isJsonObject(result)) throw new Error("Resolved model definition is not an object.");
      checked(result);
      cache.set(value, { value: result, depth: frame.maximum - frame.depth });
      return result;
    } finally {
      active.delete(value);
      frames.pop();
      const parent = frames[frames.length - 1];
      if (parent) parent.maximum = Math.max(parent.maximum, frame.maximum);
    }
  };
  try {
    const effective = resolve(root, entries[0], 0);
    if (!isJsonObject(effective)) throw new Error("The effective Thing Model is not an object.");
    const checked = parseThingModel(effective);
    return { ...checked, id: original.id, documentVersion: original.documentVersion,
      document: original.document,
      ignored: [...new Map([...(checked.ignored || []), ...ignoredSubmodels]
        .map(part => [JSON.stringify(part), part])).values()] };
  } catch (error: unknown) {
    return { ...original, fields: [], effectiveDocument: undefined,
      blockers: [error instanceof Error ? error.message : "Unable to resolve the local model dependencies."] };
  } finally {
    reportResolution?.({ ...resolution });
  }
}
