// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { JsonObject } from "./ModelAuthoring.ts";
import { isDocstringValue } from "./Docstrings.ts";

/**
 * Parses a DTDL (Digital Twin Definition Language) model file and extracts
 * telemetry and property capabilities with their Kusto column types.
 */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface DtdlCapability {
  /** Column name in the modeled table. For component capabilities this is `<component>_<leafName>`. */
  name: string;
  /** Original capability name as it appears in the device payload (the JSON key). */
  leafName: string;
  /** Name of the component this capability belongs to, or undefined for root-interface capabilities. */
  component?: string;
  displayName: string;
  /** Engineering unit declared on the DTDL capability (e.g. "degreeCelsius"), if any. */
  unit?: string;
  /** Human-readable description declared on the DTDL capability, if any. */
  description?: string;
  kind: "Telemetry" | "Property";
  schema: string;
  kustoType: string;
  included: boolean;
  semanticTypes?: string[];
  enumValues?: { value: string | number; label?: string; description?: string }[];
  declaredSchema?: unknown;
  declaredBy?: string;
  declaredAnnotations?: JsonObject;
}

export interface DtdlParseResult {
  modelId: string;
  displayName: string;
  telemetries: DtdlCapability[];
  properties: DtdlCapability[];
  description?: string;
  declaredModelMetadata?: JsonObject;
}

export class UnsupportedDtdlModelFileError extends Error {
  constructor() {
    super("Unsupported model file. Provide a DTDL interface, an array of interfaces, or an IoT Central device template export.");
    this.name = "UnsupportedDtdlModelFileError";
  }
}

// ---------------------------------------------------------------------------
// Schema → Kusto type mapping
// ---------------------------------------------------------------------------

const SCHEMA_TO_KUSTO: Record<string, string> = {
  float: "real",
  double: "real",
  integer: "int",
  long: "long",
  string: "string",
  date: "datetime",
  dateTime: "datetime",
  time: "string",
  boolean: "bool",
  duration: "string",
  Object: "dynamic",
  object: "dynamic",
  Map: "dynamic",
  map: "dynamic",
  Array: "dynamic",
  array: "dynamic",
  Geopoint: "dynamic",
  geopoint: "dynamic",
  Vector: "dynamic",
  vector: "dynamic",
  point: "dynamic",
  multiPoint: "dynamic",
  lineString: "dynamic",
  multiLineString: "dynamic",
  polygon: "dynamic",
  multiPolygon: "dynamic",
};

/**
 * Normalize a primitive schema string to its bare type name. IoT Central / DTDL primitive schemas
 * are bare words (e.g. "double"), but some representations qualify them with a path or namespace
 * (e.g. "…/double"). Take the last path segment so a qualified-but-valid primitive still resolves
 * to a concrete type instead of falling back to "string".
 */
function normalizePrimitiveSchema(schema: string): string {
  const segments = schema.split("/");
  return segments[segments.length - 1] || schema;
}

function resolveKustoType(schema: unknown): string {
  if (typeof schema === "string") {
    return SCHEMA_TO_KUSTO[normalizePrimitiveSchema(schema)] || "string";
  }

  // Complex schema — check for Enum
  if (typeof schema === "object" && schema !== null) {
    const schemaObj = schema as Record<string, unknown>;
    const schemaType = schemaObj["@type"];
    if (hasType(schemaType, "Enum")) {
      const valueSchema = schemaObj["valueSchema"];
      if (typeof valueSchema === "string") {
        return SCHEMA_TO_KUSTO[normalizePrimitiveSchema(valueSchema)] || "string";
      }
      return "string";
    }
    // Other complex schemas (Object, Map, Array, Geopoint, Vector) → dynamic
    return "dynamic";
  }

  return "string";
}

function resolveSchemaName(schema: unknown): string {
  if (typeof schema === "string") return normalizePrimitiveSchema(schema);
  if (typeof schema === "object" && schema !== null) {
    const schemaObj = schema as Record<string, unknown>;
    const schemaType = schemaObj["@type"];
    if (hasType(schemaType, "Enum")) {
      const vs = schemaObj["valueSchema"];
      return `Enum(${typeof vs === "string" ? vs : "string"})`;
    }
    if (typeof schemaType === "string") return schemaType;
    if (Array.isArray(schemaType)) {
      const name = schemaType.find((t) => typeof t === "string");
      if (typeof name === "string") return name;
    }
  }
  return "unknown";
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getDisplayName(obj: Record<string, unknown>): string {
  const dn = obj.displayName;
  if (typeof dn === "string") return dn;
  if (typeof dn === "object" && dn !== null) {
    const dnObj = dn as Record<string, string>;
    return dnObj.en || Object.values(dnObj)[0] || "";
  }
  return "";
}

/** Read a DTDL `description` field, which may be a plain string or a localized map. */
function getDescription(obj: Record<string, unknown>): string {
  const desc = obj.description;
  if (typeof desc === "string") return desc;
  if (typeof desc === "object" && desc !== null) {
    const descObj = desc as Record<string, string>;
    return descObj.en || Object.values(descObj)[0] || "";
  }
  return "";
}

function hasType(typeField: unknown, typeName: string): boolean {
  if (typeof typeField === "string") return typeField === typeName;
  if (Array.isArray(typeField)) return typeField.includes(typeName);
  return false;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function asArray<T>(value: T | T[] | undefined | null): T[] {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

// ---------------------------------------------------------------------------
// Normalization & root resolution
// ---------------------------------------------------------------------------

/**
 * Normalize any supported DTDL input into a flat array of interface objects.
 * Accepts:
 *  - a JSON array of interfaces (IoT Central UI export)
 *  - a single DTDL Interface object (e.g. a plug-and-play model file)
 *  - an IoT Central REST device template: { "@type": ["ModelDefinition","DeviceModel"], "capabilityModel": {...} }
 */
function normalizeToInterfaces(modelJson: unknown): Record<string, unknown>[] {
  let source: unknown = modelJson;

  // Unwrap the IoT Central REST device-template envelope.
  if (isRecord(source) && isRecord(source.capabilityModel)) {
    source = source.capabilityModel;
  }

  const list: unknown[] = Array.isArray(source) ? source : [source];
  const interfaces = list.filter(
    (o): o is Record<string, unknown> =>
      isRecord(o) &&
      (hasType(o["@type"], "Interface") ||
        typeof o["@id"] === "string" ||
        Array.isArray(o["contents"]))
  );

  if (interfaces.length === 0) {
    throw new UnsupportedDtdlModelFileError();
  }
  return interfaces;
}

/**
 * Determine the root interface without relying on array position. The root is the
 * interface that is NOT referenced by any other interface via `extends` or a Component
 * `schema` DTMI. When several candidates remain, prefer a DeviceModel/CapabilityModel.
 */
function resolveRootInterface(
  interfaces: Record<string, unknown>[]
): Record<string, unknown> {
  if (interfaces.length === 1) return interfaces[0];

  const referenced = new Set<string>();
  for (const iface of interfaces) {
    for (const ext of asArray(iface["extends"])) {
      if (typeof ext === "string") referenced.add(ext);
      else if (isRecord(ext) && typeof ext["@id"] === "string") {
        referenced.add(ext["@id"] as string);
      }
    }
    for (const cap of asArray(iface["contents"])) {
      if (
        isRecord(cap) &&
        hasType(cap["@type"], "Component") &&
        typeof cap["schema"] === "string"
      ) {
        referenced.add(cap["schema"] as string);
      }
    }
  }

  const roots = interfaces.filter((i) => {
    const id = i["@id"];
    return typeof id !== "string" || !referenced.has(id);
  });

  const preferred = roots.find(
    (i) =>
      hasType(i["@type"], "DeviceModel") ||
      hasType(i["@type"], "CapabilityModel")
  );
  return preferred || roots[0] || interfaces[0];
}

// ---------------------------------------------------------------------------
// Main parser
// ---------------------------------------------------------------------------

export function parseDtdlModel(modelJson: unknown): DtdlParseResult {
  const interfaces = normalizeToInterfaces(modelJson);

  // Build a lookup of all interfaces by @id (used to resolve extends/component DTMIs).
  const ifaceMap = new Map<string, Record<string, unknown>>();
  for (const iface of interfaces) {
    const id = iface["@id"];
    if (typeof id === "string") {
      if (ifaceMap.has(id)) throw new Error(`Duplicate DTDL interface ID: ${id}.`);
      ifaceMap.set(id, iface);
    }
  }
  const schemas = new Map<string, Record<string, unknown>>();
  const collectSchemas = (value: unknown) => {
    if (Array.isArray(value)) value.forEach(collectSchemas);
    else if (isRecord(value)) {
      if (typeof value["@id"] === "string" && !hasType(value["@type"], "Interface")) {
        schemas.set(value["@id"], value);
      }
      Object.values(value).forEach(collectSchemas);
    }
  };
  interfaces.forEach(collectSchemas);
  const resolveSchema = (value: unknown, active = new Set<string>()): unknown => {
    if (typeof value === "string" && value.startsWith("dtmi:")) {
      if (active.has(value)) throw new Error(`Cyclic DTDL schema reference: ${value}.`);
      const schema = schemas.get(value);
      if (!schema) throw new Error(`Required DTDL schema was not supplied: ${value}.`);
      return resolveSchema(schema, new Set([...active, value]));
    }
    if (!isRecord(value)) return value;
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [
      key,
      ["schema", "elementSchema", "valueSchema"].includes(key) ? resolveSchema(item, active)
        : key === "fields" && Array.isArray(item) ? item.map(field => resolveSchema(field, active))
        : (key === "mapKey" || key === "mapValue") ? resolveSchema(item, active) : item,
    ]));
  };

  const root = resolveRootInterface(interfaces);
  const modelId = (root["@id"] as string) || "";
  const displayName = getDisplayName(root) || modelId;
  const declaredModelMetadata: JsonObject = {};
  for (const key of ["@id", "displayName", "description"]) {
    if (!Object.prototype.hasOwnProperty.call(root, key)) continue;
    if (!isDocstringValue(root[key])) throw new Error(`Invalid declared DTDL model ${key}.`);
    declaredModelMetadata[key] = root[key];
  }

  const telemetries: DtdlCapability[] = [];
  const properties: DtdlCapability[] = [];
  const seenNames = new Set<string>();
  const bindings = new Map<string, string>();

  // Resolve an extends/component reference (a DTMI string or an inline interface) to an interface object.
  function resolveInterfaceRef(
    ref: unknown
  ): Record<string, unknown> | undefined {
    if (typeof ref === "string") {
      const resolved = ifaceMap.get(ref);
      if (!resolved) throw new Error(`Required DTDL interface was not supplied: ${ref}.`);
      return resolved;
    }
    if (isRecord(ref)) return ref;
    throw new Error("A DTDL inheritance/component reference is invalid.");
  }

  /**
   * Walk an interface's contents and `extends` chain, collecting Telemetry/Property capabilities.
   * When `componentName` is set, the capability belongs to a component: its column name is
   * prefixed as `<componentName>_<leafName>`, matching how component data arrives (telemetry with
   * cloudEvents_iothubdtsubject=<componentName>, reported properties nested under the component key with __t="c").
   *
   * `visited` guards against `extends` cycles within a single interface subtree. A fresh set is used
   * for each component subtree so that the same component interface can be reused under different
   * component names.
   */
  function extractCapabilities(
    iface: Record<string, unknown> | undefined,
    componentName: string | undefined,
    visited: Set<Record<string, unknown>>,
    active = new Set<Record<string, unknown>>()
  ) {
    if (!iface) throw new Error("A required DTDL interface is missing.");
    if (active.has(iface)) throw new Error("Cyclic DTDL inheritance or component reference.");
    if (visited.has(iface)) return;
    active.add(iface);
    visited.add(iface);

    const contents = iface["contents"];
    if (Array.isArray(contents)) {
      for (const cap of contents as unknown[]) {
        if (!isRecord(cap)) continue;
        const capType = cap["@type"];

        // Recurse into components so their nested telemetry/properties are captured, prefixing
        // their column names with the component name. A component's schema is either an inline
        // Interface object or a DTMI reference. DTDL forbids nested components; if we are already
        // inside one, keep the outer component name as the prefix.
        if (hasType(capType, "Component")) {
          if (componentName) throw new Error("Nested DTDL components are unsupported.");
          const compName = (cap["name"] as string) || componentName;
          extractCapabilities(
            resolveInterfaceRef(cap["schema"]),
            componentName || compName,
            new Set<Record<string, unknown>>(),
            active
          );
          continue;
        }

        const leafName = cap["name"] as string;
        if (!leafName) continue;

        const isTelemetry = hasType(capType, "Telemetry");
        const isProperty = hasType(capType, "Property");
        if (!isTelemetry && !isProperty) continue;

        // Skip cloud properties (@type includes "Cloud"): these are set in the cloud
        // and are not reported by the device, so they won't appear in device data.
        if (isProperty && hasType(capType, "Cloud")) continue;

        const sourceName = componentName ? `${componentName}_${leafName}` : leafName;
        const columnName = sourceName === "event_time" ? "event_time_value" : sourceName;
        const binding = JSON.stringify([componentName || "", leafName]);
        if (bindings.has(columnName) && bindings.get(columnName) !== binding) {
          throw new Error(`Distinct DTDL paths collide on output name "${columnName}". Rename the conflicting capability.`);
        }
        bindings.set(columnName, binding);
        if (seenNames.has(columnName)) continue;
        seenNames.add(columnName);
        const schema = resolveSchema(cap["schema"]);
        if ((typeof schema !== "string" && !isRecord(schema)) ||
            (typeof schema === "string" && !Object.prototype.hasOwnProperty.call(SCHEMA_TO_KUSTO, normalizePrimitiveSchema(schema)))) {
          throw new Error(`Unsupported or missing DTDL schema on ${leafName}.`);
        }
        if (isRecord(schema) && hasType(schema["@type"], "Enum") && schema.valueSchema !== "string" && schema.valueSchema !== "integer") {
          throw new Error(`Unsupported DTDL enum value schema on ${leafName}.`);
        }
        const enumValues = isRecord(schema) && hasType(schema["@type"], "Enum") && Array.isArray(schema.enumValues)
          ? schema.enumValues.map(value => {
            if (!isRecord(value) || (typeof value.enumValue !== "string" && typeof value.enumValue !== "number") ||
                (typeof value.enumValue === "number" && !Number.isFinite(value.enumValue))) {
              throw new Error(`Invalid DTDL enum value on ${leafName}.`);
            }
            return { value: value.enumValue, label: getDisplayName(value) || (typeof value.name === "string" ? value.name : undefined),
              description: getDescription(value) || undefined };
          }) : undefined;

        const declaredAnnotations: JsonObject = {};
        for (const key of ["@id", "@type", "name", "displayName", "description", "comment", "unit", "writable"]) {
          if (!Object.prototype.hasOwnProperty.call(cap, key)) continue;
          const value = cap[key];
          if (!isDocstringValue(value)) throw new Error(`Invalid declared DTDL ${key} on ${leafName}.`);
          declaredAnnotations[key] = value;
        }
        const capability: DtdlCapability = {
          name: columnName,
          leafName,
          component: componentName,
          displayName: getDisplayName(cap) || leafName,
          unit: typeof cap["unit"] === "string" ? (cap["unit"] as string) : undefined,
          description: getDescription(cap) || undefined,
          kind: isTelemetry ? "Telemetry" : "Property",
          schema: resolveSchemaName(schema),
          kustoType: resolveKustoType(schema),
          included: true,
          declaredSchema: schema,
          declaredBy: typeof iface["@id"] === "string" ? iface["@id"] : undefined,
          declaredAnnotations,
          semanticTypes: asArray(capType).filter((type): type is string =>
            typeof type === "string" && !["Telemetry", "Property", "Cloud"].includes(type)),
          enumValues,
        };
        (isTelemetry ? telemetries : properties).push(capability);
      }
    }

    // Follow inherited interfaces (extends): DTMI strings or inline interface objects.
    // Inherited capabilities belong to the same (component or root) context.
    for (const ext of asArray(iface["extends"])) {
      extractCapabilities(resolveInterfaceRef(ext), componentName, visited, active);
    }
    active.delete(iface);
  }

  extractCapabilities(root, undefined, new Set<Record<string, unknown>>());

  return { modelId, displayName, telemetries, properties, description: getDescription(root) || undefined, declaredModelMetadata };
}
