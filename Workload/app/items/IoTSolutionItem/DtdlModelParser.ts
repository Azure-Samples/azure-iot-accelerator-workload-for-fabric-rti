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
}

export interface DtdlParseResult {
  modelId: string;
  displayName: string;
  telemetries: DtdlCapability[];
  properties: DtdlCapability[];
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
  date: "date",
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
    throw new Error(
      "Unsupported model file. Provide a DTDL interface, an array of interfaces, or an IoT Central device template export."
    );
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
    if (typeof id === "string") ifaceMap.set(id, iface);
  }

  const root = resolveRootInterface(interfaces);
  const modelId = (root["@id"] as string) || "";
  const displayName = getDisplayName(root) || modelId;

  const telemetries: DtdlCapability[] = [];
  const properties: DtdlCapability[] = [];
  const seenNames = new Set<string>();

  // Resolve an extends/component reference (a DTMI string or an inline interface) to an interface object.
  function resolveInterfaceRef(
    ref: unknown
  ): Record<string, unknown> | undefined {
    if (typeof ref === "string") return ifaceMap.get(ref);
    if (isRecord(ref)) return ref;
    return undefined;
  }

  /**
   * Walk an interface's contents and `extends` chain, collecting Telemetry/Property capabilities.
   * When `componentName` is set, the capability belongs to a component: its column name is
   * prefixed as `<componentName>_<leafName>`, matching how component data arrives (telemetry with
   * IoTSubject=<componentName>, reported properties nested under the component key with __t="c").
   *
   * `visited` guards against `extends` cycles within a single interface subtree. A fresh set is used
   * for each component subtree so that the same component interface can be reused under different
   * component names.
   */
  function extractCapabilities(
    iface: Record<string, unknown> | undefined,
    componentName: string | undefined,
    visited: Set<Record<string, unknown>>
  ) {
    if (!iface || visited.has(iface)) return;
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
          const compName = (cap["name"] as string) || componentName;
          extractCapabilities(
            resolveInterfaceRef(cap["schema"]),
            componentName || compName,
            new Set<Record<string, unknown>>()
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

        const columnName = componentName ? `${componentName}_${leafName}` : leafName;
        if (seenNames.has(columnName)) continue;
        seenNames.add(columnName);

        const capability: DtdlCapability = {
          name: columnName,
          leafName,
          component: componentName,
          displayName: getDisplayName(cap) || leafName,
          unit: typeof cap["unit"] === "string" ? (cap["unit"] as string) : undefined,
          description: getDescription(cap) || undefined,
          kind: isTelemetry ? "Telemetry" : "Property",
          schema: resolveSchemaName(cap["schema"]),
          kustoType: resolveKustoType(cap["schema"]),
          included: true,
        };
        (isTelemetry ? telemetries : properties).push(capability);
      }
    }

    // Follow inherited interfaces (extends): DTMI strings or inline interface objects.
    // Inherited capabilities belong to the same (component or root) context.
    for (const ext of asArray(iface["extends"])) {
      extractCapabilities(resolveInterfaceRef(ext), componentName, visited);
    }
  }

  extractCapabilities(root, undefined, new Set<Record<string, unknown>>());

  return { modelId, displayName, telemetries, properties };
}
