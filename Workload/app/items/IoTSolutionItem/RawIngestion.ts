// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

const RAW_TABLE_COLUMNS = [
  ["data", "string"],
  ["headers", "dynamic"],
  ["user_headers", "dynamic"],
] as const;

export const RAW_TABLE_SCHEMA = RAW_TABLE_COLUMNS
  .map(([name, type]) => `${name}: ${type}`)
  .join(", ");

export function buildRawTableCommand(tableName: string): string {
  return `.create table ${tableName} (${RAW_TABLE_SCHEMA})`;
}

interface EventstreamInputSchema {
  name: string;
  schema: { columns: unknown[] };
}

export function buildEventstreamTopology(
  streamDisplayName: string,
  workspaceId: string,
  kqlDatabaseId: string,
  databaseName: string,
  tableName: string,
) {
  const sourceName = "CustomEndpoint-Source";
  const defaultStreamName = `${streamDisplayName}-stream`;
  const destName = "Eventhouse";

  const sqlQuery =
    ` SELECT\n` +
    `     JSON_STRINGIFY(stream) AS data,\n` +
    `     GETMETADATAPROPERTYVALUE(stream, '[EventHub]') AS headers,\n` +
    `     GETMETADATAPROPERTYVALUE(stream, '[User]') AS user_headers\n` +
    ` INTO [${destName}]\n` +
    ` FROM [${defaultStreamName}] AS stream`;
  const destinationInputSchemas: EventstreamInputSchema[] = [
    { name: "SqlCode", schema: { columns: [] } },
  ];
  const operatorInputSchemas: EventstreamInputSchema[] = [
    { name: defaultStreamName, schema: { columns: [] } },
  ];
  const sqlProperties: { query: string; advancedSettings: null } = {
    query: sqlQuery,
    advancedSettings: null,
  };

  return {
    sources: [
      {
        name: sourceName,
        type: "CustomEndpoint",
        properties: {},
      },
    ],
    destinations: [
      {
        name: destName,
        type: "Eventhouse",
        properties: {
          dataIngestionMode: "ProcessedIngestion",
          workspaceId,
          itemId: kqlDatabaseId,
          databaseName,
          tableName,
          inputSerialization: {
            type: "Json",
            properties: { encoding: "UTF8" },
          },
        },
        inputNodes: [{ name: "SqlCode" }],
        inputSchemas: destinationInputSchemas,
      },
    ],
    streams: [
      {
        name: defaultStreamName,
        type: "DefaultStream",
        properties: {},
        inputNodes: [{ name: sourceName }],
      },
    ],
    operators: [
      {
        name: "SqlCode",
        type: "SQL",
        inputNodes: [{ name: defaultStreamName }],
        properties: sqlProperties,
        inputSchemas: operatorInputSchemas,
      },
    ],
    compatibilityLevel: "1.1",
  };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

export function validateRawTableSchema(tableName: string, schema: unknown): void {
  if (!isRecord(schema) || !Array.isArray(schema.OrderedColumns)) {
    throw new Error(`Could not read the schema of "${tableName}". Refresh and verify access to the selected database.`);
  }

  const columns = schema.OrderedColumns;
  const incompatible = RAW_TABLE_COLUMNS.filter(([name, type]) => {
    const column = columns.find((value: unknown) => isRecord(value) && value.Name === name);
    if (!isRecord(column)) return true;
    const actualType = column.CslType || column.Type;
    const systemType = type === "string" ? "System.String" : "System.Object";
    return actualType !== type && actualType !== systemType;
  });

  if (incompatible.length > 0) {
    const required = incompatible.map(([name, type]) => `${name}: ${type}`).join(", ");
    throw new Error(
      `Raw table "${tableName}" is incompatible with IoT Hub binary CloudEvents. ` +
      `Add or correct these columns before reuse: ${required}. ` +
      "Update the existing table and its Eventstream, or create new raw tables and Eventstreams."
    );
  }
}

function hasRawProjection(query: unknown, destinationName: unknown, inputs: unknown): boolean {
  if (typeof query !== "string" || typeof destinationName !== "string" || !Array.isArray(inputs)) {
    return false;
  }
  // Validate one complete raw SELECT, not matching fragments in other outputs.
  // Preserve quoted text while removing comments, including "--" in stream names.
  const sql = query.replace(
    /'(?:''|[^'])*'|\[(?:\]\]|[^\]])*\]|\/\*[\s\S]*?\*\/|--[^\r\n]*/g,
    (token) => token.startsWith("/*") || token.startsWith("--") ? " " : token
  );
  const identifier = String.raw`(?:[a-z_]\w*|\[(?:\]\]|[^\]])+\])`;
  const unquote = (name: string) => name.startsWith("[")
    ? name.slice(1, -1).replace(/\]\]/g, "]")
    : name;
  const statement = new RegExp(
    String.raw`^\s*SELECT\s+([\s\S]+?)\s+INTO\s+(${identifier})\s+FROM\s+(${identifier})\s+(?:AS\s+)?(${identifier})\s*;?\s*$`,
    "i"
  ).exec(sql);
  if (!statement || unquote(statement[2]) !== destinationName ||
      !inputs.some((input: unknown) => isRecord(input) && input.name === unquote(statement[3]))) {
    return false;
  }

  const projection = new RegExp(
    String.raw`(JSON_STRINGIFY|GETMETADATAPROPERTYVALUE)\s*\(\s*(${identifier})\s*(?:,\s*'(\[EventHub\]|\[User\])'\s*)?\)\s+AS\s+(${identifier})`,
    "gi"
  );
  const columns = new Set<string>();
  let position = 0;
  for (const match of statement[1].matchAll(projection)) {
    const separator = statement[1].slice(position, match.index);
    if (!(position === 0 ? /^\s*$/ : /^\s*,\s*$/).test(separator) ||
        unquote(match[2]).toLowerCase() !== unquote(statement[4]).toLowerCase()) {
      return false;
    }
    const name = unquote(match[4]);
    const expressionMatches = name === "data"
      ? match[1].toUpperCase() === "JSON_STRINGIFY" && match[3] === undefined
      : match[1].toUpperCase() === "GETMETADATAPROPERTYVALUE" &&
        ((name === "headers" && match[3] === "[EventHub]") ||
         (name === "user_headers" && match[3] === "[User]"));
    if (!expressionMatches || columns.has(name)) return false;
    columns.add(name);
    position = match.index + match[0].length;
  }
  return columns.size === RAW_TABLE_COLUMNS.length && statement[1].slice(position).trim() === "";
}

export function validateRawEventstreamTopology(
  eventstreamName: string,
  topology: unknown,
  target: { tableName: string; databaseId: string; databaseName: string },
): void {
  const error = new Error(
    `Eventstream "${eventstreamName}" is not configured for IoT Hub binary CloudEvents in "${target.tableName}". ` +
    "Update and publish the SQL operator feeding this table to project JSON_STRINGIFY(stream) AS data, " +
    "GETMETADATAPROPERTYVALUE(stream, '[EventHub]') AS headers, and " +
    "GETMETADATAPROPERTYVALUE(stream, '[User]') AS user_headers, or recreate the Eventstream."
  );
  if (!isRecord(topology) || !Array.isArray(topology.destinations) || !Array.isArray(topology.operators)) {
    throw error;
  }

  const operators = topology.operators.filter(isRecord);
  const destinations = topology.destinations.filter((value: unknown) => {
    if (!isRecord(value) || value.type !== "Eventhouse" || !isRecord(value.properties)) return false;
    const properties = value.properties;
    const matchesDatabase = target.databaseId
      ? properties.itemId === target.databaseId
      : properties.databaseName === target.databaseName;
    return matchesDatabase && properties.tableName === target.tableName;
  });

  if (destinations.length === 0 || !destinations.every((destination: unknown) => {
    if (!isRecord(destination) || !isRecord(destination.properties) ||
        destination.properties.dataIngestionMode !== "ProcessedIngestion" ||
        !Array.isArray(destination.inputNodes) || destination.inputNodes.length === 0) {
      return false;
    }
    return destination.inputNodes.every((input: unknown) =>
      isRecord(input) && typeof input.name === "string" && operators.some((operator) =>
        operator.name === input.name && operator.type === "SQL" &&
        isRecord(operator.properties) &&
        hasRawProjection(operator.properties.query, destination.name, operator.inputNodes)
      )
    );
  })) {
    throw error;
  }
}
