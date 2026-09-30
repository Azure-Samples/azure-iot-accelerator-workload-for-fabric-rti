// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  RAW_TABLE_SCHEMA,
  buildRawTableCommand,
  buildEventstreamTopology,
  validateRawTableSchema,
  validateRawEventstreamTopology,
} from "./RawIngestion.ts";

const target = {
  tableName: "TelemetryRaw",
  databaseId: "database-id",
  databaseName: "IoT",
};

function topology() {
  return buildEventstreamTopology("Telemetry", "workspace-id", target.databaseId, target.databaseName, target.tableName);
}

function schema() {
  return {
    OrderedColumns: [
      { Name: "data", CslType: "string" },
      { Name: "headers", CslType: "dynamic" },
      { Name: "user_headers", CslType: "dynamic" },
    ],
  };
}

const oldQuery =
  "SELECT JSON_STRINGIFY(stream) AS data, " +
  "GETMETADATAPROPERTYVALUE(stream, '[EventHub]') AS headers " +
  "INTO [Eventhouse] FROM [Telemetry-stream] AS stream";

test("creates both raw tables with the same three-column contract", () => {
  assert.equal(RAW_TABLE_SCHEMA, "data: string, headers: dynamic, user_headers: dynamic");
  for (const name of ["TelemetryRaw", "PropertiesRaw"]) {
    assert.equal(
      buildRawTableCommand(name),
      `.create table ${name} (data: string, headers: dynamic, user_headers: dynamic)`
    );
  }
});

test("projects body, EventHub metadata, and User metadata for both raw streams", () => {
  for (const kind of ["Telemetry", "Properties"]) {
    const tableName = `${kind}Raw`;
    const result = buildEventstreamTopology(kind, "workspace-id", target.databaseId, target.databaseName, tableName);
    assert.equal(result.operators[0].properties.query,
      ` SELECT\n` +
      `     JSON_STRINGIFY(stream) AS data,\n` +
      `     GETMETADATAPROPERTYVALUE(stream, '[EventHub]') AS headers,\n` +
      `     GETMETADATAPROPERTYVALUE(stream, '[User]') AS user_headers\n` +
      ` INTO [Eventhouse]\n` +
      ` FROM [${kind}-stream] AS stream`
    );
    assert.equal(result.destinations[0].properties.tableName, tableName);
    assert.equal(result.destinations[0].properties.dataIngestionMode, "ProcessedIngestion");
    assert.deepEqual(result.destinations[0].inputNodes, [{ name: result.operators[0].name }]);
    assert.doesNotThrow(() => validateRawEventstreamTopology(kind, result, { ...target, tableName }));
  }
});

test("accepts compatible raw schemas, including additional columns and CLR types", () => {
  const value = schema();
  value.OrderedColumns.push({ Name: "extra", CslType: "long" });
  assert.doesNotThrow(() => validateRawTableSchema("TelemetryRaw", value));
  assert.doesNotThrow(() => validateRawTableSchema("PropertiesRaw", {
    OrderedColumns: [
      { Name: "data", Type: "System.String" },
      { Name: "headers", Type: "System.Object" },
      { Name: "user_headers", Type: "System.Object" },
    ],
  }));
});

test("rejects missing or wrongly typed columns with upgrade or recreate guidance", () => {
  for (const name of ["data", "headers", "user_headers"]) {
    const missing = schema();
    missing.OrderedColumns = missing.OrderedColumns.filter((column) => column.Name !== name);
    assert.throws(() => validateRawTableSchema("TelemetryRaw", missing),
      (error: Error) => error.message.includes(name) && /Update.*or create new/.test(error.message));

    const wrongType = schema();
    wrongType.OrderedColumns.find((column) => column.Name === name)!.CslType = "long";
    assert.throws(() => validateRawTableSchema("TelemetryRaw", wrongType), /incompatible/);
  }
});

test("rejects malformed table schema responses", () => {
  for (const value of [null, [], {}, { OrderedColumns: {} }]) {
    assert.throws(() => validateRawTableSchema("TelemetryRaw", value), /Could not read the schema/);
  }
});

test("accepts SQL whitespace, keyword casing, and bracketed column aliases", () => {
  const value = topology();
  value.operators[0].properties.query = value.operators[0].properties.query
    .replace("SELECT", "select")
    .replaceAll(" AS ", " as ")
    .replace("as user_headers", "as [user_headers]");
  assert.doesNotThrow(() => validateRawEventstreamTopology("Telemetry", value, target));
});

test("accepts reordered raw projections and comments without corrupting quoted stream names", () => {
  const value = buildEventstreamTopology("Telemetry--hub", "workspace-id", target.databaseId, target.databaseName, target.tableName);
  value.operators[0].properties.query = `
    -- Binary raw contract
    SELECT GETMETADATAPROPERTYVALUE(evt, '[User]') AS user_headers,
           /* retain transport metadata */ GETMETADATAPROPERTYVALUE(evt, '[EventHub]') AS headers,
           JSON_STRINGIFY(evt) AS data
    INTO [Eventhouse]
    FROM [Telemetry--hub-stream] AS evt;
  `;
  assert.doesNotThrow(() => validateRawEventstreamTopology("Telemetry--hub", value, target));
});

test("rejects old projections even when User metadata is mentioned in comments", () => {
  for (const comment of [
    "",
    "\n-- , GETMETADATAPROPERTYVALUE(stream, '[User]') AS user_headers\n",
    " /* , GETMETADATAPROPERTYVALUE(stream, '[User]') AS user_headers */ ",
  ]) {
    const value = topology();
    value.operators[0].properties.query = oldQuery + comment;
    assert.throws(
      () => validateRawEventstreamTopology("Telemetry", value, target),
      /Update and publish.*\[User\].*or recreate/
    );
  }
});

test("rejects incorrect metadata namespaces, aliases, or body projections", () => {
  const original = topology().operators[0].properties.query;
  for (const query of [
    original.replace("'[User]'", "'[EventHub]'"),
    original.replace("'[User]'", "'[user]'"),
    original.replace(" AS user_headers", " AS user_headers_extra"),
    original.replace(" AS user_headers", " AS User_Headers"),
    original.replace("JSON_STRINGIFY(stream)", "stream"),
    original.replace(" AS headers,", " AS other_headers,"),
  ]) {
    const value = topology();
    value.operators[0].properties.query = query;
    assert.throws(() => validateRawEventstreamTopology("Telemetry", value, target), /not configured/);
  }
});

test("validates the operator feeding the selected destination, not an unrelated operator", () => {
  const value = topology();
  const unused = structuredClone(value.operators[0]);
  unused.name = "UnusedSql";
  value.operators.push(unused);
  value.operators[0].properties.query = oldQuery;
  assert.throws(() => validateRawEventstreamTopology("Telemetry", value, target), /not configured/);
});

test("rejects projections from another output or SELECT statement", () => {
  const original = topology().operators[0].properties.query;
  for (const query of [
    original.replace("INTO [Eventhouse]", "INTO [Unrelated]"),
    original.replace("INTO [Eventhouse]", "INTO [Unrelated]") + ";\n" + oldQuery,
    original.replace("JSON_STRINGIFY(stream)", "JSON_STRINGIFY(other)"),
    original.replace("FROM [Telemetry-stream]", "FROM [Unrelated-stream]"),
    original.replace(" AS data,", " AS user_headers,"),
    original.replace(" AS headers,", " AS data,"),
  ]) {
    const value = topology();
    value.operators[0].properties.query = query;
    assert.throws(() => validateRawEventstreamTopology("Telemetry", value, target), /not configured/);
  }
});

test("rejects wrong databases, tables, ingestion modes, and missing input connections", () => {
  const wrongDatabase = topology();
  wrongDatabase.destinations[0].properties.itemId = "another-database";
  const wrongTable = topology();
  wrongTable.destinations[0].properties.tableName = "AnotherTable";
  const direct = topology();
  direct.destinations[0].properties.dataIngestionMode = "DirectIngestion";
  const disconnected = topology();
  disconnected.destinations[0].inputNodes = [];

  for (const value of [wrongDatabase, wrongTable, direct, disconnected, null, {}]) {
    assert.throws(() => validateRawEventstreamTopology("Telemetry", value, target), /not configured/);
  }
});

test("requires every input feeding the selected table to have the new projection", () => {
  const value = topology();
  const oldOperator = structuredClone(value.operators[0]);
  oldOperator.name = "OldSql";
  oldOperator.properties.query = oldQuery;
  value.operators.push(oldOperator);
  value.destinations[0].inputNodes.push({ name: "OldSql" });
  assert.throws(() => validateRawEventstreamTopology("Telemetry", value, target), /not configured/);
});

test("does not mutate existing schemas or Eventstream topologies during validation", () => {
  const table = schema();
  const stream = topology();
  const beforeTable = structuredClone(table);
  const beforeStream = structuredClone(stream);
  validateRawTableSchema("TelemetryRaw", table);
  validateRawEventstreamTopology("Telemetry", stream, target);
  assert.deepEqual(table, beforeTable);
  assert.deepEqual(stream, beforeStream);
});
