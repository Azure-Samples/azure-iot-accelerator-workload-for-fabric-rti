// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { decodeDoSamples, parseKqlTableResult, type JsonObject } from "./ModelAuthoring.ts";
import { fromThingModel, fromDtdlModel, setFieldRole, modelingFingerprint, fieldValidation,
  nativeIdentity, inspectFieldEvidence, clearAutomaticHints, doProjection, bindingStatus,
  type ModelingInput, type CoverageSnapshot } from "./HubModeling.ts";
import { resolveThingModel, mergeModelPatch } from "./ThingModelParser.ts";
import { parseDtdlModel } from "./DtdlModelParser.ts";
import { buildModeledDataPlan, semanticDocumentation } from "./ModeledDataPlan.ts";
import { buildDataAgentDataSource } from "./DataAgentContext.ts";
import { readColumnSemantics } from "./FieldSemantics.ts";

const tm: JsonObject = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel", id: "urn:sample:model",
  properties: { state: { type: "string" }, temperature: { type: "number" }, label: { type: "string" } } };
const row = (value: unknown, twin = false, headers: JsonObject = {}, user_headers: JsonObject = {}) => ({
  data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", contract_version: "1.0", schema_version: "1.0",
    event_kind: twin ? "lifecycle" : "measurement", event_time: "2026-10-01T12:00:00Z", payload: { value } }), headers, user_headers,
});
const eh = { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "", IoTSubject: "" };

test("V1 telemetry-then-property order is consistent in every mixed-role target and final projection", () => {
  const model = fromThingModel(resolveThingModel(tm));
  for (const roles of [["twin", "telemetry", "telemetry"], ["telemetry", "twin", "twin"], ["twin", "telemetry", "twin"]] as const) {
    for (const fields of [model.fields, [...model.fields].reverse()]) {
      const input: ModelingInput = { model, fields: fields.map((field, index) => setFieldRole(field, roles[index])),
        source: { queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databaseName: "db", telemetryTable: "Raw", twinTable: "Twins" },
        scope: { mode: "data", operator: "and", telemetry: "", twin: "" }, limit: 100 };
      const telemetry = decodeDoSamples([row({ temperature: 10 }, false, eh)], "telemetry", 100);
      const twin = decodeDoSamples([], "twin", 100);
      const snapshot: CoverageSnapshot = { fingerprint: modelingFingerprint(input), kind: "processing", telemetry, twin,
        fields: [], nativeIssues: [], checkedAt: "2026-10-01" };
      const names = { modeled: "Modeled", twinChanges: "Changes", twinState: "State" };
      const plan = buildModeledDataPlan(input, snapshot, false, names);
      const create = plan.commands.find(command => command.label === "Create Modeled")!.csl;
      const policy = plan.commands.find(command => command.label === "Set update policy on Modeled")!.csl;
      const query: string = JSON.parse(JSON.parse(policy.split(" policy update ")[1]))[0].Query;
      const actual = [...query.split("\n").at(-1)!.matchAll(/\['([^']+)'\]=/g)].map(match => match[1]);
      assert.deepEqual(actual, [...input.fields.filter(field => field.role === "telemetry"), ...input.fields.filter(field => field.role === "twin")].map(field => field.column));
      assert.deepEqual([...create.matchAll(/\['([^']+)'\]:/g)].map(match => match[1]), ["deviceId", "event_time", ...actual]);
      for (const field of input.fields) assert.ok(create.includes(`['${field.column}']:${field.kustoType}`));
    }
  }
});

const primary = { TableName: "PrimaryResult", Columns: [{ ColumnName: "value", ColumnType: "long" }], Rows: [[1]] };
const status = (severity: unknown, code: unknown = 0) => ({
  TableName: "QueryStatus", Columns: [{ ColumnName: "Severity" }, { ColumnName: "StatusCode" }], Rows: [[severity, code]],
});
test("HTTP200 partial, failure and malformed QueryStatus are never consumed", () => {
  for (const invalid of [status(2), status(1), status(0), status(4, -1), status("4"), status(4, "0"),
    { ...status(4), Rows: [] }, { ...status(4), Rows: [[4]] }, { ...status(4), Columns: [] }]) {
    assert.throws(() => parseKqlTableResult({ Tables: [primary, invalid] }));
    assert.throws(() => parseKqlTableResult({ Tables: [invalid, primary] }));
  }
  for (const valid of [status(4), status(3)]) assert.deepEqual(parseKqlTableResult({ Tables: [valid, primary] }).Rows, [[1]]);
  assert.deepEqual(parseKqlTableResult({ Tables: [primary] }).Rows, [[1]]);
  assert.throws(() => parseKqlTableResult({ Tables: [primary], Exceptions: ["partial"] }));
  assert.throws(() => parseKqlTableResult({ Tables: [primary], error: { message: "failure" } }));
});

test("v1 Table_N status is found through the result directory, including malformed status columns", () => {
  const directory = { TableName: "Table_2", Columns: [{ ColumnName: "Ordinal" }, { ColumnName: "Kind" }, { ColumnName: "Name" }],
    Rows: [[0, "QueryResult", "PrimaryResult"], [1, "QueryStatus", "QueryStatus"]] };
  const data = { ...primary, TableName: "Table_0" };
  assert.deepEqual(parseKqlTableResult({ Tables: [data, { ...status(4), TableName: "Table_1" }, directory] }).Rows, [[1]]);
  assert.throws(() => parseKqlTableResult({ Tables: [data, { ...status(2), TableName: "Table_1" }, directory] }));
  assert.throws(() => parseKqlTableResult({ Tables: [data, { TableName: "Table_1", Columns: [], Rows: [] }, directory] }));
});

test("Event Hub root/components bind the same native identities without model refs or serial numbers", () => {
  const model = fromDtdlModel(parseDtdlModel([
    { "@id": "dtmi:example:Main;1", "@type": "Interface", contents: [
      { "@type": "Telemetry", name: "temperature", schema: "double" },
      { "@type": "Component", name: "thermostat", schema: "dtmi:example:Part;1" },
    ] },
    { "@id": "dtmi:example:Part;1", "@type": "Interface", contents: [{ "@type": "Telemetry", name: "temperature", schema: "double" }] },
  ]), null);
  const telemetry = decodeDoSamples([
    row({ temperature: 20, serialNumber: "same" }, false, eh),
    row({ temperature: 21, serialNumber: "same" }, false, { ...eh, IoTSubject: "thermostat" }),
    row({ temperature: 22, serialNumber: "same" }, false, { ...eh, IoTConnectionDeviceId: "device-2" }),
  ], "telemetry", 100);
  const twin = decodeDoSamples([row({ properties: { reported: { thermostat: { __t: "c", temperature: 5 } } } }, true, eh,
    { deviceId: "device-1", hubName: "hub-1" })], "twin", 100);
  assert.deepEqual(telemetry.payloads.map(sample => nativeIdentity(sample, "telemetry")), ["device-1", "device-1", "device-2"]);
  assert.equal(telemetry.payloads[1].component, "thermostat");
  const evidence = inspectFieldEvidence(model.fields, telemetry, twin);
  assert.equal(evidence[0].telemetry, 2); assert.equal(evidence[1].telemetry, 1); assert.equal(evidence[1].twin, 1);
  assert.equal(bindingStatus(telemetry, "telemetry").identity, "Matched in this sample");
  for (const stream of ["telemetry", "twin"] as const) {
    assert.match(doProjection(stream), /IoTConnectionDeviceId/);
    assert.match(doProjection(stream), /IoTConnectionModuleId/);
    assert.ok(!doProjection(stream).includes("memberIRI"));
  }
});

test("module transport and conflicting metadata remain distinct from PnP components", () => {
  const telemetry = decodeDoSamples([row({ temperature: 10 }, false, { ...eh, IoTConnectionModuleId: "module-1", IoTSubject: "part" })], "telemetry", 100);
  const twin = decodeDoSamples([row({ properties: { reported: {} } }, true, { ...eh, IoTConnectionModuleId: "module-1" })], "twin", 100);
  assert.equal(nativeIdentity(telemetry.payloads[0], "telemetry"), "device-1/module-1");
  assert.equal(nativeIdentity(twin.payloads[0], "twin"), "device-1/module-1");
  assert.equal(telemetry.payloads[0].component, "part");
  const conflicting = decodeDoSamples([row({}, false, eh, { cloudEvents_subject: "another-device" })], "telemetry", 100);
  assert.equal(nativeIdentity(conflicting.payloads[0], "telemetry"), undefined);
  assert.equal(bindingStatus(conflicting, "telemetry").identity, "Needs review");
});

test("unrecognized component metadata cannot provide a false root-field hint", () => {
  const model = fromThingModel(resolveThingModel(tm));
  const telemetry = decodeDoSamples([row({ temperature: 10 })], "telemetry", 100);
  const twin = decodeDoSamples([], "twin", 100);
  assert.ok(inspectFieldEvidence(model.fields, telemetry, twin).every(evidence => evidence.telemetry === 0));
  assert.equal(bindingStatus(telemetry, "telemetry").component, "Needs review");
});

test("source invalidation drops only unconfirmed automatic hints", () => {
  const fields = fromThingModel(resolveThingModel(tm)).fields;
  const stale = fields.map((field, index) => ({ ...field, role: "telemetry" as const,
    roleSource: (["hint", "manual", "model"] as const)[index] }));
  const next = clearAutomaticHints(stale);
  assert.equal(next[0].role, undefined); assert.equal(next[0].roleSource, undefined);
  assert.equal(next[1].role, "telemetry"); assert.equal(next[2].role, "telemetry");
});

test("root component has disjoint telemetry, twin and durable state identities", () => {
  const document = [
    { "@id": "dtmi:example:A;1", "@type": "Interface", contents: [
      { "@type": "Telemetry", name: "temperature", schema: "double" }, { "@type": "Property", name: "state", schema: "string" },
      { "@type": "Component", name: "root", schema: "dtmi:example:B;1" }] },
    { "@id": "dtmi:example:B;1", "@type": "Interface", contents: [
      { "@type": "Telemetry", name: "temperature", schema: "double" }, { "@type": "Property", name: "state", schema: "string" }] },
  ];
  const model = fromDtdlModel(parseDtdlModel(document), document);
  assert.equal(new Set(model.fields.map(field => field.id)).size, 4);
  assert.deepEqual(fieldValidation(model.fields), []);
  assert.ok(fieldValidation([model.fields[0], { ...model.fields[0], column: "distinct" }]).some(error => error.includes("identity")));
});

test("tm:ref patches remove nulls recursively, replace arrays, retain source JSON and validate effective definitions", () => {
  const base: JsonObject = { ...tm, properties: { dim: { type: "integer", title: "Dimming",
    maximum: 100, enum: [0, 1, 2], descriptions: { en: "value", fr: "valeur" } } } };
  const child: JsonObject = { ...tm, id: "urn:child", properties: { dim: {
    "tm:ref": "urn:sample:model#/properties/dim", title: null, maximum: 80, enum: [0, 1],
    descriptions: { en: null }, unit: "%",
  } } };
  const resolved = resolveThingModel(child, [{ fileName: "base.json", document: base }], "child.json");
  assert.deepEqual(resolved.blockers, []);
  const dim = (resolved.effectiveDocument!.properties as JsonObject).dim as JsonObject;
  assert.ok(!Object.hasOwn(dim, "title")); assert.deepEqual(dim.enum, [0, 1]);
  assert.deepEqual(dim.descriptions, { fr: "valeur" }); assert.equal(dim.maximum, 80);
  assert.deepEqual(resolved.document, child);
  const invalid = resolveThingModel({ ...child, properties: { dim: { "tm:ref": "urn:sample:model#/properties/dim", maximum: "wrong" } } },
    [{ fileName: "base.json", document: base }], "child.json");
  assert.ok(invalid.blockers.length); assert.equal(invalid.effectiveDocument, undefined);
  assert.deepEqual(mergeModelPatch({ values: [1, 2], nested: { a: 1, b: 2 } }, { values: [3], nested: { a: null } }),
    { values: [3], nested: { b: 2 } });
});

test("ordinary and transitive inheritance retain new and inherited literal null facts", () => {
  const base: JsonObject = { ...tm, properties: { inheritedNil: { const: null }, inheritedDefault: { type: "string", default: null } } };
  const child: JsonObject = { ...tm, id: "urn:child", links: [{ rel: "tm:extends", href: "urn:sample:model" }], properties: {
    nil: { const: null }, state: { type: "string", default: null },
    importedNil: { "tm:ref": "urn:sample:model#/properties/inheritedNil" },
  } };
  const grandchild: JsonObject = { ...tm, id: "urn:grandchild", links: [{ rel: "tm:extends", href: "urn:child" }],
    properties: { nil: { title: "Declared null" } } };
  const resolved = resolveThingModel(grandchild, [
    { fileName: "base.json", document: base }, { fileName: "child.json", document: child },
  ], "grandchild.json");
  assert.deepEqual(resolved.blockers, []);
  assert.deepEqual(resolved.effectiveDocument!.properties, {
    inheritedNil: { const: null }, inheritedDefault: { type: "string", default: null },
    nil: { const: null, title: "Declared null" }, state: { type: "string", default: null }, importedNil: { const: null },
  });
  assert.deepEqual(resolved.document, grandchild);
  const model = fromThingModel(resolved);
  const docs = semanticDocumentation(model, model.fields.map(field => setFieldRole(field, "twin")));
  const nil = readColumnSemantics(docs.columns.nil), state = readColumnSemantics(docs.columns.state);
  assert.ok(nil?.version === 3 && state?.version === 3);
  assert.ok(Object.hasOwn(nil.field.facts, "const"));
  assert.ok(Object.hasOwn(state.field.facts, "default"));
  assert.equal(nil.field.facts.const, null);
  assert.equal(state.field.facts.default, null);
  const direct = resolveThingModel(child, [{ fileName: "base.json", document: base }], "child.json");
  assert.deepEqual((direct.effectiveDocument!.properties as JsonObject).nil, { const: null });
  assert.deepEqual((direct.effectiveDocument!.properties as JsonObject).state, { type: "string", default: null });
});

test("nested imports preserve resolved null facts without weakening actual removal patches", () => {
  const base: JsonObject = { ...tm, properties: {
    nil: { const: null, default: null },
    object: { type: "object", properties: {
      remove: { type: "string" }, inheritedNil: { const: null }, keep: { type: "integer", enum: [1, 2], title: "Remove title" },
    } },
  } };
  const child: JsonObject = { ...tm, id: "urn:child", properties: { imported: {
    "tm:ref": "urn:sample:model#/properties/object",
    properties: { remove: null, embeddedNil: { "tm:ref": "urn:sample:model#/properties/nil" },
      keep: { title: null, enum: [2] } },
  } } };
  const original = JSON.stringify({ base, child });
  const resolved = resolveThingModel(child, [{ fileName: "base.json", document: base }], "child.json");
  assert.deepEqual(resolved.blockers, []);
  const imported = (resolved.effectiveDocument!.properties as JsonObject).imported as JsonObject;
  assert.deepEqual(imported.properties, {
    inheritedNil: { const: null }, keep: { type: "integer", enum: [2] }, embeddedNil: { const: null, default: null },
  });
  assert.equal(JSON.stringify({ base, child }), original);
  assert.deepEqual(resolved.document, child);
});

test("component recognition does not claim root traffic or selected-model coverage", () => {
  const componentOnly = decodeDoSamples([row({ temperature: 22 }, false, { ...eh, IoTSubject: "other-component" })], "telemetry", 100);
  assert.equal(bindingStatus(componentOnly, "telemetry").component, "Component metadata recognized in this sample");
  const model = fromThingModel(resolveThingModel(tm));
  assert.ok(inspectFieldEvidence(model.fields, componentOnly, decodeDoSamples([], "twin", 100)).every(field => field.telemetry === 0));
  const rootOnly = decodeDoSamples([row({ temperature: 22 }, false, eh)], "telemetry", 100);
  assert.equal(bindingStatus(rootOnly, "telemetry").component, "Root traffic only in this sample");
});

test("ignored submodel composition retains usable fields and gives a concise diagnostic without resolving it", () => {
  const resolved = resolveThingModel({ ...tm, links: [{ rel: "tm:submodel", href: "urn:submodel" }] });
  assert.deepEqual(resolved.blockers, []);
  assert.ok(resolved.effectiveDocument && resolved.fields.length);
  assert.deepEqual(resolved.ignored, [{ kind: "Submodel", name: "Submodel 1", reason: "submodel" }]);
});

test("serialized source definition restores native Tables membership and observed provider leaf types", () => {
  const definition = buildDataAgentDataSource({ database: "db", table: "Table", docstring: "Telemetry events with last-known state.",
    columns: [{ name: "event_time", type: "datetime" }, { name: "custom", type: "bool", docstring: "Declared status." }] }, "workspace", "database", "table-id");
  const serialized = JSON.parse(JSON.stringify(definition));
  assert.equal(serialized.elements[0].type, "table_grouping");
  assert.equal(serialized.elements[0].display_name, "Tables");
  assert.equal(serialized.elements[0].is_selected, false);
  const table = serialized.elements[0].children[0];
  assert.equal(table.type, "kusto.table");
  assert.equal(table.is_selected, true);
  assert.equal(table.children[1].type, "kusto.column");
  assert.equal(table.children[1].data_type, "System.SByte");
  assert.ok(!serialized.dataSourceInstructions.includes("temperature"));
});
