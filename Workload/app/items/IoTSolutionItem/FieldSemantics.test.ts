// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { COLUMN_SEMANTICS_PREFIX, FIELD_SEMANTIC_INFO_PREFIX, columnDocstring, columnSemanticAnnotation,
  declaredMetadataDetails, publishedFieldFacts, readColumnSemantics, wotFieldMetadata } from "./FieldSemantics.ts";
import { fromThingModel, inspectFieldEvidence, kqlString, modelingFingerprint, setFieldRole } from "./HubModeling.ts";
import { resolveThingModel } from "./ThingModelParser.ts";
import { ModeledDataCreator } from "./ModeledDataCreation.ts";
import { buildModeledDataPlan, semanticDocumentation } from "./ModeledDataPlan.ts";
import { createCreationMock, creationRequest } from "../../../tests/creationMock.ts";
import { metadataThingModel, tupleThingModel } from "../../../tests/fieldSemanticFixtures.ts";
import { canonicalJson, isJsonObject, type JsonObject } from "./ModelAuthoring.ts";
import { buildDataAgentContext } from "./DataAgentContext.ts";

test("WoT tuple columns preserve ordered referenced nested and event schemas through current and legacy facts", () => {
  const document = { ...tupleThingModel, properties: { ...tupleThingModel.properties,
    recentAirTemperatures: { type: "array", items: { type: "number", minimum: 0 } },
    route: { type: "object", properties: {
      location: { "tm:ref": "#/schemaDefinitions/TupleArray" },
      matrix: { type: "array", items: [{ "tm:ref": "#/schemaDefinitions/TupleArray" }, { type: "boolean" }] },
    } },
  }, events: { samples: { data: { type: "array", items: [{ type: "integer" }, { type: "boolean" }] } } } };
  const before = canonicalJson(document), model = fromThingModel(resolveThingModel(document));
  assert.ok(model.fields.every(field => field.role === undefined && field.path.length === 1));
  const bay = model.fields.find(field => field.path[0] === "bayLocation")!;
  assert.equal(bay.kustoType, "dynamic"); assert.equal(bay.column, "bayLocation");
  const facts = publishedFieldFacts(bay), doc = columnDocstring(facts), parsed = readColumnSemantics(doc)!;
  assert.equal(parsed.version, 3);
  if (parsed.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(parsed.bounded, false);
  assert.deepEqual(parsed.field.facts.items, tupleThingModel.schemaDefinitions.TupleArray.items);
  assert.equal(parsed.field.facts.minItems, 3); assert.equal(parsed.field.facts.maxItems, 3);
  assert.equal(parsed.field.facts.prefixItems, undefined);
  const docs = semanticDocumentation(model, model.fields.map(field => setFieldRole(field, "telemetry")));
  const route = readColumnSemantics(docs.columns.route)!;
  assert.equal(route.version, 3);
  if (route.version !== 3 || !isJsonObject(route.field.facts.properties)) throw new Error("Expected whole object facts.");
  const { location, matrix } = route.field.facts.properties;
  assert.ok(isJsonObject(location) && isJsonObject(matrix) && Array.isArray(matrix.items) && isJsonObject(matrix.items[0]));
  assert.deepEqual(location.items, facts.items); assert.deepEqual(matrix.items[0].items, facts.items);
  assert.deepEqual(matrix.items[1], { type: "boolean" });
  const event = model.fields.find(field => field.id.startsWith("/events/"))!;
  const eventFacts = publishedFieldFacts(event);
  assert.ok(isJsonObject(eventFacts.data));
  assert.deepEqual(eventFacts.data.items, [{ type: "integer" }, { type: "boolean" }]);
  assert.ok(columnDocstring(eventFacts).length <= 1000);
  const homogeneous = readColumnSemantics(docs.columns.recentAirTemperatures)!;
  assert.equal(homogeneous.version, 3);
  if (homogeneous.version !== 3) throw new Error("Expected semantic JSON.");
  assert.deepEqual(homogeneous.field.facts.items, { type: "number", minimum: 0 });
  const legacy = readColumnSemantics(columnSemanticAnnotation({ kind: "wot" }, bay.column, bay.declaredMetadata!))!;
  assert.equal(legacy.version, 1);
  if (legacy.version !== 1 || !isJsonObject(legacy.field.metadata.schema)) throw new Error("Expected legacy WoT metadata.");
  assert.deepEqual(legacy.field.metadata.schema.items, facts.items);
  const flat = readColumnSemantics(`name: "bayLocation"\ntype: "array"\nitems: ${JSON.stringify(facts.items)}`)!;
  assert.equal(flat.version, 2);
  if (flat.version !== 2) throw new Error("Expected legacy flat facts.");
  assert.deepEqual(flat.field.facts.items, facts.items);
  assert.equal(canonicalJson(document), before);
});

test("WoT tuple validation rejects malformed entries without permitting arrays in other schema slots", () => {
  for (const item of [null, 0, true, false, "number", [], { unsupported: "schema" }]) {
    assert.throws(() => columnDocstring({ name: "x", type: "array", items: [item] }), /Malformed declared|Unsupported column/);
    assert.throws(() => readColumnSemantics(`name: "x"\ntype: "array"\nitems: ${JSON.stringify([item])}`),
      /Malformed declared|Unsupported column/);
  }
  for (const key of ["schema", "valueSchema", "elementSchema", "mapKey", "mapValue", "data"]) {
    assert.throws(() => columnDocstring({ name: "x", [key]: [{ type: "number" }] }), /Malformed declared/);
  }
  for (const items of [null, 0]) assert.throws(() => columnDocstring({ name: "x", items }), /Malformed declared/);
  assert.throws(() => columnDocstring({ name: "x", items: [{ type: "array", items: [null] }] }), /Malformed declared/);
  assert.throws(() => columnDocstring({ name: "x", items: [{ type: "number", schema: "double" }] }), /aliases/);
  assert.throws(() => columnDocstring({ name: "x", items: [{ "Declared type": "number" }] }), /aliases/);
  assert.throws(() => readColumnSemantics(FIELD_SEMANTIC_INFO_PREFIX +
    '{"name":"x","items":[{"type":"number","type":"string"}]}'), /Duplicate/);
  assert.throws(() => readColumnSemantics(columnDocstring({ name: "x", items: [{}] }) + '\nname: "injected"'), /ambiguous/);
  assert.deepEqual(readColumnSemantics(columnDocstring({ name: "x", items: [] }))?.version, 3);
  let deep: JsonObject = {};
  for (let index = 0; index < 26; index++) deep = { items: [deep] };
  assert.throws(() => columnDocstring({ name: "deep", ...deep }), /nested|nesting/);
  assert.throws(() => readColumnSemantics(FIELD_SEMANTIC_INFO_PREFIX + JSON.stringify({ name: "deep", ...deep })), /depth/);
});

test("WoT tuple budgets keep or omit complete item lists and preserve escaped semantic values", () => {
  const items = tupleThingModel.schemaDefinitions.TupleArray.items;
  const special = 'Latitude "quoted" \\ value\n' + FIELD_SEMANTIC_INFO_PREFIX + "{}\u2028\u2029";
  const escaped = { name: "location", type: "array", minItems: 3, maxItems: 3, items: [{ ...items[0], title: special }, ...items.slice(1)] };
  const doc = columnDocstring(escaped), parsed = readColumnSemantics(doc)!;
  assert.equal(parsed.version, 3);
  if (parsed.version !== 3) throw new Error("Expected semantic JSON.");
  assert.deepEqual(parsed.field.facts.items, escaped.items);
  assert.ok(!doc.includes("\n") && doc.includes("\\u2028") && doc.includes("\\u2029"));
  assert.equal(JSON.parse(kqlString(doc)), doc);
  for (const size of [999, 1000, 1001]) {
    const facts = { name: "location", type: "array", minItems: 3, maxItems: 3,
      items: [{ type: "number", title: "" }, ...items.slice(1)] };
    const baseLength = FIELD_SEMANTIC_INFO_PREFIX.length + JSON.stringify(facts).length;
    facts.items[0].title = "x".repeat(size - baseLength);
    const boundedDoc = columnDocstring(facts), result = readColumnSemantics(boundedDoc)!;
    assert.ok(boundedDoc.length <= 1000);
    assert.equal(result.version, 3);
    if (result.version !== 3) throw new Error("Expected semantic JSON.");
    assert.equal(result.field.facts.minItems, 3); assert.equal(result.field.facts.maxItems, 3);
    assert.equal(result.bounded, size > 1000);
    if (size <= 1000) {
      assert.equal(boundedDoc.length, size); assert.deepEqual(result.field.facts.items, facts.items);
    } else assert.equal(result.field.facts.items, undefined, "Never keep only a prefix of tuple positions.");
  }
  const oversized = columnDocstring({ name: "nested", type: "object", properties: {
    position: { type: "array", items: [{ type: "string", enum: ["x".repeat(2000)] }, ...items] },
  } });
  const result = readColumnSemantics(oversized)!;
  assert.equal(result.version, 3);
  if (result.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(result.bounded, true); assert.equal(result.field.facts.properties, undefined);
});

test("WoT tuple native plan guarded writes and Data Agent readback retain one dynamic column and exact tuple facts", async () => {
  const request = creationRequest(true), model = fromThingModel(resolveThingModel(tupleThingModel));
  request.input = { ...request.input, model, fields: model.fields.map(field => setFieldRole(field,
    field.path[0] === "serialNumber" ? "twin" : "telemetry")),
    scope: { ...request.input.scope, reference: { id: model.id! } } };
  request.coverage = { ...request.coverage, fingerprint: modelingFingerprint(request.input),
    fields: inspectFieldEvidence(request.input.fields, request.coverage.telemetry, request.coverage.twin) };
  const originalInput = canonicalJson(request.input), mock = createCreationMock(request);
  const withoutMetadata = request.input.fields.map(({ declaredMetadata: _metadata, ...field }) => field);
  const plainInput = { ...request.input, fields: withoutMetadata, model: { ...model, fields: withoutMetadata } };
  const plainPlan = buildModeledDataPlan(plainInput, { ...request.coverage, fingerprint: modelingFingerprint(plainInput) }, true, request.names);
  const processing = (commands: typeof mock.plan.commands) => commands.filter(command =>
    ["table", "view", "policy"].includes(command.kind)).map(command => command.csl);
  assert.deepEqual(processing(mock.plan.commands), processing(plainPlan.commands));
  const completed = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(completed.phase, "completed"); assert.equal(mock.writes.length, 9);
  const table = mock.tables.get(request.names.modeled)!;
  assert.equal(table.columns.filter(column => column.name === "bayLocation" && column.type === "dynamic").length, 1);
  assert.deepEqual(table.columns.map(column => column.name), ["deviceId", "event_time", "bayLocation", "serialNumber"]);
  const doc = table.columnDocs.bayLocation, parsed = readColumnSemantics(doc)!;
  assert.equal(parsed.version, 3);
  if (parsed.version !== 3) throw new Error("Expected semantic JSON.");
  assert.deepEqual(parsed.field.facts.items, tupleThingModel.schemaDefinitions.TupleArray.items);
  assert.equal(parsed.field.facts.role, "telemetry"); assert.equal(parsed.bounded, false);
  assert.equal(doc, mock.plan.documentation.columns.bayLocation);
  const agent = buildDataAgentContext({ database: "Database", table: request.names.modeled, docstring: table.doc,
    columns: table.columns.map(column => ({ ...column, docstring: table.columnDocs[column.name] })) });
  assert.ok(agent.sourceInstructions.includes(`Items: ${JSON.stringify(parsed.field.facts.items)}`));
  assert.ok(agent.sourceInstructions.includes("Minimum items: 3") && agent.sourceInstructions.includes("Maximum items: 3"));
  assert.equal(canonicalJson(request.input), originalInput);
});

test("declared WoT metadata preserves false zero empty null localized and resolved whole schemas without role authority", () => {
  const model = fromThingModel(resolveThingModel(metadataThingModel));
  assert.equal(model.fields.length, 5);
  assert.ok(model.fields.every(field => field.path.length === 1 && field.role === undefined));
  const temperature = model.fields.find(field => field.path[0] === "temperature")!;
  const details = Object.fromEntries(declaredMetadataDetails(temperature.declaredMetadata!).map(detail => [detail.name, detail.value]));
  assert.equal(details.minimum, 0); assert.equal(details.const, 0); assert.equal(details.default, 0);
  assert.equal(details.readOnly, false); assert.equal(details.writeOnly, false); assert.equal(details.observable, false);
  assert.deepEqual(details.titles, metadataThingModel.schemaDefinitions.reading.titles);
  const message = model.fields.find(field => field.path[0] === "message")!.declaredMetadata!;
  assert.equal((message.schema as Record<string, unknown>).const, "");
  const config = model.fields.find(field => field.path[0] === "config")!.declaredMetadata!;
  assert.equal((config.schema as Record<string, unknown>).default, null);
  assert.deepEqual((config.schema as Record<string, unknown>).properties, metadataThingModel.properties.config.properties);
  const safe = wotFieldMetadata({ type: "string", forms: [{ href: "credential-marker" }], security: ["credential-marker"] },
    { observable: false, forms: [{ href: "credential-marker" }], uriVariables: { token: "credential-marker" } }, true);
  assert.ok(!JSON.stringify(safe).includes("credential-marker"));
  assert.equal(readColumnSemantics("Existing legacy meaning.\nTelemetry."), undefined);
  assert.throws(() => readColumnSemantics(COLUMN_SEMANTICS_PREFIX + "{bad"), /Malformed/);
  assert.throws(() => readColumnSemantics(COLUMN_SEMANTICS_PREFIX + '{"version":2}'), /Unsupported/);
});

test("human reserved-marker descriptions and titles cannot shadow column annotations or embedded JSON values", () => {
  const description = COLUMN_SEMANTICS_PREFIX + '{"version":1}\nHuman "quoted" meaning with \\ characters.';
  const title = COLUMN_SEMANTICS_PREFIX + "literal title";
  const embedded = "Ordinary JSON value\n" + COLUMN_SEMANTICS_PREFIX + '{"version":2}';
  const document = { ...metadataThingModel, properties: { ...metadataThingModel.properties,
    state: { ...metadataThingModel.properties.state, description },
    message: { ...metadataThingModel.properties.message, title, default: embedded },
  } };
  const model = fromThingModel(resolveThingModel(document));
  const fields = model.fields.map(field => setFieldRole(field, "telemetry"));
  const documentation = semanticDocumentation(model, fields);
  for (const name of ["state", "message"]) {
    const field = fields.find(field => field.path[0] === name)!;
    const doc = documentation.columns[field.column], decoded = readColumnSemantics(doc)!;
    assert.equal(decoded.version, 3);
    if (decoded.version !== 3) throw new Error("Expected semantic JSON.");
    assert.ok(doc.length <= 1000 && !doc.includes("\n" + COLUMN_SEMANTICS_PREFIX));
    assert.equal(decoded.field.facts[name === "state" ? "description" : "title"], name === "state" ? description : title);
    assert.ok(!doc.split("\n").some(line => line.startsWith("@iot.meta columnSemanticsFormat")));
    assert.throws(() => readColumnSemantics(doc + '\nname: "duplicate"'), /ambiguous/);
  }
  const state = fields.find(field => field.path[0] === "state")!;
  const message = fields.find(field => field.path[0] === "message")!;
  assert.equal(state.declaredMetadata!.annotations.description, description);
  assert.equal((message.declaredMetadata!.schema as Record<string, unknown>).default, embedded);
});

test("native selected-column semantic annotation roundtrips generated quoting and verified readback without processing drift", async () => {
  const request = creationRequest(), model = fromThingModel(resolveThingModel(metadataThingModel));
  request.input = { ...request.input, model, fields: model.fields.map(field => ({
    ...setFieldRole(field, field.path[0] === "state" ? "twin" : "telemetry"),
    column: field.path[0] === "temperature" ? "temperature_output" : field.column,
  })), scope: { mode: "data", operator: "and", telemetry: "", twin: "", reference: { id: model.id! } } };
  request.coverage = { ...request.coverage, fingerprint: modelingFingerprint(request.input),
    fields: inspectFieldEvidence(request.input.fields, request.coverage.telemetry, request.coverage.twin) };
  const mock = createCreationMock(request), creator = new ModeledDataCreator();
  const withoutMetadata = request.input.fields.map(({ declaredMetadata: _metadata, ...field }) => field);
  const legacyInput = { ...request.input, fields: withoutMetadata, model: { ...model, fields: withoutMetadata } };
  const legacy = buildModeledDataPlan(legacyInput, { ...request.coverage, fingerprint: modelingFingerprint(legacyInput) },
    request.allowWaiting, request.names);
  const processing = (commands: typeof mock.plan.commands) => commands.filter(command =>
    ["table", "view", "policy"].includes(command.kind)).map(command => command.csl);
  assert.deepEqual(processing(mock.plan.commands), processing(legacy.commands));
  const completed = await creator.run(request, mock.service, true);
  assert.equal(completed.phase, "completed"); assert.equal(mock.writes.length, 9);
  const table = mock.tables.get(request.names.modeled)!;
  for (const field of request.input.fields) {
    const doc = table.columnDocs[field.column], decoded = readColumnSemantics(doc)!;
    assert.equal(decoded.version, 3);
    assert.ok(doc.length <= 1000);
    if (decoded.version !== 3) throw new Error("Expected semantic JSON.");
    assert.equal(decoded.field.name, field.path.join("."));
    assert.equal(decoded.model, undefined);
    assert.equal(doc, mock.plan.documentation.columns[field.column]);
    assert.ok(doc.includes(field.role === "twin" ? "last-known whole received value" : '"role":"telemetry"'));
  }
  assert.ok(table.doc.includes(metadataThingModel.id) && table.doc.includes("@iot.meta tableRole=modeled-data"));
  assert.ok(!mock.writes.some(command => /backfill=true|sourceHubId/.test(command)));
});
