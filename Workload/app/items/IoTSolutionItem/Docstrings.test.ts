// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { docstringControl, emitDocstring } from "./Docstrings.ts";
import { COLUMN_SEMANTICS_PREFIX, FIELD_SEMANTIC_INFO_PREFIX, FLAT_COLUMN_SEMANTICS, columnSemanticAnnotation, columnDocstring,
  publishedFieldFacts, readColumnSemantics } from "./FieldSemantics.ts";
import { canonicalJson, isJsonObject, MODEL_SEMANTIC_INFO_PREFIX, parseIotMeta, readModelIdentityMetadata } from "./ModelAuthoring.ts";
import { fromDtdlModel, fromThingModel, kqlString, modelingFingerprint, setFieldRole } from "./HubModeling.ts";
import { parseDtdlModel } from "./DtdlModelParser.ts";
import { resolveThingModel } from "./ThingModelParser.ts";
import { buildModeledDataPlan, helperObjectDocstrings, semanticDocumentation } from "./ModeledDataPlan.ts";
import { buildDataAgentContext } from "./DataAgentContext.ts";
import { creationFingerprint, ModeledDataCreator } from "./ModeledDataCreation.ts";
import { createCreationMock, creationRequest } from "../../../tests/creationMock.ts";
import { metadataDtdlModel, metadataThingModel } from "../../../tests/fieldSemanticFixtures.ts";

test("docstring emission bounds 0 999 1000 1001 and quoted Unicode without cutting values or identities", () => {
  for (const size of [0, 999, 1000, 1001]) {
    const doc = size ? emitDocstring([{ text: "x".repeat(size) }]) : emitDocstring([]);
    assert.ok(doc.length <= 1000);
    if (size <= 1000) assert.equal(doc.length, size);
    else assert.match(doc, /\ntruncated: true$/);
  }
  const text = ('😀 "quoted" \\ value\n\u2028\u2029' + COLUMN_SEMANTICS_PREFIX).repeat(100);
  const doc = columnDocstring({ name: "Temperature", description: text, unit: "degree Celsius",
    readOnly: false, minimum: 0, default: 0, type: "number", role: "telemetry" });
  assert.ok(doc.length <= 1000);
  const result = readColumnSemantics(doc)!;
  assert.equal(result.version, 3);
  if (result.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(result.bounded, true);
  assert.equal(result.field.facts.minimum, 0); assert.equal(result.field.facts.readOnly, false);
  assert.equal(result.field.facts.default, 0); assert.equal(result.field.facts.unit, "degree Celsius");
  assert.ok(String(result.field.facts.description).endsWith("..."));
  assert.ok([...String(result.field.facts.description)].every(point => point.length === 2 ||
    point.charCodeAt(0) < 0xd800 || point.charCodeAt(0) > 0xdfff));
  assert.equal(JSON.parse(kqlString(doc)), doc);
  assert.throws(() => emitDocstring([docstringControl("@iot.meta modelId: urn:" + "x".repeat(1000))]), /identity metadata exceeds/);
  const constrained = readColumnSemantics(columnDocstring({ name: "Code", type: "string",
    pattern: "^" + "x".repeat(2000) + "$", const: "exact".repeat(500), default: "value".repeat(500) }))!;
  assert.equal(constrained.version, 3);
  if (constrained.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(constrained.bounded, true);
  assert.equal(constrained.field.facts.pattern, undefined);
  assert.equal(constrained.field.facts.const, undefined);
  assert.equal(constrained.field.facts.default, undefined);
  assert.throws(() => emitDocstring([{ key: "default", value: Number.POSITIVE_INFINITY }]), /finite JSON/);
});

test("actual creation stores bounded Unicode human values and large nested facts without changing whole-value schemas", async () => {
  const request = creationRequest();
  const description = '😀 "quotes" \\ and newlines\n\u2028'.repeat(100);
  const document = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
    id: "urn:fixture:bounded-native", title: "Bounded native", description, properties: {
      temperature: { title: "Temperature", type: "number", description, minimum: 0, readOnly: false },
      state: { title: "Configuration", type: "object", properties: Object.fromEntries(Array.from({ length: 60 },
        (_, index) => [`declared-${index}`, { type: "number", minimum: 0, description: "Declared child" }])) },
    } };
  const model = fromThingModel(resolveThingModel(document));
  request.input = { ...request.input, model, fields: model.fields.map(field => setFieldRole(field,
    field.path[0] === "state" ? "twin" : "telemetry")) };
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const mock = createCreationMock(request), before = canonicalJson(document);
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "completed");
  const table = mock.tables.get(request.names.modeled)!;
  assert.ok(table.doc.length <= 1000 && JSON.parse(parseIotMeta(table.doc).modelSemanticInfo).truncated === true);
  assert.ok(mock.writes.find(command => command.includes("column-docstrings"))!.length > 1000,
    "The complete column-docstrings command must not be clipped to the per-payload limit.");
  for (const [column, doc] of Object.entries(table.columnDocs)) {
    assert.ok(doc.length <= 1000);
    assert.equal(JSON.parse(kqlString(doc)), doc);
    const parsed = readColumnSemantics(doc)!;
    assert.equal(parsed.version, 3, column);
  }
  assert.equal(readColumnSemantics(table.columnDocs.state)!.version, 3);
  assert.ok(table.columnDocs.state.startsWith(FIELD_SEMANTIC_INFO_PREFIX) && !table.columnDocs.state.includes("\n"));
  assert.equal(table.columns.find(column => column.name === "state")!.type, "dynamic");
  assert.equal(canonicalJson(model.document), before);
  const context = buildDataAgentContext({ database: "Database", table: request.names.modeled, docstring: table.doc,
    columns: table.columns.map(column => ({ ...column, docstring: table.columnDocs[column.name] })) });
  assert.ok(context.sourceInstructions.includes("truncated=true") && context.sourceInstructions.includes("Do not reconstruct missing declarations"));
});

test("declared facts publish one native-key JSON object including whole nested schemas", () => {
  const document = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
    id: "urn:fixture:facts", title: "Device", properties: {
      temperature: { title: "Temperature", description: "Current measured temperature", unit: "degree Celsius",
        type: "number", readOnly: true, minimum: 0, default: 0 },
      calibration: { title: "Calibration", type: "object", properties: {
        offset: { title: "Offset", description: "Calibration offset", type: "number", unit: "degC", minimum: 0, readOnly: false },
        nested: { type: "object", properties: { enabled: { type: "boolean", default: false } } },
        samples: { type: "array", items: { type: "object", properties: { value: { type: "number" } } } },
      } },
    } };
  const model = fromThingModel(resolveThingModel(document));
  const fields = model.fields.map(field => setFieldRole(field, "telemetry"));
  const docs = semanticDocumentation(model, fields);
  const temperature = docs.columns.temperature;
  assert.deepEqual(JSON.parse(temperature.slice(FIELD_SEMANTIC_INFO_PREFIX.length)), {
    name: "temperature", title: "Temperature", description: "Current measured temperature", unit: "degree Celsius",
    role: "telemetry", type: "number", default: 0, readOnly: true, minimum: 0,
  });
  assert.ok(!temperature.includes("Meaning:") && !temperature.includes("Telemetry.") &&
    !temperature.includes("Declared unit") && !temperature.includes(COLUMN_SEMANTICS_PREFIX));
  const calibration = docs.columns.calibration;
  const nested = JSON.parse(calibration.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.equal(nested.properties.offset.title, "Offset");
  assert.equal(nested.properties.nested.properties.enabled.default, false);
  assert.equal(nested.properties.samples.items.type, "object");
  assert.ok(!calibration.includes("\t") && calibration.length <= 1000);
  const result = readColumnSemantics(calibration)!;
  assert.equal(result.version, 3);
  if (result.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(result.bounded, false);
  assert.equal(canonicalJson(model.document), canonicalJson(document));
  assert.deepEqual(fields.map(field => [field.column, field.kustoType]), [["temperature", "real"], ["calibration", "dynamic"]]);
  assert.equal(JSON.parse(kqlString(calibration)), calibration);
  const untyped = fromThingModel(resolveThingModel({ ...document, properties: {
    reading: { "@type": "urn:quantity:Temperature", readOnly: true },
  } }));
  const unspecified = semanticDocumentation(untyped, untyped.fields.map(field => setFieldRole(field, "telemetry"))).columns.reading;
  const semantic = JSON.parse(unspecified.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.equal(semantic.type, undefined); assert.equal(semantic.unit, undefined);
  assert.equal(semantic["@type"], "urn:quantity:Temperature");
});

test("large nested declared metadata is honestly bounded with no partial keys values or orphan hierarchy", () => {
  const fields = Object.fromEntries(Array.from({ length: 100 }, (_, index) => [`child:${index} 😀`, {
    title: `Child ${index}`, description: "Long declared meaning ".repeat(20), type: "number", minimum: 0, readOnly: false,
  }]));
  const doc = columnDocstring({ name: "Configuration", role: "twin property (last-known whole received value)",
    type: "object", properties: fields, enum: Array.from({ length: 200 }, (_, index) => `value-${index}`) });
  assert.ok(doc.length <= 1000 && doc.startsWith(FIELD_SEMANTIC_INFO_PREFIX));
  const result = readColumnSemantics(doc)!;
  assert.equal(result.version, 3);
  if (result.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(result.bounded, true);
  assert.equal(result.field.facts.type, "object");
  assert.equal(result.field.facts.properties, undefined, "Oversized child collections must not become a misleading partial schema.");
  assert.equal(JSON.parse(kqlString(doc)), doc);
});

test("all modeled and helper objects use selected native model semantic JSON and bounded stored payloads", async () => {
  const request = creationRequest();
  assert.ok(isJsonObject(request.input.model.document));
  const document = { ...request.input.model.document, description: 'Device "meaning"\nSecond line', version: { model: "actual-v1" } };
  request.input.model = fromThingModel(resolveThingModel(document));
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const plan = buildModeledDataPlan(request.input, request.coverage, request.allowWaiting, request.names);
  const cluster = "mock.kusto.fabric.microsoft.com/Database";
  assert.equal(plan.documentation.table, [
    "Modeled device telemetry enriched with property LKVs",
    MODEL_SEMANTIC_INFO_PREFIX + '{"id":"urn:mock:device","title":"Device","description":"Device \\"meaning\\"\\nSecond line","version":"actual-v1"}',
    "@iot.meta tableRole=modeled-data",
    `@iot.meta source=${cluster}/Telemetry; ${cluster}/State_Output`,
  ].join("\n"));
  const helpers = helperObjectDocstrings(request.input.model, request.input.source, request.names.twinChanges);
  const semanticLine = plan.documentation.table.split("\n")[1];
  assert.equal(helpers.normalized, `Normalized device property changes in Entity-Attribute-Value (EAV) format\n${semanticLine}\n@iot.meta tableRole=properties-eav\n@iot.meta source=${cluster}/Twins`);
  assert.equal(helpers.lkv, `Property last known values\n${semanticLine}\n@iot.meta viewRole=properties-lkv\n@iot.meta source=${cluster}/Changes_Output`);
  const tags = parseIotMeta(plan.documentation.table);
  assert.equal(tags.deviceModel, undefined); assert.equal(tags.modelVersion, undefined);
  assert.equal(JSON.parse(tags.modelSemanticInfo).id, "urn:mock:device");
  assert.deepEqual(readModelIdentityMetadata(plan.documentation.table), { id: "urn:mock:device", documentVersion: "actual-v1" });
  assert.equal(plan.documentation.columns.deviceId, FIELD_SEMANTIC_INFO_PREFIX + '{"name":"deviceId","description":"IoT Hub device Id"}');
  for (const command of plan.commands) if (command.kind === "viewDoc" || command.kind === "tableDoc") {
    assert.equal(JSON.parse(parseIotMeta(command.docstring).modelSemanticInfo).version, "actual-v1");
    assert.ok(!command.docstring.includes("@iot.meta modelVersion=") && !command.docstring.includes("@iot.meta deviceModel="));
  }
  assert.ok(Object.values(plan.documentation.columns).every(doc => !doc.split("\n").some(line => line.startsWith("@iot.meta columnSemanticsFormat")) &&
    !doc.includes("\nDeclared type:") && !doc.includes("\tDeclared type:")));
  assert.equal(tags.source, `${cluster}/Telemetry; ${cluster}/State_Output`);
  const mock = createCreationMock(request);
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "completed");
  for (const table of mock.tables.values()) {
    assert.ok(table.doc.length <= 1000);
    for (const doc of Object.values(table.columnDocs)) assert.ok(doc.length <= 1000);
  }
  for (const view of mock.views.values()) assert.ok(view.doc.length <= 1000);
  assert.equal(mock.tables.get(request.names.modeled)!.doc, plan.documentation.table);
  for (const [name, doc] of Object.entries(plan.documentation.columns)) assert.equal(mock.tables.get(request.names.modeled)!.columnDocs[name], doc);
  const context = buildDataAgentContext({ database: "Database", table: request.names.modeled, docstring: plan.documentation.table,
    columns: [{ name: "event_time", type: "datetime", docstring: plan.documentation.columns.event_time },
      ...request.input.fields.map(field => ({ name: field.column, type: field.kustoType, docstring: plan.documentation.columns[field.column] }))] });
  assert.ok(context.sourceInstructions.includes("urn:mock:device") && context.sourceInstructions.includes("actual-v1"));
  assert.notEqual(JSON.parse(creationFingerprint(request))[3], "object-docstrings-v2");
  assert.notEqual(JSON.parse(creationFingerprint(request))[3], "bounded-declared-facts-v3");
});

test("new fact readers retain valid legacy envelopes and reject duplicate malformed or injected control facts", () => {
  const request = creationRequest(), field = request.input.fields[0];
  const legacy = columnSemanticAnnotation(request.input.model, field.path.join("."), field.declaredMetadata!);
  assert.equal(readColumnSemantics(legacy)!.version, 1);
  const doc = columnDocstring(publishedFieldFacts(field));
  assert.equal(readColumnSemantics(doc)!.version, 3);
  const human = 'Declared type: literal words; @iot.meta columnSemanticsFormat=facts-v2';
  const literal = readColumnSemantics(columnDocstring({ name: human, description: human, type: "string" }))!;
  assert.equal(literal.version, 3);
  if (literal.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(literal.field.facts.name, human); assert.equal(literal.field.facts.description, human);
  assert.throws(() => readColumnSemantics(doc + "\n" + FLAT_COLUMN_SEMANTICS), /ambiguous/);
  assert.equal(readColumnSemantics('name: "Old label"\ntype: "number"')!.version, 2);
  const old = 'name: "Old label"\nDeclared type: "number"\nproperties:\n\t"child":\n\t\tname: "Child"\n\t\tDeclared type: "boolean"\n' + FLAT_COLUMN_SEMANTICS;
  const oldFacts = readColumnSemantics(old)!;
  assert.equal(oldFacts.version, 2);
  if (oldFacts.version !== 2) throw new Error("Expected plain facts.");
  assert.equal(oldFacts.field.facts.type, "number");
  assert.throws(() => readColumnSemantics(doc + "\n" + FLAT_COLUMN_SEMANTICS + "\n" + FLAT_COLUMN_SEMANTICS), /Malformed/);
  assert.throws(() => readColumnSemantics('name: "X"\ntype: "number"\nDeclared type: "number"'), /duplicate/);
  assert.throws(() => readColumnSemantics('name: "X"\nproperties:\n\t"child":\n\t\ttype: "number"\n\t\tDeclared type: "number"'), /duplicate/);
  assert.throws(() => columnDocstring({ name: "X", type: "number", "Declared type": "number" }), /aliases/);
  assert.throws(() => readColumnSemantics(doc + '\nname: "duplicate"'), /ambiguous/);
  assert.throws(() => readColumnSemantics('name: "X"\nreadOnly: "yes"\n' + FLAT_COLUMN_SEMANTICS), /Malformed/);
  assert.throws(() => readColumnSemantics('name: "X"\n\t\tminimum: 0\n' + FLAT_COLUMN_SEMANTICS), /indentation/);
  assert.deepEqual(parseIotMeta('description: "quoted \\n@iot.meta modelId: bad"\n@iot.meta modelId: urn:good'),
    { modelId: "urn:good" });
  assert.throws(() => parseIotMeta("@iot.meta modelId: one\n@iot.meta modelId: two"), /Duplicate/);
  assert.throws(() => parseIotMeta("@iot.meta modelId"), /Malformed/);
  assert.deepEqual(readModelIdentityMetadata("@iot.meta modelId: urn:legacy\n@iot.meta modelVersion: old"), { id: "urn:legacy", documentVersion: "old" });
  assert.deepEqual(readModelIdentityMetadata("@iot.meta tableRole=modeled-data\n@iot.meta deviceModel=Legacy_Name"), {});
  assert.throws(() => readModelIdentityMetadata("@iot.meta tableRole=modeled-data\n@iot.meta deviceModel=urn:new\n@iot.meta modelId: urn:old"), /aliases/);
  assert.throws(() => readModelIdentityMetadata("Modeled device telemetry enriched with property LKVs\n@iot.meta tableRole=modeled-data\n@iot.meta deviceModel=not-an-id"), /Malformed/);
});

test("native DTDL semantic JSON retains inheritance components localized declarations units writable and complex schemas", () => {
  const parsed = parseDtdlModel(metadataDtdlModel), model = fromDtdlModel(parsed, metadataDtdlModel);
  const docs = semanticDocumentation(model, model.fields);
  const read = (column: string) => JSON.parse(docs.columns[column].slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  const temperature = read("temperature");
  assert.deepEqual(temperature["@type"], ["Telemetry", "Temperature"]);
  assert.deepEqual(temperature.displayName, { en: "Temperature", de: "Temperatur" });
  assert.equal(temperature.description, "Current measured temperature");
  assert.equal(temperature.schema, "double"); assert.equal(temperature.unit, "degreeCelsius");
  assert.equal(temperature.name, "temperature"); assert.equal(temperature.role, "telemetry");
  assert.equal(temperature.type, undefined); assert.equal(temperature.semanticTypes, undefined);
  assert.equal(temperature["@id"], "dtmi:fixture:SemanticBase:temperature;1");
  assert.equal(parsed.telemetries.find(field => field.leafName === "temperature")!.declaredBy, "dtmi:fixture:SemanticBase;1");
  const configuration = read("motor_configuration");
  assert.equal(configuration.component, "motor"); assert.equal(configuration.name, "configuration");
  assert.equal(configuration.writable, false); assert.equal(configuration.schema["@type"], "Object");
  assert.deepEqual(configuration.schema.fields[0], { name: "offset", schema: "double", displayName: "Offset", description: "Calibration offset" });
  assert.equal(configuration.role, "twin property (last-known whole received value)");
  const mode = read("mode");
  assert.equal(mode.writable, true); assert.equal(mode.schema.valueSchema, "integer");
  assert.equal(mode.schema.enumValues[0].enumValue, 0); assert.equal(mode.schema.enumValues[0].displayName, "Idle");
  assert.equal(mode.schema.enumValues[0].name, "idle"); assert.equal(mode.codes, undefined);
  assert.equal(read("samples").schema.elementSchema.fields[0].schema, "double");
  assert.equal(read("labels").schema.mapKey.schema, "string");
  assert.equal(read("labels").schema.mapValue.name, "value");
  assert.ok(!JSON.stringify(docs.columns).includes("cloudOnly") && !JSON.stringify(docs.columns).includes("reboot"));
});

test("IoT Central DTDL upload provenance never turns UI fallback labels into declared displayName or datatype", () => {
  const document = { "@type": ["ModelDefinition", "DeviceModel"], capabilityModel: {
    "@id": "dtmi:fixture:Central;1", "@type": "Interface", contents: [
      { "@type": ["Telemetry", "Temperature"], name: "reading", schema: "double" },
      { "@type": "Property", name: "state", description: "", displayName: "", writable: false, schema: "boolean" },
    ],
  } };
  const model = fromDtdlModel(parseDtdlModel(document), document), docs = semanticDocumentation(model, model.fields);
  const reading = JSON.parse(docs.columns.reading.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.equal(model.fields.find(field => field.path[0] === "reading")!.title, "reading");
  assert.equal(reading.displayName, undefined); assert.equal(reading.title, undefined);
  assert.equal(reading.unit, undefined); assert.equal(reading.type, undefined);
  assert.equal(reading.schema, "double"); assert.deepEqual(reading["@type"], ["Telemetry", "Temperature"]);
  const state = JSON.parse(docs.columns.state.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.equal(state.description, ""); assert.equal(state.displayName, ""); assert.equal(state.writable, false);
  const parsed = parseDtdlModel(document);
  const missing = { ...parsed, telemetries: parsed.telemetries.map(({ declaredSchema: _schema, declaredAnnotations: _annotations, ...field }) => field) };
  const legacy = fromDtdlModel(missing, document);
  const absent = JSON.parse(semanticDocumentation(legacy, legacy.fields).columns.reading.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.equal(absent.schema, undefined); assert.equal(absent.displayName, undefined);
  assert.deepEqual(absent["@type"], ["Temperature"]);
});

test("WoT semantic JSON keeps native event data and declared quantity separate from selected processing role", () => {
  const model = fromThingModel(resolveThingModel(metadataThingModel));
  const fields = model.fields.map(field => setFieldRole(field, "telemetry"));
  const docs = semanticDocumentation(model, fields);
  const alarm = JSON.parse(docs.columns.alarm.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.equal(alarm["@type"], "ex:Alarm");
  assert.equal(alarm.data["@type"], "ex:AlarmData"); assert.equal(alarm.data.type, "object");
  assert.deepEqual(alarm.data.properties.code.enum, [0, 1]); assert.equal(alarm.role, "telemetry");
  const temperature = JSON.parse(docs.columns.temperature.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.equal(temperature["@type"], "ex:Temperature"); assert.equal(temperature.readOnly, false);
  assert.equal(temperature.minimum, 0); assert.equal(temperature.default, 0); assert.equal(temperature.title, "Temperature");
  assert.ok(!docs.columns.temperature.includes("fieldSemanticsFormat") && !docs.columns.temperature.includes("Meaning:"));
});

test("semantic JSON budgets include prefix punctuation escaping and honest atomic omission at 999 1000 1001", () => {
  const overhead = (FIELD_SEMANTIC_INFO_PREFIX + JSON.stringify({ name: "x", description: "" })).length;
  for (const size of [999, 1000, 1001]) {
    const doc = columnDocstring({ name: "x", description: "a".repeat(size - overhead) });
    assert.ok(doc.length <= 1000);
    if (size <= 1000) assert.equal(doc.length, size);
    const parsed = readColumnSemantics(doc)!;
    assert.equal(parsed.version, 3);
    if (parsed.version !== 3) throw new Error("Expected semantic JSON.");
    assert.equal(parsed.bounded, size > 1000);
  }
  const doc = columnDocstring({ name: "limits", role: "telemetry", displayName: { en: "😀 title ".repeat(100), de: "Quoted \"title\" ".repeat(100) },
    description: "", minimum: 0, writable: false, default: null, const: "", enum: Array.from({ length: 300 }, (_, index) => index) });
  const json = JSON.parse(doc.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
  assert.ok(doc.length <= 1000); assert.equal(json.truncated, true); assert.equal(json.enum, undefined);
  assert.equal(json.default, null); assert.equal(json.const, ""); assert.equal(json.description, "");
  assert.equal(json.writable, false); assert.equal(json.minimum, 0);
  assert.ok(typeof json.displayName.en === "string" && typeof json.displayName.de === "string");
  assert.equal(JSON.parse(kqlString(doc)), doc);
  assert.throws(() => columnDocstring({ name: "x".repeat(1000), role: "telemetry" }), /identity\/role metadata exceeds/);
});

test("semantic JSON rejects ambiguous duplicate malformed and reserved control forms while preserving literal words", () => {
  const literal = 'Declared type: @iot.meta fieldSemanticInfo=literal\n@iot.meta columnSemanticsFormat=facts-v2';
  const doc = columnDocstring({ name: "literal", description: literal, type: "object",
    properties: { [literal]: { type: "string", default: literal } } });
  const decoded = readColumnSemantics(doc)!;
  assert.equal(decoded.version, 3);
  if (decoded.version !== 3) throw new Error("Expected semantic JSON.");
  assert.equal(decoded.field.facts.description, literal);
  assert.ok(!doc.includes("\n"));
  for (const text of ['[]', 'null', '{"name":"X","name":"Y"}', '{"name":"X","properties":{"a":{"type":"string","type":"number"}}}',
    '{"name":"X","role":["telemetry"]}', '{"name":"X","readOnly":"yes"}', '{"name":"X","type":"number","schema":"double"}',
    '{"name":"X","@type":"Temperature","semanticTypes":["Temperature"]}', '{"name":"X","forms":[{"href":"not-analytic"}]}']) {
    assert.throws(() => readColumnSemantics(FIELD_SEMANTIC_INFO_PREFIX + text), /semantic|Semantic/);
  }
  assert.throws(() => readColumnSemantics("@iot.meta fieldSemanticInfo: {}"), /Malformed/);
  assert.throws(() => readColumnSemantics(doc + "\n" + doc), /ambiguous/);
  assert.throws(() => readColumnSemantics('name: "Old"\n' + doc), /ambiguous/);
  assert.throws(() => readColumnSemantics(COLUMN_SEMANTICS_PREFIX + '{"version":1}\n' + doc), /ambiguous/);
});

test("both source families persist exact current JSON plans and consumer roles without crediting old native metadata", async () => {
  for (const kind of ["wot", "dtdl"] as const) {
    const request = creationRequest();
    const model = kind === "wot" ? fromThingModel(resolveThingModel(metadataThingModel))
      : fromDtdlModel(parseDtdlModel(metadataDtdlModel), metadataDtdlModel);
    request.input = { ...request.input, model, fields: model.fields.map(field => field.role ? field : setFieldRole(field, "telemetry")) };
    request.coverage.fingerprint = modelingFingerprint(request.input);
    const mock = createCreationMock(request), before = canonicalJson(model.document);
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "completed");
    const table = mock.tables.get(request.names.modeled)!;
    assert.deepEqual(table.columnDocs, mock.plan.documentation.columns);
    assert.ok(Object.values(table.columnDocs).every(doc => doc.startsWith(FIELD_SEMANTIC_INFO_PREFIX) && doc.length <= 1000));
    const context = buildDataAgentContext({ database: "Database", table: request.names.modeled, docstring: table.doc,
      columns: table.columns.map(column => ({ ...column, docstring: table.columnDocs[column.name] })) });
    assert.ok(context.examples.some(example => example.query.includes("avg(")));
    assert.equal(canonicalJson(model.document), before);
    const old = JSON.parse(creationFingerprint(request)); old[3] = "plain-declared-facts-v4";
    assert.notEqual(creationFingerprint(request), JSON.stringify(old));
    const changed = createCreationMock(request), original = changed.service.request;
    changed.service.request = async (csl, operation) => {
      const value = await original(csl, operation);
      if (csl.includes(" column-docstrings ")) {
        changed.tables.get(request.names.modeled)!.columnDocs[request.input.fields[0].column] = 'name: "Legacy value"\nrole: "telemetry"';
      }
      return value;
    };
    const rejected = await new ModeledDataCreator().run(request, changed.service, true);
    assert.equal(rejected.phase, "failed");
    assert.match(rejected.error!, /documentation could not be verified/);
  }
});
