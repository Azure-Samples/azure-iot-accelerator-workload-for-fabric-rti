// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { MODEL_SEMANTIC_INFO_PREFIX, canonicalJson, filterTablesByRole, modelMeaningDocstring, parseIotMeta,
  readModelIdentityMetadata, type JsonObject } from "./ModelAuthoring.ts";
import { emitJsonDocstring } from "./Docstrings.ts";
import { fromDtdlModel, fromThingModel, kqlString, modelingFingerprint, setFieldRole, type HubModel } from "./HubModeling.ts";
import { resolveThingModel } from "./ThingModelParser.ts";
import { parseDtdlModel } from "./DtdlModelParser.ts";
import { buildModeledDataPlan, helperObjectDocstrings, semanticDocumentation } from "./ModeledDataPlan.ts";
import { buildDataAgentContext } from "./DataAgentContext.ts";
import { creationFingerprint, ModeledDataCreator } from "./ModeledDataCreation.ts";
import { FIELD_SEMANTIC_INFO_PREFIX, readColumnSemantics } from "./FieldSemantics.ts";
import { createCreationMock, creationRequest } from "../../../tests/creationMock.ts";
import { metadataDtdlModel } from "../../../tests/fieldSemanticFixtures.ts";

const opening = "Modeled device telemetry enriched with property LKVs";
const tm = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
  id: "urn:fixture:ThermoProX200", title: "ThermoPro X200", description: "Thermostat model", version: { model: "1.0" },
  properties: { temperature: { title: "Temperature", type: "number", unit: "degC" }, state: { type: "string" } } };
function requestFor(model: HubModel) {
  const request = creationRequest();
  request.input = { ...request.input, model, fields: model.fields.map(field => field.role ? field
    : setFieldRole(field, field.path[0] === "state" ? "twin" : "telemetry")) };
  request.coverage.fingerprint = modelingFingerprint(request.input);
  return request;
}
const json = (doc: string): JsonObject => JSON.parse(parseIotMeta(doc).modelSemanticInfo);
const modeledDoc = (model: HubModel) => {
  const request = requestFor(model);
  return buildModeledDataPlan(request.input, request.coverage, true, request.names).documentation.table;
};

test("model object docs use exactly agreed WoT JSON and unchanged opening role/source lines", () => {
  const model = fromThingModel(resolveThingModel(tm)), doc = modeledDoc(model);
  assert.equal(doc, [
    opening,
    MODEL_SEMANTIC_INFO_PREFIX + '{"id":"urn:fixture:ThermoProX200","title":"ThermoPro X200","description":"Thermostat model","version":"1.0"}',
    "@iot.meta tableRole=modeled-data",
    "@iot.meta source=mock.kusto.fabric.microsoft.com/Database/Telemetry; mock.kusto.fabric.microsoft.com/Database/State_Output",
  ].join("\n"));
  assert.deepEqual(readModelIdentityMetadata(doc), { id: tm.id, documentVersion: "1.0" });
  assert.deepEqual(filterTablesByRole(["Raw", "Modeled"], { Raw: "@iot.meta tableRole=iothub-telemetry-raw", Modeled: doc }, "modeled-data"), ["Modeled"]);
  assert.ok(!/\n(?:title:|description:|@iot.meta (?:deviceModel|modelVersion|modelId|tableSemanticInfo)=)/.test(doc));
});

test("declared WoT localizations and scalar version come from chosen root rather than UI or inherited fallback", () => {
  const base = { ...tm, id: "urn:fixture:parent", title: "Parent title", description: "Parent description" };
  const child = { "@context": tm["@context"], "@type": tm["@type"], id: "urn:fixture:child",
    links: [{ rel: "tm:extends", href: base.id }], version: { model: "child-version" } };
  const resolved = fromThingModel(resolveThingModel([base, child]));
  assert.notEqual(resolved.title, child.id, "The fixture must exercise an inherited UI label.");
  assert.deepEqual(json(modeledDoc(resolved)), { id: child.id, version: "child-version" });
  const translated = { ...tm, titles: { de: "Gerät", en: "English title", "en-US": "US title" },
    descriptions: { de: "Beschreibung", en: "English meaning" } };
  const doc = modeledDoc(fromThingModel(resolveThingModel(translated)));
  assert.deepEqual(json(doc).titles, translated.titles); assert.deepEqual(json(doc).descriptions, translated.descriptions);
  const meaning = modelMeaningDocstring(doc);
  assert.ok(meaning.includes('title: "US title"') && meaning.includes('description: "English meaning"'));
  const otherLanguage = modelMeaningDocstring(modeledDoc(fromThingModel(resolveThingModel({ ...tm, titles: { de: "Gerät" } }))));
  assert.ok(otherLanguage.includes('title: "ThermoPro X200"') && !otherLanguage.includes('title: "Gerät"'),
    "Compatibility must keep the existing WoT preferred-language/default-title behavior.");
  const { title: _title, description: _description, version: _version, ...absent } = tm;
  assert.deepEqual(json(modeledDoc(fromThingModel(resolveThingModel(absent)))), { id: tm.id });
});

test("DTDL arrays inheritance component and IoT Central models retain chosen-root native @id and localized declarations", () => {
  const root = { ...metadataDtdlModel[2], displayName: { en: "Root thermostat", fr: "Thermostat racine" },
    description: { en: "Root model meaning", fr: "Sens racine" } };
  const documents = [metadataDtdlModel[0], metadataDtdlModel[1], root];
  const centralRoot = { ...root, extends: metadataDtdlModel[0],
    contents: root.contents!.map(capability => capability["@type"] === "Component" ? { ...capability, schema: metadataDtdlModel[1] } : capability) };
  for (const document of [documents, { displayName: "WRAPPER IS NOT MODEL TITLE", capabilityModel: centralRoot }]) {
    const parsed = parseDtdlModel(document), model = fromDtdlModel(parsed, document), doc = modeledDoc(model), info = json(doc);
    assert.deepEqual(info, { "@id": root["@id"], displayName: root.displayName, description: root.description });
    assert.equal(info.id, undefined); assert.equal(info.version, undefined); assert.equal(info.title, undefined);
    assert.deepEqual(readModelIdentityMetadata(doc), { id: root["@id"], documentVersion: "2" });
    assert.ok(modelMeaningDocstring(doc).includes('title: "Root thermostat"'));
    assert.ok(!doc.includes("WRAPPER") && !doc.includes("contents") && !doc.includes("schemas"));
  }
  const { displayName: _title, description: _description, ...bare } = root;
  const absentDocument = [metadataDtdlModel[0], metadataDtdlModel[1], bare];
  const absentModel = fromDtdlModel(parseDtdlModel(absentDocument), absentDocument);
  assert.equal(absentModel.semanticTitle, undefined);
  assert.deepEqual(json(modeledDoc(absentModel)), { "@id": root["@id"] });
});

test("aggregate model docstring budget preserves complete JSON identity role/source at 999 1000 1001 and unicode boundaries", () => {
  const model = fromThingModel(resolveThingModel(tm));
  const short = { ...model, declaredModelMetadata: { id: tm.id, description: "" } };
  const overhead = modeledDoc(short).length;
  for (const length of [999, 1000, 1001]) {
    const doc = modeledDoc({ ...short, declaredModelMetadata: { id: tm.id, description: "a".repeat(length - overhead) } });
    assert.ok(doc.length <= 1000);
    if (length <= 1000) assert.equal(doc.length, length);
    assert.equal(json(doc).truncated === true, length > 1000);
    assert.equal(json(doc).id, tm.id);
    assert.ok(doc.endsWith("@iot.meta source=mock.kusto.fabric.microsoft.com/Database/Telemetry; mock.kusto.fabric.microsoft.com/Database/State_Output"));
    assert.equal(JSON.parse(kqlString(doc)), doc);
  }
  const prose = '😀 "quotes" \\ newline\n\u2028\u2029'.repeat(100);
  const doc = modeledDoc({ ...model, declaredModelMetadata: { id: tm.id, titles: { en: prose, de: prose }, description: prose, version: "0" } });
  assert.ok(doc.length <= 1000); assert.equal(json(doc).truncated, true); assert.equal(json(doc).version, "0");
  assert.ok(!doc.includes("\ntruncated:") && !doc.includes("\\ud83d\""));
  assert.equal(JSON.parse(JSON.stringify({ csl: kqlString(doc) })).csl, kqlString(doc));
  const escaped = 'exact @iot.meta modelSemanticInfo=literal\\n@iot.meta modelId: not-a-tag';
  const literalDoc = modeledDoc({ ...model, declaredModelMetadata: { id: tm.id, title: escaped, description: escaped } });
  assert.equal(json(literalDoc).description, escaped);
  assert.equal(readModelIdentityMetadata(literalDoc).id, tm.id);
  assert.deepEqual(json(modeledDoc({ ...model, declaredModelMetadata: { id: tm.id, title: "", description: "", version: "" } })),
    { id: tm.id, title: "", description: "", version: "" });
});

test("impossible aggregate identity/source/version budgets fail before any creation mutation", async () => {
  const model = fromThingModel(resolveThingModel(tm)), longId = "urn:fixture:" + "i".repeat(1000);
  assert.throws(() => modeledDoc({ ...model, id: longId, document: { ...tm, id: longId }, declaredModelMetadata: { id: longId } }),
    /Required model semantic identity\/role metadata exceeds/);
  assert.throws(() => modeledDoc({ ...model, declaredModelMetadata: { id: tm.id, version: "v".repeat(1000) } }),
    /Required model semantic identity\/role metadata exceeds/);
  const request = requestFor(model); request.input.source.databaseName = "D".repeat(1000);
  request.coverage.fingerprint = modelingFingerprint(request.input);
  let writes = 0;
  const result = await new ModeledDataCreator().run(request, {
    inspect: async () => request.coverage, request: async () => { writes++; throw new Error("No database call is allowed."); },
  }, true);
  assert.equal(result.phase, "blocked"); assert.equal(writes, 0); assert.match(result.error!, /Required model semantic/);
  assert.throws(() => semanticDocumentation({ ...model, id: "urn:changed" }, model.fields), /selected root/);
});

test("new table parser rejects duplicate malformed reserved and mixed aliases while valid marker-like literals stay data", () => {
  const tag = (text: string) => opening + "\n" + MODEL_SEMANTIC_INFO_PREFIX + text + "\n@iot.meta tableRole=modeled-data";
  for (const value of ['[]', 'null', '{"id":"urn:x","id":"urn:y"}',
    '{"id":"urn:x","titles":{"en":"one","en":"two"}}', '{"id":"bad id"}', '{"id":"urn:x","@id":"dtmi:x;1"}',
    '{"id":"urn:x","displayName":"Alias"}', '{"@id":"dtmi:x;1","title":"Alias"}',
    '{"@id":"dtmi:x;1","version":"1"}', '{"id":"urn:x","version":{"model":"1"}}',
    '{"id":"urn:x","title":22}', '{"id":"urn:x","description":{"en":22}}', '{"id":"urn:x","truncated":false}',
    '{"id":"urn:x","contents":[]}']) assert.throws(() => parseIotMeta(tag(value)), /model|Model/);
  const valid = tag('{"id":"urn:x","title":"@iot.meta modelVersion=not-a-tag","description":"line\\\\n@iot.meta tableRole=not-a-role"}');
  assert.equal(parseIotMeta(valid).tableRole, "modeled-data"); assert.equal(readModelIdentityMetadata(valid).id, "urn:x");
  for (const suffix of ["@iot.meta modelId: urn:x", "@iot.meta deviceModel=urn:x", "@iot.meta modelVersion=1", 'title: "Duplicate"']) {
    assert.throws(() => parseIotMeta(valid + "\n" + suffix), /aliases/);
  }
  assert.throws(() => parseIotMeta(valid + "\n" + MODEL_SEMANTIC_INFO_PREFIX + "{}"), /Duplicate/);
  assert.throws(() => parseIotMeta("@iot.meta modelSemanticInfo: {}"), /Malformed/);
  assert.throws(() => parseIotMeta("@iot.meta modelSemanticInfoOther={}"), /Malformed/);
  assert.throws(() => parseIotMeta("@iot.meta modelSemanticInfo"), /Malformed/);
  assert.equal(readColumnSemantics(valid), undefined, "Table tags must not become fabricated column semantics.");
});

test("legacy table identities still read and equivalent model metadata preserves existing agent instructions and examples", () => {
  const request = requestFor(fromThingModel(resolveThingModel(tm)));
  const plan = buildModeledDataPlan(request.input, request.coverage, true, request.names);
  const old = [opening, 'title: "ThermoPro X200"', 'description: "Thermostat model"', `@iot.meta deviceModel=${tm.id}`,
    "@iot.meta modelVersion=1.0", "@iot.meta tableRole=modeled-data", parseIotMeta(plan.documentation.table).source
      ? "@iot.meta source=" + parseIotMeta(plan.documentation.table).source : ""].join("\n");
  const source = { database: "Database", table: request.names.modeled,
    columns: [{ name: "deviceId", type: "string", docstring: plan.documentation.columns.deviceId },
      { name: "event_time", type: "datetime", docstring: plan.documentation.columns.event_time },
      { name: "temperature", type: "real", docstring: plan.documentation.columns.temperature }] };
  assert.deepEqual(buildDataAgentContext({ ...source, docstring: plan.documentation.table }), buildDataAgentContext({ ...source, docstring: old }));
  assert.deepEqual(readModelIdentityMetadata("@iot.meta modelId: urn:legacy\n@iot.meta modelVersion: 0"), { id: "urn:legacy", documentVersion: "0" });
  assert.deepEqual(readModelIdentityMetadata("@iot.meta tableRole=modeled-data\n@iot.meta deviceModel=urn:legacy\n@iot.meta modelVersion=1"),
    { id: "urn:legacy", documentVersion: "1" });
  assert.deepEqual(readModelIdentityMetadata("@iot.meta tableRole=modeled-data\n@iot.meta deviceModel=Old_Model"), {});
  assert.equal(modelMeaningDocstring(old), old);
  const dtdlLocalized = MODEL_SEMANTIC_INFO_PREFIX + JSON.stringify({
    "@id": "dtmi:fixture:Local;1", displayName: { fr: "Première traduction", "en-US": "Later translation" },
  });
  assert.equal(modelMeaningDocstring(dtdlLocalized), 'title: "Première traduction"',
    "DTDl compatibility must retain its existing first-declared localization fallback.");
});

test("both real plans persist exact modeled object JSON and reject previous object-doc readback without changing column JSON", async () => {
  for (const model of [fromThingModel(resolveThingModel(tm)), fromDtdlModel(parseDtdlModel(metadataDtdlModel), metadataDtdlModel)]) {
    const request = requestFor(model), mock = createCreationMock(request), plan = mock.plan;
    const before = canonicalJson(model.document), result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "completed");
    const table = mock.tables.get(request.names.modeled)!;
    assert.equal(table.doc, plan.documentation.table);
    assert.deepEqual(table.columnDocs, plan.documentation.columns);
    assert.ok(Object.values(table.columnDocs).every(doc => doc.startsWith(FIELD_SEMANTIC_INFO_PREFIX)));
    assert.equal(mock.writes.length, 9); assert.equal(canonicalJson(model.document), before);
    const old = JSON.parse(creationFingerprint(request)); old[3] = "field-semantic-info-json-v5";
    assert.notEqual(creationFingerprint(request), JSON.stringify(old));
    const altered = createCreationMock(request), original = altered.service.request;
    altered.service.request = async (csl, kind) => {
      const value = await original(csl, kind);
      if (csl.startsWith(`.alter table ['${request.names.modeled}'] docstring `)) {
        altered.tables.get(request.names.modeled)!.doc = opening + "\n@iot.meta tableRole=modeled-data\n@iot.meta deviceModel=" + model.id;
      }
      return value;
    };
    const failed = await new ModeledDataCreator().run(request, altered.service, true);
    assert.equal(failed.phase, "failed"); assert.match(failed.error!, /Object docstring/);
  }
  const surrounded = emitJsonDocstring(MODEL_SEMANTIC_INFO_PREFIX, { id: "urn:x" },
    { before: [opening], after: ["@iot.meta tableRole=modeled-data"], requiredKeys: ["id"] });
  assert.deepEqual(readModelIdentityMetadata(surrounded), { id: "urn:x" });
});

test("both helpers publish the same actual selected WoT DTDL or IoT Central model rather than a sanitized UI identity", async () => {
  const central = { displayName: "Not the model title", capabilityModel: {
    "@id": "dtmi:fixture:CentralHelper;1", "@type": "Interface", displayName: { en: "Declared central device", de: "Gerät" },
    description: { en: "Declared meaning" }, contents: [
      { "@type": "Telemetry", name: "temperature", schema: "double" }, { "@type": "Property", name: "state", schema: "string" },
    ],
  } };
  const wot = fromThingModel(resolveThingModel({ ...tm, titles: { en: "Declared device" }, descriptions: { en: "Meaning" } }));
  const models = [wot, fromDtdlModel(parseDtdlModel(metadataDtdlModel), metadataDtdlModel),
    fromDtdlModel(parseDtdlModel(central), central)];
  for (const original of models) {
    const model = { ...original, title: "SANITIZED_UI_NAME_IS_NOT_ID" }, request = requestFor(model), mock = createCreationMock(request);
    const info = json(mock.plan.documentation.table), docs = helperObjectDocstrings(model, request.input.source, request.names.twinChanges);
    assert.deepEqual(json(docs.normalized), info); assert.deepEqual(json(docs.lkv), info);
    assert.equal(parseIotMeta(docs.normalized).tableRole, "properties-eav");
    assert.equal(parseIotMeta(docs.lkv).viewRole, "properties-lkv");
    assert.equal(parseIotMeta(docs.normalized).source, "mock.kusto.fabric.microsoft.com/Database/Twins");
    assert.equal(parseIotMeta(docs.lkv).source, "mock.kusto.fabric.microsoft.com/Database/Changes_Output");
    assert.ok(docs.normalized.startsWith("Normalized device property changes in Entity-Attribute-Value (EAV) format\n"));
    assert.ok(docs.lkv.startsWith("Property last known values\n"));
    for (const doc of Object.values(docs)) {
      assert.ok(doc.length <= 1000 && !doc.includes("@iot.meta deviceModel=") && !doc.includes("@iot.meta modelVersion="));
      assert.ok(!doc.includes("SANITIZED_UI_NAME_IS_NOT_ID") && !doc.includes("Not the model title"));
      assert.equal(readModelIdentityMetadata(doc).id, model.id);
      assert.equal(JSON.parse(kqlString(doc)), doc);
    }
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "completed"); assert.equal(mock.writes.length, 9);
    assert.equal(mock.tables.get(request.names.twinChanges)!.doc, docs.normalized);
    assert.equal(mock.views.get(request.names.twinState)!.doc, docs.lkv);
    assert.deepEqual(mock.tables.get(request.names.twinChanges)!.columnDocs, {});
    assert.equal(mock.tables.get(request.names.modeled)!.doc, mock.plan.documentation.table);
    assert.deepEqual(mock.tables.get(request.names.modeled)!.columnDocs, mock.plan.documentation.columns);
  }
});

test("each helper aggregate JSON budget includes its own opening control lines and escaping without cutting identities", () => {
  const model = fromThingModel(resolveThingModel(tm)), request = requestFor(model);
  const source = request.input.source, normalized = request.names.twinChanges;
  const empty = { ...model, declaredModelMetadata: { id: tm.id, description: "" } };
  const short = helperObjectDocstrings(empty, source, normalized);
  for (const kind of ["normalized", "lkv"] as const) for (const size of [999, 1000, 1001]) {
    const doc = helperObjectDocstrings({ ...empty, declaredModelMetadata: { id: tm.id, description: "x".repeat(size - short[kind].length) } },
      source, normalized)[kind];
    assert.ok(doc.length <= 1000);
    if (size <= 1000) assert.equal(doc.length, size);
    assert.equal(json(doc).truncated === true, size > 1000);
    assert.equal(json(doc).id, tm.id);
    assert.equal(parseIotMeta(doc).source, parseIotMeta(short[kind]).source);
    assert.equal(JSON.parse(kqlString(doc)), doc);
  }
  const description = '😀 "quotes" \\\n@iot.meta deviceModel=literal\u2028\u2029'.repeat(100);
  const docs = helperObjectDocstrings({ ...model, declaredModelMetadata: {
    id: tm.id, version: "actual-v1", title: "", description, titles: { en: description, de: description },
  } }, source, normalized);
  for (const doc of Object.values(docs)) {
    assert.ok(doc.length <= 1000 && !doc.includes("\ntruncated:"));
    assert.equal(json(doc).id, tm.id); assert.equal(json(doc).version, "actual-v1"); assert.equal(json(doc).title, "");
    assert.equal(json(doc).truncated, true); assert.equal(JSON.parse(kqlString(doc)), doc);
  }
  const absent = fromThingModel(resolveThingModel({ "@context": tm["@context"], "@type": tm["@type"], id: tm.id, properties: tm.properties }));
  for (const doc of Object.values(helperObjectDocstrings(absent, source, normalized))) assert.deepEqual(json(doc), { id: tm.id });
});

test("unrepresentable helper source budgets fail before any native request while modeled-table docs can still fit", async () => {
  const request = requestFor(fromThingModel(resolveThingModel(tm)));
  request.input.source.twinTable = "RawProperties_" + "p".repeat(1000);
  request.coverage.fingerprint = modelingFingerprint(request.input);
  assert.ok(semanticDocumentation(request.input.model, request.input.fields, { source: request.input.source, names: request.names }).table.length < 1000);
  let nativeRequests = 0;
  const result = await new ModeledDataCreator().run(request, { inspect: async () => request.coverage,
    request: async () => { nativeRequests++; throw new Error("Helper budget must fail before any native operation."); } }, true);
  assert.equal(result.phase, "blocked"); assert.equal(nativeRequests, 0); assert.match(result.error!, /Required model semantic/);
  const long = "urn:fixture:" + "i".repeat(1000);
  const model = { ...request.input.model, id: long, document: { ...tm, id: long }, declaredModelMetadata: { id: long } };
  assert.throws(() => helperObjectDocstrings(model, { ...request.input.source, twinTable: "Twins" }, "Changes"), /Required model semantic/);
});

test("legacy sanitized helper aliases remain legacy and helper JSON uses the existing malformed and collision guards", () => {
  const request = requestFor(fromThingModel(resolveThingModel(tm)));
  const oldNormalized = "Normalized device property changes in Entity-Attribute-Value (EAV) format\n" +
    "@iot.meta tableRole=properties-eav\n@iot.meta source=cluster/db/Twins\n@iot.meta deviceModel=ThermoPro_X200\n@iot.meta modelVersion=1.0";
  const oldView = "Property last known values\n@iot.meta viewRole=properties-lkv\n@iot.meta source=cluster/db/Changes\n" +
    "@iot.meta deviceModel=ThermoPro_X200\n@iot.meta modelVersion: 1.0";
  for (const doc of [oldNormalized, oldView]) {
    assert.deepEqual(readModelIdentityMetadata(doc), { documentVersion: "1.0" });
    assert.equal(parseIotMeta(doc).deviceModel, "ThermoPro_X200");
  }
  const docs = helperObjectDocstrings(request.input.model, request.input.source, request.names.twinChanges);
  for (const doc of Object.values(docs)) {
    assert.deepEqual(readModelIdentityMetadata(doc), { id: tm.id, documentVersion: "1.0" });
    assert.throws(() => parseIotMeta(doc + "\n@iot.meta deviceModel=ThermoPro_X200"), /aliases/);
    assert.throws(() => parseIotMeta(doc + "\n" + MODEL_SEMANTIC_INFO_PREFIX + "{}"), /Duplicate/);
    assert.throws(() => parseIotMeta(doc.replace(/@iot.meta modelSemanticInfo=.*/, MODEL_SEMANTIC_INFO_PREFIX + "{broken")), /Malformed/);
  }
  const marker = 'Actual "model"\n@iot.meta tableRole=not-control';
  const literal = helperObjectDocstrings({ ...request.input.model, declaredModelMetadata: { id: tm.id, title: marker } },
    request.input.source, request.names.twinChanges);
  assert.equal(json(literal.normalized).title, marker);
  assert.equal(parseIotMeta(literal.normalized).tableRole, "properties-eav");
  assert.equal(parseIotMeta(literal.lkv).viewRole, "properties-lkv");
});

test("both helper expected-document readbacks reject old metadata and invalidate old completion credit without adopting resources", async () => {
  const models = [fromThingModel(resolveThingModel(tm)), fromDtdlModel(parseDtdlModel(metadataDtdlModel), metadataDtdlModel)];
  for (const model of models) for (const helper of ["normalized", "lkv"] as const) {
    const request = requestFor(model), mock = createCreationMock(request), original = mock.service.request;
    const target = helper === "normalized" ? request.names.twinChanges : request.names.twinState;
    mock.service.request = async (csl, kind) => {
      const value = await original(csl, kind);
      if (csl.startsWith(`.alter ${helper === "normalized" ? "table" : "materialized-view"} ['${target}'] docstring `)) {
        const doc = "@iot.meta deviceModel=Old_Sanitized_Name\n@iot.meta modelVersion=1.0";
        if (helper === "normalized") mock.tables.get(target)!.doc = doc; else mock.views.get(target)!.doc = doc;
      }
      return value;
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed"); assert.match(result.error!, /Object docstring/);
    assert.ok(!mock.writes.some(command => command.includes(" policy update ") || command.startsWith(".drop")));
    const old = JSON.parse(creationFingerprint(request)); old[3] = "model-semantic-info-json-v6";
    assert.notEqual(creationFingerprint(request), JSON.stringify(old));
  }
});
