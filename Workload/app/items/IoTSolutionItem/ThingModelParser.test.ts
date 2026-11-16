// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import { parseThingModel, resolveThingModel, THING_MODEL_RESOLUTION_LIMITS, type ThingModelResolutionStats, type UploadedThingModel } from "./ThingModelParser.ts";
import { fieldValidation, fromThingModel, inspectFieldEvidence } from "./HubModeling.ts";
import { decodeDoSamples } from "./ModelAuthoring.ts";
import { isJsonObject, type JsonObject, type JsonValue } from "./ModelAuthoring.ts";
import { localPlainThingModel, localReferenceWorkModel, localRepeatedReferenceModel } from "../../../tests/localWotFixtures.ts";

const resolutionResults: Record<string, unknown>[] = [];
process.once("exit", () => {
  if (process.env.THING_MODEL_RESOLUTION_RESULTS) fs.writeFileSync(process.env.THING_MODEL_RESOLUTION_RESULTS,
    JSON.stringify(resolutionResults, null, 2) + "\n");
});

const tm = {
  "@context": "https://www.w3.org/2022/wot/td/v1.1",
  "@type": "tm:ThingModel", id: "urn:example:thermostat", title: "Thermostat",
  version: { model: "2.1" },
  properties: {
    temperature: { type: "number", unit: "degC", observable: true, readOnly: true },
    operatingMode: { type: "string", writeOnly: true },
    compressor: { type: "object", properties: { temperature: { type: "number" } } },
  },
  actions: { reboot: {} },
  events: { overheating: { data: { type: "object", properties: { temperature: { type: "number" } } } } },
};

test("validates a real TM structure while keeping document and registry versions distinct", () => {
  const result = parseThingModel(tm);
  assert.equal(result.id, "urn:example:thermostat");
  assert.equal(result.documentVersion, "2.1");
  assert.equal(result.fields.length, 4);
  assert.deepEqual(result.blockers, []);
  assert.ok(!("registryVersion" in result));
  assert.ok(result.fields.every(field => !("role" in field)));
  assert.ok(!result.fields.some(field => field.name === "reboot"));
});

test("does not infer Hub roles from observable, readOnly or writeOnly", () => {
  const result = parseThingModel(tm);
  assert.deepEqual(result.fields.slice(0, 2), [
    { name: "temperature", kind: "Property", title: "temperature", declaredType: "number" },
    { name: "operatingMode", kind: "Property", title: "operatingMode", declaredType: "string" },
  ]);
});

test("rejects TD, JSON Schema, DTDL and malformed structure", () => {
  for (const value of [
    { ...tm, "@type": "Thing" },
    { $schema: "https://json-schema.org/draft/2020-12/schema", type: "object" },
    { "@type": "Interface", "@id": "dtmi:example:Thermostat;1", contents: [] },
    [],
  ]) assert.throws(() => parseThingModel(value), /Thing Model/);
  for (const value of [
    { ...tm, "@context": "https://invalid.example/context" },
    { ...tm, properties: { temperature: { type: "float" } } },
    { ...tm, properties: { temperature: { observable: "yes" } } },
    { ...tm, version: { instance: "1" } },
  ]) assert.throws(() => parseThingModel(value), /WoT 1.1 format/);
});

test("required root ID supersedes local no-ID acceptance without inferring an identifier", () => {
  const { id: _id, ...local } = tm;
  assert.throws(() => parseThingModel(local), /doesn't have an ID/);
  for (const id of [null, "", " ", "\t", 3, {}, [], false]) {
    assert.throws(() => parseThingModel({ ...tm, id }), /doesn't have an ID/);
  }
  assert.throws(() => parseThingModel({ ...local, $id: tm.id, version: { model: tm.id } }), /doesn't have an ID/);
  assert.throws(() => parseThingModel({ ...tm, id: "thermostat.json" }), /absolute IRI/);
  assert.throws(() => parseThingModel({ ...tm, id: " urn:example:thermostat " }), /absolute IRI/);
  assert.throws(() => parseThingModel({ ...tm, id: "{{modelId}}" }), /absolute IRI/);
  for (const id of ["urn:example:Device;1", "https://example.org/设备#Model", `urn:example:${"x".repeat(520)}`]) {
    assert.equal(parseThingModel({ ...tm, id }).id, id);
  }
});

test("ID guidance gives a URN example without restricting valid absolute IRIs or identity", () => {
  const { id: _id, ...missing } = tm;
  assert.throws(() => parseThingModel(missing), { message:
    "This Thing Model doesn't have an ID. Add an ID, like urn:example:thermostat, then upload the file again." });
  assert.throws(() => parseThingModel({ ...tm, id: "relative/model.json" }), { message:
    "The model ID must be an absolute IRI, like urn:example:thermostat. Fix the ID, then upload the file again." });
  for (const id of ["urn:example:thermostat", "https://models.example.test/thermostat#v1", "tag:example.org,2026:thermostat"]) {
    const parsed = parseThingModel({ ...tm, id });
    assert.equal(parsed.id, id); assert.equal((parsed.document as Record<string, unknown>).id, id);
    assert.deepEqual(parsed.blockers, []);
  }
});
test("WoT format diagnostics retain all useful validation paths", () => {
  assert.throws(() => parseThingModel({ ...tm, properties: {
    temperature: { type: "float", observable: "yes" }, state: { type: 42, readOnly: "yes" },
  } }), error => error instanceof Error && error.message.startsWith("This Thing Model doesn't follow the WoT 1.1 format: ") &&
    error.message.includes("/properties/temperature") && error.message.includes("/properties/state") &&
    error.message.split("; ").length > 4);
});
test("accepts optional title, untyped schemas and language-specific supplied titles without inventing types", () => {
  const result = parseThingModel({
    "@context": ["https://www.w3.org/2022/wot/td/v1.1"],
    "@type": ["tm:ThingModel", "saref:Device"],
    id: "urn:fixture:optional-title",
    titles: { en: "Declared title" },
    properties: { status: { titles: { "en-US": "Operating state" } } },
  });
  assert.equal(result.title, "Declared title");
  assert.deepEqual(result.fields, [{ name: "status", kind: "Property", title: "Operating state" }]);
});

test("surfaces required imports/inheritance as unsupported and exposes no partial effective fields", () => {
  for (const document of [
    { ...tm, links: [{ rel: "tm:extends", href: "https://example.org/base.tm.jsonld" }] },
    { ...tm, properties: { ...tm.properties, imported: { "tm:ref": "#/schemaDefinitions/value" } } },
    { ...tm, properties: { ...tm.properties, imported: { "tm:ref": "https://example.org/base#/properties/value" } } },
    { ...tm, properties: { ...tm.properties, imported: { "$ref": "#/schemaDefinitions/value" } } },
    { ...tm, base: "{{baseURL}}" },
  ]) {
    const result = parseThingModel(document);
    assert.ok(result.blockers.length > 0);
    assert.deepEqual(result.fields, []);
  }
});

test("does not claim validation of a remapped Thing Model namespace", () => {
  const result = parseThingModel({
    ...tm, "@context": [tm["@context"], { tm: "https://example.org/not-wot#" }],
  });
  assert.ok(result.blockers.some(blocker => blocker.includes("namespace")));
  assert.deepEqual(result.fields, []);
});
  test("TM ignored null and oneOf candidates retain valid siblings original content and null-valued typed data", () => {
    const document: JsonObject = { ...tm, properties: {
      reading: { type: "number", unit: "degC", description: "Valid nullable sample." },
      nullType: { type: "null" }, literalNull: { type: null },
      choice: { oneOf: [{ type: "number" }, { type: "string" }] },
      unusedChoice: { oneOf: [{ "tm:ref": "https://unused.invalid/never-fetch" }] },
    }, events: { alarm: { data: { oneOf: [{ type: "boolean" }, { type: "null" }] } } } };
    const before = JSON.stringify(document), parsed = resolveThingModel(document);
    assert.deepEqual(parsed.blockers, []);
    assert.deepEqual(parsed.fields.map(field => field.name), ["reading"]);
    assert.deepEqual(parsed.ignored?.map(part => [part.name, part.reason]), [
      ["nullType", "null type"], ["literalNull", "null type"], ["choice", "oneOf"], ["unusedChoice", "oneOf"], ["alarm", "oneOf"],
    ]);
    assert.equal(JSON.stringify(document), before);
    assert.deepEqual(parsed.document, document);
    const model = fromThingModel(parsed), sample = decodeDoSamples([{
      data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
        event_kind: "measurement", event_time: "2026-10-01T12:00:00Z", payload: { value: { reading: null } } }),
      headers: { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "", IoTSubject: "" }, user_headers: {},
    }], "telemetry", 100);
    assert.equal(model.fields[0].kustoType, "real");
    assert.equal(model.fields[0].unit, "degC");
    assert.equal(inspectFieldEvidence(model.fields, sample, decodeDoSamples([], "twin", 100))[0].telemetry, 1);
    const empty = fromThingModel(resolveThingModel({ ...tm, properties: { unavailable: { type: null } }, events: {} }));
    assert.equal(empty.fields.length, 0);
    assert.ok(fieldValidation(empty.fields).length);
  });
  test("TM ignored fields are determined after supported local inheritance and imports including effective oneOf", () => {
    const base: JsonObject = { ...tm, id: "urn:fixture:ignored-base", events: {}, properties: {
      inheritedChoice: { oneOf: [{ type: "integer" }, { type: "string" }] }, replacedNull: { type: "null" },
    } };
    const document: JsonObject = { ...tm, links: [{ rel: "tm:extends", href: base.id }], events: {}, properties: {
      replacedNull: { type: "number", unit: "kPa" },
      importedChoice: { "tm:ref": "#/schemaDefinitions/choice" },
      importedNull: { "tm:ref": "#/schemaDefinitions/nullValue" },
      patched: { "tm:ref": "#/schemaDefinitions/choice", oneOf: null, type: "string" },
    }, schemaDefinitions: { choice: { oneOf: [{ type: "number" }, { type: "boolean" }] }, nullValue: { type: "null" } } };
    const parsed = resolveThingModel(document, [{ fileName: "base.tm.json", document: base }]);
    assert.deepEqual(parsed.blockers, []);
    assert.deepEqual(parsed.fields.map(field => field.name), ["replacedNull", "patched"]);
    assert.deepEqual(parsed.ignored?.map(part => [part.name, part.reason]), [
      ["inheritedChoice", "oneOf"], ["importedChoice", "oneOf"], ["importedNull", "null type"],
    ]);
    assert.equal(fromThingModel(parsed).fields[0].unit, "kPa");
    assert.deepEqual(parsed.document, document);
  });
  test("TM ignored submodels preserve connected non-first packaged roots inheritance and safe diagnostic names", () => {
    const child: JsonObject = { ...tm, id: "urn:fixture:ignored-child", properties: { childOnly: { type: "number" } }, events: {} };
    const base: JsonObject = { ...tm, id: "urn:fixture:ignored-parent", properties: { inherited: { type: "boolean" } }, events: {} };
    const root: JsonObject = { ...tm, id: "urn:fixture:ignored-root", properties: { own: { type: "string" } }, events: {},
      links: [{ rel: "tm:extends", href: base.id }, { rel: "tm:submodel", href: child.id, instanceName: "motor" },
        { rel: "https://www.w3.org/2022/wot/tm#submodel", href: "https://private.invalid/unused?credential=fixture" }],
      "tm:submodel": { ignored: "urn:fixture:construct" } };
    const document = [child, base, root], parsed = resolveThingModel(document);
    assert.deepEqual(parsed.blockers, []);
    assert.equal(parsed.id, root.id);
    assert.deepEqual(parsed.fields.map(field => field.name), ["inherited", "own"]);
    assert.deepEqual(parsed.ignored?.map(part => [part.kind, part.name]), [
      ["Submodel", "motor"], ["Submodel", "Submodel 3"], ["Submodel", "tm:submodel"],
    ]);
    assert.ok(!JSON.stringify(parsed.ignored).includes("private.invalid") && !JSON.stringify(parsed.ignored).includes("credential"));
    assert.deepEqual(parsed.document, document);
  });
  test("TM ignored projection never strips nested declarations from valid whole object and array fields", () => {
    const object = { type: "object", properties: { nullable: { type: "null" },
      choice: { oneOf: [{ type: "number" }, { type: "string" }] }, "tm:submodel": { type: "string" } } };
    const array = { type: "array", items: { oneOf: [{ type: "number" }, { type: "null" }] } };
    const parsed = resolveThingModel({ ...tm, properties: { object, array }, events: {} });
    assert.deepEqual(parsed.blockers, []);
    assert.deepEqual(parsed.ignored, []);
    const model = fromThingModel(parsed);
    assert.deepEqual(model.fields.map(field => [field.path[0], field.kustoType, field.schema]), [
      ["object", "dynamic", object], ["array", "dynamic", array],
    ]);
  });
test("accepts the WoT 1.1 2022 prefix and equivalent explicit JSON-LD 1.1 prefix definitions", () => {
  const prefixes: JsonValue[] = [
    "https://www.w3.org/2022/wot/tm#",
    { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true },
    { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true, "@protected": true },
    { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true, "@protected": false },
  ];
  for (const prefix of prefixes) {
    const document: JsonObject = { ...tm, "@context": [tm["@context"], {
      tm: prefix, dtdl: "urn:microsoft:dtdl:", sourceKind: "dtdl:sourceKind",
    }], properties: { ...tm.properties, temperature: { ...tm.properties.temperature, sourceKind: "Telemetry" } } };
    for (const result of [parseThingModel(document), resolveThingModel(document)]) {
      assert.deepEqual(result.blockers, []);
      assert.equal(result.fields.length, 4);
      assert.equal(result.id, tm.id);
      assert.deepEqual(result.document, document);
      assert.ok(result.fields.every(field => !("role" in field) && !("component" in field)));
    }
  }
});

test("accepts the absolute 2022 ThingModel type without rewriting the uploaded document", () => {
  const document = { ...tm, "@type": "https://www.w3.org/2022/wot/tm#ThingModel" };
  assert.deepEqual(parseThingModel(document).document, document);
  assert.deepEqual(resolveThingModel(document).blockers, []);
  assert.throws(() => parseThingModel({ ...tm, "@type": "https://www.w3.org/2019/wot/tm#ThingModel" }), /not a WoT Thing Model/);
});

test("rejects other namespaces and expanded definitions that do not establish the tm prefix", () => {
  for (const prefix of [
    "https://www.w3.org/2019/wot/tm#",
    "http://www.w3.org/2022/wot/tm#",
    "https://example.org/not-wot#",
    { "@id": "https://www.w3.org/2022/wot/tm#" },
    { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": false },
    { "@id": "https://www.w3.org/2019/wot/tm#", "@prefix": true },
    { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true, "@reverse": "https://example.org/reverse" },
  ]) {
    const result = parseThingModel({ ...tm, "@context": [tm["@context"], { tm: prefix }] });
    assert.ok(result.blockers.some(blocker => blocker.includes("namespace")));
    assert.deepEqual(result.fields, []);
  }
});

test("explicit protected metadata must be boolean and never exposes partial fields on failure", () => {
  const invalid: JsonValue[] = ["yes", "false", 0, 1, {}, [], null];
  for (const value of invalid) {
    const document: JsonObject = { ...tm, "@context": [tm["@context"], {
      tm: { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true, "@protected": value },
    }] };
    for (const result of [parseThingModel(document), resolveThingModel(document)]) {
      assert.ok(result.blockers.some(blocker => blocker.includes("@protected")));
      assert.deepEqual(result.fields, []);
      assert.deepEqual(result.document, document);
    }
  }
});

test("rejects pathological depth before recursive schema validation", () => {
  let nested: Record<string, unknown> = { type: "string" };
  for (let i = 0; i < 60; i++) nested = { type: "object", properties: { child: nested } };
  assert.throws(() => parseThingModel({ ...tm, properties: { nested } }), /nesting depth/);
});

function measuredResolution(document: JsonValue, dependencies: UploadedThingModel[] = []) {
  let stats: Readonly<ThingModelResolutionStats> | undefined;
  const result = resolveThingModel(document, dependencies, "bounded.tm.json", value => { stats = value; });
  assert.ok(stats);
  assert.ok(stats.workUnits <= THING_MODEL_RESOLUTION_LIMITS.workUnits);
  assert.ok(stats.expandedNodes <= THING_MODEL_RESOLUTION_LIMITS.expandedNodes);
  assert.ok(stats.expandedBytes <= THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  assert.ok(!("resolution" in result), "Bounded diagnostics must not become persisted model facts.");
  resolutionResults.push({ id: Array.isArray(document) ? "bundle" : document !== null && typeof document === "object" ? document.id : undefined,
    stats, blockers: result.blockers, fieldCount: result.fields.length, hasEffectiveDocument: !!result.effectiveDocument });
  return { result, stats };
}

function expandedNodeCount(value: JsonValue): number {
  return 1 + (value !== null && typeof value === "object"
    ? Object.values(value).reduce<number>((count, child) => count + expandedNodeCount(child), 0) : 0);
}

test("valid plain many-property models below the upload and output bounds retain every field without an import work error", () => {
  for (const count of [3000, 8000]) {
    const document = localPlainThingModel(count), text = JSON.stringify(document, null, 2);
    const inputBytes = Buffer.byteLength(text, "utf8");
    assert.ok(inputBytes <= THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
    if (count === 8000) assert.ok(inputBytes > THING_MODEL_RESOLUTION_LIMITS.expandedBytes * 0.98);
    assert.doesNotMatch(text, /tm:ref|tm:extends/);
    assert.equal(parseThingModel(document).fields.length, count);
    const { result, stats } = measuredResolution(JSON.parse(text));
    assert.deepEqual(result.blockers, []);
    assert.ok(stats.workUnits > 100_000 && stats.workUnits < THING_MODEL_RESOLUTION_LIMITS.workUnits);
    assert.ok(stats.expandedNodes < THING_MODEL_RESOLUTION_LIMITS.expandedNodes);
    assert.ok(stats.expandedBytes < THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
    assert.equal(stats.expandedBytes, Buffer.byteLength(JSON.stringify(document), "utf8"));
    assert.equal(stats.expandedNodes, expandedNodeCount(document));
    assert.equal(stats.cacheHits, 0);
    assert.deepEqual(result.document, document);
    assert.deepEqual(result.effectiveDocument, document);
    assert.ok(isJsonObject(document.properties));
    assert.deepEqual(result.fields, Object.entries(document.properties).map(([name, value]) => {
      assert.ok(isJsonObject(value));
      return { name, kind: "Property", title: value.title, declaredType: "number" };
    }));
    const model = fromThingModel(result);
    assert.equal(model.fields.length, count);
    assert.ok(model.fields.every((field, index) => field.path[0] === `measurement${index}` &&
      field.title === `Measurement ${index}` && field.kustoType === "real" && field.unit === "degC" &&
      field.description === "Observed sensor reading for operational monitoring."));
  }
});

test("plain models reach the exact expanded-node limit before the raised work budget and reject one further property", () => {
  const document = (count: number): JsonObject => ({ "@context": tm["@context"], "@type": tm["@type"],
    id: "urn:fixture:plain-node-boundary", title: "Plain node boundary",
    properties: Object.fromEntries(Array.from({ length: count }, (_, index) => [`reading${index}`, { type: "number" }])) });
  const count = (THING_MODEL_RESOLUTION_LIMITS.expandedNodes - expandedNodeCount(document(0))) / 2;
  assert.ok(Number.isInteger(count));
  const input = document(count);
  assert.ok(Buffer.byteLength(JSON.stringify(input), "utf8") < THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  const accepted = measuredResolution(input);
  assert.deepEqual(accepted.result.blockers, []);
  assert.equal(accepted.result.fields.length, count);
  assert.equal(accepted.stats.expandedNodes, THING_MODEL_RESOLUTION_LIMITS.expandedNodes);
  assert.ok(accepted.stats.workUnits < THING_MODEL_RESOLUTION_LIMITS.workUnits);
  const rejected = measuredResolution(document(count + 1));
  assert.match(rejected.result.blockers.join(" "), /expanded Thing Model.*node limit \(100000\)/);
  assert.ok(rejected.stats.workUnits < THING_MODEL_RESOLUTION_LIMITS.workUnits);
  assert.deepEqual(rejected.result.fields, []);
  assert.equal(rejected.result.effectiveDocument, undefined);
});

test("repeated acyclic imports reuse resolved references while counting every expanded output occurrence", () => {
  const document = localRepeatedReferenceModel(4), original = JSON.stringify(document);
  const { result, stats } = measuredResolution(document);
  assert.deepEqual(result.blockers, []);
  assert.ok(stats.cacheHits > 10);
  assert.ok(stats.workUnits < 1000 && stats.workUnits < stats.expandedNodes);
  assert.equal(stats.expandedBytes, Buffer.byteLength(JSON.stringify(result.effectiveDocument), "utf8"));
  assert.equal(stats.expandedNodes, expandedNodeCount(result.effectiveDocument!));
  assert.equal(JSON.stringify(document), original);
  assert.equal(result.fields.length, 1);
  assert.equal(fromThingModel(result).fields[0].kustoType, "dynamic");
});

test("small acyclic fanout cannot exceed the aggregate expanded node limit or expose partial fields", () => {
  const document = localRepeatedReferenceModel(7), original = JSON.stringify(document);
  assert.ok(Buffer.byteLength(original, "utf8") < 3000);
  const { result, stats } = measuredResolution(document);
  assert.match(result.blockers.join(" "), /expanded Thing Model.*node limit \(100000\)/);
  assert.deepEqual(result.fields, []);
  assert.equal(result.effectiveDocument, undefined);
  assert.ok(stats.cacheHits > 0 && stats.workUnits < 1000);
  assert.equal(JSON.stringify(document), original);
  assert.throws(() => fromThingModel(result), /dependencies/);
});

test("expanded nodes permit the exact limit and reject one further occurrence even when the target is cached", () => {
  const base = localRepeatedReferenceModel(6), initial = measuredResolution(base);
  assert.deepEqual(initial.result.blockers, []);
  const remaining = THING_MODEL_RESOLUTION_LIMITS.expandedNodes - initial.stats.expandedNodes - 1;
  assert.ok(remaining > 0 && remaining < 10000);
  const document: JsonObject = { ...base, "x:boundary": Array.from({ length: remaining }, () => true) };
  const accepted = measuredResolution(document);
  assert.deepEqual(accepted.result.blockers, []);
  assert.equal(accepted.stats.expandedNodes, THING_MODEL_RESOLUTION_LIMITS.expandedNodes);
  assert.equal(expandedNodeCount(accepted.result.effectiveDocument!), THING_MODEL_RESOLUTION_LIMITS.expandedNodes);
  const rejected = measuredResolution({ ...document, "x:boundary": Array.from({ length: remaining + 1 }, () => true) });
  assert.match(rejected.result.blockers.join(" "), /node limit/);
  assert.deepEqual(rejected.result.fields, []);
  assert.equal(rejected.result.effectiveDocument, undefined);
});

test("expanded JSON size counts UTF-8 and escapes exactly at the byte limit before final schema validation", () => {
  const prefix = 'Multibyte é😀 with "quotes", \\slashes and\nnewlines. ';
  const document = { "@context": tm["@context"], "@type": tm["@type"], id: "urn:fixture:byte-boundary",
    properties: { value: { type: "string", description: prefix } } };
  const overhead = Buffer.byteLength(JSON.stringify(document), "utf8");
  document.properties.value.description += "x".repeat(THING_MODEL_RESOLUTION_LIMITS.expandedBytes - overhead);
  assert.equal(Buffer.byteLength(JSON.stringify(document), "utf8"), THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  const accepted = measuredResolution(document);
  assert.deepEqual(accepted.result.blockers, []);
  assert.equal(accepted.stats.expandedBytes, THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  assert.equal(Buffer.byteLength(JSON.stringify(accepted.result.effectiveDocument), "utf8"), THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  const rejected = measuredResolution({ ...document,
    properties: { value: { ...document.properties.value, description: document.properties.value.description + "x" } } });
  assert.equal(rejected.result.blockers[0],
    "The expanded Thing Model exceeds the supported size (2 MB). Reduce repeated imports or split the model.");
  assert.equal(rejected.result.effectiveDocument, undefined);
  assert.deepEqual(rejected.result.fields, []);
});

test("cached large literal sizes are measured once rather than re-encoded for every import overlay", () => {
  const description = "x".repeat(256 * 1024);
  const document: JsonObject = { "@context": tm["@context"], "@type": tm["@type"], id: "urn:fixture:literal-cache",
    schemaDefinitions: { value: { type: "string", description } },
    properties: Object.fromEntries(Array.from({ length: 10 }, (_, index) =>
      [`field${index}`, { "tm:ref": "#/schemaDefinitions/value", title: `Reading ${index}` }])) };
  const encode = TextEncoder.prototype.encode;
  let largeEncodings = 0;
  TextEncoder.prototype.encode = function (input?: string) {
    if (input && input.length >= description.length) largeEncodings++;
    return encode.call(this, input);
  };
  try {
    const { result, stats } = measuredResolution(document);
    assert.match(result.blockers.join(" "), /supported size/);
    assert.equal(result.effectiveDocument, undefined);
    assert.deepEqual(result.fields, []);
    assert.equal(largeEncodings, 1);
    assert.ok(stats.workUnits < 1000 && stats.cacheHits > 0);
  } finally { TextEncoder.prototype.encode = encode; }
});

test("reference-heavy resolution permits its exact work bound and rejects one further item before node byte or depth bounds", () => {
  const empty = measuredResolution(localReferenceWorkModel(0)), one = measuredResolution(localReferenceWorkModel(1));
  assert.deepEqual(empty.result.blockers, []);
  assert.deepEqual(one.result.blockers, []);
  const perItem = one.stats.workUnits - empty.stats.workUnits;
  const count = Math.floor((THING_MODEL_RESOLUTION_LIMITS.workUnits - empty.stats.workUnits) / perItem) - 1;
  assert.ok(count > 1000 && count < 10000);
  const document = localReferenceWorkModel(count);
  assert.ok(Array.isArray(document) && isJsonObject(document[0]));
  const unaligned = measuredResolution(document);
  assert.deepEqual(unaligned.result.blockers, []);
  const remaining = THING_MODEL_RESOLUTION_LIMITS.workUnits - unaligned.stats.workUnits;
  const flags = remaining % 4, padding = (remaining - flags * 5) / 4;
  assert.ok(Number.isInteger(padding) && padding >= 0 && padding < 100);
  for (let index = 0; index < flags; index++) document[0][`x:flag${index}`] = true;
  document[0]["x:boundary"] = Array.from({ length: padding }, () => true);
  assert.ok(Buffer.byteLength(JSON.stringify(document), "utf8") < THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  const accepted = measuredResolution(document);
  assert.deepEqual(accepted.result.blockers, []);
  assert.equal(accepted.result.fields.length, 1);
  assert.equal(accepted.stats.workUnits, THING_MODEL_RESOLUTION_LIMITS.workUnits);
  assert.ok(accepted.stats.expandedNodes < THING_MODEL_RESOLUTION_LIMITS.expandedNodes);
  assert.ok(accepted.stats.expandedBytes < THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  document[0]["x:boundary"] = Array.from({ length: padding + 1 }, () => true);
  const rejected = measuredResolution(document);
  assert.equal(rejected.result.blockers[0],
    "Thing Model dependency resolution exceeds the supported work limit. Simplify the model or reduce repeated imports.");
  assert.equal(rejected.stats.workUnits, THING_MODEL_RESOLUTION_LIMITS.workUnits);
  assert.ok(rejected.stats.expandedNodes < THING_MODEL_RESOLUTION_LIMITS.expandedNodes);
  assert.ok(rejected.stats.expandedBytes < THING_MODEL_RESOLUTION_LIMITS.expandedBytes);
  assert.deepEqual(rejected.result.fields, []);
  assert.equal(rejected.result.effectiveDocument, undefined);
});

test("cached imports retain independent recursive overlays, deletion, array replacement and resolved null facts", () => {
  const definition: JsonObject = { type: "object", title: "Original", default: null,
    properties: { state: { type: "string", title: "State", enum: ["a", "b"] } },
    descriptions: { en: "Remove", fr: "Keep" } };
  const document: JsonObject = { "@context": tm["@context"], "@type": tm["@type"], id: "urn:fixture:cached-overlays",
    schemaDefinitions: { definition }, properties: {
      first: { "tm:ref": "#/schemaDefinitions/definition", title: null,
        descriptions: { en: null }, properties: { state: { enum: ["b"] } } },
      second: { "tm:ref": "#/schemaDefinitions/definition", title: "Second" },
    } };
  const original = JSON.stringify(document), { result, stats } = measuredResolution(document);
  assert.deepEqual(result.blockers, []);
  const properties = result.effectiveDocument!.properties as JsonObject;
  assert.deepEqual(properties.first, { type: "object", default: null,
    properties: { state: { type: "string", title: "State", enum: ["b"] } }, descriptions: { fr: "Keep" } });
  assert.deepEqual(properties.second, { ...definition, title: "Second" });
  assert.deepEqual((result.effectiveDocument!.schemaDefinitions as JsonObject).definition, definition);
  assert.ok(stats.cacheHits > 0);
  assert.equal(JSON.stringify(document), original);
});

test("reference reuse is scoped to the owning document and preserves relative fragment identity", () => {
  const shared: JsonObject = { "tm:ref": "#/schemaDefinitions/leaf" };
  const helper = (id: string, unit: string): JsonObject => ({ "@context": tm["@context"], "@type": tm["@type"], id,
    schemaDefinitions: { leaf: { type: "number", unit }, wrapper: shared } });
  const document: JsonObject = { "@context": tm["@context"], "@type": tm["@type"], id: "urn:fixture:cache-owners",
    properties: { a: { "tm:ref": "a.tm.json#/schemaDefinitions/wrapper" },
      b: { "tm:ref": "b.tm.json#/schemaDefinitions/wrapper" } } };
  const { result } = measuredResolution(document, [
    { fileName: "a.tm.json", document: helper("urn:fixture:helper:a", "degC") },
    { fileName: "b.tm.json", document: helper("urn:fixture:helper:b", "kPa") },
  ]);
  assert.deepEqual(result.blockers, []);
  assert.deepEqual(result.effectiveDocument!.properties, { a: { type: "number", unit: "degC" }, b: { type: "number", unit: "kPa" } });
});

test("warm cached definitions cannot bypass resolution depth or ignored-candidate context checks", () => {
  const definitions: JsonObject = {};
  let nested: JsonObject = { type: "number" };
  for (let index = 0; index < 12; index++) nested = { type: "object", properties: { child: nested } };
  definitions.leaf = nested;
  for (let index = 0; index < 30; index++) definitions[`chain${index}`] =
    { "tm:ref": `#/schemaDefinitions/${index === 29 ? "leaf" : `chain${index + 1}`}` };
  const document: JsonObject = { "@context": tm["@context"], "@type": tm["@type"], id: "urn:fixture:cached-depth",
    properties: { warm: { "tm:ref": "#/schemaDefinitions/leaf" }, deep: { "tm:ref": "#/schemaDefinitions/chain0" } },
    schemaDefinitions: definitions };
  const deep = measuredResolution(document);
  assert.match(deep.result.blockers.join(" "), /depth/);
  assert.equal(deep.result.effectiveDocument, undefined);
  const candidate = measuredResolution({ ...document, properties: { choice: { "tm:ref": "#/schemaDefinitions/choice" } },
    schemaDefinitions: { choice: { oneOf: [{ "tm:ref": "https://not-supplied.invalid/definition" }] } } });
  assert.match(candidate.result.blockers.join(" "), /Required model file missing/);
  assert.deepEqual(candidate.result.fields, []);
});
