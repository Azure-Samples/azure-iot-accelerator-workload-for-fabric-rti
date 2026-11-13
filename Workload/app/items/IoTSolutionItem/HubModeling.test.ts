// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { canonicalJson, decodeDoSamples, isJsonObject, type JsonObject, type StreamSample } from "./ModelAuthoring.ts";
import { resolveThingModel } from "./ThingModelParser.ts";
import { parseDtdlModel } from "./DtdlModelParser.ts";
import {
  applyRoleHints, assignableReference, coverageReady, effectiveScope, fieldValidation, fieldValue, fromDtdlModel,
  fromThingModel, inspectFieldEvidence, inspectScopedCoverage, inspectStreamCoverage, modelingFingerprint, nativeBindingIssues,
  nativeIdentity, scopedDoProjection, setFieldRole, validatePredicate, type CoverageSnapshot, type ModelingInput, type ProcessingScope,
} from "./HubModeling.ts";
import { buildModeledDataPlan, modeledDataScript, semanticDocumentation } from "./ModeledDataPlan.ts";
import { readColumnSemantics } from "./FieldSemantics.ts";
import { buildDataAgentContext, documentedColumns } from "./DataAgentContext.ts";
import { definitionWithoutModel } from "./IoTSolutionItemDefinition.ts";
import { mixedDtdlRequest } from "../../../tests/mixedSourceDtdlFixtures.ts";

const tm: JsonObject = {
  "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
  id: "urn:example:thermostat", title: "Thermostat", description: "Thermostat operating measurements and state.",
  version: { model: "2.1" },
  properties: {
    temperature: { type: "number", title: "Temperature", description: "Ambient temperature.", unit: "degC", "@type": "saref:Temperature" },
    compressor: { type: "object", properties: { temperature: { type: "number", description: "Compressor temperature.", unit: "degC" } } },
    state: { type: "string", description: "Reported operating state.", enum: ["idle", "running", "fault"] },
  },
};
const row = (value: unknown, stream: "telemetry" | "twin", user: JsonObject = {}, reference = "") => ({
  data: JSON.stringify({
    event_kind: stream === "telemetry" ? "measurement" : "lifecycle",
    event_time: "2026-10-01T12:00:00Z", schema_id: "com.do.telemetry.observation.v1",
    schema_version: "1.0", contract_version: "1.0",
    payload: { value, memberTypeIRI: reference, memberIRI: "urn:adr:not-native-device-identity" },
  }), headers: {}, user_headers: user,
});
const empty = (): StreamSample => ({ rawRows: 0, observations: 0, skippedWithoutReportedProperties: 0, payloads: [] });
function input(): ModelingInput {
  const model = fromThingModel(resolveThingModel(tm));
  return {
    model, fields: model.fields.map(field => setFieldRole(field, field.path[0] === "state" ? "twin" : "telemetry")),
    source: { queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databaseName: "Samples",
      telemetryTable: "Raw Telemetry", twinTable: "Raw Twins" },
    scope: { mode: "data", operator: "and", telemetry: "", twin: "" }, limit: 100,
  };
}
function snapshot(config: ModelingInput, telemetry = empty(), twin = empty()): CoverageSnapshot {
  return { fingerprint: modelingFingerprint(config), kind: "processing", telemetry, twin,
    fields: inspectFieldEvidence(config.fields, telemetry, twin), nativeIssues: nativeBindingIssues(telemetry, twin),
    checkedAt: "2026-10-01T12:00:00Z" };
}
const names = { modeled: "ModeledData_Thermostat", twinChanges: "TwinChanges_Thermostat", twinState: "TwinState_Thermostat" };

test("local TM inheritance/imports retain the full original document and preserve distinct effective paths", () => {
  const base: JsonObject = { ...tm, id: "urn:example:base", properties: {
    serial: { type: "string", description: "Manufacturer serial number." },
    reused: { type: "number", unit: "kPa", description: "Declared pressure." },
  } };
  const child: JsonObject = { ...tm, links: [{ rel: "tm:extends", href: "./base.tm.json" }],
    properties: { ...isJsonObject(tm.properties) ? tm.properties : {},
      pressure: { "tm:ref": "./base.tm.json#/properties/reused", title: "Pressure" } } };
  const result = resolveThingModel(child, [{ fileName: "base.tm.json", document: base }], "child.tm.json");
  assert.deepEqual(result.blockers, []);
  assert.deepEqual(result.document, child);
  const model = fromThingModel(result);
  assert.ok(model.fields.some(field => field.id === "/properties/compressor" && field.kustoType === "dynamic"));
  assert.ok(model.fields.some(field => field.id === "/properties/temperature"));
  assert.equal(model.fields.find(field => field.path[0] === "pressure")?.unit, "kPa");
  assert.ok(model.fields.every(field => field.role === undefined));
  assert.equal(model.documentVersion, "2.1");
});

test("local resolver blocks missing, cyclic, conflicting, unsupported and malformed required dependencies", () => {
  const missing = resolveThingModel({ ...tm, links: [{ rel: "tm:extends", href: "./missing.json" }] });
  assert.match(missing.blockers.join(" "), /Required model file missing/);
  assert.equal(missing.effectiveDocument, undefined);
  assert.deepEqual(missing.fields, []);
  const a: JsonObject = { ...tm, id: "urn:a", links: [{ rel: "tm:extends", href: "urn:b" }] };
  const b: JsonObject = { ...tm, id: "urn:b", links: [{ rel: "tm:extends", href: "urn:a" }] };
  assert.match(resolveThingModel(a, [{ fileName: "b.json", document: b }], "a.json").blockers.join(" "), /Cyclic/);
  assert.throws(() => resolveThingModel(tm, [{ fileName: "duplicate.json", document: tm }]), /Duplicate model ID/);
  const submodel = resolveThingModel({ ...tm, links: [{ rel: "tm:submodel", href: "urn:component" }] });
  assert.deepEqual(submodel.blockers, []);
  assert.ok(submodel.fields.length);
  assert.equal(submodel.ignored?.[0].reason, "submodel");
});

test("a missing child ID is rejected rather than inherited from a base model", () => {
  const child: JsonObject = {
    "@context": tm["@context"], "@type": "tm:ThingModel", links: [{ rel: "tm:extends", href: "urn:example:thermostat" }],
  };
  assert.throws(() => resolveThingModel(child, [{ fileName: "base.json", document: tm }]), /doesn't have an ID/);
  assert.ok(!("id" in child));
});

test("local dependency resolution does not drop a reference query/version selector", () => {
  const child: JsonObject = { ...tm, links: [{ rel: "tm:extends", href: "./base.json?version=7" }] };
  const result = resolveThingModel(child, [{ fileName: "base.json", document: { ...tm, id: "urn:base" } }], "child.json");
  assert.equal(result.effectiveDocument, undefined);
  assert.match(result.blockers.join(" "), /exact document binding/);
});

test("absolute Thing Model type is validated without changing the original JSON", () => {
  const document = { ...tm, "@type": "https://www.w3.org/2022/wot/tm#ThingModel" };
  const result = resolveThingModel(document);
  assert.deepEqual(result.blockers, []);
  assert.equal(result.document["@type"], document["@type"]);
});

test("event-name binding keeps scalar and unspecified event schemas as named declared values", () => {
  const result = resolveThingModel({ ...tm, events: { alarm: { data: { type: "string" } }, untyped: { title: "Untyped event" } } });
  const event = fromThingModel(result).fields.find(field => field.id === "/events/alarm/data")!;
  assert.deepEqual(event.path, ["alarm"]);
  assert.equal(event.column, "alarm"); assert.equal(event.kustoType, "string");
  assert.equal(event.role, undefined);
  const unspecified = fromThingModel(result).fields.find(field => field.column === "untyped")!;
  assert.equal(unspecified.kustoType, "dynamic"); assert.equal(unspecified.role, undefined);
  assert.deepEqual(unspecified.schema, {}); assert.deepEqual(unspecified.path, ["untyped"]);
  assert.ok(result.fields.some(field => field.name === "untyped" && !field.declaredType));
});

test("DTDL roles, exact enums and required schema dependencies remain authoritative", () => {
  const document = {
    "@id": "dtmi:example:Thermostat;2", "@type": "Interface",
    schemas: [{ "@id": "dtmi:example:State;1", "@type": "Enum", valueSchema: "integer", enumValues: [
      { name: "idle", enumValue: 0, displayName: "Idle" }, { name: "running", enumValue: 1, displayName: "Running" },
    ] }],
    contents: [
      { "@type": ["Telemetry", "Temperature"], name: "temperature", schema: "double", unit: "degreeCelsius" },
      { "@type": "Property", name: "state", schema: "dtmi:example:State;1" },
      { "@type": "Property", name: "date", schema: "date" },
      { "@type": "Command", name: "reboot" }, { "@type": ["Property", "Cloud"], name: "customer", schema: "string" },
    ],
  };
  const model = fromDtdlModel(parseDtdlModel(document), document);
  assert.equal(model.fields.length, 3);
  assert.equal(model.fields.find(field => field.column === "date")?.kustoType, "datetime");
  const state = model.fields.find(field => field.column === "state")!;
  assert.deepEqual(state.codes.map(code => [code.value, code.label]), [[0, "Idle"], [1, "Running"]]);
  assert.deepEqual(setFieldRole(state, state.role), { ...state, roleSource: "manual" },
    "Same-role programmatic calls retain the existing helper metadata semantics.");
  assert.equal(state.roleSource, "model");
  assert.throws(() => setFieldRole(state, "telemetry"), /declared/);
  assert.throws(() => parseDtdlModel({ ...document, extends: "dtmi:missing;1" }), /not supplied/);
  assert.throws(() => parseDtdlModel({ ...document, contents: [{ "@type": "Component", name: "motor", schema: "dtmi:missing;1" }] }), /not supplied/);
});

test("role hints keep manual/model choices and leave both/neither inconclusive", () => {
  const model = fromThingModel(resolveThingModel(tm));
  const telemetry = decodeDoSamples([row({ temperature: 20, state: "idle" }, "telemetry", { cloudEvents_subject: "device-1" })], "telemetry", 100);
  const twin = decodeDoSamples([row({ properties: { reported: { state: "running" } } }, "twin")], "twin", 100);
  const evidence = inspectFieldEvidence(model.fields, telemetry, twin);
  const hinted = applyRoleHints(model.fields, evidence);
  assert.equal(hinted.find(field => field.path[0] === "temperature")?.role, "telemetry");
  assert.equal(hinted.find(field => field.path[0] === "state")?.role, undefined);
  assert.equal(hinted.find(field => field.path[0] === "compressor")?.role, undefined);
  const manual = hinted.map(field => field.path[0] === "state" ? setFieldRole(field, "twin") : field);
  const after = applyRoleHints(manual, evidence.map(item => ({ ...item, state: "telemetry" })));
  assert.equal(after.find(field => field.path[0] === "state")?.role, "twin");
});

test("root whole values are distinct from explicit component telemetry, not inferred from an object name", () => {
  const config = input();
  const root = config.fields.find(field => field.path.length === 1 && field.path[0] === "temperature")!;
  const compressor = config.fields.find(field => field.path[0] === "compressor")!;
  const component = decodeDoSamples([row({ temperature: 50 }, "telemetry",
    { cloudEvents_subject: "d1", cloudEvents_iothubdtsubject: "compressor" })], "telemetry", 100).payloads[0];
  assert.equal(fieldValue(component, root, "telemetry").found, false);
  assert.deepEqual(fieldValue(component, compressor, "telemetry"), { found: false });
  const nested = decodeDoSamples([row({ compressor: { temperature: null } }, "telemetry", { cloudEvents_subject: "device-1" })], "telemetry", 100).payloads[0];
  assert.deepEqual(fieldValue(nested, compressor, "telemetry"), { found: true, value: { temperature: null } });
});

test("reference scopes fully parenthesize AND/OR and empty OR extras retain the type filter", () => {
  const scope: ProcessingScope = { mode: "reference", operator: "or", telemetry: "tostring(headers[\"IoTConnectionDeviceId\"]) == \"d1\" or data.payload.value.temperature > 10",
    twin: "", reference: { id: "urn:model" } };
  assert.equal(effectiveScope(scope, "telemetry"), "(tostring(data[\"payload\"][\"memberTypeIRI\"]) == \"urn:model\") or (tostring(headers[\"IoTConnectionDeviceId\"]) == \"d1\" or data.payload.value.temperature > 10)");
  assert.equal(effectiveScope(scope, "twin"), "(tostring(data[\"payload\"][\"memberTypeIRI\"]) == \"urn:model\")");
  assert.equal(effectiveScope({ ...scope, operator: "and" }, "telemetry"), "(tostring(data[\"payload\"][\"memberTypeIRI\"]) == \"urn:model\") and (tostring(headers[\"IoTConnectionDeviceId\"]) == \"d1\" or data.payload.value.temperature > 10)");
  assert.equal(effectiveScope(scope, "telemetry", true), scope.telemetry);
  validatePredicate("tostring(headers[\"IoTConnectionDeviceId\"]) == \"x;|y\"");
  assert.throws(() => validatePredicate("true | union Other"), /predicate only/);
  assert.throws(() => validatePredicate("(true"), /Close/);
});
test("unselected model reference choice cannot query processing coverage or authorize a plan", async () => {
  const config = input();
  config.scope = { ...config.scope, mode: "unselected" };
  let queries = 0;
  const fetcher: typeof fetch = async () => { queries++; throw new Error("Must not query."); };
  assert.throws(() => effectiveScope(config.scope, "telemetry"), /Choose a model reference filter/);
  for (const stream of ["telemetry", "twin"] as const) {
    await assert.rejects(() => inspectStreamCoverage(config, "synthetic", stream, undefined, fetcher), /Choose a model reference filter/);
  }
  await assert.rejects(() => inspectScopedCoverage(config, "synthetic", "processing", undefined, undefined, fetcher), /Choose a model reference filter/);
  const current = snapshot(config);
  assert.equal(coverageReady(config, current, true), false);
  assert.throws(() => buildModeledDataPlan(config, current, true, names), /processing scope/);
  assert.equal(queries, 0);
});
test("additional conditions require explicit choice and None ignores retained drafts in both generated predicates", async () => {
  const config = input();
  config.scope = { ...config.scope, additional: "unselected" };
  let queries = 0;
  const fetcher: typeof fetch = async () => { queries++; throw new Error("No queries before an additional-condition choice."); };
  for (const stream of ["telemetry", "twin"] as const) {
    await assert.rejects(() => inspectStreamCoverage(config, "fixture", stream, undefined, fetcher), /Choose whether to specify additional conditions/);
  }
  assert.equal(coverageReady(config, snapshot(config), true), false);
  assert.throws(() => buildModeledDataPlan(config, snapshot(config), true, names), /processing scope/);
  assert.equal(queries, 0);
  const telemetry = decodeDoSamples([row({ temperature: 22 }, "telemetry", { cloudEvents_subject: "d1" })], "telemetry", 100);
  const twin = decodeDoSamples([row({ properties: { reported: { state: "on" } } }, "twin", { cloudEvents_deviceid: "d1" })], "twin", 100);
  for (const mode of ["data", "reference"] as const) for (const operator of ["and", "or"] as const) {
    config.scope = { mode, operator, additional: "none", reference: { id: config.model.id! },
      telemetry: 'invalid retained draft | union Other', twin: 'payload.privateField == "retained-draft"' };
    const plan = buildModeledDataPlan(config, snapshot(config, telemetry, twin), false, names);
    for (const stream of ["telemetry", "twin"] as const) {
      const predicate = mode === "data" ? "true" : `(tostring(data["payload"]["memberTypeIRI"]) == "${config.model.id}")`;
      assert.equal(effectiveScope(config.scope, stream), predicate);
      const table = stream === "telemetry" ? config.source.telemetryTable : config.source.twinTable;
      assert.ok(plan.commands.some(command => command.kind === "policy" && command.source === table &&
        command.query.includes(`| where ${predicate}\n`) && !command.query.includes("retained") && !command.query.includes("privateField")));
    }
  }
  config.scope = { ...config.scope, additional: "specified", telemetry: "", twin: "" };
  assert.equal(effectiveScope(config.scope, "telemetry"), `(tostring(data["payload"]["memberTypeIRI"]) == "${config.model.id}")`);
});
test("original raw scope is shared before normalization in samples and both uncapped runtime policies", async () => {
  const config = input();
  const telemetryRow = row({ temperature: 22, compressor: { temperature: 23 } }, "telemetry", { cloudEvents_subject: "d1" }, config.model.id);
  const twinRow = row({ properties: { reported: { state: "running" } } }, "twin", { cloudEvents_deviceid: "d1" }, config.model.id);
  const telemetry = decodeDoSamples([telemetryRow], "telemetry", 100), twin = decodeDoSamples([twinRow], "twin", 100);
  const rawPredicate = 'data["schema_id"] == "com.do.telemetry.observation.v1"';
  assert.equal(JSON.parse(telemetryRow.data).schema_id, "com.do.telemetry.observation.v1");
  assert.equal(JSON.parse(telemetryRow.data).payload.value.schema_id, undefined, "This predicate must not run after data is replaced.");
  for (const mode of ["data", "reference"] as const) for (const operator of ["and", "or"] as const) {
    config.scope = { mode, operator, additional: "specified", reference: { id: config.model.id! },
      telemetry: rawPredicate, twin: 'isnotempty(tostring(user_headers["cloudEvents_deviceid"]))' };
    const plan = buildModeledDataPlan(config, snapshot(config, telemetry, twin), false, names);
    for (const stream of ["telemetry", "twin"] as const) {
      const projection = scopedDoProjection(config.scope, stream);
      assert.ok(projection.startsWith(`| extend data = parse_json(tostring(data))\n| where ${effectiveScope(config.scope, stream)}\n| extend __observation = data`));
      assert.ok(projection.indexOf("| where ") < projection.indexOf("| extend data = __observation.payload.value"));
      assert.ok(!projection.includes("mv-expand __observation") && !projection.includes("__iotRawConditionMatched"));
      const table = stream === "telemetry" ? config.source.telemetryTable : config.source.twinTable;
      const policy = plan.commands.find(command => command.kind === "policy" && command.source === table)!;
      assert.ok(policy.query!.includes(`['${table}']\n${projection}\n`));
      assert.ok(!/\| (top|take) \d+/.test(policy.query!), "Production processing must not inherit the author's sample cap.");
      const queries: string[] = [];
      await inspectStreamCoverage(config, "fixture", stream, undefined, async (_url, init) => {
        queries.push(JSON.parse(String(init?.body)).csl);
        const value = stream === "telemetry" ? telemetryRow : twinRow;
        return new Response(JSON.stringify({ Tables: [{
          Columns: [{ ColumnName: "data", ColumnType: "string" }, { ColumnName: "headers", ColumnType: "dynamic" },
            { ColumnName: "user_headers", ColumnType: "dynamic" }, { ColumnName: "__coverage_ingestion_time", ColumnType: "datetime" }],
          Rows: [[value.data, value.headers, value.user_headers, "2026-10-01T12:00:00Z"]],
        }] }));
      });
      assert.equal(queries.length, 2);
      assert.ok(queries[1].includes(`| top 10000 by __coverage_ingestion_time desc nulls first\n${projection}\n| top 100 by __coverage_ingestion_time desc nulls first\n`));
      assert.ok(queries[1].indexOf("| top 10000") < queries[1].indexOf("| extend data = parse_json"));
    }
    assert.ok(plan.commands.some(command => command.kind === "view" &&
      command.query.includes("ingestion_time() > ago(30d)") && command.query.includes("arg_max(event_time, propertyValue)")));
  }
  const current = snapshot(config, telemetry, twin), obsolete = JSON.parse(current.fingerprint);
  obsolete.predicateContract = "original-raw-row-v1";
  assert.equal(coverageReady(config, { ...current, fingerprint: canonicalJson(obsolete) }, false), false);
  delete obsolete.predicateContract;
  assert.equal(coverageReady(config, { ...current, fingerprint: canonicalJson(obsolete) }, false), false);
});
test("single-observation scoping rejects top-level batches and supports string and dynamic data with whole array-valued fields", async () => {
  const config = input();
  const value = row({ history: [{ value: 1 }, { value: 2 }] }, "telemetry", { cloudEvents_subject: "d1" });
  const batch = { ...value, data: JSON.stringify([JSON.parse(value.data)]) };
  const generic = decodeDoSamples([batch], "telemetry", 100);
  assert.equal(generic.payloads.length, 1, "The unrelated generic batch reader must remain usable.");
  assert.equal(generic.batchedRows, 1);
  let queries = 0;
  const response = (data: unknown, type = "string") => new Response(JSON.stringify({ Tables: [{
    Columns: [{ ColumnName: "data", ColumnType: type }, { ColumnName: "headers", ColumnType: "dynamic" },
      { ColumnName: "user_headers", ColumnType: "dynamic" }, { ColumnName: "__coverage_ingestion_time", ColumnType: "datetime" }],
    Rows: [[data, value.headers, value.user_headers, "2026-10-01T12:00:00Z"]],
  }] }));
  await assert.rejects(() => inspectStreamCoverage(config, "fixture", "telemetry", undefined,
    async () => { queries++; return response(batch.data); }), /top-level observation batches/);
  assert.equal(queries, 1, "Unsupported input must fail at the unfiltered probe, not return fabricated no-match results.");
  assert.throws(() => buildModeledDataPlan(config, snapshot(config, generic), true, names), /top-level observation batches/);
  await assert.rejects(() => inspectStreamCoverage(config, "fixture", "telemetry", undefined,
    async () => response(value.data, "long")), /data:string or dynamic/);
  for (const [data, type] of [[value.data, "string"], [JSON.parse(value.data), "dynamic"], [value.data, "dynamic"]] as const) {
    const supported = await inspectStreamCoverage(config, "fixture", "telemetry", undefined, async () => response(data, type));
    assert.deepEqual(supported.sample.payloads[0].value.history, [{ value: 1 }, { value: 2 }]);
  }
  await assert.rejects(() => inspectStreamCoverage(config, "fixture", "telemetry", undefined,
    async () => response([JSON.parse(value.data)], "dynamic")), /top-level observation batches/);
});
test("original raw scope leaves component extraction whole arrays and last-known enrichment after the filter", () => {
  const request = mixedDtdlRequest();
  request.input.scope = { ...request.input.scope, additional: "specified",
    telemetry: 'isnotempty(data)', twin: 'data["schema_id"] == "com.do.telemetry.observation.v1"' };
  request.coverage!.fingerprint = modelingFingerprint(request.input);
  const plan = buildModeledDataPlan(request.input, request.coverage!, false, request.names);
  const telemetry = plan.commands.find(command => command.kind === "policy" &&
    command.source === request.input.source.telemetryTable)!.query!;
  const twin = plan.commands.find(command => command.kind === "policy" &&
    command.source === request.input.source.twinTable)!.query!;
  assert.ok(telemetry.indexOf("| where isnotempty(data)") < telemetry.indexOf("| extend component"));
  assert.ok(telemetry.includes('iff(component == "motor", payload["measurement"], dynamic(null))') &&
    telemetry.includes('iff(component == "power", payload["samples"], dynamic(null))') &&
    telemetry.includes("| lookup kind=leftouter"));
  assert.ok(!telemetry.includes("mv-expand") && !/\| (top|take) \d+/.test(telemetry));
  assert.ok(twin.indexOf(`| where ${request.input.scope.twin}`) < twin.indexOf("| mv-expand key"));
  assert.ok(twin.includes('strcat(key, "_", subkey)') && twin.includes("val[subkey]"));
  assert.ok(plan.commands.some(command => command.kind === "view" && command.query.includes("ago(30d)") &&
    command.query.includes("arg_max(event_time, propertyValue)")));
});
test("original raw scope surfaces native alias/schema errors without translating or dropping the user's condition", async () => {
  const config = input();
  config.scope.telemetry = "payload.temperature > 10";
  const value = row({ temperature: 22 }, "telemetry", { cloudEvents_subject: "d1" });
  let queries = 0;
  await assert.rejects(() => inspectStreamCoverage(config, "fixture", "telemetry", undefined, async (_url, init) => {
    const query = JSON.parse(String(init?.body)).csl;
    queries++;
    if (query.includes("__observation")) {
      assert.ok(query.includes("| extend data = parse_json(tostring(data))\n| where payload.temperature > 10\n| extend __observation"));
      return new Response(JSON.stringify({ error: { code: "SemanticError", message: "Failed to resolve raw column 'payload'." } }), { status: 400 });
    }
    return new Response(JSON.stringify({ Tables: [{
      Columns: [{ ColumnName: "data", ColumnType: "string" }, { ColumnName: "headers", ColumnType: "dynamic" },
        { ColumnName: "user_headers", ColumnType: "dynamic" }, { ColumnName: "__coverage_ingestion_time", ColumnType: "datetime" }],
      Rows: [[value.data, value.headers, value.user_headers, "2026-10-01T12:00:00Z"]],
    }] }));
  }), /400|payload|SemanticError/);
  assert.equal(queries, 2);
  assert.equal(config.scope.telemetry, "payload.temperature > 10");
});
test("model coverage none AND OR selections reach both native update-policy predicates unchanged", () => {
  const config = input(), id = config.model.id!;
  const telemetry = decodeDoSamples([row({ temperature: 22 }, "telemetry", { cloudEvents_subject: "d1" })], "telemetry", 100);
  for (const mode of ["data", "reference"] as const) for (const operator of ["and", "or"] as const) {
    config.scope = { mode, operator, reference: { id }, telemetry: "data.payload.value.temperature > 20", twin: "" };
    const plan = buildModeledDataPlan(config, snapshot(config, telemetry), false, names);
    const telemetryPolicy = plan.commands.find(command => command.kind === "policy" && command.source === config.source.telemetryTable)!;
    const twinPolicy = plan.commands.find(command => command.kind === "policy" && command.source === config.source.twinTable)!;
    assert.ok(telemetryPolicy.query!.includes(`| where ${mode === "data" ? "data.payload.value.temperature > 20"
      : `(tostring(data["payload"]["memberTypeIRI"]) == "${id}") ${operator} (data.payload.value.temperature > 20)`}\n`));
    assert.ok(twinPolicy.query!.includes(`| where ${mode === "data" ? "true" : `(tostring(data["payload"]["memberTypeIRI"]) == "${id}")`}\n`));
    assert.ok(!twinPolicy.query!.includes("or (true)"));
    assert.ok(plan.commands.some(command => command.kind === "view" && command.query!.includes("arg_max(event_time, propertyValue)")));
  }
});

test("local WoT assignment uses only the supplied stable ID, while DTDL assignment and missing IDs are rejected", () => {
  const model = input().model;
  assert.deepEqual(assignableReference({ ...model, documentVersion: "2.3" }), { id: model.id });
  assert.throws(() => assignableReference({ ...model, kind: "dtdl" }), /only for WoT/);
  assert.throws(() => assignableReference({ ...model, id: undefined }), { message: "The model ID has changed. Upload the file again." });
  const dtdl = { mode: "reference" as const, operator: "and" as const, telemetry: "", twin: "", reference: { id: "dtmi:example:Device;1" } };
  assert.equal(effectiveScope(dtdl, "telemetry"), "(tostring(data[\"payload\"][\"memberTypeIRI\"]) == \"dtmi:example:Device;1\")");
});

test("zero processing telemetry requires explicit waiting; preview cannot establish processing readiness", () => {
  const config = input();
  const current = snapshot(config);
  assert.equal(coverageReady(config, current, false), false);
  assert.equal(coverageReady(config, current, true), true);
  assert.equal(coverageReady(config, { ...current, kind: "preview" }, true), false);
  const changed = { ...config, scope: { ...config.scope, operator: "or" as const } };
  assert.equal(coverageReady(changed, current, true), false);
  assert.equal(changed.fields[0].role, config.fields[0].role);
});

test("native metadata never falls back to ADR refs or telemetry deviceid-first heuristics", () => {
  const telemetry = decodeDoSamples([row({ temperature: 22 }, "telemetry",
    { cloudEvents_subject: "device-1/module-1", cloudEvents_deviceid: "wrong" })], "telemetry", 100);
  const twin = decodeDoSamples([row({ properties: { reported: { state: "idle" } } }, "twin",
    { cloudEvents_deviceid: "device-1", cloudEvents_moduleid: "module-1" })], "twin", 100);
  assert.equal(nativeIdentity(telemetry.payloads[0], "telemetry"), "device-1/module-1");
  assert.equal(nativeIdentity(twin.payloads[0], "twin"), "device-1/module-1");
  assert.deepEqual(nativeBindingIssues(telemetry, twin), []);
  const noNative = decodeDoSamples([row({ temperature: 22 }, "telemetry")], "telemetry", 100);
  assert.equal(nativeIdentity(noNative.payloads[0], "telemetry"), undefined);
  assert.ok(nativeBindingIssues(noNative, twin).length);
});

test("scoped coverage bounds raw input before filters and keeps telemetry/twin evidence isolated", async () => {
  const config = input();
  config.scope = { mode: "reference", operator: "or", telemetry: "tostring(headers[\"IoTConnectionDeviceId\"]) == \"d1\"", twin: "",
    reference: { id: config.model.id! } };
  const queries: string[] = [];
  const fetcher: typeof fetch = async (_url, init) => {
    const query = JSON.parse(String(init?.body)).csl;
    queries.push(query);
    const telemetry = query.includes("Raw Telemetry");
    const value = row({ temperature: 22, state: "idle" }, "telemetry", { cloudEvents_subject: "d1" });
    return new Response(JSON.stringify({ Tables: [{
      Columns: [{ ColumnName: "data", ColumnType: "string" }, { ColumnName: "headers", ColumnType: "dynamic" },
        { ColumnName: "user_headers", ColumnType: "dynamic" }, { ColumnName: "__coverage_ingestion_time", ColumnType: "datetime" }],
      Rows: telemetry ? [[value.data, value.headers, value.user_headers, "2026-10-01T12:00:00Z"]] : [],
    }] }));
  };
  const current = await inspectScopedCoverage(config, "mock", "processing", undefined, undefined, fetcher);
  assert.equal(current.telemetry.payloads.length, 1);
  assert.equal(current.twin.payloads.length, 0);
  const state = current.fields.find(field => field.fieldId.endsWith("/state"))!;
  assert.equal(state.telemetry, 1);
  assert.equal(state.twin, 0);
  const scoped = queries.filter(query => query.includes("__observation"));
  assert.equal(scoped.length, 2);
  assert.ok(scoped.every(query => query.indexOf("| top 10000") < query.indexOf("| where ") &&
    query.includes("| top 100 by __coverage_ingestion_time desc nulls first") && !query.includes("by event_time desc")));
  assert.ok(scoped.every(query => query.indexOf(`| where ${effectiveScope(config.scope, query.includes("Raw Telemetry") ? "telemetry" : "twin")}`) <
    query.indexOf("| extend __observation")));
  assert.ok(scoped[0].includes("or (tostring(headers[\"IoTConnectionDeviceId\"]) == \"d1\")"));
  assert.ok(scoped[1].includes(`| where (tostring(data["payload"]["memberTypeIRI"]) == "${config.model.id}")\n| extend`));
});

test("semantic docs contain meanings, exact model facts and native columns, not bindings or prototypes", () => {
  const config = input();
  const docs = semanticDocumentation(config.model, config.fields);
  assert.match(docs.table, /^Modeled device telemetry enriched with property LKVs/);
  assert.match(docs.columns.temperature, /Ambient temperature/);
  const temperature = readColumnSemantics(docs.columns.temperature), state = readColumnSemantics(docs.columns.state);
  assert.ok(temperature?.version === 3 && state?.version === 3);
  assert.equal(temperature.field.facts.unit, "degC");
  assert.deepEqual(state.field.facts.enum, ["idle", "running", "fault"]);
  assert.ok(!docs.columns.state.includes("Idle"));
  assert.match(docs.columns.state, /twin property \(last-known whole received value\)/);
  assert.ok(docs.columns.deviceId && docs.columns.event_time);
  const all = JSON.stringify(docs);
  for (const forbidden of ["rowKind", "Raw Telemetry", "Raw Twins", "registry", "Kusto", "cast", "prototype", "binding", "no conversion"]) {
    assert.ok(!all.includes(forbidden), forbidden);
  }
  assert.ok(!docs.columns.temperature.includes(config.model.id!));
});

test("unknown title/version facts are omitted while the required ID and supported enum codes remain exact", () => {
  const result = resolveThingModel({
    "@context": tm["@context"], "@type": "tm:ThingModel",
    id: "urn:fixture:code-labels",
    properties: { state: { type: "string", enum: ["idle", "run"] } },
  });
  const model = fromThingModel(result);
  const fields = model.fields.map(field => setFieldRole(field, "twin"));
  const docs = semanticDocumentation(model, fields);
  assert.ok(!docs.table.includes("Untitled"));
  assert.match(docs.table, /@iot.meta modelSemanticInfo=\{"id":"urn:fixture:code-labels"\}/);
  assert.ok(!docs.table.includes("version"));
  assert.match(docs.columns.state, /"enum":\["idle","run"\]/);
});

test("semantic column dictionaries retain special property names without prototype mutation", () => {
  const model = input().model;
  const field = { ...model.fields[0], column: "__proto__", role: "telemetry" as const };
  const docs = semanticDocumentation(model, [field]);
  assert.ok(Object.prototype.hasOwnProperty.call(docs.columns, "__proto__"));
  assert.equal(Object.getPrototypeOf(docs.columns), null);
});

test("setup plan keeps column semantics on ModeledData and restores plain owned helper object docs without ingestion mutation", () => {
  const config = input();
  const telemetry = decodeDoSamples([row({ temperature: 22 }, "telemetry", { cloudEvents_subject: "d1" })], "telemetry", 100);
  const current = snapshot(config, telemetry);
  const plan = buildModeledDataPlan(config, current, false, names);
  const script = modeledDataScript(plan);
  assert.equal(plan.commands.filter(command => /column-docstrings/.test(command.csl)).length, 1);
  assert.equal(plan.commands.filter(command => command.kind === "viewDoc").length, 1);
  assert.equal(plan.commands.filter(command => command.kind === "tableDoc").length, 2);
  assert.ok(plan.commands.filter(command => command.kind === "columnDocs").every(command => command.target === names.modeled));
  assert.ok(!plan.commands.some(command => /policy streamingingestion/.test(command.csl)));
  assert.ok(!script.includes("sourceHubId"));
  assert.ok(!script.includes("rowKind"));
  assert.ok(script.includes("ingestion_time() > ago(30d)"));
  assert.match(script, /arg_max\(event_time, propertyValue\)/);
  assert.match(script, /lookup kind=leftouter/);
  const twinPolicy = plan.commands.find(command => command.label === `Set update policy on ${names.twinChanges}`)!.csl;
  const literal = twinPolicy.slice(twinPolicy.indexOf(" policy update ") + " policy update ".length);
  assert.match(JSON.parse(JSON.parse(literal))[0].Query, /bag_keys\(rp\)/);
  assert.ok(!JSON.parse(JSON.parse(literal))[0].Query.includes("__selected"));
  assert.ok(plan.prerequisites.some(item => item.includes("streaming")));
  assert.match(plan.dataAgent.sourceInstructions, /last-known/);
  assert.ok(!plan.dataAgent.agentInstructions.includes("temperature"));
  assert.throws(() => buildModeledDataPlan(config, { ...current, kind: "preview" }, true, names), /processing scope/);
  assert.throws(() => buildModeledDataPlan(config, current, false, { ...names, modeled: config.source.telemetryTable }), /distinct output/);
});

test("consumer context snapshots native docs and separates shared purpose from source semantics/query rules", () => {
  const columns = documentedColumns({ OrderedColumns: [
    { Name: "deviceId", CslType: "string", DocString: "Native device identity." },
    { Name: "event_time", CslType: "datetime", DocString: "Event timestamp." },
    { Name: "temperature", CslType: "real", DocString: "Ambient temperature.\nTelemetry.\nUnit: degC." },
    { Name: "state", CslType: "string", DocString: 'Twin property (last-known value).\nValues: 0 = Idle; 1 = Running.' },
  ] });
  const context = buildDataAgentContext({ database: "Samples", table: "Modeled Data", docstring: "Thermostat data.\n@iot.meta tableRole=modeled-data", columns });
  assert.match(context.sourceInstructions, /unit: degC/);
  assert.match(context.sourceInstructions, /0 = Idle/);
  assert.match(context.sourceInstructions, /not an event-time\/as-of snapshot/);
  assert.ok(!context.agentInstructions.includes("temperature"));
  assert.ok(!context.agentDescription.includes("Modeled Data"));
  assert.ok(context.examples.some(example => example.query.includes("where isnotnull") && example.query.includes("arg_max")));
  assert.ok(context.examples.every(example => example.query.startsWith("['Modeled Data']")));
  assert.throws(() => documentedColumns({ OrderedColumns: [{ Name: "x" }] }), /unsupported/);
});

test("model completion does not retain model JSON, bindings or output config in the accelerator item", () => {
  const existing = { wizardStep: "completed", deviceModel: { fileName: "model.json", modelName: "Thermostat", modelJson: tm,
    modeledDataName: "ModeledData" }, tables: { telemetryTableName: "RawTelemetry", propertiesTableName: "RawTwins" } };
  const next = definitionWithoutModel(existing);
  assert.ok(!("deviceModel" in next));
  assert.deepEqual(next.tables, existing.tables);
  assert.ok(existing.deviceModel);
});

test("selected fields require roles and unique native-safe output names", () => {
  const config = input();
  assert.deepEqual(fieldValidation(config.fields), []);
  assert.ok(fieldValidation(config.fields.map(field => ({ ...field, column: "deviceId" }))).length);
  assert.ok(fieldValidation(config.fields.map(field => ({ ...field, role: undefined }))).length);
});
