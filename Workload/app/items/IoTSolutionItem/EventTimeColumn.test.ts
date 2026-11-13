// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { fromDtdlModel, fromThingModel, fieldValidation, modelingFingerprint, coverageReady, setFieldRole } from "./HubModeling.ts";
import { parseDtdlModel } from "./DtdlModelParser.ts";
import { resolveThingModel } from "./ThingModelParser.ts";
import { buildModeledDataPlan } from "./ModeledDataPlan.ts";
import { creationFingerprint, ModeledDataCreator } from "./ModeledDataCreation.ts";
import { canonicalJson, type JsonObject } from "./ModelAuthoring.ts";
import { FIELD_SEMANTIC_INFO_PREFIX } from "./FieldSemantics.ts";
import { buildDataAgentContext, buildDataAgentDataSource, requireEventTimeColumn } from "./DataAgentContext.ts";
import { buildDataAgentQueryGuidelines, buildLatestValueQuery } from "./dataAgentQueryGuidance.ts";
import { currentStreamChecks } from "./StreamCoverage.ts";
import { createCreationMock, creationRequest } from "../../../tests/creationMock.ts";
import { metadataDtdlModel } from "../../../tests/fieldSemanticFixtures.ts";

test("both model families create event_time-only EAV LKV and modeled schemas with unchanged DO clock and horizon", async () => {
  for (const kind of ["wot", "dtdl"] as const) {
    const request = creationRequest();
    if (kind === "dtdl") {
      request.input.model = fromDtdlModel(parseDtdlModel(metadataDtdlModel), metadataDtdlModel);
      request.input.fields = request.input.model.fields;
      request.coverage.fingerprint = modelingFingerprint(request.input);
    }
    const mock = createCreationMock(request), plan = mock.plan;
    const eav = plan.commands.find(command => command.kind === "table" && command.target === request.names.twinChanges)!;
    const lkv = plan.commands.find(command => command.kind === "view")!;
    const modeled = plan.commands.find(command => command.kind === "table" && command.target === request.names.modeled)!;
    assert.ok(eav.kind === "table" && lkv.kind === "view" && modeled.kind === "table");
    assert.deepEqual(eav.columns.map(column => [column.name, column.type]),
      [["deviceId", "string"], ["propertyName", "string"], ["propertyValue", "dynamic"], ["event_time", "datetime"]]);
    assert.deepEqual(lkv.columns.map(column => column.name), ["deviceId", "propertyName", "event_time", "propertyValue"]);
    assert.deepEqual(modeled.columns.slice(0, 2), [{ name: "deviceId", type: "string" }, { name: "event_time", type: "datetime" }]);
    assert.ok(lkv.query.includes("ingestion_time() > ago(30d)") && lkv.query.includes("arg_max(event_time, propertyValue)"));
    for (const command of plan.commands.filter(command => command.kind === "policy")) {
      assert.ok(command.kind === "policy" && command.query.includes("event_time = todatetime(__observation.event_time)"));
      assert.ok(!command.csl.includes("enqueuedTime") && !command.csl.includes("IoTEnqueueTime"));
    }
    const system = JSON.parse(plan.documentation.columns.event_time.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
    assert.deepEqual(system, { name: "event_time", description: "Event timestamp" });
    const credit = JSON.parse(creationFingerprint(request));
    assert.equal(credit[4], plan.documentation.columns.event_time);
    const oldDescription = [...credit];
    oldDescription[4] = FIELD_SEMANTIC_INFO_PREFIX + '{"name":"event_time","description":"Event timestamp from the DO observation."}';
    assert.notEqual(creationFingerprint(request), JSON.stringify(oldDescription));
    assert.equal(plan.documentation.columns.enqueuedTime, undefined);
    assert.equal(JSON.parse(plan.documentation.columns.deviceId.slice(FIELD_SEMANTIC_INFO_PREFIX.length)).description, "IoT Hub device Id");
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "completed"); assert.equal(mock.writes.length, 9);
    for (const name of [request.names.twinChanges, request.names.modeled]) {
      assert.ok(mock.tables.get(name)!.columns.some(column => column.name === "event_time" && column.type === "datetime"));
      assert.ok(!mock.tables.get(name)!.columns.some(column => column.name === "enqueuedTime"));
    }
    assert.equal(result.fingerprint, creationFingerprint(request));
  }
});

test("actual event_time model fields retain declared identity in distinct output columns and cannot overwrite system time", () => {
  const wot: JsonObject = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
    id: "urn:fixture:event-time-field", title: "Time field", properties: {
      event_time: { type: "string", description: "Declared literal event_time field." },
      enqueuedTime: { type: "string", description: "Declared literal enqueuedTime field." },
    } };
  const wotModel = fromThingModel(resolveThingModel(wot));
  const fields = wotModel.fields.map(field => setFieldRole(field, "telemetry"));
  assert.deepEqual(fields.map(field => [field.path[0], field.column]),
    [["event_time", "event_time_value"], ["enqueuedTime", "enqueuedTime"]]);
  assert.deepEqual(fieldValidation(fields), []);
  assert.ok(fieldValidation(fields.map((field, index) => index ? field : { ...field, column: "event_time" })).length);
  const dtdl: JsonObject = { "@id": "dtmi:fixture:TimeField;1", "@type": "Interface", contents: [
    { "@type": "Telemetry", name: "event_time", schema: "string", description: "Literal event_time meaning." },
  ] };
  const model = fromDtdlModel(parseDtdlModel(dtdl), dtdl);
  assert.equal(model.fields[0].column, "event_time_value"); assert.equal(model.fields[0].path[0], "event_time");
  assert.deepEqual(fieldValidation(model.fields), []);
  for (const selected of [{ model: wotModel, fields }, { model, fields: model.fields }]) {
    const request = creationRequest();
    request.input = { ...request.input, ...selected }; request.coverage.fingerprint = modelingFingerprint(request.input);
    const plan = buildModeledDataPlan(request.input, request.coverage, true, request.names);
    assert.ok(plan.commands.some(command => command.kind === "policy" && command.query.includes('payload["event_time"]')));
    const semantic = JSON.parse(plan.documentation.columns.event_time_value.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
    assert.equal(semantic.name, "event_time");
  }
  assert.equal(canonicalJson(wotModel.document), canonicalJson(wot));
  assert.equal(canonicalJson(model.document), canonicalJson(dtdl));
});

test("old native time schemas cannot pass current exact output readback for EAV LKV or modeled data", async () => {
  for (const target of ["normalized", "lkv", "modeled"] as const) {
    const request = creationRequest(), mock = createCreationMock(request), original = mock.service.request;
    const name = target === "normalized" ? request.names.twinChanges : target === "modeled" ? request.names.modeled : request.names.twinState;
    mock.service.request = async (csl, kind) => {
      if (target !== "lkv" && csl === `.show table ['${name}'] schema as json` && mock.tables.has(name)) {
        mock.tables.get(name)!.columns = mock.tables.get(name)!.columns.map(column => column.name === "event_time" ? { ...column, name: "enqueuedTime" } : column);
      }
      const result = await original(csl, kind);
      if (target === "lkv" && kind === "query" && csl.startsWith(`['${name}']\n`)) {
        const index = result.Columns.findIndex(column => column.ColumnName === "ColumnName");
        for (const row of result.Rows) if (row[index] === "event_time") row[index] = "enqueuedTime";
      }
      return result;
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed"); assert.match(result.error!, /schema|column/i);
    assert.ok(!mock.writes.some(command => command.includes(" policy update ") || command.startsWith(".drop")));
  }
});

test("old coverage tickets and old completed metadata cannot certify the event_time schema or trigger replay", () => {
  const request = creationRequest(), snapshot = JSON.parse(modelingFingerprint(request.input));
  delete snapshot.timeColumn;
  assert.equal(coverageReady(request.input, { ...request.coverage, fingerprint: canonicalJson(snapshot) }, true), false);
  const checks = currentStreamChecks(request.input, 1), base = JSON.parse(checks.base);
  delete base.timeColumn;
  const next = currentStreamChecks(request.input, 1, { ...checks, base: canonicalJson(base) });
  assert.notEqual(next.base, canonicalJson(base)); assert.equal(next.telemetry.phase, "unchecked");
  const prior = JSON.parse(creationFingerprint(request)); prior[3] = "all-object-model-semantic-info-json-v7";
  assert.notEqual(creationFingerprint(request), JSON.stringify(prior));
});

test("active consumers require event_time with no old-time fallback while preserving genuine other field names", () => {
  const request = creationRequest(), plan = buildModeledDataPlan(request.input, request.coverage, true, request.names);
  {
    const time = "event_time";
    const source = { database: "Database", table: "Inspected", docstring: plan.documentation.table, columns: [
      { name: "deviceId", type: "string", docstring: plan.documentation.columns.deviceId },
      { name: time, type: "datetime", docstring: time === "event_time" ? plan.documentation.columns.event_time : "Existing legacy time meaning." },
      { name: "temperature", type: "real", docstring: plan.documentation.columns.temperature },
    ] };
    assert.equal(requireEventTimeColumn(source.columns), time);
    const context = buildDataAgentContext(source);
    assert.ok(context.sourceInstructions.includes(`Use \`${time}\` for time-range filters`));
    assert.ok(context.examples.some(example => example.query.includes(`arg_max(${time},`)));
    assert.ok(context.examples.some(example => example.query.includes(`bin(${time}, 5m)`)));
    const dataSource = buildDataAgentDataSource(source, "workspace", "database", "table");
    assert.equal(dataSource.elements[0].children[0].children[1].id, time);
    assert.ok(buildLatestValueQuery("Inspected", "temperature").includes(`arg_max(${time}, *)`));
    assert.ok(buildDataAgentQueryGuidelines().includes(`bin(${time}, interval)`));
  }
  assert.throws(() => requireEventTimeColumn([{ name: "other", type: "datetime" }]), /event_time:datetime/);
  assert.throws(() => buildDataAgentContext({ database: "Database", table: "OldSchema", docstring: "", columns: [
    { name: "enqueuedTime", type: "datetime" },
  ] }), /event_time:datetime/);
  assert.throws(() => requireEventTimeColumn([{ name: "event_time", type: "string" }]), /event_time:datetime/);
});
