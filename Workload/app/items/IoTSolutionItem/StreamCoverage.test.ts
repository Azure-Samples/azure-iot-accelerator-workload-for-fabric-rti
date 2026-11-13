// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { COVERAGE_SAMPLING_CONTRACT, coverageReady, effectiveScope, inspectStreamCoverage, modelingFingerprint, setFieldRole, validatePredicate, type ModelingInput } from "./HubModeling.ts";
import { assembleStreamCoverage, beginStreamCheck, cancelStreamCheck, coverageConditionExample, currentStreamChecks, currentStreamResult, editStreamChecks, finishStreamCheck, modelCoverageReady, streamCoverageReady, type StreamChecks } from "./StreamCoverage.ts";
import { createCreationMock, creationRequest } from "../../../tests/creationMock.ts";
import { ModeledDataCreator } from "./ModeledDataCreation.ts";
import { canonicalJson, type SampleStream } from "./ModelAuthoring.ts";

function result(request: ReturnType<typeof creationRequest>, stream: SampleStream) {
  return { sample: request.coverage[stream], probe: request.coverage.sourceProbe![stream], checkedAt: "2026-10-21T12:00:00Z" };
}
function complete(request = creationRequest(), revision = 1): StreamChecks {
  let checks = currentStreamChecks(request.input, revision);
  for (const stream of ["telemetry", "twin"] as const) {
    const begun = beginStreamCheck(request.input, revision, stream, checks);
    checks = finishStreamCheck(request.input, begun.checks, begun.ticket, { result: result(request, stream) });
  }
  return checks;
}
test("coverage examples use selected whole field paths and exact native component guards with safe KQL escaping", () => {
    const fields = creationRequest().input.fields;
    const telemetryField = fields.find(field => field.role === "telemetry")!;
    assert.equal(coverageConditionExample(fields, "telemetry"), 'isempty(tostring(headers["IoTSubject"])) and isempty(tostring(user_headers["cloudEvents_iothubdtsubject"])) and isnotnull(data.payload.value.temperature)');
    assert.equal(coverageConditionExample(fields, "twin"), 'isnotnull(data.payload.value.properties.reported.state)');
    const component = { ...telemetryField, component: 'motor"\\one', path: ['whole/value~"key'] };
    const telemetry = coverageConditionExample([component], "telemetry");
    assert.equal(telemetry, '(tostring(headers["IoTSubject"]) == "motor\\"\\\\one" or tostring(user_headers["cloudEvents_iothubdtsubject"]) == "motor\\"\\\\one") and isnotnull(data.payload.value["whole/value~\\"key"])');
    const twin = coverageConditionExample([{ ...component, role: "twin" }], "twin");
    assert.equal(twin, 'tostring(data.payload.value.properties.reported["motor\\"\\\\one"].__t) == "c" and isnotnull(data.payload.value.properties.reported["motor\\"\\\\one"]["whole/value~\\"key"])');
    const root = coverageConditionExample([...fields, { ...component, role: "twin" }], "twin");
    assert.equal(root, 'tostring(data.payload.value.properties.reported.state.__t) != "c" and isnotnull(data.payload.value.properties.reported.state)');
    const array = { ...telemetryField, path: ["readings"], kustoType: "dynamic", schema: { type: "array" } };
    assert.equal(coverageConditionExample([array], "telemetry"), 'isempty(tostring(headers["IoTSubject"])) and isempty(tostring(user_headers["cloudEvents_iothubdtsubject"])) and isnotnull(data.payload.value.readings)');
    assert.equal(coverageConditionExample(fields.map(field => ({ ...field, included: false })), "telemetry"), 'isnotempty(tostring(headers["IoTConnectionDeviceId"])) or isnotempty(tostring(user_headers["cloudEvents_subject"]))');
    for (const expression of [telemetry, twin, root]) assert.doesNotThrow(() => validatePredicate(expression));
  });
  test("coverage result reuse requires current per-stream proof rather than a checked phase or old owner", () => {
    const request = creationRequest(), checks = complete(request);
    assert.equal(currentStreamResult(request.input, 1, checks, "telemetry"), checks.telemetry.proof!.result);
    assert.equal(currentStreamResult(request.input, 1, checks, "twin"), checks.twin.proof!.result);
    assert.equal(currentStreamResult(request.input, 2, checks, "twin"), undefined);
    assert.equal(currentStreamResult(request.input, 1, { ...checks, telemetry: { ...checks.telemetry, proof: undefined } }, "telemetry"), undefined);
    assert.equal(currentStreamResult(request.input, 1, { ...checks, twin: { ...checks.twin, proof: checks.telemetry.proof } }, "twin"), undefined);
    const changed = { ...request.input, scope: { ...request.input.scope, telemetry: "true" } };
    const edited = editStreamChecks(request.input, changed.scope, 1, checks);
    assert.equal(currentStreamResult(changed, 1, edited, "telemetry"), undefined);
    assert.equal(currentStreamResult(changed, 1, edited, "twin"), checks.twin.proof!.result);
  });
test("independent coverage requires both actual stream proofs and rejects wrong-stream or old combined credit", () => {
  const request = creationRequest(), started = beginStreamCheck(request.input, 1, "telemetry");
  const one = finishStreamCheck(request.input, started.checks, started.ticket, { result: result(request, "telemetry") });
  assert.equal(assembleStreamCoverage(request.input, 1, one), undefined);
  assert.equal(streamCoverageReady(request.input, 1, one, false), false);
  const both = complete(request);
  assert.equal(streamCoverageReady(request.input, 1, both, false), true);
  assert.equal(assembleStreamCoverage(request.input, 1, both)!.fingerprint, modelingFingerprint(request.input));
  assert.equal(assembleStreamCoverage(request.input, 1, { ...both, twin: { ...both.twin, proof: both.telemetry.proof } }), undefined);
  assert.equal(streamCoverageReady(request.input, 1, undefined, false), false);
});
test("unselected reference choice and reference-option ABA retire both stream proofs", () => {
  const request = creationRequest(), original = complete(request);
  const pending: ModelingInput = { ...request.input, scope: { ...request.input.scope, mode: "unselected" } };
  assert.equal(assembleStreamCoverage(pending, 1, original), undefined);
  assert.equal(streamCoverageReady(pending, 1, original, true), false);
  assert.throws(() => beginStreamCheck(pending, 1, "telemetry"), /Choose a model reference filter/);
  const away = editStreamChecks(request.input, pending.scope, 1, original);
  const back = editStreamChecks(pending, request.input.scope, 1, away);
  assert.equal(streamCoverageReady(request.input, 1, back, true), false);
  assert.equal(back.telemetry.proof, undefined); assert.equal(back.twin.proof, undefined);
  assert.ok(back.revision > original.revision);
});
test("additional choice and None Specify ABA retire both proofs while specified single-stream edits retain the peer", () => {
  const request = creationRequest(), original = complete(request);
  const oldBase = JSON.parse(original.base);
  delete oldBase.predicateContract;
  const obsolete = { ...original, base: canonicalJson(oldBase) };
  assert.equal(currentStreamResult(request.input, 1, obsolete, "telemetry"), undefined);
  assert.equal(streamCoverageReady(request.input, 1, obsolete, true), false);
  const none = { ...request.input.scope, additional: "none" as const };
  const away = editStreamChecks(request.input, none, 1, original);
  assert.equal(away.telemetry.proof, undefined); assert.equal(away.twin.proof, undefined);
  const back = editStreamChecks({ ...request.input, scope: none }, request.input.scope, 1, away);
  assert.equal(streamCoverageReady(request.input, 1, back, true), false);
  assert.ok(back.revision > original.revision);
  const pending = { ...request.input, scope: { ...request.input.scope, additional: "unselected" as const } };
  assert.equal(currentStreamResult(pending, 1, original, "telemetry"), undefined);
  assert.equal(assembleStreamCoverage(pending, 1, original), undefined);
  const edit = editStreamChecks(request.input, { ...request.input.scope, telemetry: "true" }, 1, original);
  assert.equal(edit.twin.proof, original.twin.proof);
});
test("editing one filter preserves current peer proof but invalidates assembled coverage until only that stream is rechecked", () => {
  const request = creationRequest(), both = complete(request);
  const scope = { ...request.input.scope, telemetry: "tostring(headers[\"IoTConnectionDeviceId\"]) == \"d2\"" }, input = { ...request.input, scope };
  const edited = editStreamChecks(request.input, scope, 1, both);
  assert.equal(edited.twin.proof, both.twin.proof);
  assert.equal(edited.telemetry.proof, undefined);
  assert.equal(edited.telemetry.edited, true);
  assert.equal(assembleStreamCoverage(input, 1, edited), undefined);
  const begun = beginStreamCheck(input, 1, "telemetry", edited);
  const done = finishStreamCheck(input, begun.checks, begun.ticket, { result: result(request, "telemetry") });
  assert.equal(done.twin.proof, both.twin.proof);
  assert.equal(streamCoverageReady(input, 1, done, false), true);
});
test("concurrent local checks can finish out of order without awarding the unchecked peer proof", () => {
  const request = creationRequest();
  const telemetry = beginStreamCheck(request.input, 1, "telemetry");
  const twin = beginStreamCheck(request.input, 1, "twin", telemetry.checks);
  const first = finishStreamCheck(request.input, twin.checks, twin.ticket, { result: result(request, "twin") });
  assert.equal(first.telemetry.phase, "checking");
  assert.equal(assembleStreamCoverage(request.input, 1, first), undefined);
  const both = finishStreamCheck(request.input, first, telemetry.ticket, { result: result(request, "telemetry") });
  assert.equal(streamCoverageReady(request.input, 1, both, false), true);
});
test("filter ABA failure retry and out-of-order callbacks never restore old or failed proof", () => {
  const request = creationRequest(), original = request.input;
  const first = beginStreamCheck(original, 1, "telemetry", complete(request));
  const changed = { ...original, scope: { ...original.scope, telemetry: "true" } };
  const away = editStreamChecks(original, changed.scope, 1, first.checks);
  const back = editStreamChecks(changed, original.scope, 1, away);
  assert.equal(finishStreamCheck(original, back, first.ticket, { result: result(request, "telemetry") }), back);
  const retry = beginStreamCheck(original, 1, "telemetry", back);
  const failed = finishStreamCheck(original, retry.checks, retry.ticket, { error: "Synthetic HTTP403" });
  assert.equal(assembleStreamCoverage(original, 1, failed), undefined);
  assert.equal(finishStreamCheck(original, failed, retry.ticket, { result: result(request, "telemetry") }), failed);
  const next = beginStreamCheck(original, 1, "telemetry", failed);
  assert.equal(finishStreamCheck(original, next.checks, retry.ticket, { result: result(request, "telemetry") }), next.checks);
  assert.equal(finishStreamCheck(original, next.checks, next.ticket, { result: result(request, "telemetry") }).telemetry.phase, "checked");
});
test("cancel retires only the owned pending stream and retains a completed peer across remount and retry", () => {
  const request = creationRequest();
  for (const stream of ["telemetry", "twin"] as const) {
    const peer = stream === "telemetry" ? "twin" : "telemetry";
    const initial = complete(request);
    const started = beginStreamCheck(request.input, 1, stream, initial);
    const canceled = cancelStreamCheck(started.checks, started.ticket);
    assert.equal(canceled.base, initial.base);
    assert.equal(canceled.revision, initial.revision);
    assert.equal(canceled[stream].phase, "unchecked");
    assert.ok(canceled[stream].generation > started.ticket.generation);
    assert.equal(canceled[peer].proof, initial[peer].proof);
    assert.equal(currentStreamChecks(request.input, 1, canceled), canceled, "Unchanged remount must retain the recheckable state and peer.");
    assert.equal(streamCoverageReady(request.input, 1, canceled, false), false);
    assert.equal(finishStreamCheck(request.input, canceled, started.ticket, { result: result(request, stream) }), canceled);
    const retry = beginStreamCheck(request.input, 1, stream, canceled);
    const done = finishStreamCheck(request.input, retry.checks, retry.ticket, { result: result(request, stream) });
    assert.equal(streamCoverageReady(request.input, 1, done, false), true);
    assert.equal(finishStreamCheck(request.input, done, started.ticket, { error: "Late old request" }), done);
  }
});
test("stale ownership tickets cannot cancel a newer request or a completed stream", () => {
  const request = creationRequest();
  const first = beginStreamCheck(request.input, 1, "telemetry", complete(request));
  const newer = beginStreamCheck(request.input, 1, "telemetry", first.checks);
  assert.equal(cancelStreamCheck(newer.checks, first.ticket), newer.checks);
  assert.equal(cancelStreamCheck(newer.checks, { ...newer.ticket, revision: newer.ticket.revision - 1 }), newer.checks);
  assert.equal(cancelStreamCheck(newer.checks, { ...newer.ticket, base: "other-source" }), newer.checks);
  const completed = finishStreamCheck(request.input, newer.checks, newer.ticket, { result: result(request, "telemetry") });
  assert.equal(cancelStreamCheck(completed, newer.ticket), completed);
});
test("retiring two owned pending checks advances each without fabricating either proof", () => {
  const request = creationRequest();
  const telemetry = beginStreamCheck(request.input, 1, "telemetry");
  const twin = beginStreamCheck(request.input, 1, "twin", telemetry.checks);
  const one = cancelStreamCheck(twin.checks, telemetry.ticket);
  assert.equal(one.twin, twin.checks.twin);
  const both = cancelStreamCheck(one, twin.ticket);
  assert.equal(both.telemetry.phase, "unchecked"); assert.equal(both.twin.phase, "unchecked");
  assert.equal(both.telemetry.proof, undefined); assert.equal(both.twin.proof, undefined);
  assert.equal(assembleStreamCoverage(request.input, 1, both), undefined);
});
test("shared source model roles depth reference mode operator and source revision invalidate both proofs including ABA", () => {
  const request = creationRequest(), input = request.input, original = complete(request);
  const variants: ModelingInput[] = [
    { ...input, source: { ...input.source, twinTable: "OtherProperties" } },
    { ...input, model: { ...input.model, description: "Changed meaning" } },
    { ...input, fields: input.fields.map(field => ({ ...field, included: !field.included })) },
    { ...input, fields: input.fields.map(field => setFieldRole(field, field.role === "telemetry" ? "twin" : "telemetry")) },
    { ...input, limit: 500 },
    { ...input, scope: { ...input.scope, mode: "reference" } },
    { ...input, scope: { ...input.scope, operator: "or" } },
    { ...input, scope: { ...input.scope, reference: { id: "urn:other" } } },
  ];
  for (const changed of variants) {
    const away = currentStreamChecks(changed, 1, original);
    assert.equal(away.telemetry.proof, undefined); assert.equal(away.twin.proof, undefined);
    const back = currentStreamChecks(input, 1, away);
    assert.equal(streamCoverageReady(input, 1, back, false), false);
    assert.ok(back.revision > original.revision);
  }
  assert.equal(streamCoverageReady(input, 2, original, false), false);
});
test("explicit waiting still needs both checks and valid raw binding probes; empty properties stay valid sparse evidence", () => {
  const waiting = creationRequest(true), both = complete(waiting);
  assert.equal(streamCoverageReady(waiting.input, 1, both, false), false);
  assert.equal(streamCoverageReady(waiting.input, 1, both, true), true);
  assert.equal(streamCoverageReady(waiting.input, 1, { ...both, twin: { generation: 100, phase: "unchecked" } }, true), false);
  const invalid = JSON.parse(JSON.stringify(both)) as StreamChecks;
  invalid.telemetry.proof!.result.probe.payloads[0].headers = {};
  assert.equal(streamCoverageReady(waiting.input, 1, invalid, true), false);
  const sparse = creationRequest();
  sparse.coverage.twin = { rawRows: 0, observations: 0, skippedWithoutReportedProperties: 0, payloads: [] };
  assert.equal(streamCoverageReady(sparse.input, 1, complete(sparse), false), true);
});
test("verified zero page policy requires current both-stream proof and nonempty valid raw sources without consent", async () => {
  const request = creationRequest(true), checks = complete(request);
  assert.equal(modelCoverageReady(request.input, 1, checks), true);
  const obsolete = structuredClone(checks), base = JSON.parse(obsolete.base);
  base.predicateContract = "original-raw-row-v1";
  obsolete.base = canonicalJson(base);
  assert.equal(modelCoverageReady(request.input, 1, obsolete), false);
  assert.equal(modelCoverageReady(request.input, 2, checks), false);
  assert.equal(modelCoverageReady(request.input, 1, { ...checks, twin: { generation: 100, phase: "unchecked" } }), false);
  for (const stream of ["telemetry", "twin"] as const) {
    const invalid = structuredClone(checks);
    invalid[stream].proof!.result.probe.payloads = [];
    assert.equal(modelCoverageReady(request.input, 1, invalid), false);
    const mock = createCreationMock(request);
    const fresh = structuredClone(request.coverage);
    fresh.sourceProbe![stream].payloads = [];
    mock.options.inspect = async () => fresh;
    const creation = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(creation.phase, "blocked");
    assert.match(creation.error!, /raw-source proof is empty, invalid or incomplete/);
    assert.equal(mock.writes.length, 0);
  }
  const mock = createCreationMock(request);
  for (const table of mock.tables.values()) table.columns[0].type = "dynamic";
  const created = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(created.phase, "completed");
  assert.equal(created.awaitingData, true);
  assert.equal(mock.writes.length, 9);
});
function rawResponse(stream: SampleStream, empty = false, ingestionTime: unknown = "2026-10-17T12:00:00Z"): Response {
  const value = stream === "telemetry" ? { temperature: 20 } : { properties: { reported: { state: "running" } } };
  return new Response(JSON.stringify({ Tables: [{
    Columns: [{ ColumnName: "data", ColumnType: "string" }, { ColumnName: "headers", ColumnType: "dynamic" },
      { ColumnName: "user_headers", ColumnType: "dynamic" }, { ColumnName: "__coverage_ingestion_time", ColumnType: "datetime" }],
    Rows: empty ? [] : [[JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
      event_kind: stream === "telemetry" ? "measurement" : "lifecycle", event_time: "2026-10-16T12:00:00Z", payload: { value } }),
    { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "", IoTSubject: "" }, {}, ingestionTime]],
  }] }));
}
test("each local reader bounds raw input then scopes only its named stream in ingestion order with cancellation", async () => {
  const input = creationRequest().input;
  for (const stream of ["telemetry", "twin"] as const) {
    const queries: string[] = [], abort = new AbortController();
    const read: typeof fetch = async (_url, init) => {
      assert.equal(init?.signal, abort.signal);
      const headers = new Headers(init?.headers);
      assert.equal(headers.get("Accept"), "application/json");
      const body = JSON.parse(String(init?.body)); queries.push(body.csl);
      return rawResponse(stream);
    };
    const inspected = await inspectStreamCoverage(input, "mock-only", stream, abort.signal, read);
    assert.equal(queries.length, 2);
    const selected = stream === "telemetry" ? "Telemetry" : "Twins", other = stream === "telemetry" ? "Twins" : "Telemetry";
    assert.ok(queries.every(query => query.includes(`['${selected}']`) && !query.includes(`['${other}']`)));
    assert.ok(queries.every(query => query.includes("| top 10000 by __coverage_ingestion_time desc nulls first")));
    assert.match(queries[0], /top 100 by __coverage_ingestion_time desc nulls first/);
    assert.ok(queries[1].indexOf("| where true\n") < queries[1].indexOf("| extend __observation"));
    assert.match(queries[1], /\| top 100 by __coverage_ingestion_time desc nulls first/);
    assert.ok(queries[1].indexOf("| top 10000") < queries[1].indexOf("| extend data = parse_json"));
    assert.equal(inspected.sample.payloads.length, 1);
    assert.equal(inspected.probe.payloads.length, 1);
    abort.abort(); await assert.rejects(inspectStreamCoverage(input, "mock-only", stream, abort.signal, read), /cancel|abort/i);
  }
});
test("local reader keeps empty results distinct from permission failure and incompatible raw schema", async () => {
  const input = creationRequest().input;
  const empty = await inspectStreamCoverage(input, "mock-only", "twin", undefined, async () => rawResponse("twin", true));
  assert.equal(empty.sample.payloads.length, 0);
  await assert.rejects(inspectStreamCoverage(input, "mock-only", "telemetry", undefined,
    async () => new Response("Forbidden", { status: 403 })), /access|permission|403/i);
  await assert.rejects(inspectStreamCoverage(input, "mock-only", "telemetry", undefined,
    async () => new Response(JSON.stringify({ Tables: [{ Columns: [{ ColumnName: "wrong", ColumnType: "string" },
      { ColumnName: "headers", ColumnType: "dynamic" }, { ColumnName: "user_headers", ColumnType: "dynamic" },
      { ColumnName: "__coverage_ingestion_time", ColumnType: "datetime" }], Rows: [["x", {}, {}, "2026-10-17T12:00:00Z"]] }] }))),
  /data:string|device data/);
});

test("coverage refuses unavailable ingestion clocks instead of treating them as latest or widening the sample", async () => {
  const input = creationRequest().input;
  for (const clock of [null, "", "invalid", 123]) {
    let queries = 0;
    await assert.rejects(inspectStreamCoverage(input, "mock-only", "telemetry", undefined, async (_url, init) => {
      queries++;
      assert.match(JSON.parse(String(init?.body)).csl, /top 10000 by __coverage_ingestion_time desc nulls first/);
      return rawResponse("telemetry", false, clock);
    }), /ingestion timestamps are unavailable or invalid/);
    assert.equal(queries, 1, "An unverifiable source must not dispatch the filtered coverage query.");
  }
});

test("recent coverage uses exactly 10000 raw rows before model and predicate filters, then at most 100 matches by ingestion", async () => {
  assert.deepEqual(COVERAGE_SAMPLING_CONTRACT, { rawRowLimit: 10000, observationLimit: 100, order: "ingestion-time-desc" });
  for (const stream of ["telemetry", "twin"] as const) {
    for (const matching of [0, 5, 200]) {
      const request = creationRequest();
      const input = { ...request.input, limit: 1000, scope: { ...request.input.scope, mode: "reference" as const,
        telemetry: 'data["payload"]["eligible"] == true', twin: 'data["payload"]["eligible"] == true',
        reference: { id: request.input.model.id! } } };
      const total = 12000;
      const rows = Array.from({ length: total }, (_, index) => {
        const recent = index >= total - 10000;
        const value = stream === "telemetry" ? { temperature: index } : { properties: { reported: { state: String(index) } } };
        return {
          // Old malformed JSON must not be parsed merely to find matching data.
          data: !recent ? "{invalid old JSON" : JSON.stringify({
            schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
            event_kind: stream === "telemetry" ? "measurement" : "lifecycle",
            event_time: index === total - 1 ? "2020-01-01T00:00:00Z" : "2099-01-01T00:00:00Z",
            payload: { value, memberTypeIRI: index >= total - matching ? input.model.id : "urn:other",
              eligible: index >= total - matching },
          }),
          ingestedAt: new Date(Date.UTC(2026, 9, 17) + index * 1000).toISOString(),
        };
      });
      const queries: string[] = [];
      const read: typeof fetch = async (_url, init) => {
        const query = JSON.parse(String(init?.body)).csl as string;
        queries.push(query);
        const tops = [...query.matchAll(/\| top (\d+) by ([^\n]+)/g)];
        assert.deepEqual(tops.map(match => Number(match[1])), [10000, 100]);
        assert.ok(tops.every(match => match[2] === "__coverage_ingestion_time desc nulls first"));
        const rawBound = query.indexOf("| top 10000");
        const parse = query.indexOf("| extend data = parse_json");
        if (parse >= 0) {
          assert.ok(rawBound < parse && parse < query.indexOf(`| where ${effectiveScope(input.scope, stream)}`));
          assert.ok(query.includes("event_time = todatetime(__observation.event_time)"));
        }
        let selected = [...rows].reverse().slice(0, Number(tops[0][1]));
        if (parse >= 0) selected = selected.filter(row => {
          const observation = JSON.parse(row.data);
          return observation.payload.memberTypeIRI === input.model.id && observation.payload.eligible === true;
        });
        selected = selected.slice(0, Number(tops[1][1]));
        return new Response(JSON.stringify({ Tables: [{
          Columns: [{ ColumnName: "data", ColumnType: "string" }, { ColumnName: "headers", ColumnType: "dynamic" },
            { ColumnName: "user_headers", ColumnType: "dynamic" }, { ColumnName: "__coverage_ingestion_time", ColumnType: "datetime" }],
          Rows: selected.map(row => [row.data, { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "", IoTSubject: "" },
            {}, row.ingestedAt]),
        }] }));
      };
      const inspected = await inspectStreamCoverage(input, "mock-only", stream, undefined, read);
      assert.equal(queries.length, 2);
      assert.equal(inspected.probe.payloads.length, 100);
      assert.equal(inspected.sample.payloads.length, Math.min(matching, 100));
      const actual = inspected.sample.payloads.map(sample => stream === "telemetry" ? sample.value.temperature : Number(sample.value.state));
      assert.deepEqual(actual, Array.from({ length: Math.min(matching, 100) }, (_, index) => total - 1 - index));
      if (matching) assert.equal(inspected.sample.payloads[0].eventTime, "2020-01-01T00:00:00Z",
        "Coverage ordering must not rewrite or sort by the DO event clock.");
    }
  }
});

test("old all-history sampler proofs and recheck results cannot authorize native creation", async () => {
  const request = creationRequest(), oldFingerprint = JSON.parse(request.coverage.fingerprint);
  delete oldFingerprint.coverageSampling;
  const obsolete = { ...request.coverage, fingerprint: canonicalJson(oldFingerprint) };
  assert.equal(coverageReady(request.input, obsolete, false), false);
  const checks = complete(request), oldBase = JSON.parse(checks.base);
  delete oldBase.coverageSampling;
  assert.equal(modelCoverageReady(request.input, 1, { ...checks, base: canonicalJson(oldBase) }), false);
  for (const legacyAtRecheck of [false, true]) {
    const mock = createCreationMock(request);
    if (legacyAtRecheck) mock.options.inspect = async () => obsolete;
    const result = await new ModeledDataCreator().run(legacyAtRecheck ? request : { ...request, coverage: obsolete }, mock.service, true);
    assert.equal(result.phase, "blocked");
    assert.equal(mock.writes.length, 0);
    assert.match(result.error!, /processing scope|current model coverage/);
  }
});
test("single direct Create always runs full preflight before first mutation with no prior false-run prerequisite", async () => {
  const creator = new ModeledDataCreator();
  for (const suffix of ["", "_Again"]) {
    const request = creationRequest();
    request.names = Object.fromEntries(Object.entries(request.names).map(([key, value]) => [key, value + suffix])) as typeof request.names;
    const mock = createCreationMock(request);
    const outcome = await creator.run(request, mock.service, true);
    assert.equal(outcome.phase, "completed");
    const firstWrite = mock.calls.findIndex(call => /^\.(create|alter)\b/.test(call.csl));
    assert.ok(firstWrite > 0);
    const reads = mock.calls.slice(0, firstWrite);
    assert.ok(reads.some(call => call.csl.includes("principal roles")));
    assert.ok(reads.some(call => call.csl.includes(".show functions")));
    assert.ok(reads.some(call => call.csl.includes("['Telemetry'] schema as json")));
    assert.ok(reads.some(call => call.csl.includes("['Twins'] schema as json")));
    assert.equal(reads.filter(call => call.csl.endsWith("| getschema")).length, 2);
    assert.ok(!mock.writes.some(command => /policy streamingingestion|backfill=true/.test(command)));
  }
  const request = creationRequest(), blocked = createCreationMock(request);
  blocked.tables.get("Telemetry")!.streaming = true;
  const outcome = await creator.run(request, blocked.service, true);
  assert.equal(outcome.phase, "blocked"); assert.equal(blocked.writes.length, 0); assert.equal(outcome.uncertainCommand, undefined);
});
