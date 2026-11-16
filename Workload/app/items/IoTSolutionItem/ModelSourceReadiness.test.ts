// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { currentSourceAnalysis, modelSourceKey, SOURCE_ANALYSIS_LIMIT, SOURCE_PENDING, sourceActionCurrent,
  sourceSelectionStatus, type SourceAnalysis } from "./ModelSourceReadiness.ts";
import { readPayloadSamples } from "./ModelAuthoring.ts";
import { applyRoleHints, fieldValidation, inspectFieldEvidence, matchRoleEvidence, prepareRoleEvidence, setFieldRole } from "./HubModeling.ts";
import { creationRequest } from "../../../tests/creationMock.ts";

test("source-first readiness requires inspected current membership distinct tables and no loading or error", () => {
  const input = { key: "source", loading: false, error: "", eventhouseId: "eh", databaseId: "db",
    queryServiceUri: "https://mock.kusto.fabric.microsoft.com",
    eventhouseIds: ["eh"], databaseIds: ["db"], tables: ["ManualTelemetry", "ManualProperties"],
    telemetry: "ManualTelemetry", properties: "ManualProperties" };
  assert.equal(sourceSelectionStatus(input).reason, "Source tables selected.");
  assert.equal(sourceSelectionStatus(input).ready, true);
  for (const change of [{ eventhouseIds: [] }, { databaseIds: [] }, { tables: [] },
    { telemetry: "" }, { properties: "" }, { queryServiceUri: "" }]) {
    assert.equal(sourceSelectionStatus({ ...input, ...change }).ready, false);
    assert.equal(sourceSelectionStatus({ ...input, ...change }).reason, SOURCE_PENDING);
  }
  assert.equal(sourceSelectionStatus({ ...input, loading: true }).reason, "Wait for the source tables to finish loading.");
  assert.equal(sourceSelectionStatus({ ...input, error: "HTTP403" }).reason,
    "Resolve the source selection error before inspecting samples or uploading a model.");
  assert.equal(sourceSelectionStatus({ ...input, properties: input.telemetry }).reason,
    "Choose different tables for telemetry and reported properties.");
});

test("source-first action snapshots reject stale readiness source identities and same-source refresh revisions", () => {
  const start = { key: "source-a", ready: true, reason: "Source tables selected.", revision: 1 };
  assert.equal(sourceActionCurrent(start, { ...start }), true);
  assert.equal(sourceActionCurrent({ ...start, ready: false }, start), false);
  assert.equal(sourceActionCurrent(start, { ...start, ready: false }), false);
  assert.equal(sourceActionCurrent(start, { ...start, key: "source-b" }), false);
  assert.equal(sourceActionCurrent(start, { ...start, revision: 2 }), false);
  assert.equal(sourceActionCurrent(start, { ...start, revision: 3, ready: true }), false,
    "Returning to ready after a refresh must not revive an older asynchronous action.");
});

test("missing-role summary can omit repeated sentences without weakening validation or hiding other blockers", () => {
  const fields = creationRequest().input.fields.map(field => setFieldRole(field, undefined));
  const missing = fields.filter(field => field.included && !field.role);
  assert.equal(missing.length, 2);
  assert.equal(fieldValidation(fields).length, 2);
  assert.deepEqual(fieldValidation(fields, { includeRoleErrors: false }), []);
  const collision = fields.map(field => ({ ...field, column: "duplicate" }));
  assert.ok(fieldValidation(collision, { includeRoleErrors: false }).some(error => error.includes('map to "duplicate"')));
  const excluded = fields.map(field => ({ ...field, included: false }));
  assert.deepEqual(fieldValidation(excluded, { includeRoleErrors: false }), ["Select at least one field to continue"]);
  assert.equal(fieldValidation(fields.map(field => setFieldRole(field, "telemetry"))).length, 0);
});

test("analyzed source samples require the complete current source key and selection revision, including away-and-back", () => {
  const request = creationRequest();
  const context = { modelEhId: "eh", modelDbId: "db", modelDbName: "Database", modelQueryServiceUri: request.input.source.queryServiceUri,
    rawTelemetryTable: "Telemetry", rawPropertiesTable: "Twins", hubSourceSelectionRevision: 3 };
  const analysis: SourceAnalysis = { key: modelSourceKey("workspace", context), selectionRevision: 3, limit: SOURCE_ANALYSIS_LIMIT,
    telemetry: request.coverage.telemetry, twin: request.coverage.twin,
    roleEvidence: prepareRoleEvidence(request.coverage.telemetry, request.coverage.twin), analyzedAt: "2026-10-25T00:00:00Z" };
  const selected = { ...context, hubSourceAnalysis: analysis };
  assert.equal(currentSourceAnalysis("workspace", selected), analysis);
  for (const change of [{ modelEhId: "other" }, { modelDbId: "other" }, { modelDbName: "other" },
    { modelQueryServiceUri: "https://other.example" }, { rawTelemetryTable: "OtherTelemetry" },
    { rawPropertiesTable: "OtherTwins" }, { hubSourceSelectionRevision: 5 }]) {
    assert.equal(currentSourceAnalysis("workspace", { ...selected, ...change }), undefined);
  }
  assert.equal(currentSourceAnalysis("other-workspace", selected), undefined);
  assert.equal(currentSourceAnalysis("workspace", { ...selected, hubSourceSelectionRevision: 4 }), undefined);
  assert.equal(currentSourceAnalysis("workspace", { ...selected, hubSourceAnalysis: { ...analysis, limit: 100 } }), undefined);
  assert.equal(currentSourceAnalysis("workspace", { ...selected, hubSourceAnalysis: { ...analysis, roleEvidence: undefined } }), undefined);
});

test("normal source analysis requests latest1000 raw rows independently for both streams and does not infer processing coverage", async () => {
  const source = creationRequest().input.source, queries: string[] = [];
  const result = await readPayloadSamples(source, "synthetic-only", SOURCE_ANALYSIS_LIMIT, undefined, async (_url, init) => {
    const body = JSON.parse(String(init?.body)); queries.push(body.csl);
    assert.equal(body.db, source.databaseName);
    assert.match(body.csl, /\| top 1000 by ingestion_time\(\) desc\n\| project data, headers, user_headers$/);
    assert.doesNotMatch(body.csl, /\| where|memberTypeIRI|mv-expand/);
    const twin = body.csl.includes("['Twins']");
    const row = [JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
      event_time: "2026-10-25T00:00:00Z", event_kind: twin ? "lifecycle" : "measurement",
      payload: { value: twin ? { properties: { reported: { state: "idle" } } } : { temperature: 22 } } }),
      { IoTConnectionDeviceId: "unit-device" }, {}];
    return new Response(JSON.stringify({ Tables: [{ Columns: [["data", "string"], ["headers", "dynamic"], ["user_headers", "dynamic"]]
      .map(([ColumnName, ColumnType]) => ({ ColumnName, ColumnType })), Rows: Array.from({ length: 1000 }, () => row) }] }));
  });

  assert.equal(queries.length, 2);
  assert.equal(queries.filter(query => query.includes("['Telemetry']")).length, 1);
  assert.equal(queries.filter(query => query.includes("['Twins']")).length, 1);
  assert.equal(result.telemetry.rawRows, 1000); assert.equal(result.twin.rawRows, 1000);
  assert.equal(result.telemetry.payloads.length, 1000); assert.equal(result.twin.payloads.length, 1000);
  assert.ok(!("fingerprint" in result) && !("kind" in result), "Unfiltered samples must not be represented as processing coverage.");
});

test("model-free source role profile reuses native whole-value matching for roots components and ambiguity", () => {
    const request = creationRequest(), source = request.coverage.sourceProbe!;
    const base = source.telemetry.payloads[0], reported = {
      measurement: { nested: 10 }, a: { __t: "c", measurement: { nested: 20 } }, b: { __t: "c", measurement: [1, 2] },
    };
    const telemetry = { ...source.telemetry, payloads: [
      { ...base, value: { measurement: { nested: 1 } } },
      { ...base, component: "a", value: { measurement: { nested: 2 } } },
      { ...base, value: { unknownOnly: 3 }, native: { ...base.native!, profile: "unknown" as const } },
    ] };
    const twin = { ...source.twin, payloads: [{ ...source.twin.payloads[0], value: reported, reported }] };
    const template = request.input.fields[0];
    const fields = [
      { ...template, id: "root-measurement", path: ["measurement"], component: undefined },
      { ...template, id: "component-a", path: ["measurement"], component: "a" },
      { ...template, id: "component-b", path: ["measurement"], component: "b" },
      { ...template, id: "root-envelope", path: ["a"], component: undefined },
      { ...template, id: "unknown-only", path: ["unknownOnly"], component: undefined },
    ].map(field => ({ ...field, role: undefined, roleSource: undefined }));
    const profile = prepareRoleEvidence(telemetry, twin);
    assert.ok(profile.bindings.every(binding => binding.path.length === 1));
    assert.ok(!JSON.stringify(profile).includes('"nested"'), "Complex values must not become leaf bindings.");
    for (const candidates of [fields, fields.map(field => ({ ...field, included: !field.component }))]) {
      assert.deepEqual(matchRoleEvidence(candidates, profile), inspectFieldEvidence(candidates, telemetry, twin));
    }
    const evidence = matchRoleEvidence(fields, profile);
    assert.deepEqual(evidence.map(item => item.state), ["both", "both", "twin", "missing", "missing"]);
    assert.equal(matchRoleEvidence(fields.map(field => ({ ...field, included: !field.component })), profile)[3].twin, 1);
    const explicit = [setFieldRole(fields[2], "telemetry"), { ...fields[2], role: "telemetry" as const, roleSource: "model" as const }];
    assert.ok(applyRoleHints(explicit, evidence).every(field => field.role === "telemetry"));
    assert.deepEqual(matchRoleEvidence(fields, prepareRoleEvidence({ ...telemetry, payloads: [] }, { ...twin, payloads: [] }))
      .map(item => item.state), ["missing", "missing", "missing", "missing", "missing"]);
});
