// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import crypto from "node:crypto";
import Ajv from "ajv";
import addFormats from "ajv-formats";
import tmSchema from "./schemas/wot-tm.schema.json" with { type: "json" };
import {
  applyRoleHints, effectiveScope, fieldValidation, fieldValue, fromThingModel, inspectFieldEvidence,
  modelValidation, orderedModeledFields, setFieldRole, validatePredicate,
} from "./HubModeling.ts";
import { canonicalJson, decodeDoSamples, inferPayloadSchema, isJsonObject } from "./ModelAuthoring.ts";
import { parseThingModel, resolveThingModel } from "./ThingModelParser.ts";
import { buildModeledDataPlan, semanticDocumentation } from "./ModeledDataPlan.ts";
import { ModeledDataCreator } from "./ModeledDataCreation.ts";
import { createCreationMock } from "../../../tests/creationMock.ts";
import { readColumnSemantics } from "./FieldSemantics.ts";
import { mountedWotVariationKeys, wotVariations, wotVariationMockPayloads, wotVariationRequest } from "../../../tests/wotVariationFixtures.ts";

const outcomes: Record<string, unknown>[] = [];
const guardOutcomes: Record<string, unknown>[] = [];
const codecOutcomes: Record<string, unknown>[] = [];
const ajv = new Ajv({ allErrors: true, strict: false });
addFormats(ajv);
const validateTmSchema = ajv.compile(tmSchema);
process.once("exit", () => {
  if (process.env.WOT_VARIATION_RESULTS) fs.writeFileSync(process.env.WOT_VARIATION_RESULTS,
    JSON.stringify({ provenance: "agent-authored", count: outcomes.length, cases: outcomes, creationGuards: guardOutcomes,
      mountedFixtureCodec: codecOutcomes,
      schemaBoundary: "Vendored W3C TM JSON Schema validation only; not full JSON-LD or every normative WoT conformance assertion." }, null, 2) + "\n");
});

test("mounted variation fixture codec retains raw lifecycle reported envelopes and exact whole fields in all six cases", () => {
  for (const key of mountedWotVariationKeys) {
    const variation = wotVariations.find(candidate => candidate.key === key)!;
    const request = wotVariationRequest(variation), payloads = wotVariationMockPayloads(request);
    const original = canonicalJson(variation.document);
    const roundtrip = (stream: "telemetry" | "twin") => decodeDoSamples([{
      data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
        event_time: "2026-10-01T12:00:00Z", event_kind: stream === "telemetry" ? "measurement" : "lifecycle",
        payload: { value: payloads[stream] } }),
      headers: { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "", IoTSubject: "" },
      user_headers: { hubName: "mock-hub" },
    }], stream, 100);
    const telemetry = roundtrip("telemetry"), twin = roundtrip("twin");
    assert.equal(telemetry.payloads.length, 1, `${key}: telemetry fixture disappeared`);
    assert.equal(twin.payloads.length, 1, `${key}: raw lifecycle fixture needs properties.reported`);
    assert.equal(twin.skippedWithoutReportedProperties, 0);
    assert.deepEqual(twin.payloads[0].reported, request.coverage.twin.payloads[0].reported);
    const hints = applyRoleHints(request.input.model.fields,
      inspectFieldEvidence(request.input.model.fields, telemetry, twin));
    assert.deepEqual(hints.map(field => field.role), request.input.fields.map(field => field.role), key);
    for (const [index, field] of request.input.fields.entries()) {
      const actual = fieldValue((field.role === "twin" ? twin : telemetry).payloads[0], field, field.role!);
      assert.equal(actual.found, true, `${key}: ${field.path[0]} missing`);
      assert.deepEqual(actual.value, variation.fields![index].value);
    }
    assert.deepEqual(payloads.telemetry, request.coverage.telemetry.payloads[0].value, "Telemetry fixture must be unchanged.");
    assert.equal(canonicalJson(variation.document), original);
    codecOutcomes.push({ key, rawPayloadValues: payloads, telemetryPayloads: telemetry.payloads.length,
      twinPayloads: twin.payloads.length, skippedWithoutReportedProperties: twin.skippedWithoutReportedProperties,
      fields: hints.map(field => ({ path: field.path, role: field.role })), originalDocumentUnchanged: true, passed: true });
  }
});

for (const variation of wotVariations) {
  test(`WoT variation ${variation.key}: expected ${variation.expected}`, async () => {
    const original = canonicalJson(variation.document), fetcher = globalThis.fetch;
    let networkCalls = 0;
    globalThis.fetch = async () => { networkCalls++; throw new Error("A model variation must not perform external fetch."); };
    const result: Record<string, unknown> = { key: variation.key, title: variation.title, provenance: variation.provenance,
      input: variation.document, inputSha256: crypto.createHash("sha256").update(original).digest("hex"), expected: variation.expected,
      validAgainstTmJsonSchema: validateTmSchema(variation.document), schemaErrors: validateTmSchema.errors };
    try {
      if (variation.expected === "root-id-required" || variation.expected === "rejected") {
        assert.throws(() => resolveThingModel(variation.document), (error: unknown) => error instanceof Error &&
          error.message.toLowerCase().includes(variation.error!.toLowerCase()));
        result.actual = variation.expected; result.expectedRejection = true;
      } else {
        const parsed = resolveThingModel(variation.document);
        if (variation.expected === "unsupported") {
          assert.ok(parsed.blockers.some(blocker => blocker.toLowerCase().includes(variation.error!.toLowerCase())), parsed.blockers.join("; "));
          assert.deepEqual(parsed.fields, []); assert.equal(parsed.effectiveDocument, undefined);
          assert.throws(() => fromThingModel(parsed), /dependencies/);
          result.actual = "unsupported"; result.blockers = parsed.blockers;
        } else {
          assert.deepEqual(parsed.blockers, []);
          const model = fromThingModel(parsed);
          assert.equal(model.id, variation.document.id);
          assert.deepEqual(model.document, variation.document);
          assert.deepEqual(model.fields.map(field => ({ name: field.path[0], type: field.kustoType })),
            variation.fields!.map(field => ({ name: field.name, type: field.type })));
          assert.ok(model.fields.every(field => field.path.length === 1 && field.role === undefined &&
            field.roleSource === undefined && field.component === undefined));
          if (variation.expected === "selection-conflict") {
            const fields = model.fields.map(field => setFieldRole(field, "telemetry"));
            assert.ok(fieldValidation(fields).some(error => error.includes("Two selected fields")));
            assert.equal(new Set(fields.map(field => field.id)).size, fields.length);
            assert.deepEqual(fieldValidation(fields.map((field, index) => ({ ...field, included: index === 0 }))), []);
            result.actual = "selection-conflict";
          } else if (variation.expected === "no-data-fields") {
            assert.equal(model.fields.length, 0);
            assert.deepEqual(fieldValidation(model.fields), ["Select at least one field to continue"]);
            if (variation.key === "actions-only") assert.ok(isJsonObject(model.document) && "actions" in model.document);
            else assert.ok(model.ignored?.length, "The all-ignored fixture must identify its actual skipped parts.");
            result.actual = "no-data-fields";
          } else {
            const request = wotVariationRequest(variation), input = request.input;
            assert.deepEqual(modelValidation(input.model), []);
            assert.deepEqual(fieldValidation(input.fields), []);
            assert.deepEqual(request.coverage.nativeIssues, []);
            for (const [index, field] of input.fields.entries()) {
              const value = fieldValue(request.coverage[field.role!].payloads[0], field, field.role!);
              assert.equal(value.found, true);
              assert.deepEqual(value.value, variation.fields![index].value);
            }
            const hints = applyRoleHints(model.fields, inspectFieldEvidence(model.fields, request.coverage.telemetry, request.coverage.twin));
            assert.deepEqual(hints.map(field => field.role), input.fields.map(field => field.role));
            const ambiguous = input.fields.map(field => ({ fieldId: field.id, telemetry: 1, twin: 1, state: "both" as const }));
            assert.ok(applyRoleHints(model.fields, ambiguous).every(field => field.role === undefined));
            assert.deepEqual(applyRoleHints(input.fields, ambiguous).map(field => field.role), input.fields.map(field => field.role));
            const plan = buildModeledDataPlan(input, request.coverage, false, request.names);
            const modeled = plan.commands.find(command => command.kind === "table" && command.target === request.names.modeled)!;
            assert.equal(modeled.kind, "table");
            if (modeled.kind !== "table") throw new Error("The modeled schema was not generated.");
            assert.deepEqual(modeled.columns, [{ name: "deviceId", type: "string" }, { name: "event_time", type: "datetime" },
              ...orderedModeledFields(input.fields).map(field => ({ name: field.column, type: field.kustoType }))]);
            assert.equal(plan.commands.filter(command => command.kind === "table" || command.kind === "view").length, 3);
            assert.equal(plan.commands.filter(command => command.kind === "policy").length, 2);
            const view = plan.commands.find(command => command.kind === "view")!;
            assert.equal(view.kind, "view");
            if (view.kind !== "view") throw new Error("Missing LKV view.");
            assert.match(view.query, /ingestion_time\(\) > ago\(30d\)/);
            assert.match(view.query, /arg_max\(event_time, propertyValue\) by deviceId, propertyName/);
            assert.deepEqual(plan.commands.filter(command =>
              command.kind === "tableDoc" || command.kind === "viewDoc" || command.kind === "columnDocs")
              .map(command => [command.kind, command.target]).sort(), [
              ["tableDoc", request.names.twinChanges], ["viewDoc", request.names.twinState],
              ["tableDoc", request.names.modeled], ["columnDocs", request.names.modeled],
            ].sort());
            const docs = semanticDocumentation(input.model, input.fields);
            assert.ok(docs.table.includes(String(variation.document.id)));
            for (const field of input.fields) {
              const info = readColumnSemantics(docs.columns[field.column]);
              assert.ok(info?.version === 3);
              assert.equal(info.field.facts.role, field.role === "twin" ? "twin property (last-known whole received value)" : "telemetry");
              assert.ok(!/binding|selected fields|output mapping/i.test(docs.columns[field.column]));
            }
            if (variation.key === "number-unit") {
              const info = readColumnSemantics(docs.columns.pressure);
              assert.ok(info?.version === 3);
              assert.equal(info.field.facts.unit, "kPa");
            }
            if (variation.key === "localized") {
              assert.equal(input.model.title, "US model"); assert.equal(input.model.documentVersion, "3.4.0");
              assert.equal(input.fields[0].title, "US temperature"); assert.ok(!("registryVersion" in input.model));
            }
            const schema = inferPayloadSchema([request.coverage.telemetry.payloads[0].value]);
            assert.equal(schema.$schema, "https://json-schema.org/draft/2020-12/schema");
            assert.ok(!JSON.stringify(schema).includes(String(variation.document.id)));
            const mock = createCreationMock(request);
            const created = await new ModeledDataCreator().run(request, mock.service, true);
            assert.equal(created.phase, "completed", created.error);
            assert.equal(created.objects.length, 3);
            assert.ok(mock.writes.every(command => !/backfill\s*=\s*true|set-or-append|sourceHubId/.test(command)));
            result.actual = "supported"; result.fields = input.fields.map(field => ({
              id: field.id, path: field.path, column: field.column, type: field.kustoType, role: field.role,
            }));
            result.nativePlan = { outputs: 3, policies: 2, ingestionWindowDays: 30, fakeCreator: created.phase };
          }
        }
      }
      assert.equal(networkCalls, 0);
      assert.equal(canonicalJson(variation.document), original);
      result.sourceUnchanged = true; result.implicitFetches = networkCalls; result.passed = true;
    } catch (error: unknown) {
      result.actual = "unexpected-failure"; result.passed = false;
      result.error = error instanceof Error ? error.message : "Unknown test failure";
      throw error;
    } finally {
      globalThis.fetch = fetcher;
      outcomes.push(result);
    }
  });
}

test("variation reference scope preserves supplied IDs, ANDs independent stream predicates and rejects query injection", () => {
  const document = wotVariations.find(variation => variation.key === "localized")!.document;
  assert.equal(parseThingModel(document).id, document.id);
  const scope = { mode: "reference" as const, operator: "and" as const, reference: { id: String(document.id) },
    telemetry: 'deviceId == "device-1"', twin: 'user_headers.opType == "replaceTwin"' };
  assert.equal(effectiveScope(scope, "telemetry"), `(tostring(data["payload"]["memberTypeIRI"]) == ${JSON.stringify(document.id)}) and (${scope.telemetry})`);
  assert.equal(effectiveScope(scope, "twin"), `(tostring(data["payload"]["memberTypeIRI"]) == ${JSON.stringify(document.id)}) and (${scope.twin})`);
  assert.equal(effectiveScope(scope, "telemetry", true), scope.telemetry);
  for (const predicate of ['true | take 1', 'true; .drop table x', 'true // comment', 'deviceId == "unclosed']) {
    assert.throws(() => validatePredicate(predicate));
  }
});

test("variation creation guards reject missing roles, stale reference scope, raw-name collision and invalid native identities before writes", async () => {
  const supported = wotVariations.find(variation => variation.key === "complex-whole")!;
  for (const negative of ["role", "scope", "name", "native"] as const) {
    const request = wotVariationRequest(supported), originalDocument = canonicalJson(request.input.model.document);
    const mock = createCreationMock(request);
    if (negative === "role") request.input.fields[0] = setFieldRole(request.input.fields[0], undefined);
    if (negative === "scope") request.input.scope = { ...request.input.scope, mode: "reference", reference: { id: "urn:copilot:test:changed" } };
    if (negative === "name") request.names.modeled = request.input.source.telemetryTable;
    if (negative === "native") request.coverage.telemetry.payloads[0].headers = { IoTConnectionDeviceId: 42 };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "blocked", `${negative}: ${result.error}`);
    assert.equal(mock.writes.length, 0);
    assert.equal(canonicalJson(request.input.model.document), originalDocument);
    guardOutcomes.push({ negative, expected: "blocked/zero writes", actual: result.phase, writes: mock.writes.length,
      error: result.error, originalDocumentUnchanged: true });
  }
});
