// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import { requireThingModelId, resolveThingModel, selectThingModelRoot } from "./ThingModelParser.ts";
import { fromThingModel, modelValidation, modelingFingerprint, assignableReference } from "./HubModeling.ts";
import { canonicalJson, isJsonObject, type JsonObject, type JsonValue } from "./ModelAuthoring.ts";
import { assignmentOwner } from "./HubAssignment.ts";
import { ModeledDataCreator } from "./ModeledDataCreation.ts";
import { createCreationMock } from "../../../tests/creationMock.ts";
import { localBundle, localBundleRoot, localDefinitions, localFixtureRequest, localNegativeFixtures,
  localPositiveFixtures, localSingleModel } from "../../../tests/localWotFixtures.ts";

const outcomes: Record<string, unknown>[] = [];
process.once("exit", () => {
  if (process.env.LOCAL_WOT_RESULTS) fs.writeFileSync(process.env.LOCAL_WOT_RESULTS,
    JSON.stringify({ cases: outcomes, count: outcomes.length, scope: "Agent-authored pure parser/model/creator tests; no external fetch." }, null, 2) + "\n");
});
for (const fixture of localPositiveFixtures) test(`local WoT ${fixture.key} preserves full upload own root ID roles and native creation`, async () => {
  const original = canonicalJson(fixture.document), fetcher = globalThis.fetch;
  let fetches = 0;
  globalThis.fetch = async () => { fetches++; throw new Error("Local models must not fetch."); };
  try {
    const parsed = resolveThingModel(fixture.document);
    assert.deepEqual(parsed.blockers, []);
    const model = fromThingModel(parsed);
    assert.equal(model.id, fixture.rootId);
    assert.deepEqual(parsed.document, fixture.document); assert.deepEqual(model.document, fixture.document);
    assert.deepEqual(modelValidation(model), []); assert.deepEqual(assignableReference(model), { id: fixture.rootId });
    assert.deepEqual(model.fields.map(field => [field.path[0], field.kustoType]), [
      ["measurement", "dynamic"], ["history", "dynamic"], ["alarm", "string"],
    ]);
    assert.ok(model.fields.every(field => !field.role && !field.roleSource && field.path.length === 1));
    assert.equal(model.documentVersion, "root-version");
    const request = localFixtureRequest(fixture.document), mock = createCreationMock(request);
    const created = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(created.phase, "completed", created.error); assert.equal(mock.writes.length, 9);
    assert.equal(canonicalJson(fixture.document), original); assert.equal(fetches, 0);
    outcomes.push({ key: fixture.key, input: fixture.document, rootId: model.id, documentVersion: model.documentVersion,
      fields: model.fields, fullUploadUnchanged: true, implicitFetches: fetches, fakeCreator: created.phase, passed: true });
  } finally { globalThis.fetch = fetcher; }
});
for (const fixture of localNegativeFixtures) test(`local WoT ${fixture.key} rejection exposes no effective fields`, () => {
  try {
    const parsed = resolveThingModel(fixture.document);
    assert.ok(parsed.blockers.some(blocker => blocker.includes(fixture.error)), parsed.blockers.join("; "));
    assert.equal(parsed.effectiveDocument, undefined); assert.deepEqual(parsed.fields, []);
    assert.throws(() => fromThingModel(parsed), /dependencies/);
  } catch (error) {
    if (error instanceof assert.AssertionError) throw error;
    assert.ok(error instanceof Error && error.message.includes(fixture.error), String(error));
  }
  outcomes.push({ key: fixture.key, expectedRejection: fixture.error, input: fixture.document, passed: true });
});
test("bundle root and forward imports are independent of array order and never inherited from helper version", () => {
  assert.ok(Array.isArray(localBundle));
  for (const input of [localBundle, [...localBundle].reverse()]) {
    assert.equal(selectThingModelRoot(input).id, localBundleRoot.id);
    const parsed = resolveThingModel(input);
    assert.equal(parsed.id, localBundleRoot.id); assert.equal(parsed.documentVersion, "root-version");
  }
  const own: JsonObject = { ...localBundleRoot, id: "urn:copilot:test:no-own-version" }; delete own.version;
  const parsed = resolveThingModel([own, localDefinitions]);
  assert.equal(parsed.id, own.id); assert.equal(parsed.documentVersion, undefined);
  assert.equal(parsed.effectiveDocument!.version, undefined);
  outcomes.push({ key: "root-order-and-own-version", expectedRoot: localBundleRoot.id, orderIndependent: true,
    ownIdAndVersionNotInherited: true, passed: true });
});
test("self-member pointers and exact own-ID pointers do not remove the bundle root but whole-root self-cycles fail", () => {
  for (const location of ["", String(localBundleRoot.id)]) {
    const own: JsonObject = { ...localBundleRoot, schemaDefinitions: localDefinitions.schemaDefinitions,
      properties: { reading: { "tm:ref": `${location}#/schemaDefinitions/state` } }, events: {} };
    const input = [localDefinitions, own];
    assert.equal(selectThingModelRoot(input).id, own.id);
    const parsed = resolveThingModel(input);
    assert.deepEqual(parsed.blockers, []); assert.equal(parsed.fields[0].declaredType, "string");
    assert.equal(requireThingModelId(parsed.document), own.id);
  }
  const own: JsonObject = { ...localSingleModel, properties: { self: { "tm:ref": localSingleModel.id! } } };
  const parsed = resolveThingModel([own]);
  assert.ok(parsed.blockers.some(blocker => blocker.includes("Cyclic"))); assert.deepEqual(parsed.fields, []);
  outcomes.push({ key: "self-member-root-graph", pointerLocations: ["fragment-only", "exact-own-ID"],
    rootKept: true, wholeRootCycleBlocked: true, passed: true });
});
test("same-object escaped and URI-encoded JSON pointers and valid array indices resolve whole DataSchemas", () => {
  const input: JsonObject = { ...localSingleModel, schemaDefinitions: {
    "a/b~c": { type: "object", properties: { n: { type: "number" } } },
    choices: { oneOf: [{ type: "string" }, { type: "array", items: { type: "number" } }] },
  }, properties: { escaped: { "tm:ref": "#/schemaDefinitions/a~1b~0c" },
    encoded: { "tm:ref": "#%2FschemaDefinitions%2Fa~1b~0c" },
    array: { "tm:ref": "#/schemaDefinitions/choices/oneOf/1" } }, events: {} };
  const parsed = resolveThingModel(input);
  assert.deepEqual(parsed.blockers, []);
  assert.deepEqual(parsed.fields.map(field => field.declaredType), ["object", "object", "array"]);
  for (const pointer of ["#/schemaDefinitions/choices/oneOf/01", "#/schemaDefinitions/a~2b", "#not-a-pointer", "#/schemaDefinitions/choices/oneOf/9"]) {
    const result = resolveThingModel({ ...input, properties: { bad: { "tm:ref": pointer } } });
    assert.ok(result.blockers.length); assert.equal(result.effectiveDocument, undefined); assert.deepEqual(result.fields, []);
  }
  outcomes.push({ key: "escaped-encoded-array-pointers", input, actualTypes: parsed.fields.map(field => field.declaredType),
    invalidArrayIndexEscapeAndFragmentBlocked: true, passed: true });
});
test("local imports patch metadata recursively without mutating definitions or treating JSON Schema ref as merge patch", () => {
  const input: JsonObject = { ...localSingleModel, schemaDefinitions: { reading: { type: "string", title: "Remove",
    enum: ["a", "b"], descriptions: { en: "Remove", fr: "Keep" } } },
  properties: { state: { "tm:ref": "#/schemaDefinitions/reading", title: null, enum: ["b"], descriptions: { en: null }, unit: "supplied" } }, events: {} };
  const before = canonicalJson(input), parsed = resolveThingModel(input);
  assert.deepEqual(parsed.blockers, []);
  assert.deepEqual((parsed.effectiveDocument!.properties as JsonObject).state,
    { type: "string", enum: ["b"], descriptions: { fr: "Keep" }, unit: "supplied" });
  assert.equal(canonicalJson(input), before);
  outcomes.push({ key: "merge-patch-metadata", input, actual: parsed.effectiveDocument!.properties, originalUnchanged: true, passed: true });
});
test("exact URL query Unicode and fragment-bearing local IDs are not normalized or fetched", () => {
  for (const id of ["https://example.org/定义?revision=1", "urn:copilot:definition#member"]) {
    const definitions: JsonObject = { ...localDefinitions, id };
    const root: JsonObject = { ...localSingleModel, links: [{ rel: "tm:extends", href: id }] };
    const parsed = resolveThingModel([definitions, root]);
    assert.deepEqual(parsed.blockers, []); assert.equal(parsed.id, root.id);
    assert.deepEqual(parsed.document, [definitions, root]);
  }
  const mismatch = { ...localBundleRoot, properties: { value: { "tm:ref": "https://example.org/local/definitions?version=2#/schemaDefinitions/state" } },
    links: [{ rel: "tm:extends", href: localDefinitions.id! }] };
  const parsed = resolveThingModel([localDefinitions, mismatch]);
  assert.ok(parsed.blockers.some(blocker => blocker.includes("Required model file missing")));
  assert.deepEqual(parsed.fields, []);
  outcomes.push({ key: "exact-query-unicode-fragment-ID", exactIds: ["https://example.org/定义?revision=1", "urn:copilot:definition#member"],
    queryMismatchBlocked: true, passed: true });
});
test("missing IDs mixed non-TMs graph wrappers ambiguous roots and cycles never select the first definition", () => {
  const { id: _id, ...missingId } = localDefinitions;
  for (const input of [[], [missingId, localBundleRoot], [localDefinitions, { "@type": "Interface", "@id": "dtmi:test:Device;1" }],
    { "@graph": [localDefinitions, localBundleRoot] }] as JsonValue[]) assert.throws(() => resolveThingModel(input));
  const a: JsonObject = { ...localSingleModel, id: "urn:cycle:a", links: [{ rel: "tm:extends", href: "urn:cycle:b" }] };
  const b: JsonObject = { ...localSingleModel, id: "urn:cycle:b", links: [{ rel: "tm:extends", href: "urn:cycle:a" }] };
  assert.throws(() => resolveThingModel([a, b]), /cyclic/);
  assert.throws(() => resolveThingModel([localSingleModel, a, b]), /unrelated or cyclic/);
  const self = resolveThingModel({ ...localSingleModel, properties: { loop: { "tm:ref": "#/properties/loop" } } });
  assert.ok(self.blockers.some(blocker => blocker.includes("Cyclic"))); assert.deepEqual(self.fields, []);
  outcomes.push({ key: "invalid-packaging-and-cycles", missingIdsMixedTypesGraphAndEmptyBlocked: true,
    cyclicModelGraphBlocked: true, recursivePointerBlocked: true, passed: true });
});
test("depth bounds and invalid resolved patches remain fail-closed", () => {
  let definition: JsonObject = { type: "string" };
  for (let index = 0; index < 55; index++) definition = { type: "object", properties: { child: definition } };
  assert.throws(() => resolveThingModel({ ...localSingleModel, schemaDefinitions: { definition } }), /depth/);
  const invalid = resolveThingModel({ ...localSingleModel, properties: { bad: { "tm:ref": "#/schemaDefinitions/history", type: "float" } } });
  assert.ok(invalid.blockers.length); assert.equal(invalid.effectiveDocument, undefined);
  const chain: JsonObject[] = Array.from({ length: 55 }, (_, index) => ({
    "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel", id: `urn:copilot:test:depth:${index}`,
    schemaDefinitions: { reading: index === 54 ? { type: "string" }
      : { "tm:ref": `urn:copilot:test:depth:${index + 1}#/schemaDefinitions/reading` } },
  }));
  const root: JsonObject = { ...localSingleModel, properties: { value: { "tm:ref": "urn:copilot:test:depth:0#/schemaDefinitions/reading" } }, events: {} };
  const deep = resolveThingModel([root, ...chain]);
  assert.ok(deep.blockers.some(blocker => blocker.includes("depth")), deep.blockers.join("; "));
  assert.equal(deep.effectiveDocument, undefined); assert.deepEqual(deep.fields, []);
  outcomes.push({ key: "depth-and-invalid-effective-schema", nestedDepth: 55, dependencyChain: 55,
    bothDepthBoundsEnforced: true, invalidEffectiveSchemaBlocked: true, passed: true });
});
test("a valid root cannot mask remapped tm namespace authority on an imported bundle member", () => {
  const input: JsonValue = [{ ...localDefinitions, "@context": ["https://www.w3.org/2022/wot/td/v1.1",
    { tm: "https://example.org/not-wot#" }] }, localBundleRoot];
  const parsed = resolveThingModel(input);
  assert.ok(parsed.blockers.some(blocker => blocker.includes("namespace")));
  assert.equal(parsed.effectiveDocument, undefined); assert.deepEqual(parsed.fields, []);
  outcomes.push({ key: "bundle-member-namespace-authority", validRootCannotMaskInvalidHelper: true,
    originalBundlePreserved: canonicalJson(parsed.document) === canonicalJson(input), passed: true });
});
test("whole bundle mutations invalidate root validation ownership and processing fingerprint without stored alternate identity", () => {
  const model = fromThingModel(resolveThingModel(localBundle)), request = localFixtureRequest(localBundle);
  const source = { key: "local", telemetry: { table: "Telemetry", docstring: "" }, twin: { table: "Twins", docstring: "" } };
  const beforeOwner = assignmentOwner(source, 1, model), beforeFingerprint = modelingFingerprint(request.input);
  assert.equal(requireThingModelId(model.document), localBundleRoot.id);
  const changed = JSON.parse(JSON.stringify(localBundle)) as JsonValue;
  assert.ok(Array.isArray(changed)); assert.ok(isJsonObject(changed[0]));
  changed[0].title = "Changed helper facts";
  const next = fromThingModel(resolveThingModel(changed));
  assert.equal(next.id, model.id); assert.notEqual(assignmentOwner(source, 1, next), beforeOwner);
  assert.notEqual(modelingFingerprint({ ...request.input, model: next }), beforeFingerprint);
  assert.ok(modelValidation({ ...model, id: String(localDefinitions.id) }).length);
  assert.ok(modelValidation({ ...model, document: [localDefinitions] }).length);
  outcomes.push({ key: "bundle-owner-fingerprint-root-validation", helperMutationInvalidatesOwnerAndCoverage: true,
    helperIdCannotReplaceSelectedRoot: true, passed: true });
});
