// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import { HubAssignmentApi, ASSIGNMENT_ARM_SCOPE } from "./HubAssignmentApi.ts";
import { sourceHubs, assignmentBlocker, ModelAssignmentOperation, type AssignmentReceipt } from "./HubAssignment.ts";
import { applyRoleHints, fieldValue, fieldValidation, inspectFieldEvidence } from "./HubModeling.ts";
import { canonicalJson, decodeDoSamples, type JsonObject } from "./ModelAuthoring.ts";
import { bindNativeMetadata } from "./HubTransport.ts";
import { buildModeledDataPlan } from "./ModeledDataPlan.ts";
import { ModeledDataCreator } from "./ModeledDataCreation.ts";
import { createCreationMock, creationRequest } from "../../../tests/creationMock.ts";
import { mixedAssignmentFixture, mixedAssignmentSource, mixedDtdlDocument, mixedDtdlRequest,
  mixedHubs, mixedNamespace, mixedRawSample, mixedReportedValues } from "../../../tests/mixedSourceDtdlFixtures.ts";

const outcomes: Record<string, unknown>[] = [];
process.once("exit", () => {
  if (process.env.MIXED_SOURCE_RESULTS) fs.writeFileSync(process.env.MIXED_SOURCE_RESULTS,
    JSON.stringify({ provenance: "Synthetic fixtures; not live combined-table or ARM proof",
      cases: outcomes, count: outcomes.length }, null, 2) + "\n");
});
test("combined source docstrings preserve regular-first linked-second order and tolerate duplicate case and whitespace", () => {
  const telemetry = "Raw device telemetry from IoT Hub\n@iot.meta tableRole=iothub-telemetry-raw\n" +
    "@iot.meta source=regular-hub-1.azure-devices.net; linked-hub-1.azure-devices.net";
  const twin = "Raw device property changes from IoT Hub\n@iot.meta tableRole=iothub-properties-raw\n" +
    "@iot.meta source=regular-hub-1.azure-devices.net; linked-hub-1.azure-devices.net";
  const source = { key: "combined-metadata-fixture", telemetry: { table: "TelemetryRawIoTHub_combined_hubs", docstring: telemetry },
    twin: { table: "PropertiesRawIoTHub_combined_hubs", docstring: twin } };
  const expected = ["regular-hub-1.azure-devices.net", "linked-hub-1.azure-devices.net"];
  assert.deepEqual(sourceHubs(source), { sources: expected, issues: [] });
  source.twin.docstring = "  @iot.meta source=  LINKED-HUB-1.azure-devices.net ; regular-hub-1.azure-devices.net ; LINKED-HUB-1.azure-devices.net  ";
  assert.deepEqual(sourceHubs(source), { sources: expected, issues: [] });
  outcomes.push({ name: "combined docstrings and duplicate/case/whitespace", expected, actual: sourceHubs(source),
    provenance: "Metadata fixture; no resource requests or actual provisioned resources.", passed: true });
});
const cases: { name: string; telemetry: (keyof typeof mixedHubs)[]; twin: (keyof typeof mixedHubs)[];
  devices: string[]; hubs: string[] }[] = [
  { name: "same mixed Hub sets", telemetry: ["a", "u"], twin: ["a", "u"], devices: ["a", "both"], hubs: ["a", "u"] },
  { name: "same sets reordered and duplicated", telemetry: ["u", "a", "a"], twin: ["a", "u"], devices: ["a", "both"], hubs: ["u", "a"] },
  { name: "different stream sets use supported union", telemetry: ["a"], twin: ["u"], devices: ["a", "both"], hubs: ["a", "u"] },
  { name: "two linked Hubs share one namespace", telemetry: ["a", "u"], twin: ["b"], devices: ["a", "b", "both"], hubs: ["a", "u", "b"] },
  { name: "all unlinked is informative and nonmutating", telemetry: ["u"], twin: ["u"], devices: [], hubs: ["u"] },
  { name: "different linked stream sets are modelable", telemetry: ["a"], twin: ["b"], devices: ["a", "b", "both"], hubs: ["a", "b"] },
];
for (const scenario of cases) test(`mixed assignment ${scenario.name}`, async () => {
  const source = mixedAssignmentSource(scenario.telemetry, scenario.twin), fixture = mixedAssignmentFixture();
  assert.deepEqual(sourceHubs(source), { sources: scenario.hubs.map(key => mixedHubs[key].host), issues: [] });
  const inventory = await new HubAssignmentApi(fixture).discover(source, new AbortController().signal);
  assert.deepEqual(inventory.hubs.map(hub => hub.id), scenario.hubs.map(key => mixedHubs[key].id));
  assert.deepEqual(inventory.hubs.map(hub => hub.state), scenario.hubs.map(key => key === "u" ? "unlinked" : "linked"));
  assert.deepEqual(inventory.issues, []);
  assert.deepEqual(inventory.targets.map(target => target.name).sort(), scenario.devices.slice().sort());
  assert.equal(new Set(inventory.targets.map(target => target.id)).size, scenario.devices.length);
  assert.ok(inventory.targets.every(target => target.hubIds.length && target.hubIds.every(id =>
    id !== mixedHubs.u.id && scenario.hubs.some(key => mixedHubs[key].id === id)) && !assignmentBlocker(target)));
  assert.ok(fixture.scopes.every(scope => scope === ASSIGNMENT_ARM_SCOPE));
  assert.ok(fixture.calls.every(call => !/provisioning|enrollment|listKeys/i.test(call.url + (call.body || "")) &&
    !["PATCH", "PUT", "DELETE"].includes(call.method)));
  assert.equal(fixture.calls.filter(call => new URL(call.url).pathname === `${mixedNamespace}/registryDevices`).length,
    scenario.devices.length ? 1 : 0, "A shared namespace must not duplicate device-list reads.");
  outcomes.push({ name: scenario.name, source, expectedDevices: scenario.devices,
    actualTargets: inventory.targets, hubStates: inventory.hubs, issues: inventory.issues, dpsRequests: 0, writes: 0, passed: true });
});

for (const scenario of ["missing-link", "pending-link", "failed-link", "forbidden-read"] as const) {
  test(`mixed source summary reports the current ${scenario} contract alongside an unlinked Hub`, async () => {
    const fixture = mixedAssignmentFixture(), fetcher: typeof fetch = async (url, init) => {
      const response = await fixture.fetcher(url, init);
      if (new URL(String(url)).pathname !== mixedHubs.a.id) return response;
      if (scenario === "forbidden-read") return new Response("{}", { status: 403 });
      const body = await response.json();
      if (scenario === "missing-link") delete body.properties.deviceRegistry;
      else body.properties.deviceRegistry.linkingProperties.state = scenario === "pending-link" ? "Pending" : "Failed";
      return new Response(JSON.stringify(body), { headers: { "Content-Type": "application/json" } });
    };
    const inventory = await new HubAssignmentApi({ ...fixture, fetcher }).discover(mixedAssignmentSource(["u", "a"]), new AbortController().signal);
    const expectedState = scenario === "pending-link" ? "pending" : scenario === "forbidden-read" ? "unverified" : "unlinked";
    assert.deepEqual(inventory.hubs.map(hub => hub.state), ["unlinked", expectedState]);
    assert.equal(inventory.targets.length, 0);
    if (scenario === "forbidden-read") assert.equal(inventory.issues[0].kind, "permission");
    if (scenario !== "forbidden-read") assert.deepEqual(inventory.issues, []);
    assert.ok(fixture.calls.every(call => !/provisioning|enrollment/i.test(call.url)));
    outcomes.push({ name: `source state ${scenario}`, expected: ["unlinked", expectedState],
      actual: inventory.hubs, issues: inventory.issues, targetCount: 0, passed: true });
  });
}

test("mixed absent registry link does not block a verified linked target or weaken conditional assignment guards", async () => {
  const fixture = mixedAssignmentFixture(), source = mixedAssignmentSource(["u", "a"]);
  const targetId = `${mixedNamespace}/registryDevices/a`, model = creationRequest().input.model;
  const fakePatches: { url: string; body: unknown; ifMatch: string | null }[] = [];
  const fetcher: typeof fetch = async (url, init) => {
    const path = new URL(String(url)).pathname;
    if (init?.method === "PATCH") {
      const body: unknown = JSON.parse(String(init.body)), ifMatch = new Headers(init.headers).get("If-Match");
      assert.equal(path, targetId); assert.equal(ifMatch, '"mixed-a-1"');
      assert.deepEqual(body, { properties: { deviceTypeRefs: [model.id] } });
      fakePatches.push({ url: String(url), body, ifMatch });
      return new Response("", { status: 202, headers: { Location:
        `https://management.azure.com/subscriptions/${targetId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/mixed-operation?api-version=2026-11-02-preview` } });
    }
    if (path.endsWith("/asyncOperationStatuses/mixed-operation")) {
      return new Response(JSON.stringify({ status: "Succeeded" }));
    }
    const response = await fixture.fetcher(url, init);
    if (path === mixedHubs.u.id) {
      const body = await response.json();
      delete body.properties.deviceRegistry;
      return new Response(JSON.stringify(body), { headers: { "Content-Type": "application/json" } });
    }
    if (path === `${mixedNamespace}/registryDevices`) return new Response(JSON.stringify({ value: [{ id: targetId }] }),
      { headers: { "Content-Type": "application/json" } });
    return response;
  };
  const api = new HubAssignmentApi({ ...fixture, fetcher }), signal = new AbortController().signal;
  const inventory = await api.discover(source, signal);
  assert.deepEqual(inventory.hubs.map(hub => hub.state), ["unlinked", "linked"]);
  assert.deepEqual(inventory.issues, []);
  assert.equal(inventory.targets.length, 1);
  const target = inventory.targets[0];
  assert.equal(target.id, targetId); assert.deepEqual(target.hubIds, [mixedHubs.a.id]);
  assert.equal(target.reference.kind, "unassigned"); assert.equal(target.updatePermission, "granted");
  assert.equal(assignmentBlocker(target), undefined);
  const services = api.assignmentServices(source), operation = new ModelAssignmentOperation();
  await assert.rejects(() => operation.run([target], model, false, services, () => true, signal, () => undefined),
    /Confirm the selected first assignments/);
  assert.equal(fakePatches.length, 0, "A source issue must not weaken explicit confirmation.");
  const published: AssignmentReceipt[] = [];
  const receipts = await operation.run([target], model, true, services, () => true, signal, receipt => published.push(receipt));
  assert.equal(receipts.length, 1); assert.equal(receipts[0].state, "confirmed");
  assert.deepEqual(published.map(receipt => receipt.state).filter((state, index, states) =>
    index === 0 || state !== states[index - 1]), ["checking", "assigning", "accepted", "confirmed"]);
  assert.equal(fakePatches.length, 1);
  assert.ok(fixture.calls.filter(call => new URL(call.url).pathname === targetId).length >= 3,
    "Discovery, operation recheck and pre-PATCH recheck must each read the linked target.");
  assert.equal(fixture.calls.filter(call => new URL(call.url).pathname === mixedHubs.u.id).length, 1,
    "An unrelated unlinked source must not be a target's assignment prerequisite.");
  assert.ok(fixture.scopes.every(scope => scope === ASSIGNMENT_ARM_SCOPE));
  assert.ok([...fixture.calls.map(call => call.url), ...fakePatches.map(call => call.url)].every(url => !/provisioning|enrollment/i.test(url)));
  outcomes.push({ name: "mixed actual-profile absent registry plus one linked permitted target",
    input: { source, absentRegistryHub: mixedHubs.u.id, linkedHub: mixedHubs.a.id, targetId },
    expected: "An unlinked source coexists with one unassigned/granted target; explicit confirmation permits exactly one conditional fake PATCH202",
    hubStates: inventory.hubs, sourceIssues: inventory.issues, targetBlocker: null, target,
    unconfirmedWrites: 0, fakePatches, receipt: receipts[0], phases: published.map(receipt => receipt.state),
    dpsRequests: 0, actualCloudRequests: 0, actualCloudWrites: 0, passed: true });
});

test("DTDL components keep eighteen whole fields and model-declared roles despite contrary stream hints", async () => {
  const original = canonicalJson(mixedDtdlDocument), request = mixedDtdlRequest(), fields = request.input.fields;
  assert.equal(fields.length, 18); assert.equal(new Set(fields.map(field => field.id)).size, 18);
  assert.deepEqual(fieldValidation(fields), []);
  assert.deepEqual(request.coverage.nativeIssues, []);
  for (const component of ["", "motor", "power"]) for (const [name, type, role] of [
    ["temperature", "real", "telemetry"], ["measurement", "dynamic", "telemetry"], ["samples", "dynamic", "telemetry"],
    ["state", "string", "twin"], ["configuration", "dynamic", "twin"], ["history", "dynamic", "twin"],
  ]) {
    const field = fields.find(value => (value.component || "") === component && value.path[0] === name)!;
    assert.ok(field); assert.equal(field.column, component ? `${component}_${name}` : name);
    assert.equal(field.kustoType, type); assert.equal(field.role, role); assert.equal(field.roleSource, "model");
    assert.equal(field.path.length, 1);
  }
  const contrary = fields.map(field => ({ fieldId: field.id, telemetry: field.role === "twin" ? 1 : 0,
    twin: field.role === "telemetry" ? 1 : 0, state: field.role === "twin" ? "telemetry" as const : "twin" as const }));
  assert.deepEqual(applyRoleHints(fields, contrary), fields);
  assert.ok(request.coverage.fields.every(field => field.telemetry > 0 || field.twin > 0));
  const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
  assert.match(plan.commands.find(command => command.kind === "view")!.csl, /by deviceId, propertyName/);
  assert.ok(!plan.commands.some(command => /sourceHubId/.test(command.csl)));
  const mock = createCreationMock(request), created = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(created.phase, "completed", created.error);
  assert.equal(mock.writes.length, 9);
  assert.equal(canonicalJson(mixedDtdlDocument), original);
  outcomes.push({ name: "DTDL whole fields and declared roles", fields, creator: created.phase, fakeWrites: 9,
    commandCount: plan.commands.length, originalUnchanged: true, passed: true });
});

test("DTDL absent and unknown components never bind equal root/component leaf names or missing markers", () => {
  const request = mixedDtdlRequest();
  for (const component of ["", "motor", "power", "unknown"]) {
    const sample = decodeDoSamples([mixedRawSample("telemetry", component)], "telemetry", 100).payloads[0];
    for (const field of request.input.fields.filter(field => field.role === "telemetry"))
      assert.equal(fieldValue(sample, field, "telemetry").found, (field.component || "") === component);
  }
  const bad = mixedRawSample("twin");
  bad.data = JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
    event_time: "2026-10-23T21:00:00Z", event_kind: "lifecycle", payload: { value: { properties: { reported: {
      ...mixedReportedValues, motor: { state: "not-a-component" }, unknown: { __t: "c", state: "unknown" },
    } } } } });
  const sample = decodeDoSamples([bad], "twin", 100).payloads[0];
  for (const field of request.input.fields.filter(field => field.role === "twin" && field.component === "motor"))
    assert.equal(fieldValue(sample, field, "twin").found, false);
  outcomes.push({ name: "absent/unknown components and missing marker", expected: "no cross-binding", passed: true });
});

test("DTDL native module identities remain distinct and conflicting module/component metadata is fail-closed", () => {
  const identities: string[] = [];
  for (const module of ["", "module-1", "module-2"]) {
    for (const stream of ["telemetry", "twin"] as const) {
      const row = mixedRawSample(stream, stream === "telemetry" ? "motor" : "", module);
      const sample = decodeDoSamples([row], stream, 100).payloads[0];
      assert.equal(sample.native!.identity, module ? `fixture-device/${module}` : "fixture-device");
      if (stream === "telemetry") identities.push(sample.native!.identity!);
    }
  }
  assert.equal(new Set(identities).size, 3);
  for (const user of [
    { cloudEvents_subject: "fixture-device/other", cloudEvents_moduleid: "other", cloudEvents_iothubdtsubject: "motor" },
    { cloudEvents_subject: "fixture-device/module-1", cloudEvents_moduleid: "module-1", cloudEvents_iothubdtsubject: "power" },
  ]) {
    const row = mixedRawSample("telemetry", "motor", "module-1"), bound = bindNativeMetadata(row.headers, user, "telemetry");
    assert.equal(bound.identity, undefined); assert.ok(bound.issues.length);
  }
  const evidence = inspectFieldEvidence(mixedDtdlRequest().input.fields,
    decodeDoSamples([mixedRawSample("telemetry", "unknown")], "telemetry", 100),
    decodeDoSamples([], "twin", 100));
  assert.ok(evidence.every(field => field.telemetry === 0 && field.twin === 0));
  outcomes.push({ name: "module identity and metadata conflict guards", identities, conflictingProfiles: 2, passed: true });
});

test("DTDL component decoding accepts actual blank EH plus CE profile without treating synthetic module as a component", () => {
  const request = mixedDtdlRequest();
  for (const module of ["", "module-1"]) for (const component of ["", "motor", "power"]) {
    const decoded = (["telemetry", "twin"] as const).map(stream => {
      const row = mixedRawSample(stream, component, module);
      const user: JsonObject = { cloudEvents_subject: module ? `fixture-device/${module}` : "fixture-device",
        ...(module ? { cloudEvents_moduleid: module } : {}),
        ...(stream === "telemetry" ? { cloudEvents_iothubdtsubject: component } : { cloudEvents_deviceid: "fixture-device" }) };
      const sample = decodeDoSamples([{ ...row,
        headers: { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" }, user_headers: user }], stream, 100);
      assert.equal(sample.payloads.length, 1); assert.equal(sample.payloads[0].native!.profile, "cloudEvents");
      assert.deepEqual(sample.payloads[0].native!.issues, []);
      assert.equal(sample.payloads[0].native!.identity, module ? `fixture-device/${module}` : "fixture-device");
      return sample;
    });
    for (const field of request.input.fields.filter(field => field.role === "telemetry"))
      assert.equal(fieldValue(decoded[0].payloads[0], field, "telemetry").found, (field.component || "") === component);
    for (const field of request.input.fields.filter(field => field.role === "twin"))
      assert.equal(fieldValue(decoded[1].payloads[0], field, "twin").found, true);
  }
  outcomes.push({ name: "actual blank EH/CE profile with DTDL native components and synthetic modules", combinations: 6,
    streamsPerCombination: 2, nativeShape: "Parent-attributed metadata profile, agent-authored payload/device/module values", passed: true });
});
