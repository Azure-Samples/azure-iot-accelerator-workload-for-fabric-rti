// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test, { type TestContext } from "node:test";
import { requireThingModelId, resolveThingModel } from "./ThingModelParser.ts";
import { coverageReady, fromDtdlModel, fromThingModel, modelValidation, inspectStreamCoverage } from "./HubModeling.ts";
import { parseDtdlModel } from "./DtdlModelParser.ts";
import { canonicalJson } from "./ModelAuthoring.ts";
import {
  ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE, ASSIGNMENT_OPERATION_TIMEOUT_MS, ASSIGNMENT_POLL_TIMEOUT_MS, ASSIGNMENT_WAVE_SIZE, ASSIGNMENT_WAVE_INTERVAL_MS, AssignmentListLoading, ModelAssignmentOperation, allAssignmentHubsUnavailable, assignmentBlocker, assignmentIntentOwner, currentAssignmentHubSummary, currentAssignmentIntent, currentAssignmentListsLoading,
  confirmedAssignmentCount, inventoryHasModelAssignments, modelReferenceFilterAvailability,
  assignmentIssuesForView, assignmentReference, assignmentOwner, readReference, sourceHubs,
  DPS_INFORMATIONAL_ONLY, CURRENT_REFERENCE_FRESHNESS, permissionGrants, referenceIsCurrent, specificEtag,
  type AssignmentInventory, type AssignmentSource, type AssignmentTarget, type AssignmentServices, type AssignmentReceipt, type AssignmentIssue,
} from "./HubAssignment.ts";
import {
  AssignmentApiError, HubAssignmentApi, HubAssignmentReader, assignmentTokenFailure, ASSIGNMENT_ARM_SCOPE,
  pollAssignmentReadback, ASSIGNMENT_POLL_INTERVAL_MS,
} from "./HubAssignmentApi.ts";
import { ModeledDataCreator } from "./ModeledDataCreation.ts";
import { buildModeledDataPlan } from "./ModeledDataPlan.ts";
import { modelingFingerprint } from "./HubModeling.ts";
import { creationRequest } from "../../../tests/creationMock.ts";
import { getPersistableWizardContext } from "../../components/Wizard/wizardUtils.ts";

const root = "/subscriptions/00000000-0000-0000-0000-000000000001/resourceGroups/fixture";
const hubId = `${root}/providers/Microsoft.Devices/IotHubs/fixture-hub`;
const nsId = `${root}/providers/Microsoft.DeviceRegistry/namespaces/fixture-ns`;
const deviceId = `${nsId}/registryDevices/fixture-device`;
const dpsId = `${root}/providers/Microsoft.Devices/ProvisioningServices/fixture-dps`;
const document = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
  id: "urn:fixture:Assignment;1", properties: { value: { type: "number" } } };
const model = () => fromThingModel(resolveThingModel(document));
const source = (): AssignmentSource => ({ key: "fixture-source",
  telemetry: { table: "RawTelemetry", docstring: "@iot.meta tableRole=iothub-telemetry-raw\n@iot.meta source=fixture-hub.azure-devices.net" },
  twin: { table: "RawProperties", docstring: "@iot.meta source=fixture-hub.azure-devices.net" } });
const target = (): AssignmentTarget => ({ id: deviceId, name: "Fixture device", kind: "device",
  hubIds: [hubId], namespaceId: nsId, etag: '"fixture-etag-1"', reference: { kind: "unassigned", ids: [] } });
const response = (body: unknown, status = 200) => new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
const signal = () => new AbortController().signal;
function trackedAssignmentTimers(context: TestContext) {
  const set = globalThis.setTimeout, clear = globalThis.clearTimeout;
  const pending = new Set<ReturnType<typeof setTimeout>>();
  context.mock.method(globalThis, "setTimeout", (callback: (...args: unknown[]) => void, delay?: number, ...args: unknown[]) => {
    const timer = set(() => { pending.delete(timer); callback(...args); }, delay);
    pending.add(timer); return timer;
  });
  context.mock.method(globalThis, "clearTimeout", (timer: ReturnType<typeof setTimeout>) => { pending.delete(timer); clear(timer); });
  return pending;
}

test("assignment list loading bridges phases with owned disposable leases and never persists or gates readback alone", () => {
  const context = { hubModel: model(), hubAssignmentSource: source() };
  const ready = { key: source().key, revision: 1, ready: true, reason: "Validated." };
  const owner = assignmentIntentOwner(context, ready);
  const loading = new AssignmentListLoading(owner, () => undefined);
  const chosen = { ...context, hubAssignmentIntent: { owner, mode: "assign", listLoading: loading } };
  loading.reserve();
  assert.equal(currentAssignmentListsLoading(chosen, ready), true);
  const hubDone = loading.begin();
  loading.reserve(); hubDone();
  assert.equal(currentAssignmentListsLoading(chosen, ready), true, "Hub-to-device handoff must have no idle gap.");
  const deviceDone = loading.begin();
  assert.equal(currentAssignmentListsLoading(chosen, ready), true);
  deviceDone(); deviceDone();
  assert.equal(currentAssignmentListsLoading(chosen, ready), false);
  assert.equal(currentAssignmentListsLoading({ ...chosen, hubAssignmentIntent: { ...chosen.hubAssignmentIntent,
    attempt: { operationId: "accepted-polling", modelId: document.id, targetIds: [deviceId] } } }, ready), false);
  const oldDone = loading.begin();
  const replacement = new AssignmentListLoading(owner, () => undefined), newDone = replacement.begin();
  const next = { ...chosen, hubAssignmentIntent: { ...chosen.hubAssignmentIntent, listLoading: replacement } };
  loading.dispose(); oldDone();
  assert.equal(currentAssignmentListsLoading(next, ready), true, "Old disposal/completion cannot clear a newer lease.");
  assert.equal(currentAssignmentListsLoading(next, { ...ready, revision: 3 }), false);
  assert.equal(currentAssignmentListsLoading({ ...next, hubAssignmentIntent: { ...next.hubAssignmentIntent, mode: "skip" } }, ready), false);
  assert.equal(Object.hasOwn(getPersistableWizardContext(next), "hubAssignmentIntent"), false);
  newDone();
  const replayedOldDone = replacement.begin();
  replacement.dispose(); replacement.activate();
  const replayedCurrentDone = replacement.begin();
  replayedOldDone();
  assert.equal(currentAssignmentListsLoading(next, ready), true, "Effect replay cannot let an old release clear current work.");
  replayedCurrentDone(); replacement.dispose();
  assert.equal(currentAssignmentListsLoading(next, ready), false);
});

test("assignment intent is source and model revision bound, explicit, and never enables DTDL assignment", () => {
  const context = { hubModel: model(), hubAssignmentSource: source(), hubSourceSelectionRevision: 3, hubModelRevision: 2 };
  const ready = { key: source().key, revision: 4, ready: true, reason: "Validated." };
  const owner = assignmentIntentOwner(context, ready);
  assert.ok(owner);
  assert.equal(currentAssignmentIntent(context, ready), undefined);
  const chosen = { ...context, hubAssignmentIntent: { owner, mode: "assign" } };
  assert.equal(currentAssignmentIntent(chosen, ready)?.mode, "assign");
  assert.equal(currentAssignmentIntent({ ...chosen, hubAssignmentIntent: { owner, mode: "skip" } }, ready)?.mode, "skip");
  for (const changed of [{ ...ready, key: "other" }, { ...ready, revision: 5 }, { ...ready, ready: false }]) {
    assert.equal(currentAssignmentIntent(chosen, changed), undefined);
  }
  for (const changed of [
    { ...chosen, hubSourceSelectionRevision: 5 },
    { ...chosen, hubModelRevision: 4 },
    { ...chosen, hubAssignmentSource: { ...source(), twin: { ...source().twin, docstring: "" } } },
    { ...chosen, hubModel: fromThingModel(resolveThingModel({ ...document, description: "Different document." })) },
    { ...chosen, hubAssignmentIntent: { owner, mode: "unrecognized" } },
  ]) assert.equal(currentAssignmentIntent(changed, ready), undefined);
  const dtdl = { "@id": "dtmi:fixture:Device;1", "@type": "Interface",
    contents: [{ "@type": "Telemetry", name: "value", schema: "double" }] };
  const dtdlContext = { ...context, hubModel: fromDtdlModel(parseDtdlModel(dtdl), dtdl) };
  const dtdlOwner = assignmentIntentOwner(dtdlContext, ready);
  assert.equal(currentAssignmentIntent({ ...dtdlContext, hubAssignmentIntent: { owner: dtdlOwner, mode: "assign" } }, ready), undefined);
  assert.equal(currentAssignmentIntent({ ...dtdlContext, hubAssignmentIntent: { owner: dtdlOwner, mode: "skip" } }, ready)?.mode, "skip");
  assert.equal(currentAssignmentIntent({ ...chosen, hubModel: undefined }, ready), undefined);
});
test("model reference filter availability requires current positive Hub evidence and never invents all-unlinked status", () => {
  const context = { hubModel: model(), hubAssignmentSource: source() };
  const ready = { key: source().key, revision: 1, ready: true, reason: "Validated." };
  const owner = assignmentIntentOwner(context, ready);
  const summary = { sourceSignature: canonicalJson(source()), hubs: [{ id: hubId, name: "Hub", namespaceId: nsId, state: "linked" as const }],
    issues: [], complete: true };
  const cached = { ...context, hubAssignmentHubSummary: { owner, value: summary } };
  assert.equal(modelReferenceFilterAvailability(context, ready), "unconfirmed");
  assert.equal(modelReferenceFilterAvailability(cached, ready), "available");
  assert.equal(modelReferenceFilterAvailability({ ...cached, hubAssignmentHubSummary: { owner, value: { ...summary,
    hubs: [...summary.hubs, { id: hubId + "-other", name: "Other", state: "unlinked" }] } } }, ready), "available");
  for (const state of ["pending", "unverified"] as const) {
    assert.equal(modelReferenceFilterAvailability({ ...cached, hubAssignmentHubSummary: { owner, value: { ...summary,
      hubs: [{ ...summary.hubs[0], state }] } } }, ready), "unconfirmed");
  }
  const unlinked = { ...summary, hubs: [{ ...summary.hubs[0], state: "unlinked" as const }] };
  assert.equal(modelReferenceFilterAvailability({ ...cached, hubAssignmentHubSummary: { owner, value: unlinked } }, ready), "unlinked");
  for (const value of [
    { ...unlinked, complete: false },
    { ...unlinked, hubs: [] },
    { ...unlinked, issues: [{ resource: hubId, message: "Read failed.", retryable: true }] },
    { ...unlinked, sourceSignature: "old-source" },
  ]) assert.equal(modelReferenceFilterAvailability({ ...cached, hubAssignmentHubSummary: { owner, value } }, ready), "unconfirmed");
  assert.equal(modelReferenceFilterAvailability(cached, { ...ready, revision: 3 }), "unconfirmed");
  assert.equal(modelReferenceFilterAvailability({ ...cached, hubModelRevision: 2 }, ready), "unconfirmed");
  assert.equal(modelReferenceFilterAvailability(cached, { ...ready, ready: false }), "unconfirmed");
  assert.equal(modelReferenceFilterAvailability({ ...context, hubModel: { kind: "dtdl" } }, ready), "dtdl");
});

test("known nonlinked Hub results qualify summary severity but never replace explicit assignment intent", () => {
  const context = { hubModel: model(), hubAssignmentSource: source() };
  const ready = { key: source().key, revision: 1, ready: true, reason: "Validated." };
  const owner = assignmentIntentOwner(context, ready);
  const summary = { sourceSignature: canonicalJson(source()),
    hubs: [{ id: hubId, name: "fixture-hub", state: "unlinked" as const }], issues: [], complete: true };
  const unavailable = { ...context, hubAssignmentHubSummary: { owner, value: summary },
    hubAssignmentIntent: { owner, mode: "unavailable" } };
  assert.equal(currentAssignmentIntent(unavailable, ready), undefined);
  assert.equal(allAssignmentHubsUnavailable(currentAssignmentHubSummary(unavailable, ready)), true);
  for (const value of [
    { ...summary, complete: false },
    { ...summary, issues: [{ resource: hubId, message: "Failed read.", retryable: true }] },
    { ...summary, hubs: [] },
    { ...summary, hubs: [{ ...summary.hubs[0], state: "unverified" }] },
    { ...summary, sourceSignature: "another-source" },
  ]) {
    const changed = { ...unavailable, hubAssignmentHubSummary: { owner, value } };
    assert.equal(currentAssignmentIntent(changed, ready), undefined);
    assert.equal(currentAssignmentIntent({ ...changed, hubAssignmentIntent: { owner, mode: "assign" } }, ready)?.mode, "assign");
  }
  assert.equal(currentAssignmentIntent(unavailable, { ...ready, revision: 2 }), undefined);
  assert.equal(currentAssignmentIntent({ ...unavailable, hubModelRevision: 2 }, ready), undefined);
  assert.equal(currentAssignmentIntent({ ...unavailable, hubAssignmentIntent: { owner, mode: "skip" },
    hubAssignmentHubSummary: undefined }, ready)?.mode, "skip");
});

test("source metadata reads whole semicolon lines, both streams and no arbitrary addresses", () => {
  const value = source();
  value.telemetry.docstring += "; second-hub.azure-devices.net\n@iot.meta source=THIRD-HUB.azure-devices.net";
  value.twin.docstring = `@iot.meta source=${hubId}; https://not-a-hub.example/path`;
  assert.deepEqual(sourceHubs(value).sources,
    ["fixture-hub.azure-devices.net", "second-hub.azure-devices.net", "third-hub.azure-devices.net", hubId.toLowerCase()]);
  assert.equal(sourceHubs(value).issues.length, 1);
  value.telemetry.docstring = ""; value.twin.docstring = "";
  assert.equal(sourceHubs(value).issues.length, 2);
});
const pollingTarget = (): AssignmentTarget => ({ ...target(),
  deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: "fixture-device" } });
test("initial conditional PATCH429 is a definite failed write with no observation or automatic replay", async () => {
  const fixture = fixtureClient({ completePoint: true, currentProperties: {} });
  let writes = 0, statusReads = 0;
  const dispatchedRequests: { url: string; method: string }[] = [];
  const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
    const address = String(url), method = init?.method || "GET";
    dispatchedRequests.push({ url: address, method });
    if (method === "PATCH") {
      writes++;
      assert.equal(new Headers(init?.headers).get("If-Match"), '"fixture-etag-1"');
      assert.deepEqual(JSON.parse(String(init?.body)), { properties: { deviceTypeRefs: [document.id] } });
      return writes === 1 ? new Response("", { status: 429, headers: { "Retry-After": "3600" } })
        : new Response(null, { status: 204 });
    }
    if (new URL(address).pathname.includes("operationStatuses")) statusReads++;
    return fixture.fetcher(url, init);
  } });
  const selected = (await api.discover(source(), signal())).targets[0], operation = new ModelAssignmentOperation();
  const first = await operation.run([selected], model(), true, api.assignmentServices(source()),
    () => true, signal(), () => undefined, "throttled-original-write");
  assert.equal(first[0].state, "failed");
  assert.equal(first[0].writeDispatched, true, "The rejected HTTP request was actually dispatched.");
  assert.equal(first[0].message, "The Azure request failed (429).");
  assert.equal(first[0].readbackDeadline, undefined); assert.equal(first[0].trackingId, undefined);
  assert.equal(first[0].readbackFailure, undefined);
  assert.equal(writes, 1); assert.equal(statusReads, 0);
  const afterRefusal = dispatchedRequests.length;
  await pollAssignmentReadback({ operationId: first[0].operationId, modelId: document.id,
    targetIds: [selected.id], targets: [selected] }, first, {
    current: () => true,
    read: async () => { throw new Error("A rejected write has no accepted operation to observe."); },
    publish: () => { throw new Error("A rejected write cannot restart or rewrite itself."); },
  }, signal());
  await assert.rejects(() => operation.run([selected], model(), false, api.assignmentServices(source()),
    () => true, signal(), () => undefined), /Confirm/);
  assert.equal(dispatchedRequests.length, afterRefusal); assert.equal(writes, 1);
  const refreshed = await api.recheck(selected, source(), signal());
  const beforeSecondAttempt = dispatchedRequests.length;
  const second = await operation.run([refreshed], model(), true, api.assignmentServices(source()),
    () => true, signal(), () => undefined, "new-explicit-confirmation");
  assert.equal(second[0].state, "confirmed"); assert.equal(second[0].operationId, "new-explicit-confirmation");
  assert.equal(writes, 2); assert.equal(statusReads, 0);
  const secondRequests = dispatchedRequests.slice(beforeSecondAttempt);
  for (const resource of [hubId, nsId, deviceId, `${deviceId}/capabilities`,
    `${deviceId}/providers/Microsoft.Authorization/permissions`])
    assert.ok(secondRequests.some(request => request.method === "GET" && new URL(request.url).pathname === resource),
      "A second explicitly confirmed attempt must repeat current canonical safety prechecks: " + resource);
  const retainedUnknown: AssignmentReceipt = { ...first[0], state: "unknown",
    readbackFailure: { kind: "request", retryable: false }, message: "The Azure request failed (429)." };
  await pollAssignmentReadback({ operationId: retainedUnknown.operationId, modelId: document.id,
    targetIds: [selected.id], targets: [selected] }, [retainedUnknown], {
    current: () => true,
    read: async () => { throw new Error("Old unknown receipts cannot infer a rejected write from message text."); },
    publish: () => { throw new Error("Old unknown receipts cannot be reclassified from message text."); },
  }, signal());
  assert.equal(retainedUnknown.state, "unknown");
});
test("native operation GET429 remains a read throttle until Succeeded or the original polling deadline", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const pendingTimers = trackedAssignmentTimers(context), originalNow = Date.now, originalRandom = Math.random;
  let time = 0;
  Date.now = () => time; Math.random = () => 0.5;
  try {
    for (const permanent of [false, true]) {
      const start = time, selected = { ...pollingTarget(), updatePermission: "granted" as const };
      let writes = 0, settled = false;
      const reads: number[] = [], updates: AssignmentReceipt[] = [];
      const address = `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/throttled-read?api-version=2026-11-02-preview`;
      const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
        if (init?.method === "PATCH") {
          writes++;
          return new Response(null, { status: 202, headers: { Location: address, "Retry-After": "3600" } });
        }
        assert.equal(String(url), address, "Completion cannot fall back to a mutable device GET.");
        reads.push(time - start);
        return !permanent && reads.length > 1 ? response({ status: "Succeeded" })
          : new Response("", { status: 429, headers: { "Retry-After": "3600" } });
      } });
      api.recheck = async value => value;
      const operation = new ModelAssignmentOperation().run([selected], model(), true, api.assignmentServices(source()),
        () => true, signal(), receipt => updates.push(receipt)).finally(() => { settled = true; });
      for (let second = 0; second <= 31 && !settled; second++) {
        for (let turn = 0; turn < 60; turn++) await Promise.resolve();
        if (!settled) { time += 1000; context.mock.timers.tick(1000); }
      }
      const result = await operation;
      assert.equal(writes, 1);
      assert.equal(result[0].state, permanent ? "unknown" : "confirmed");
      assert.equal(result[0].readbackDeadline, start + 30000);
      assert.ok(updates.every(receipt => receipt.state !== "failed"), "A status-read throttle is not rejection of the original assignment.");
      assert.ok(updates.some(receipt => receipt.state === "unknown" && receipt.readbackFailure?.retryable === true));
      assert.deepEqual(reads, permanent ? [3000, 6000, 9000, 14000, 19000, 24000] : [3000, 6000]);
      if (permanent) assert.equal(result[0].message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
      assert.equal(pendingTimers.size, 0);
      await pollAssignmentReadback({ operationId: result[0].operationId, modelId: document.id,
        targetIds: [selected.id], targets: [selected] }, result, {
        current: () => true,
        read: async () => { throw new Error("A settled operation cannot restart reads or gain another deadline."); },
        publish: () => { throw new Error("A settled operation cannot gain another polling window."); },
      }, signal());
    }
  } finally { Date.now = originalNow; Math.random = originalRandom; }
});
test("a429 thrown after write acceptance is not classified as an initial mutation rejection", async () => {
  const selected = pollingTarget();
  const receipts = await new ModelAssignmentOperation().run([selected], model(), true, {
    read: async value => value,
    write: async (_target, _id, _signal, dispatch) => { dispatch(); return "accepted"; },
    trackingId: () => "original-accepted-operation",
    observe: async () => { throw new AssignmentApiError("request", "The Azure request failed (429).", true, 429); },
  }, () => true, signal(), () => undefined);
  assert.equal(receipts[0].state, "unknown"); assert.equal(receipts[0].writeDispatched, true);
  assert.ok(receipts[0].readbackDeadline); assert.equal(receipts[0].trackingId, "original-accepted-operation");
});
test("dispatched untracked transport uncertainty stays terminal unknown through observation and retention", async () => {
  const guidance = "The assignment request may have reached Azure, but its result couldn't be confirmed. Check the device later. No assignment will be retried.";
  for (const outcome of ["500", "503", "network"] as const) {
    const fixture = fixtureClient(), selected = (await fixture.api.discover(source(), signal())).targets[0];
    let writes = 0, afterWriteReads = 0, dispatched = false;
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
      if (init?.method === "PATCH") {
        writes++; dispatched = true;
        if (outcome === "network") throw new Error("Synthetic private transport detail.");
        return response({ error: { message: "Synthetic private service detail." } }, Number(outcome));
      }
      if (dispatched) afterWriteReads++;
      return fixture.fetcher(url, init);
    } });
    const receipts = await new ModelAssignmentOperation().run([selected], model(), true, api.assignmentServices(source()),
      () => true, signal(), () => undefined, "untracked-" + outcome);
    assert.equal(receipts[0].state, "unknown");
    assert.deepEqual(receipts[0].readbackFailure, { kind: "request", retryable: false });
    assert.equal(receipts[0].message, guidance);
    const updates: AssignmentReceipt[] = [];
    await pollAssignmentReadback({ operationId: receipts[0].operationId, modelId: document.id,
      targetIds: [selected.id], targets: [selected] }, receipts, {
      current: () => true, wait: async () => undefined,
      read: (value, id, abort, trackingId) => api.readAssignment(value, id, abort, trackingId),
      publish: receipt => updates.push(receipt),
    }, signal());
    const retained = updates.at(-1) || receipts[0];
    assert.equal(retained.state, "unknown");
    assert.deepEqual(retained.readbackFailure, { kind: "request", retryable: false });
    assert.equal(retained.message, guidance);
    assert.equal(retained.trackingId, undefined);
    assert.equal(retained.readbackDeadline, undefined, "Untracked transport cannot acquire a polling deadline.");
    await pollAssignmentReadback({ operationId: retained.operationId, modelId: document.id,
      targetIds: [selected.id], targets: [selected] }, [JSON.parse(JSON.stringify(retained))], {
      current: () => true, wait: async () => { throw new Error("Terminal uncertainty cannot schedule another observation."); },
      read: async () => { throw new Error("No operation status exists to read."); },
      publish: () => { throw new Error("Retained terminal uncertainty must remain unchanged."); },
      now: () => retained.preparationDeadline! + 60000,
    }, signal());
    await pollAssignmentReadback({ operationId: retained.operationId, modelId: document.id,
      targetIds: [selected.id], targets: [selected] }, [{ ...retained,
        message: "Earlier uncertain request guidance.", readbackFailure: undefined }], {
      current: () => true, wait: async () => { throw new Error("Untracked receipts cannot start observation."); },
      read: async () => { throw new Error("Untracked receipts cannot perform reads."); },
      publish: receipt => { assert.equal(receipt.state, "unknown"); assert.equal(receipt.message, guidance);
        assert.deepEqual(receipt.readbackFailure, { kind: "request", retryable: false }); },
    }, signal());
    assert.equal(writes, 1); assert.equal(afterWriteReads, 0);
    assert.doesNotMatch(retained.message!, /private|operation failed|30 seconds/);
  }
});
test("missing unsafe and lost operation tracking stay terminal unknown without observation authorization or replay", async () => {
  for (const outcome of ["missing", "unsafe", "lost"] as const) {
    const fixture = fixtureClient(), selected = (await fixture.api.discover(source(), signal())).targets[0];
    let dispatched = false, writes = 0, afterWriteRequests = 0, afterWriteAuthorizations = 0;
    const api = new HubAssignmentApi({ token: async () => {
      if (dispatched) afterWriteAuthorizations++;
      return "synthetic";
    }, fetcher: async (url, init) => {
      if (init?.method === "PATCH") {
        dispatched = true; writes++;
        const address = outcome === "unsafe" ? "https://unsafe.example.invalid/operation?opaque=fixture"
          : `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/fixture?api-version=2026-11-02-preview`;
        return new Response(null, { status: 202, headers: outcome === "missing" ? {} : { Location: address } });
      }
      if (dispatched) afterWriteRequests++;
      return fixture.fetcher(url, init);
    } });
    const services = api.assignmentServices(source());
    const receipts = await new ModelAssignmentOperation().run([selected], model(), true,
      { ...services, observe: undefined }, () => true, signal(), () => undefined);
    if (outcome === "lost") {
      assert.equal(receipts[0].state, "accepted"); assert.ok(receipts[0].trackingId);
      receipts[0].trackingId = "retained-handle-no-longer-in-memory";
    } else {
      assert.equal(receipts[0].state, "unknown"); assert.equal(receipts[0].trackingId, undefined);
    }
    const updates: AssignmentReceipt[] = [];
    const attempt = { operationId: receipts[0].operationId, modelId: document.id, targetIds: [selected.id], targets: [selected] };
    await pollAssignmentReadback(attempt, receipts, { current: () => true, wait: async () => undefined,
      read: (value, id, abort, trackingId) => api.readAssignment(value, id, abort, trackingId),
      publish: receipt => updates.push(receipt),
    }, signal());
    const retained = updates.at(-1) || receipts[0];
    assert.equal(retained.state, "unknown");
    assert.deepEqual(retained.readbackFailure, { kind: "shape", retryable: false });
    assert.match(retained.message!, outcome === "missing" ? /did not expose operation tracking information/
      : outcome === "unsafe" ? /unsafe or unsupported operation address/ : /tracking information is unavailable/);
    assert.doesNotMatch(JSON.stringify(retained), /https:|opaque=|30 seconds|operation failed|was canceled/);
    await pollAssignmentReadback(attempt, [JSON.parse(JSON.stringify(retained))], {
      current: () => true, now: () => (retained.readbackDeadline ?? retained.preparationDeadline)! + 60000,
      wait: async () => { throw new Error("Unavailable tracking cannot resume."); },
      read: async () => { throw new Error("Unavailable tracking cannot be read."); },
      publish: () => { throw new Error("Terminal metadata uncertainty cannot be rewritten."); },
    }, signal());
    if (outcome === "lost") {
      let repaired = false;
      await pollAssignmentReadback(attempt, [{ ...retained, state: "failed" }], {
        current: () => true, now: () => retained.readbackDeadline! + 60000,
        wait: async () => { throw new Error("Legacy unavailable tracking cannot start observation."); },
        read: async () => { throw new Error("Legacy unavailable tracking cannot authorize reads."); },
        publish: receipt => { assert.deepEqual(receipt, retained); repaired = true; },
      }, signal());
      assert.equal(repaired, true, "Only the retained missing-tracking failure should normalize back to uncertainty.");
    }
    assert.equal(writes, 1); assert.equal(afterWriteRequests, 0); assert.equal(afterWriteAuthorizations, 0);
  }
});
test("unsupported retained reference evidence cannot fall through as current canonical evidence", () => {
  const canonical: AssignmentTarget = { ...pollingTarget(), referenceEvidence: {
    source: "adr-point-get", observedAt: "2026-11-05T00:00:00Z", apiVersion: "2026-11-02-preview",
    rawCarrier: "present", currentArmEtag: '"fixture-etag-1"',
  } };
  assert.equal(referenceIsCurrent(canonical), true);
  for (const evidence of [
    { ...canonical.referenceEvidence, source: "retired-unknown" },
    { ...canonical.referenceEvidence, source: "arg" },
    { ...canonical.referenceEvidence, source: "dps-api" },
    { ...canonical.referenceEvidence, rawCarrier: "unsupported" },
    null, {},
  ]) {
    const retained = { ...canonical };
    Reflect.set(retained, "referenceEvidence", evidence);
    assert.equal(referenceIsCurrent(retained), false);
    assert.ok(assignmentBlocker(retained));
  }
  for (const change of [
    { apiVersion: undefined }, { rawCarrier: undefined }, { currentArmEtag: '"changed"' },
    { currentArmEtag: "*" }, { currentArmEtag: '"*"' },
  ]) assert.equal(referenceIsCurrent({ ...canonical,
    referenceEvidence: { ...canonical.referenceEvidence!, ...change } }), false);
  for (const identity of [undefined, { uuid: "invalid", externalDeviceId: "fixture-device" },
    { uuid: "00000000-0000-0000-0000-000000000000", externalDeviceId: "fixture-device" },
    { ...canonical.deviceIdentity!, externalDeviceId: "" }]) {
    assert.equal(referenceIsCurrent({ ...canonical, deviceIdentity: identity }), false);
  }
  assert.equal(referenceIsCurrent({ ...canonical, referenceEvidence: { ...canonical.referenceEvidence!, rawCarrier: "missing" } }), true);
  assert.equal(referenceIsCurrent(pollingTarget()), true, "Existing evidence-free synthetic contracts remain supported.");
});
test("an undispatched deadline expiry reports a known not-sent outcome without uncertain mutation claims", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const originalNow = Date.now;
  let time = 0, dispatches = 0, rechecks = 0;
  Date.now = () => time;
  try {
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async () => {
      dispatches++; return response({});
    } });
    api.recheck = async value => {
      if (++rechecks === 2) await new Promise<void>(resolve => setTimeout(resolve, 31000));
      return value;
    };
    const operation = new ModelAssignmentOperation().run([{ ...pollingTarget(), updatePermission: "granted" }], model(), true,
      api.assignmentServices(source()), () => true, signal(), () => undefined);
    for (let index = 0; index < 20; index++) await Promise.resolve();
    time = 30000; context.mock.timers.tick(30000);
    const receipts = await operation;
    assert.equal(dispatches, 0);
    assert.equal(rechecks, 2, "The fixture must exhaust the budget during the API's internal pre-dispatch recheck.");
    assert.equal(receipts[0].writeDispatched, false);
    assert.equal(receipts[0].state, "failed");
    assert.match(receipts[0].message!, /not sent|not dispatched/i);
    assert.doesNotMatch(receipts[0].message!, /unknown|may have reached/i);
    time = 31000; context.mock.timers.tick(1000);
    for (let index = 0; index < 10; index++) await Promise.resolve();
    assert.equal(dispatches, 0, "A late internal recheck cannot dispatch after its original deadline.");
  } finally { Date.now = originalNow; }
});
test("the wave gate waits for its own whole wave or ten seconds, not one target or an older wave", async context => {
  assert.equal(ASSIGNMENT_WAVE_SIZE, 5); assert.equal(ASSIGNMENT_WAVE_INTERVAL_MS, 10000);
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const pendingTimers = trackedAssignmentTimers(context);
  const originalNow = Date.now, originalRandom = Math.random, controller = new AbortController();
  let time = 0;
  Date.now = () => time;
  Math.random = () => { throw new Error("Wave admission must not add jitter."); };
  const selected = Array.from({ length: 15 }, (_, index) => ({ ...target(), id: `${deviceId}-wave-${index}` }));
  const admitted: { index: number; at: number }[] = [], release = new Map<number, () => void>();
  const operation = new ModelAssignmentOperation();
  const running = operation.run(selected, model(), true, {
    read: async value => { admitted.push({ index: selected.findIndex(target => target.id === value.id), at: time }); return value; },
    write: async (value, _model, _signal, dispatch) => {
      dispatch();
      await new Promise<void>(resolve => release.set(selected.findIndex(target => target.id === value.id), resolve));
      return "completed";
    },
  }, () => true, controller.signal, () => undefined, "one-wave-attempt");
  const flush = async () => { for (let turn = 0; turn < 40; turn++) await Promise.resolve(); };
  try {
    await flush(); assert.equal(admitted.length, 5);
    release.get(0)!(); await flush();
    assert.equal(admitted.length, 5, "One settled target must not refill an individual slot or release a whole new wave.");
    time = 9999; context.mock.timers.tick(9999); await flush(); assert.equal(admitted.length, 5);
    time = 10000; context.mock.timers.tick(1); await flush(); assert.equal(admitted.length, 10);
    time = 11000; context.mock.timers.tick(1000);
    for (const index of [1, 2, 3, 4]) release.get(index)!();
    await flush();
    assert.equal(admitted.length, 10, "An older wave settling cannot open the newer wave's ten-second gate.");
    time = 12000; context.mock.timers.tick(1000); release.get(5)!(); await flush(); assert.equal(admitted.length, 10);
    time = 14000; context.mock.timers.tick(2000);
    for (const index of [6, 7, 8, 9]) release.get(index)!();
    await flush();
    assert.equal(admitted.length, 15, "The entire current wave settling must release its successor before its ten-second timer.");
    assert.deepEqual(admitted.map(value => value.at), [...Array(5).fill(0), ...Array(5).fill(10000), ...Array(5).fill(14000)]);
    for (const index of [10, 11, 12, 13, 14]) release.get(index)!();
    const receipts = await running;
    assert.ok(receipts.every(receipt => receipt.state === "confirmed" && receipt.operationId === "one-wave-attempt"));
    assert.deepEqual(receipts.map(receipt => receipt.targetId), selected.map(target => target.id));
    assert.equal(pendingTimers.size, 0, "Early-finished waves must remove their10s gates and every per-target timer.");
    time = 60000; context.mock.timers.runAll(); await flush();
    assert.equal(admitted.length, 15, "Early completion must retire the old gate rather than trigger later work.");
  } finally { controller.abort(); await running; Date.now = originalNow; Math.random = originalRandom; }
});
test("a whole first wave settling early admits the next five early and clears its gate without cancelling started outcomes", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const pendingTimers = trackedAssignmentTimers(context), originalNow = Date.now;
  let time = 0;
  Date.now = () => time;
  const selected = Array.from({ length: 10 }, (_, index) => ({ ...target(), id: `${deviceId}-early-${index}` }));
  selected[0].writeBlocker = "Current access is read-only.";
  const admissions: number[] = [];
  const running = new ModelAssignmentOperation().run(selected, model(), true, {
    read: async value => { admissions.push(time); return value; },
    write: async (value, _model, _signal, dispatch) => {
      if (value.id === selected[1].id) throw new AssignmentApiError("permission", "Current access denied.");
      dispatch(); await new Promise<void>(resolve => setTimeout(resolve, time ? 500 : 1000)); return "completed";
    },
  }, () => true, signal(), () => undefined);
  const flush = async () => { for (let turn = 0; turn < 50; turn++) await Promise.resolve(); };
  try {
    await flush(); assert.deepEqual(admissions, [0, 0, 0, 0]);
    time = 1000; context.mock.timers.tick(1000); await flush();
    assert.deepEqual(admissions, [0, 0, 0, 0, 1000, 1000, 1000, 1000, 1000]);
    time = 1500; context.mock.timers.tick(500); await flush();
    const receipts = await running;
    assert.deepEqual(receipts.map(receipt => receipt.state), ["blocked", "failed", ...Array(8).fill("confirmed")]);
    assert.ok(receipts.every(receipt => receipt.readbackDeadline === undefined));
    assert.ok(receipts.every(receipt => receipt.preparationDeadline !== undefined),
      "A started synchronous/blocked/failed target must not borrow undefined polling time as the never-admitted sentinel.");
    assert.equal(pendingTimers.size, 0);
  } finally { Date.now = originalNow; }
});
test("cancel and source retirement abort every admitted wave phase and clear gates without later writes", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const pendingTimers = trackedAssignmentTimers(context), originalNow = Date.now;
  let time = 0;
  Date.now = () => time;
  const flush = async () => { for (let turn = 0; turn < 60; turn++) await Promise.resolve(); };
  try {
    for (const phase of ["gate", "preflight", "dispatch", "poll"] as const) {
      const start = time, controller = new AbortController();
      let current = true, settled = false, writes = 0, aborted = 0;
      const selected = Array.from({ length: 12 }, (_, index) => ({ ...target(), id: `${deviceId}-${phase}-${index}` }));
      const releases: (() => void)[] = [], receipts: AssignmentReceipt[] = [];
      const wait = (signal: AbortSignal) => {
        signal.addEventListener("abort", () => { aborted++; }, { once: true });
        return new Promise<void>(resolve => releases.push(resolve));
      };
      const running = new ModelAssignmentOperation().run(selected, model(), true, {
        read: async (value, signal) => { if (phase === "preflight") await wait(signal); return value; },
        write: async (_target, _id, signal, dispatch) => {
          dispatch(); writes++;
          if (phase === "dispatch") await wait(signal);
          return "accepted";
        },
        observe: async (_target, receipt, signal, _current, publish) => {
          if (phase === "gate" && receipt.targetId === selected[0].id) { publish({ ...receipt, state: "confirmed" }); return; }
          await wait(signal);
          publish({ ...receipt, state: "confirmed" });
        },
      }, () => current, controller.signal, receipt => receipts.push(receipt)).finally(() => { settled = true; });
      await flush();
      if (phase === "poll") {
        time = start + 10000; context.mock.timers.tick(10000); await flush();
        assert.equal(writes, 10, "The second wave must really overlap the first before source retirement.");
      } else assert.equal(writes, phase === "preflight" ? 0 : 5);
      current = false; controller.abort();
      const final = await running;
      assert.equal(settled, true); assert.equal(pendingTimers.size, 0);
      const admitted = phase === "poll" ? 10 : 5;
      assert.ok(final.slice(admitted).every(receipt => receipt.state === "cancelled" && receipt.writeDispatched === false &&
        receipt.preparationDeadline === undefined && receipt.readbackDeadline === undefined));
      assert.equal(aborted, phase === "gate" ? 4 : admitted);
      if (phase === "preflight") assert.ok(final.slice(0, admitted).every(receipt =>
        receipt.state === "cancelled" && receipt.writeDispatched === false));
      else assert.ok(final.slice(phase === "gate" ? 1 : 0, admitted).every(receipt =>
        receipt.state === "unknown" && receipt.writeDispatched));
      if (phase === "gate") assert.equal(final[0].state, "confirmed");
      const publications = receipts.length, previousWrites = writes;
      releases.forEach(release => release()); await flush();
      time += 60000; context.mock.timers.tick(60000); await flush();
      assert.equal(receipts.length, publications); assert.equal(writes, previousWrites); assert.equal(pendingTimers.size, 0);
    }
  } finally { Date.now = originalNow; }
});
test("preparation auth and dispatched transport waits stay bounded without inventing polling expiry", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const pendingTimers = trackedAssignmentTimers(context), originalNow = Date.now;
  let time = 0;
  Date.now = () => time;
  const flush = async () => { for (let turn = 0; turn < 60; turn++) await Promise.resolve(); };
  try {
    for (const phase of ["authorization", "dispatched"] as const) {
      const start = time, selected = { ...pollingTarget(), updatePermission: "granted" as const };
      let writes = 0, tokens = 0, release: (() => void) | undefined;
      const api = new HubAssignmentApi({ token: async () => {
        tokens++;
        if (phase === "authorization") await new Promise<void>(resolve => { release = resolve; });
        return "synthetic";
      }, fetcher: async (_url, init) => {
        assert.equal(init?.method, "PATCH"); writes++;
        await new Promise<void>(resolve => { release = resolve; });
        return new Response(null, { status: 204 });
      } });
      api.recheck = async value => value;
      const running = new ModelAssignmentOperation().run([selected], model(), true, api.assignmentServices(source()),
        () => true, signal(), () => undefined);
      await flush(); assert.ok(release);
      time = start + 30000; context.mock.timers.tick(30000);
      const receipts = await running;
      assert.equal(receipts[0].state, phase === "authorization" ? "failed" : "unknown");
      assert.equal(receipts[0].writeDispatched, phase === "dispatched");
      assert.equal(receipts[0].readbackDeadline, undefined);
      assert.notEqual(receipts[0].message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
      assert.equal(writes, phase === "dispatched" ? 1 : 0); assert.equal(tokens, 1);
      if (phase === "authorization") assert.match(receipts[0].message!, /not sent|not dispatched/i);
      else assert.match(receipts[0].message!, /may have reached/);
      release!(); await flush(); assert.equal(writes, phase === "dispatched" ? 1 : 0); assert.equal(pendingTimers.size, 0);
    }
  } finally { Date.now = originalNow; }
});
test("a fresh polling window starts after delayed native acceptance and outlives the preparation deadline", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const originalNow = Date.now, originalRandom = Math.random;
  let time = 0;
  Date.now = () => time; Math.random = () => 0.5;
  try {
    for (const pending of [false, true]) {
      const start = time, selected = { ...pollingTarget(), updatePermission: "granted" as const };
      const events: { kind: string; at: number }[] = [];
      const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
        if (init?.method === "PATCH") {
          events.push({ kind: "write", at: time });
          await new Promise<void>(resolve => setTimeout(resolve, 29000));
          events.push({ kind: "accepted", at: time });
          return new Response(null, { status: 202, headers: { "Azure-AsyncOperation":
            `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/delayed?api-version=2026-11-02-preview&sig=synthetic` } });
        }
        assert.match(String(url), /operationStatuses\/delayed/);
        events.push({ kind: "poll", at: time });
        return response({ status: !pending && time - start >= 53000 ? "Succeeded" : "Running" });
      } });
      api.recheck = async value => value;
      let settled = false;
      const receipts: AssignmentReceipt[] = [];
      const running = new ModelAssignmentOperation().run([selected], model(), true, api.assignmentServices(source()),
        () => true, signal(), receipt => receipts.push(receipt)).finally(() => { settled = true; });
      for (let second = 0; second < 65 && !settled; second++) {
        for (let turn = 0; turn < 60; turn++) await Promise.resolve();
        if (settled) break;
        time += 1000; context.mock.timers.tick(1000);
      }
      const result = await running;
      assert.equal(result[0].readbackDeadline, start + 59000, "Preparation/PATCH time cannot consume the fresh polling window.");
      assert.equal(result[0].preparationDeadline, start + ASSIGNMENT_OPERATION_TIMEOUT_MS);
      assert.equal(api.assignmentServices(source()).trackingDeadline!(selected), result[0].readbackDeadline,
        "The original private operation and published polling receipt must share the same deadline.");
      assert.equal(result[0].state, pending ? "unknown" : "confirmed");
      assert.equal(events.filter(event => event.kind === "write").length, 1);
      assert.deepEqual(events.filter(event => event.kind === "poll").map(event => event.at - start),
        [32000, 35000, 38000, 43000, 48000, 53000]);
      assert.equal(receipts.find(receipt => receipt.state === "accepted")?.readbackDeadline, start + 59000);
      assert.doesNotMatch(JSON.stringify(receipts), /sig=|https:/);
      if (pending) {
        assert.equal(result[0].message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
        assert.equal(time, start + 59000, "Pending observation must expire at its own elapsed30s, without one final read.");
      }
    }
  } finally { Date.now = originalNow; Math.random = originalRandom; }
});
test("delayed accepted native tracking retains one polling deadline and scheduled wait across remount and read-only retry", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const originalNow = Date.now;
  let time = 0, writes = 0, reads = 0;
  Date.now = () => time;
  const selected = { ...pollingTarget(), updatePermission: "granted" as const };
  const address = `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/retained-window?api-version=2026-11-02-preview`;
  try {
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (_url, init) => {
      assert.equal(init?.method, "PATCH"); writes++;
      await new Promise<void>(resolve => setTimeout(resolve, 29000));
      return new Response(null, { status: 202, headers: { Location: address } });
    } });
    api.recheck = async value => value;
    const services = api.assignmentServices(source());
    const running = new ModelAssignmentOperation().run([selected], model(), true, { ...services, observe: undefined },
      () => true, signal(), () => undefined, "original-operation");
    for (let turn = 0; turn < 40; turn++) await Promise.resolve();
    time = 29000; context.mock.timers.tick(29000);
    let receipt = (await running)[0];
    assert.equal(receipt.readbackDeadline, 59000); assert.equal(services.trackingDeadline!(selected), 59000);
    const attempt = { operationId: receipt.operationId, modelId: receipt.modelId, targetIds: [selected.id], targets: [selected] };
    const unmount = new AbortController();
    await assert.rejects(() => pollAssignmentReadback(attempt, [receipt], {
      current: () => true, random: () => 0.5,
      wait: async delay => { assert.equal(delay, 3000); unmount.abort(); },
      read: async () => { throw new Error("No read before the original wait."); }, publish: value => { receipt = value; },
    }, unmount.signal), { name: "AbortError" });
    assert.equal(receipt.readbackNotBefore, 32000);
    time = 31000;
    const resumed = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async url => {
      assert.equal(String(url), address); reads++;
      if (reads === 1) return response({}, 403);
      return response({ status: "Succeeded" });
    } });
    await pollAssignmentReadback(attempt, [JSON.parse(JSON.stringify(receipt))], {
      current: () => true, now: () => time, random: () => { throw new Error("A retained scheduled wait cannot be rerolled."); },
      wait: async delay => { assert.equal(delay, 1000); time += delay; },
      read: (target, id, abort, trackingId) => resumed.readAssignment(target, id, abort, trackingId),
      publish: value => { receipt = value; },
    }, signal());
    assert.equal(receipt.state, "failed"); assert.equal(receipt.readbackPollCount, 1); assert.equal(receipt.readbackDeadline, 59000);
    receipt = { ...receipt, state: "unknown", readbackFailure: undefined };
    await pollAssignmentReadback(attempt, [receipt], {
      current: () => true, now: () => time, random: () => 0.5,
      wait: async delay => { assert.equal(delay, 3000); time += delay; },
      read: (target, id, abort, trackingId) => resumed.readAssignment(target, id, abort, trackingId),
      publish: value => { receipt = value; },
    }, signal());
    assert.equal(receipt.state, "confirmed"); assert.equal(receipt.readbackDeadline, 59000);
    assert.equal(receipt.readbackPollCount, 2); assert.equal(services.trackingDeadline!(selected), 59000);
    assert.equal(writes, 1); assert.equal(reads, 2); assert.equal(time, 35000);
  } finally { Date.now = originalNow; }
});
test("delayed native assignments start five-device waves at zero ten and twenty seconds while older operations remain active", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const originalNow = Date.now, originalRandom = Math.random;
  let time = 0;
  Date.now = () => time; Math.random = () => 0.5;
  const events: { kind: string; device: string; at: number }[] = [];
  const selected = Array.from({ length: 12 }, (_, index) => ({ ...pollingTarget(),
    id: `${deviceId}-${index}`, name: `Delayed device ${index + 1}`,
    updatePermission: "granted" as const,
    deviceIdentity: { ...pollingTarget().deviceIdentity!, externalDeviceId: `delayed-${index}` } }));
  const active = new Set<string>(), admissions = new Map<string, number>(), deadlines = new Map<string, number>();
  const completionOrder: string[] = [];
  let maxActive = 0;
  const controller = new AbortController();
  try {
    const api = new HubAssignmentApi({ token: async () => "synthetic",
      fetcher: async (url, init) => {
        const address = new URL(String(url));
        if (init?.method === "PATCH") {
          const device = selected.find(value => address.pathname === value.id)!;
          assert.ok(device);
          assert.equal(new Headers(init.headers).get("If-Match"), device.etag);
          assert.deepEqual(JSON.parse(String(init.body)), { properties: { deviceTypeRefs: [document.id] } });
          events.push({ kind: "write", device: device.id, at: time });
          await new Promise<void>(resolve => setTimeout(resolve, [29000, 18000, 16000, 14000, 8000][selected.indexOf(device)] || 8000));
          events.push({ kind: "accepted", device: device.id, at: time });
          return new Response(null, { status: 202, headers: { "Azure-AsyncOperation":
            `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/${selected.indexOf(device)}?api-version=2026-11-02-preview`,
            "Retry-After": "3600" } });
        }
        assert.ok(address.pathname.includes("/asyncOperationStatuses/"), "No post-write device GET is allowed.");
        const device = selected[Number(address.pathname.split("/").at(-1))];
        events.push({ kind: "status", device: device.id, at: time });
        return response({ status: "Succeeded" });
      } });
    const rechecks: string[] = [];
    api.recheck = async value => {
      if (!rechecks.includes(value.id)) {
        admissions.set(value.id, time); active.add(value.id); maxActive = Math.max(maxActive, active.size);
      }
      rechecks.push(value.id); return value;
    };
    let settled = false;
    const operation = new ModelAssignmentOperation();
    const running = operation.run(selected, model(), true,
      api.assignmentServices(source()), () => true, controller.signal, receipt => {
        if (receipt.readbackDeadline) deadlines.set(receipt.targetId, receipt.readbackDeadline);
        if (receipt.state === "confirmed") {
          active.delete(receipt.targetId); completionOrder.push(receipt.targetId);
        }
      }).finally(() => { settled = true; });
    for (let turn = 0; turn < 30; turn++) await Promise.resolve();
    assert.equal(active.size, 5); assert.equal(admissions.size, 5);
    await assert.rejects(() => operation.run(selected, model(), true, api.assignmentServices(source()),
      () => true, controller.signal, () => undefined), /already running/);
    for (let index = 0; index < 100 && !settled; index++) {
      for (let turn = 0; turn < 60; turn++) await Promise.resolve();
      time += 1000; context.mock.timers.tick(1000);
    }
    if (!settled) controller.abort();
    const receipts = await running;
    assert.equal(events.filter(event => event.kind === "status").length, 12);
    assert.ok(receipts.every(receipt => receipt.state === "confirmed"));
    assert.deepEqual(receipts.map(receipt => receipt.targetId), selected.map(value => value.id));
    assert.ok(maxActive > 5, "Overlapping waves must not retain an implicit five-in-flight cap."); assert.equal(active.size, 0);
    assert.notDeepEqual(completionOrder, selected.map(value => value.id), "Different device durations must complete out of order.");
    assert.equal(events.filter(event => event.kind === "write").length, 12);
    for (const value of selected) {
      const acceptance = events.find(event => event.kind === "accepted" && event.device === value.id)!;
      const status = events.find(event => event.kind === "status" && event.device === value.id)!;
      assert.ok(status.at - acceptance.at < ASSIGNMENT_POLL_TIMEOUT_MS);
      assert.equal(deadlines.get(value.id), acceptance.at + ASSIGNMENT_POLL_TIMEOUT_MS);
      assert.equal(rechecks.filter(id => id === value.id).length, 2, "Both pre-write checks must remain.");
    }
    assert.deepEqual(selected.map(value => admissions.get(value.id)), [...Array(5).fill(0), ...Array(5).fill(10000), ...Array(2).fill(20000)]);
    assert.ok(time > ASSIGNMENT_OPERATION_TIMEOUT_MS, "The fixture must cover a batch longer than one device's cutoff.");
    assert.ok(selected.every(value => value.reference.kind === "unassigned"));
  } finally { controller.abort(); Date.now = originalNow; Math.random = originalRandom; }
});
test("assignment verification without the original native operation never reads the mutable device", async () => {
  let tokens = 0, requests = 0;
  const api = new HubAssignmentApi({ token: async () => { tokens++; return "synthetic"; },
    fetcher: async () => { requests++; throw new Error("Must not fetch a device as fallback."); } });
  await assert.rejects(() => api.readAssignment(pollingTarget(), document.id, signal()), /tracking information is unavailable/);
  await assert.rejects(() => api.readAssignment(pollingTarget(), document.id, signal(), "lost-ephemeral-handle"), /tracking information is unavailable/);
  assert.equal(tokens, 0); assert.equal(requests, 0);
});
test("parallel native devices settle failures and timeouts independently without consuming queued deadlines or replaying writes", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const originalNow = Date.now, originalRandom = Math.random;
  let time = 0;
  Date.now = () => time; Math.random = () => 0.5;
  const selected = Array.from({ length: 8 }, (_, index) => ({ ...pollingTarget(),
    id: `${deviceId}-${index}`, name: `Mixed device ${index + 1}`, updatePermission: "granted" as const,
    deviceIdentity: { ...pollingTarget().deviceIdentity!, externalDeviceId: `mixed-${index}` } }));
  const admissions = new Map<string, number>(), latest = new Map<string, AssignmentReceipt>();
  const rechecks = new Map<string, number>(), writes: string[] = [], polls: string[] = [];
  const controller = new AbortController();
  try {
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
      const address = new URL(String(url));
      if (init?.method === "PATCH") {
        const index = selected.findIndex(value => value.id === address.pathname);
        assert.ok(index >= 0);
        assert.equal(new Headers(init.headers).get("If-Match"), selected[index].etag);
        assert.deepEqual(JSON.parse(String(init.body)), { properties: { deviceTypeRefs: [document.id] } });
        writes.push(selected[index].id);
        if (index === 1) return response({}, 403);
        return new Response(null, { status: 202, headers: { Location:
          `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/${index}?api-version=2026-11-02-preview` } });
      }
      assert.ok(address.pathname.includes("/operationStatuses/"), "Parallel completion cannot fall back to a root GET.");
      const index = Number(address.pathname.split("/").at(-1));
      polls.push(selected[index].id);
      return response({ status: index === 0 ? "Failed" : index === 2 ? "Running" : "Succeeded" });
    } });
    api.recheck = async value => {
      const count = (rechecks.get(value.id) || 0) + 1; rechecks.set(value.id, count);
      if (count === 1) admissions.set(value.id, time);
      if (value.id === selected[7].id) return new Promise(() => undefined);
      return value.id === selected[3].id && count === 2 ? { ...value, etag: '"changed-before-dispatch"' } : value;
    };
    let settled = false;
    const running = new ModelAssignmentOperation().run(selected, model(), true, api.assignmentServices(source()),
      () => true, controller.signal, receipt => latest.set(receipt.targetId, receipt)).finally(() => { settled = true; });
    for (let index = 0; index < 80 && !settled; index++) {
      for (let turn = 0; turn < 60; turn++) await Promise.resolve();
      time += 1000; context.mock.timers.tick(1000);
    }
    if (!settled) controller.abort();
    const receipts = await running;
    assert.deepEqual(receipts.map(receipt => receipt.state),
      ["failed", "failed", "unknown", "blocked", "confirmed", "confirmed", "confirmed", "failed"]);
    assert.deepEqual(receipts.map(receipt => receipt.targetId), selected.map(value => value.id));
    assert.deepEqual(receipts, selected.map(value => latest.get(value.id)), "Concurrent publications must not lose another device's final receipt.");
    assert.equal(new Set(writes).size, writes.length);
    assert.ok(!writes.includes(selected[3].id) && !writes.includes(selected[7].id));
    assert.equal(receipts[2].writeDispatched, true);
    assert.equal(receipts[2].message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
    assert.equal(receipts[7].writeDispatched, false);
    assert.match(receipts[7].message!, /not sent|not dispatched/i);
    assert.ok(admissions.get(selected[7].id)! > 0);
    for (const receipt of receipts) {
      assert.equal(receipt.preparationDeadline, admissions.get(receipt.targetId)! + ASSIGNMENT_OPERATION_TIMEOUT_MS);
      assert.equal(receipt.readbackDeadline, [0, 2, 4, 5, 6].some(index => selected[index].id === receipt.targetId)
        ? admissions.get(receipt.targetId)! + ASSIGNMENT_POLL_TIMEOUT_MS : undefined);
    }
    assert.ok(polls.includes(selected[0].id) && !polls.includes(selected[1].id));
    assert.ok(polls.filter(id => id === selected[2].id).length > 1, "The pending native operation must reach its own observation cutoff.");
    assert.ok(time >= admissions.get(selected[7].id)! + ASSIGNMENT_OPERATION_TIMEOUT_MS);
  } finally { controller.abort(); Date.now = originalNow; Math.random = originalRandom; }
});
test("canceling parallel native work stops queued admission and retains confirmed receipts without claiming untouched sends", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const originalNow = Date.now, originalRandom = Math.random;
  let time = 0, current = true;
  Date.now = () => time; Math.random = () => 0.5;
  const selected = Array.from({ length: 8 }, (_, index) => ({ ...pollingTarget(),
    id: `${deviceId}-${index}`, name: `Canceled device ${index + 1}`, updatePermission: "granted" as const }));
  const writes: string[] = [], rechecks: string[] = [], publications: AssignmentReceipt[] = [];
  const controller = new AbortController();
  let releasePreflight: (() => void) | undefined, abortedReads = 0, releaseStatus: (() => void) | undefined;
  try {
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
      const address = new URL(String(url));
      if (init?.method === "PATCH") {
        const index = selected.findIndex(value => value.id === address.pathname);
        writes.push(selected[index].id);
        return new Response(null, { status: 202, headers: { Location:
          `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/${index}?api-version=2026-11-02-preview` } });
      }
      assert.ok(address.pathname.includes("/operationStatuses/"));
      if (address.pathname.endsWith("/0")) return response({ status: "Succeeded" });
      init?.signal?.addEventListener("abort", () => { abortedReads++; }, { once: true });
      return new Promise(resolve => { releaseStatus = () => resolve(response({ status: "Succeeded" })); });
    } });
    api.recheck = async value => {
      rechecks.push(value.id);
      if (value.id === selected[4].id) await new Promise<void>(resolve => { releasePreflight = resolve; });
      return value;
    };
    const running = new ModelAssignmentOperation().run(selected, model(), true, api.assignmentServices(source()),
      () => current, controller.signal, receipt => {
        publications.push(receipt);
        if (receipt.targetId === selected[0].id && receipt.state === "confirmed") {
          current = false; controller.abort();
        }
      });
    for (let turn = 0; turn < 60; turn++) await Promise.resolve();
    assert.equal(writes.length, 4); assert.ok(releasePreflight);
    time = 3000; context.mock.timers.tick(3000);
    const receipts = await running;
    assert.equal(receipts[0].state, "confirmed", "A cancellation racing operation success must not erase confirmed history.");
    assert.ok(receipts.slice(1, 4).every(receipt => receipt.state === "unknown" && receipt.writeDispatched));
    assert.ok(receipts.slice(4).every(receipt => receipt.state === "cancelled" && receipt.writeDispatched === false));
    assert.ok(receipts.slice(5).every(receipt => receipt.readbackDeadline === undefined));
    assert.ok(!rechecks.some(id => selected.slice(5).some(value => value.id === id)));
    assert.equal(writes.length, 4);
    const beforeLate = publications.length;
    releasePreflight!(); releaseStatus?.();
    for (let turn = 0; turn < 30; turn++) await Promise.resolve();
    assert.equal(writes.length, 4); assert.equal(publications.length, beforeLate);
    assert.ok(abortedReads <= 3);
    await assert.rejects(() => new ModelAssignmentOperation().run([selected[0], { ...selected[0], id: selected[0].id.toUpperCase() }],
      model(), true, api.assignmentServices(source()), () => true, signal(), () => undefined), /distinct devices/);
  } finally {
    controller.abort(); releasePreflight?.(); releaseStatus?.();
    Date.now = originalNow; Math.random = originalRandom;
  }
});
test("native operation success survives a later cleared device without a post-write device GET", async () => {
  const fixture = fixtureClient({ completePoint: true });
  const selected = (await fixture.api.discover(source(), signal())).targets[0];
  const operation = `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/native-1?api-version=2026-11-02-preview&sig=synthetic-only`;
  let patched = false, polls = 0, patchCount = 0, tokenCount = 0;
  const api = new HubAssignmentApi({ token: async () => { tokenCount++; return "synthetic"; },
    fetcher: async (url, init) => {
      if (init?.method === "PATCH") {
        patched = true; patchCount++;
        return new Response(null, { status: 202, headers: { "Azure-AsyncOperation": operation, Location: operation, "Retry-After": "3600" } });
      }
      if (String(url) === operation) {
        polls++;
        return new Response(JSON.stringify({ status: polls === 1 ? "Running" : "Succeeded" }),
          { headers: { "Retry-After": "7" } });
      }
      assert.equal(patched, false, "Completion must not fetch a mutable device that another actor has cleared.");
      return fixture.fetcher(url, init);
    } });
  const services = { ...api.assignmentServices(source()), observe: undefined };
  const receipts = await new ModelAssignmentOperation().run([selected], model(), true, services, () => true, signal(), () => undefined, "local-attempt");
  assert.equal(receipts[0].state, "accepted"); assert.ok(receipts[0].trackingId);
  assert.notEqual(receipts[0].trackingId, "local-attempt");
  assert.ok(!JSON.stringify(receipts).includes("sig=") && !JSON.stringify(receipts).includes("asyncOperationStatuses"));
  const resumed = new HubAssignmentApi({ token: async () => { tokenCount++; return "synthetic"; },
    fetcher: async (url, init) => {
      assert.equal(String(url), operation); assert.ok(!init?.method || init.method === "GET");
      polls++; return new Response(JSON.stringify({ status: polls === 1 ? "Running" : "Succeeded" }), { headers: { "Retry-After": "3600" } });
    } });
  const updates: AssignmentReceipt[] = [], delays: number[] = [];
  await pollAssignmentReadback({ operationId: "local-attempt", modelId: document.id, targetIds: [selected.id], targets: [selected] },
    JSON.parse(JSON.stringify(receipts)), { current: () => true, wait: async delay => { delays.push(delay); },
      read: (target, id, abort, trackingId) => resumed.readAssignment(target, id, abort, trackingId), publish: value => updates.push(value) }, signal());
  assert.equal(updates.at(-1)?.state, "confirmed");
  assert.equal(patchCount, 1); assert.equal(polls, 2);
  assert.ok(delays.every(delay => delay >= 1900 && delay <= 4000));
  assert.equal(selected.reference.kind, "unassigned", "An operation receipt must not synthesize current resource state.");
  assert.ok(tokenCount > 0);
});
test("unsafe missing and unreadable mutation operation addresses never receive polling authorization", async () => {
  const fixture = fixtureClient({ completePoint: true });
  const selected = (await fixture.api.discover(source(), signal())).targets[0];
  for (const address of [undefined, "http://management.azure.com/unsafe", "https://evil.invalid/operation",
    `https://management.azure.com${deviceId}?api-version=2026-11-02-preview`,
    `https://management.azure.com/subscriptions/other/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/123`,
    "https://management.azure.com@evil.invalid/operation",
    `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/%2e%2e/resource`,
    `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/other/../operationStatuses/123`,
    `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/123#fragment`,
    `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus\\operationStatuses\\123`,
  ]) {
    let patched = false, posts = 0, pollAuth = 0;
    const api = new HubAssignmentApi({ token: async () => { if (patched) pollAuth++; return "synthetic"; },
      fetcher: async (url, init) => {
        if (init?.method === "PATCH") {
          posts++; patched = true;
          return new Response(null, { status: 202, headers: address ? { Location: address } : {} });
        }
        assert.equal(patched, false, "Invalid metadata must not trigger device or operation fallback reads.");
        return fixture.fetcher(url, init);
      } });
    await assert.rejects(() => api.patchReference(selected, document.id, signal(), source()), /operation (address|tracking)/);
    assert.equal(posts, 1); assert.equal(pollAuth, 0);
  }
});
test("documented Location-only operationStatuses and observed Azure-AsyncOperation asyncOperationStatuses are both supported", async () => {
  const subscription = "00000000-0000-0000-0000-000000000000";
  const selected: AssignmentTarget = { ...pollingTarget(),
    updatePermission: "granted",
    id: deviceId.replace(deviceId.split("/")[2], subscription),
    namespaceId: nsId.replace(deviceId.split("/")[2], subscription),
    hubIds: [hubId.replace(deviceId.split("/")[2], subscription)] };
  const documented = "https://management.azure.com/subscriptions/00000000-0000-0000-0000-000000000000/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/00000000-0000-0000-0000-000000000001?api-version=2026-11-02-preview";
  const observed = `https://management.azure.com/subscriptions/${subscription}/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/00000000-0000-0000-0000-000000000002?api-version=2026-11-02-preview&sig=synthetic%2Fopaque`;
  for (const [headers, expected] of [
    [{ Location: documented, "Retry-After": "3600" }, documented],
    [{ "Azure-AsyncOperation": observed, Location: documented, "Retry-After": "3600" }, observed],
  ] as const) {
    let writes = 0, polls = 0, patched = false;
    const api = new HubAssignmentApi({ token: async scope => {
      assert.equal(scope, ASSIGNMENT_ARM_SCOPE); return "synthetic";
    }, fetcher: async (url, init) => {
      if (init?.method === "PATCH") {
        writes++; patched = true;
        assert.equal(new Headers(init.headers).get("If-Match"), selected.etag);
        assert.deepEqual(JSON.parse(String(init.body)), { properties: { deviceTypeRefs: [document.id] } });
        return new Response(null, { status: 202, headers });
      }
      assert.equal(patched, true);
      assert.equal(String(url), expected);
      assert.equal(init?.redirect, "error");
      polls++; return response({ status: "Succeeded" });
    } });
    api.recheck = async value => value;
    const services = api.assignmentServices(source());
    assert.equal(await services.write(selected, document.id, signal(), () => undefined), "accepted");
    const tracking = services.trackingId!(selected);
    assert.ok(tracking && !tracking.includes("sig=") && !tracking.includes("https:"));
    assert.equal((await api.readAssignment(selected, document.id, signal(), tracking)).confirmed, true);
    assert.equal(writes, 1); assert.equal(polls, 1);
    assert.equal(selected.reference.kind, "unassigned");
  }
});
test("operation endpoints report failed canceled malformed and throttled results without querying device state", async () => {
  const fixture = fixtureClient({ completePoint: true });
  const selected = (await fixture.api.discover(source(), signal())).targets[0];
  for (const status of ["Failed", "Canceled", "Succeeded", "Unexpected", "unreadable", "throttled"]) {
    const operation = `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/${status}?api-version=2026-11-02-preview`;
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
      if (init?.method === "PATCH") return new Response(null, { status: 202, headers: { Location: operation } });
      if (String(url) === operation) return status === "unreadable" ? new Response("not JSON")
        : status === "throttled" ? new Response("", { status: 429, headers: { "Retry-After": "8" } })
          : response({ status });
      return fixture.fetcher(url, init);
    } });
    const services = api.assignmentServices(source());
    await services.write(selected, document.id, signal(), () => undefined);
    const read = () => api.readAssignment(selected, document.id, signal(), services.trackingId!(selected));
    if (status === "Unexpected" || status === "unreadable") await assert.rejects(read, /status|response could not be read/);
    else if (status === "throttled") await assert.rejects(read, error => error instanceof AssignmentApiError && error.retryable && error.retryAfterMs === 8000);
    else {
      const result = await read();
      assert.equal(result.confirmed, status === "Succeeded");
      assert.equal(result.terminalState, status === "Failed" ? "failed" : status === "Canceled" ? "cancelled" : undefined);
    }
  }
});
test("operation polling keeps the original deadline across retry and bounds an in-flight request", async () => {
  const selected = pollingTarget(), attempt = { operationId: "deadline", modelId: document.id, targetIds: [selected.id], targets: [selected] };
  const receipt: AssignmentReceipt = { operationId: attempt.operationId, modelId: document.id, targetId: selected.id,
    name: selected.name, kind: "device", state: "accepted", readbackDeadline: Date.now() + 30 };
  const updates: AssignmentReceipt[] = [];
  let reads = 0;
  await pollAssignmentReadback(attempt, [receipt], { current: () => true, wait: async () => undefined,
    read: async () => { reads++; return new Promise(() => undefined); }, publish: value => updates.push(value) }, signal());
  assert.equal(updates.at(-1)?.state, "unknown");
  assert.equal(updates.at(-1)!.message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
  assert.equal(updates.at(-1)!.readbackDeadline, receipt.readbackDeadline);
  await pollAssignmentReadback(attempt, [{ ...updates.at(-1)!, state: "unknown" }], {
    current: () => true, wait: async () => undefined, read: async () => { reads++; return { confirmed: true }; },
    publish: value => updates.push(value),
  }, signal());
  assert.equal(reads, 1); assert.equal(updates.at(-1)!.message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
});
test("native synchronous 200 and 204 assignment completion never fetches device confirmation", async () => {
  for (const status of [200, 204]) {
    const fixture = fixtureClient({ completePoint: true });
    const selected = (await fixture.api.discover(source(), signal())).targets[0];
    let patched = false, writes = 0;
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) => {
      if (init?.method === "PATCH") { writes++; patched = true; return new Response(null, { status }); }
      assert.equal(patched, false, "A synchronous completion must not perform post-write device confirmation.");
      return fixture.fetcher(url, init);
    } });
    const receipts = await new ModelAssignmentOperation().run([selected], model(), true,
      api.assignmentServices(source()), () => true, signal(), () => undefined);
    assert.equal(receipts[0].state, "confirmed"); assert.equal(writes, 1);
    assert.equal(receipts[0].readbackDeadline, undefined);
    assert.ok(receipts[0].preparationDeadline);
    assert.equal(selected.reference.kind, "unassigned");
  }
});
test("terminal native operation results settle receipts without fabricating confirmation", async () => {
  const selected = pollingTarget();
  for (const terminalState of ["failed", "cancelled"] as const) {
    const updates: AssignmentReceipt[] = [];
    await pollAssignmentReadback({ operationId: "terminal", modelId: document.id, targetIds: [selected.id], targets: [selected] },
      [{ operationId: "terminal", modelId: document.id, targetId: selected.id, name: selected.name, kind: "device",
        state: "accepted", readbackDeadline: Date.now() + 10000 }],
      { current: () => true, wait: async () => undefined, read: async () => ({ confirmed: false, terminalState }),
        publish: receipt => updates.push(receipt) }, signal());
    assert.equal(updates.at(-1)?.state, terminalState);
    assert.equal(updates.some(receipt => receipt.state === "confirmed"), false);
  }
});
test("operation polling uses tiered jittered waits across remount and ignores all service delays", async () => {
  const selected = pollingTarget(), attempt = { operationId: "cadence", modelId: document.id,
    targetIds: [selected.id], targets: [selected] };
  let time = 1000, polls = 0, rolls = 0;
  let receipt: AssignmentReceipt = { operationId: "cadence", modelId: document.id, targetId: selected.id,
    name: selected.name, kind: "device", state: "accepted", readbackDeadline: 601000 };
  const delays: number[] = [];
  const invoke = async (count: number, random: () => number) => {
    const controller = new AbortController();
    await assert.rejects(() => pollAssignmentReadback(attempt, [receipt], {
      now: () => time, random: () => { rolls++; return random(); }, current: () => true,
      wait: async delay => { delays.push(delay); time += delay; },
      read: async () => {
        polls++;
        if (polls === count) controller.abort();
        if (polls % 2) throw new AssignmentApiError("request", "Throttled.", true, 429, undefined, 3600000);
        return { confirmed: false, retryAfterMs: 3600000 };
      },
      publish: value => { receipt = value; },
    }, controller.signal), { name: "AbortError" });
  };
  await invoke(4, () => 0.5);
  assert.deepEqual(delays, [3000, 3000, 3000, 5000]);
  await invoke(8, () => 0.5);
  assert.deepEqual(delays, [3000, 3000, 3000, 5000, 5000, 5000, 10000, 10000]);
  assert.equal(rolls, 8); assert.equal(receipt.readbackDeadline, 601000);

  for (const [randomValue, expected] of [[0, 2000], [1, 4000]] as const) {
    const controller = new AbortController();
    let saved: AssignmentReceipt = { ...receipt, readbackPollCount: 0, readbackNotBefore: undefined, readbackDeadline: time + 100000 };
    await assert.rejects(() => pollAssignmentReadback(attempt, [saved], {
      now: () => time, random: () => randomValue, current: () => true,
      wait: async delay => { assert.equal(delay, expected); controller.abort(); },
      read: async () => { throw new Error("No read before canceled wait."); },
      publish: value => { saved = value; },
    }, controller.signal), { name: "AbortError" });
    const next = new AbortController();
    time += 500;
    await assert.rejects(() => pollAssignmentReadback(attempt, [saved], {
      now: () => time, random: () => { throw new Error("Must not reroll a scheduled wait."); }, current: () => true,
      wait: async delay => { assert.equal(delay, expected - 500); next.abort(); },
      read: async () => { throw new Error("No canceled read."); }, publish: () => undefined,
    }, next.signal), { name: "AbortError" });
  }
  const deadlineUpdates: AssignmentReceipt[] = [];
  const deadline = time + 50;
  await pollAssignmentReadback(attempt, [{ ...receipt, readbackNotBefore: undefined, readbackDeadline: deadline }], {
    now: () => time, random: () => 1, current: () => true,
    wait: async delay => { assert.equal(delay, 50); time += delay; },
    read: async () => { throw new Error("No operation request after expiry."); },
    publish: value => deadlineUpdates.push(value),
  }, signal());
  assert.equal(deadlineUpdates.at(-1)?.state, "unknown");
  assert.equal(deadlineUpdates.at(-1)!.message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
});
test("permanently pending tracking stops at the original 30 seconds with fresh jitter and no final request", async () => {
  assert.equal(ASSIGNMENT_OPERATION_TIMEOUT_MS, 30000);
  const selected = pollingTarget(), attempt = { operationId: "bounded", modelId: document.id, targetIds: [selected.id], targets: [selected] };
  let time = 0, polls = 0, rolls = 0;
  const waits: number[] = [], receipts: AssignmentReceipt[] = [];
  await pollAssignmentReadback(attempt, [{ operationId: "bounded", modelId: document.id, targetId: selected.id,
    name: selected.name, kind: "device", state: "accepted", readbackDeadline: 30000 }], {
    now: () => time, random: () => rolls++ % 2, current: () => true,
    wait: async delay => { waits.push(delay); time += delay; },
    read: async () => { polls++; return { confirmed: false, retryAfterMs: 3600000 }; },
    publish: receipt => receipts.push(receipt),
  }, signal());
  assert.deepEqual(waits, [2000, 4000, 2000, 6000, 4000, 6000, 6000]);
  assert.equal(rolls, 7); assert.equal(polls, 6); assert.equal(time, 30000);
  assert.equal(receipts.at(-1)?.state, "unknown");
  assert.equal(receipts.at(-1)!.message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
  assert.ok(!receipts.at(-1)!.message!.includes("operation failed"));
  const before = polls;
  await pollAssignmentReadback(attempt, [receipts.at(-1)!], { now: () => time, current: () => true,
    read: async () => { polls++; return { confirmed: true }; }, publish: () => undefined,
  }, signal());
  assert.equal(polls, before);
});
test("a hung status request is locally aborted at the original 30 second cutoff without remote cancellation", async context => {
  context.mock.timers.enable({ apis: ["setTimeout"] });
  const selected = pollingTarget(), updates: AssignmentReceipt[] = [];
  let time = 0, calls = 0, aborted = false;
  const running = pollAssignmentReadback({ operationId: "hung", modelId: document.id,
    targetIds: [selected.id], targets: [selected] }, [{ operationId: "hung", modelId: document.id,
      targetId: selected.id, name: selected.name, kind: "device", state: "accepted", readbackDeadline: 30000 }], {
    now: () => time, random: () => 0.5, current: () => true,
    wait: async delay => { time += delay; },
    read: async (_target, _id, abort) => {
      calls++; abort.addEventListener("abort", () => { aborted = true; });
      return new Promise(() => undefined);
    }, publish: receipt => updates.push(receipt),
  }, signal());
  for (let index = 0; index < 10; index++) await Promise.resolve();
  assert.equal(calls, 1); assert.equal(time, 3000);
  time = 30000; context.mock.timers.tick(27000);
  await running;
  assert.equal(aborted, true);
  assert.equal(updates.at(-1)?.state, "unknown");
  assert.equal(updates.at(-1)!.message, ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
  assert.equal(calls, 1);
});
test("readback polling waits 3000ms, never overlaps and stops reading each confirmed device", async () => {
  const first = pollingTarget(), second = { ...first, id: first.id + "-second", name: "Second" };
  const attempt = { operationId: "owned", modelId: document.id, targetIds: [first.id, second.id], targets: [first, second] };
  const receipts: AssignmentReceipt[] = attempt.targets.map(value => ({ operationId: attempt.operationId, modelId: document.id,
    targetId: value.id, name: value.name, kind: "device", state: "accepted", readbackDeadline: Date.now() + 600000 }));
  const calls: string[] = [], delays: number[] = [], updates: AssignmentReceipt[] = [];
  const waits: (() => void)[] = [];
  let release: (() => void) | undefined, firstReads = 0;
  const running = pollAssignmentReadback(attempt, receipts, { current: () => true, random: () => 0.5,
    wait: milliseconds => { delays.push(milliseconds); return new Promise(resolve => waits.push(resolve)); },
    read: async value => {
      calls.push(value.id);
      if (value.id === first.id && firstReads++ === 0) await new Promise<void>(resolve => { release = resolve; });
      return { target: { ...value, reference: { kind: "assigned", ids: [document.id] } }, confirmed: true };
    }, publish: value => updates.push(value) }, signal());
  assert.equal(ASSIGNMENT_POLL_INTERVAL_MS, 3000); assert.ok(delays[0] >= 2900 && delays[0] <= 3000); assert.equal(calls.length, 0);
  waits.shift()!(); for (let i = 0; i < 10; i++) await Promise.resolve();
  assert.deepEqual(calls, [first.id]); assert.equal(waits.length, 0);
  release!(); await running;
  assert.deepEqual(calls, [first.id, second.id]); assert.equal(updates.filter(value => value.state === "confirmed").length, 2);
  assert.equal(waits.length, 0);
});
test("pending and transient verification ignore service delay while terminal errors stop affected devices", async () => {
  const first = pollingTarget(), second = { ...first, id: first.id + "-second" };
  const attempt = { operationId: "owned", modelId: document.id, targetIds: [first.id, second.id], targets: [first, second] };
  const receipts: AssignmentReceipt[] = attempt.targets.map(value => ({ operationId: "owned", modelId: document.id,
    targetId: value.id, name: value.name, kind: "device", state: "accepted", readbackDeadline: Date.now() + 600000 }));
  const waits: number[] = [], updates: AssignmentReceipt[] = [], calls: string[] = [];
  let round = 0;
  await pollAssignmentReadback(attempt, receipts, {
    current: () => true, random: () => 0.5, wait: async delay => { waits.push(delay); round++; },
    read: async value => {
      calls.push(value.id);
      if (value.id === second.id) throw new AssignmentApiError("permission", "Current access denied.");
      if (round === 1) throw new AssignmentApiError("request", "Throttled.", true, 429, undefined, 7000);
      return { target: { ...value, reference: round >= 3 ? { kind: "assigned", ids: [document.id] } : value.reference },
        confirmed: round >= 3 };
    }, publish: value => updates.push(value),
  }, signal());
  assert.ok(waits.every(delay => delay >= 2900 && delay <= 3000)); assert.deepEqual(calls, [first.id, second.id, first.id, first.id]);
  assert.ok(updates.some(value => value.targetId === first.id && value.state === "unknown" && value.message?.includes("no assignment")));
  assert.ok(updates.some(value => value.targetId === second.id && value.state === "failed" && !value.readbackFailure?.retryable));
  assert.equal(updates.filter(value => value.state === "confirmed").length, 1);
});
test("aborted or stale result polling cannot publish, and old-operation receipts cannot start reads", async () => {
  const selected = pollingTarget(), controller = new AbortController();
  const attempt = { operationId: "current", modelId: document.id, targetIds: [selected.id], targets: [selected] };
  const receipt: AssignmentReceipt = { operationId: "current", modelId: document.id, targetId: selected.id, name: selected.name,
    kind: "device", state: "accepted", readbackDeadline: Date.now() + 600000 };
  let reads = 0, publications = 0;
  await pollAssignmentReadback(attempt, [{ ...receipt, operationId: "older" }], {
    read: async () => { reads++; throw new Error("Must not read."); }, current: () => true, publish: () => publications++,
  }, signal());
  assert.equal(reads, 0);
  let release: ((value: { target: AssignmentTarget; confirmed: boolean }) => void) | undefined;
  const running = pollAssignmentReadback(attempt, [receipt], {
    wait: async () => undefined, current: () => true,
    read: () => { reads++; return new Promise(resolve => { release = resolve; }); },
    publish: () => publications++,
  }, controller.signal);
  for (let index = 0; index < 10; index++) await Promise.resolve();
  const publicationsBeforeAbort = publications;
  controller.abort();
  await assert.rejects(running, { name: "AbortError" });
  release!({ target: { ...selected, reference: { kind: "assigned", ids: [document.id] } }, confirmed: true });
  await Promise.resolve(); assert.equal(publications, publicationsBeforeAbort);
  await assert.rejects(() => pollAssignmentReadback(attempt, [receipt], {
    wait: async () => undefined, current: () => false,
    read: async () => { reads++; throw new Error("Must not read stale owner."); }, publish: () => publications++,
  }, signal()), { name: "AbortError" });
  assert.equal(reads, 1);
});
test("interrupted untracked dispatch stops unknown while a captured operation resumes its original scheduled wait", async () => {
  const selected = pollingTarget(), controller = new AbortController(), updates: AssignmentReceipt[] = [];
  let release: (() => void) | undefined;
  const running = new ModelAssignmentOperation().run([selected], model(), true, {
    read: async value => value,
    write: async (_target, _id, _signal, dispatch) => {
      dispatch(); await new Promise<void>(resolve => { release = resolve; }); return "accepted";
    },
  }, () => true, controller.signal, value => updates.push(value), "explicit-operation");
  for (let index = 0; index < 20; index++) await Promise.resolve();
  assert.equal(updates.at(-1)?.state, "assigning"); assert.equal(updates.at(-1)?.writeDispatched, true);
  controller.abort();
  const result = await running;
  assert.equal(result[0].state, "unknown"); assert.equal(result[0].writeDispatched, true);
  assert.deepEqual(result[0].readbackFailure, { kind: "request", retryable: false });
  const beforeRelease = updates.length;
  release!();
  for (let index = 0; index < 10; index++) await Promise.resolve();
  assert.equal(updates.length, beforeRelease, "A late mutation response cannot publish or resurrect tracking.");
  const delays: number[] = [];
  const attempt = { operationId: "explicit-operation", modelId: document.id, targetIds: [selected.id], targets: [selected] };
  await pollAssignmentReadback(attempt, result, {
    current: () => true, wait: async () => { throw new Error("No untracked operation can resume."); },
    read: async () => { throw new Error("No untracked operation can be read."); },
    publish: () => { throw new Error("Terminal untracked results cannot change."); },
  }, signal());
  const fixture = fixtureClient({ accepted: true });
  const current = (await fixture.api.discover(source(), signal())).targets[0];
  const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async (url, init) =>
    String(url).includes("/asyncOperationStatuses/") ? response({ status: "Succeeded" }) : fixture.fetcher(url, init) });
  const accepted = await new ModelAssignmentOperation().run([current], model(), true,
    { ...api.assignmentServices(source()), observe: undefined }, () => true, signal(), () => undefined, attempt.operationId);
  assert.equal(accepted[0].state, "accepted"); assert.ok(accepted[0].trackingId);
  await pollAssignmentReadback({ ...attempt, targets: [current] }, [{ ...accepted[0], readbackNotBefore: Date.now() + 10000 }], {
    current: () => true, wait: async delay => { delays.push(delay); },
    read: (value, id, abort, trackingId) => api.readAssignment(value, id, abort, trackingId),
    publish: value => updates.push(value),
  }, signal());
  assert.ok(delays[0] > 9000 && delays[0] <= 10000);
  assert.equal(updates.at(-1)?.state, "confirmed");
});
test("assignment success requires every selected device receipt from the current explicit attempt", () => {
  const ids = [deviceId, deviceId + "-second"];
  const intent = { owner: "current", mode: "assign" as const,
    attempt: { operationId: "current-operation", modelId: document.id, targetIds: ids } };
  const history: AssignmentReceipt[] = ids.map(targetId => ({ operationId: "current-operation", targetId,
    name: targetId, modelId: document.id, kind: "device", state: "confirmed" }));
  assert.equal(confirmedAssignmentCount(intent, history, document.id), 2);
  for (const state of ["accepted", "checking", "assigning", "unknown", "failed", "blocked", "cancelled"] as const) {
    assert.equal(confirmedAssignmentCount(intent, [history[0], { ...history[1], state }], document.id), 0);
  }
  assert.equal(confirmedAssignmentCount(intent, history.map(receipt => ({ ...receipt, operationId: "older-operation" })), document.id), 0);
  assert.equal(confirmedAssignmentCount(intent, history.slice(0, 1), document.id), 0);
  assert.equal(confirmedAssignmentCount(intent, [...history, { ...history[0], state: "failed" }], document.id), 0);
  assert.equal(confirmedAssignmentCount(intent, history, "another-model"), 0);
  assert.equal(confirmedAssignmentCount(undefined, history, document.id), 0);
  assert.equal(confirmedAssignmentCount({ ...intent, attempt: { ...intent.attempt, targetIds: [] } }, history, document.id), 0);
  assert.equal(confirmedAssignmentCount({ ...intent, attempt: { ...intent.attempt, targetIds: [deviceId, deviceId.toUpperCase()] } }, history, document.id), 0);
});
test("already-assigned inventory requires nonempty complete current references regardless of model ID", () => {
  const inventory: AssignmentInventory = { hubs: [{ id: hubId, name: "Hub", namespaceId: nsId, state: "linked" }],
    targets: [{ ...target(), reference: { kind: "assigned", ids: [document.id] } }], issues: [], hasMore: false };
  assert.equal(inventoryHasModelAssignments(inventory), true);
  const other: AssignmentTarget = { ...inventory.targets[0], id: deviceId + "-other", updatePermission: "unknown",
    reference: { kind: "assigned", ids: ["urn:another-model"] } };
  assert.equal(inventoryHasModelAssignments({ ...inventory, targets: [other] }), true);
  assert.equal(inventoryHasModelAssignments({ ...inventory, targets: [...inventory.targets, other] }), true);
  assert.equal(assignmentBlocker(other), "Existing model assignments are read-only.");
  assert.equal(inventoryHasModelAssignments({ ...inventory, hubs: [...inventory.hubs,
    { id: hubId + "-unlinked", name: "Known nonlinked Hub", state: "unlinked" }] }), true);
  for (const value of [
    { ...inventory, hasMore: true },
    { ...inventory, targets: [] },
    { ...inventory, issues: [{ resource: hubId, message: "Incomplete read.", retryable: true }] },
    { ...inventory, hubs: [] },
    { ...inventory, hubs: [{ ...inventory.hubs[0], state: "unlinked" as const }] },
    { ...inventory, hubs: [...inventory.hubs, { id: hubId + "-pending", name: "Pending Hub", state: "pending" as const }] },
    { ...inventory, hubs: [...inventory.hubs, { id: hubId + "-unverified", name: "Unknown Hub", state: "unverified" as const }] },
    { ...inventory, targets: [{ ...inventory.targets[0], hubIds: [hubId + "-unknown"] }] },
    { ...inventory, targets: [{ ...target(), reference: { kind: "unknown" as const, reason: "Unknown" } }] },
    { ...inventory, targets: [target()] },
    { ...inventory, targets: [{ ...target(), reference: { kind: "assigned" as const, ids: [""] as [string] } }] },
    { ...inventory, targets: [{ ...target(), reference: { kind: "assigned" as const, ids: ["x".repeat(513)] as [string] } }] },
    { ...inventory, targets: [{ ...inventory.targets[0], referenceEvidence: { source: "adr-point-get" as const, observedAt: "fixture",
      apiVersion: "2026-11-02-preview" as const, rawCarrier: "present" as const } }] },
  ]) assert.equal(inventoryHasModelAssignments(value), false);
});
test("reference omission/null/malformed are unknown, only explicit empty is unassigned", () => {
  for (const value of [{}, { deviceTypeRefs: null }, { deviceTypeRefs: "" }, { deviceTypeRefs: ["a", "b"] },
    { deviceTypeRefs: [""] }, { deviceTypeRefs: ["x".repeat(513)] }]) assert.equal(readReference(value).kind, "unknown");
  assert.deepEqual(readReference({ deviceTypeRefs: [] }), { kind: "unassigned", ids: [] });
  assert.deepEqual(readReference({ deviceTypeRefs: [document.id] }), { kind: "assigned", ids: [document.id] });
});
test("common model boundary rejects retained old models and assignment limits are not global TM limits", async () => {
  const stale = model(); const { id: _id, ...old } = document; stale.document = old; stale.id = undefined;
  assert.ok(modelValidation(stale).length);
  assert.throws(() => assignmentOwner(source(), 1, stale), /doesn't have an ID/);
  let reads = 0;
  await assert.rejects(() => inspectStreamCoverage({ source: { queryServiceUri: "", databaseName: "", telemetryTable: "", twinTable: "" },
    model: stale, fields: [], limit: 100, scope: { mode: "data", operator: "and", telemetry: "", twin: "" } }, "fixture-token",
    "telemetry", undefined, async () => { reads++; throw new Error("Must not read."); }), /doesn't have an ID/);
  assert.equal(reads, 0);
  const long = fromThingModel(resolveThingModel({ ...document, id: `urn:fixture:${"x".repeat(520)}` }));
  assert.equal(modelValidation(long).length, 0);
  assert.throws(() => assignmentReference(long), /512/);
  assert.equal(requireThingModelId(document), document.id);
  assert.notEqual(assignmentOwner(source(), 1, model()), assignmentOwner(source(), 2, model()));
});
function legacyEnrollment(kind: "group" | "individual"): AssignmentTarget {
  return { ...target(), id: `${dpsId}/${kind === "group" ? "enrollmentGroups" : "enrollments"}/stale-selection`,
    name: "Stale enrollment selection", kind };
}
test("retained enrollment selections are blocked at the operation boundary before any stale read/write adapter or token callback", async () => {
  for (const kind of ["group", "individual"] as const) {
    const legacy = legacyEnrollment(kind);
    let reads = 0, writes = 0;
    assert.equal(assignmentBlocker(legacy), DPS_INFORMATIONAL_ONLY);
    const receipts = await new ModelAssignmentOperation().run([legacy], model(), true, {
      read: async target => { reads++; return target; },
      write: async (_target, _id, _signal, dispatch) => { writes++; dispatch(); return "completed"; },
    }, () => true, signal(), () => undefined);
    assert.equal(reads, 0); assert.equal(writes, 0);
    assert.equal(receipts[0].state, "blocked"); assert.equal(receipts[0].message, DPS_INFORMATIONAL_ONLY);
  }
  const forged = { ...legacyEnrollment("group"), kind: "device" as const };
  let callbacks = 0;
  const result = await new ModelAssignmentOperation().run([forged], model(), true, {
    read: async target => { callbacks++; return target; },
    write: async () => { callbacks++; return "completed"; },
  }, () => true, signal(), () => undefined);
  assert.equal(callbacks, 0); assert.equal(result[0].state, "blocked");
  assert.match(result[0].message!, /ADR registry-device/);
});
test("ADR API services and reader reject enrollment kinds before network/auth and expose no retired DPS data-plane methods", async () => {
  let tokens = 0, requests = 0, acknowledgments = 0;
  const api = new HubAssignmentApi({ token: async () => { tokens++; return "synthetic"; },
    fetcher: async () => { requests++; throw new Error("No request should be made."); } });
  const services = api.assignmentServices(source()), reader = new HubAssignmentReader(api, source(), () => {
    throw new Error("Informational enrollment interaction must not publish inventory.");
  }, () => true);
  for (const kind of ["group", "individual"] as const) {
    const legacy = legacyEnrollment(kind);
    await assert.rejects(() => services.read(legacy, signal()), /cannot be checked/);
    await assert.rejects(() => services.write(legacy, document.id, signal(), () => acknowledgments++), /not supported/);
    await assert.rejects(() => reader.load(kind, "view", signal()), /outside this workload/);
    await assert.rejects(() => reader.load(kind, "refresh", signal()), /outside this workload/);
  }
  for (const method of ["readKnownGroup", "readKnownIndividual", "readEnrollmentPage", "loadEnrollments",
    "readEnrollmentForAssignment", "putEnrollmentReference"]) assert.equal(Reflect.get(api, method), undefined);
  assert.equal(tokens, 0); assert.equal(requests, 0); assert.equal(acknowledgments, 0); reader.dispose();
});
test("source-owned initial load, explicit frontend recovery, refresh and continuation use only ARM and never discover DPS", async () => {
  const fixture = fixtureClient(), scopes: string[] = [];
  let authorized = false;
  const api = new HubAssignmentApi({ token: async scope => {
    scopes.push(scope); assert.equal(scope, ASSIGNMENT_ARM_SCOPE);
    if (!authorized) throw assignmentTokenFailure({ error: 1 });
    return "synthetic";
  }, fetcher: fixture.fetcher });
  const reader = new HubAssignmentReader(api, source(), () => undefined, () => true);
  await assert.rejects(() => reader.load("device", "view", signal()), /authorization/);
  const attempted = scopes.length;
  await assert.rejects(() => reader.load("device", "view", signal()));
  assert.equal(scopes.length, attempted);
  authorized = true;
  const recovered = await reader.load("device", "refresh", signal());
  assert.equal(recovered.targets.length, 1);
  await reader.load("device", "more", signal());
  await reader.load("device", "refresh", signal());
  assert.ok(scopes.length > attempted && scopes.every(scope => scope === ASSIGNMENT_ARM_SCOPE));
  assert.ok(fixture.calls.every(call => new URL(call.url).origin === "https://management.azure.com" &&
    !/provisioningServices|enrollment|azure-devices-provisioning/i.test(call.url + String(call.init?.body || ""))));
  reader.dispose();
});
test("device inventory discards retained enrollment rows/auth issues and cannot reintroduce them through readback", async () => {
  const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async () => { throw new Error("No network."); } });
  const legacy = legacyEnrollment("group"), publications: AssignmentInventory[] = [];
  api.discover = async () => ({ hubs: [], targets: [target(), legacy], hasMore: false, issues: [
    { resource: dpsId, scope: "enrollments", authScope: "dps", kind: "authentication", retryable: true, message: "Old DPS issue." },
  ] });
  const reader = new HubAssignmentReader(api, source(), value => publications.push(value), () => true);
  const current = await reader.load("device", "view", signal());
  assert.deepEqual(current.targets, [target()]); assert.deepEqual(current.issues, []);
  const before = publications.length; reader.updateTarget(legacy);
  assert.equal(publications.length, before); reader.dispose();
});
test("a retained current-looking coverage proof cannot bypass the required ID at plan or Create", async () => {
  const request = creationRequest();
  const { id: _id, ...old } = document;
  request.input.model = { ...request.input.model, id: undefined, document: old };
  request.coverage.fingerprint = modelingFingerprint(request.input);
  assert.equal(coverageReady(request.input, request.coverage, false), false);
  assert.throws(() => buildModeledDataPlan(request.input, request.coverage, false, request.names), /doesn't have an ID/);
  let calls = 0;
  const result = await new ModeledDataCreator().run(request, {
    inspect: async () => { calls++; return request.coverage; },
    request: async () => { calls++; throw new Error("No native operation permitted."); },
  }, true);
  assert.equal(result.phase, "blocked"); assert.match(result.error!, /doesn't have an ID/); assert.equal(calls, 0);
});

function fixtureClient(overrides: { omit?: boolean; capability?: string; failDevice?: boolean; denyWrite?: boolean;
  hubProperties?: Record<string, unknown>; hubStatus?: number;
  malformedPermissions?: boolean; accepted?: boolean; capabilityType?: string; omitCapabilityType?: boolean; mixedCapabilities?: boolean;
  failNamespace?: boolean; collectionStatus?: number; malformedCollection?: boolean;
  indexedRows?: unknown[]; indexedStatus?: number; currentProperties?: Record<string, unknown>;
  completePoint?: boolean; pointStatus?: number; pointType?: unknown; pointEtag?: unknown; pointId?: string } = {}) {
  const calls: { url: string; init?: RequestInit }[] = [];
  const completePoint = overrides.completePoint !== false;
  const fetcher: typeof fetch = async (input, init) => {
    const url = String(input); calls.push({ url, init });
    if (url.includes("ResourceGraph")) {
      const query = JSON.parse(String(init?.body)).query;
      if (query.includes("registrydevices")) return response({ data: overrides.indexedRows || [{ id: deviceId,
        propertyFieldNames: overrides.omit ? ["externalDeviceId"] : ["deviceTypeRefs"],
        reference: overrides.omit ? null : [] }], resultTruncated: false }, overrides.indexedStatus || 200);
      assert.ok(query.includes("microsoft.devices/iothubs"), "ADR assignment must not discover provisioning services.");
      return response({ data: [{ id: hubId, name: "fixture-hub", hostName: "fixture-hub.azure-devices.net" }], resultTruncated: "false" });
    }
    const path = new URL(url).pathname;
    if (path === hubId) return response({ id: hubId, name: "fixture-hub", properties: overrides.hubProperties || {
      hostName: "fixture-hub.azure-devices.net",
      deviceRegistry: { namespaceResourceId: nsId, linkingProperties: { state: "Succeeded" } } } }, overrides.hubStatus || 200);
    if (path === nsId) return overrides.failNamespace ? response({}, 403) : response({ id: nsId, properties: { messaging: { endpoints: {
      "custom-hub-key": { endpointType: "Microsoft.Devices/IotHubs", resourceId: hubId.toUpperCase(), address: "fixture-hub.service.azure-devices.net" },
    } } } });
    if (path === `${nsId}/registryDevices`) return overrides.collectionStatus ? response({}, overrides.collectionStatus)
      : overrides.malformedCollection ? response({ value: "Invalid fixture page" })
        : response({ value: [{ id: deviceId, name: "fixture-device" }] });
    if (path === `${deviceId}/capabilities`) return response({ value: overrides.mixedCapabilities ? [
      { id: `${deviceId}/capabilities/hub-wrong-endpoint`, properties: { capabilityType: "Microsoft.IoTHub", messagingEndpointName: "wrong-endpoint" } },
      { id: `${deviceId}/capabilities/other-right-endpoint`, properties: { capabilityType: "Microsoft.SoftwareUpdate", messagingEndpointName: "custom-hub-key" } },
    ] : [{ id: `${deviceId}/capabilities/custom-capability`,
      properties: { ...(!overrides.omitCapabilityType ? { capabilityType: overrides.capabilityType || "Microsoft.IoTHub" } : {}),
        messagingEndpointName: overrides.capability || "custom-hub-key" } }] });
    if (path === `${deviceId}/providers/Microsoft.Authorization/permissions`) return response({
      value: overrides.malformedPermissions ? [{ actions: "*" }] : [{ actions: overrides.denyWrite ? ["*/read"] : ["*"], notActions: [] }],
    });
    if (path === deviceId && init?.method === "PATCH") return new Response("{}", { status: overrides.accepted ? 202 : 200,
      headers: overrides.accepted ? { Location: `https://management.azure.com/subscriptions/${deviceId.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/fixture-operation?api-version=2026-11-02-preview` } : {} });
    if (path === deviceId) return overrides.failDevice ? response({}, 403) : response({
      id: overrides.pointId || deviceId, name: "fixture-device",
      ...((completePoint || "pointType" in overrides)
        ? { type: "pointType" in overrides ? overrides.pointType : "Microsoft.DeviceRegistry/namespaces/registryDevices" } : {}),
      etag: "pointEtag" in overrides ? overrides.pointEtag : '"fixture-etag-1"',
      properties: { ...(completePoint ? {
        uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: "fixture-device", provisioningState: "Succeeded",
      } : {}), ...(overrides.currentProperties || (overrides.omit ? {} : { deviceTypeRefs: [] })) },
    }, overrides.pointStatus || 200);
    throw new Error(`Unexpected fixture request: ${url}`);
  };
  return { calls, fetcher, api: new HubAssignmentApi({ token: async scope => { assert.equal(scope, ASSIGNMENT_ARM_SCOPE); return "synthetic-token"; }, fetcher }) };
}
test("canonical current references ignore stale unavailable device ARG and issue no per-device index queries", async () => {
  for (const [currentRefs, indexedRefs, expected] of [
    [[], [document.id], "unassigned"],
    [[document.id], [], "assigned"],
    [["urn:fixture:other"], [document.id], "assigned"],
    [undefined, ["urn:fixture:stale"], "unassigned"],
  ] as const) {
    const currentProperties = currentRefs === undefined ? {} : { deviceTypeRefs: currentRefs };
    const fixture = fixtureClient({ completePoint: true, currentProperties,
      indexedRows: [{ id: deviceId, propertyFieldNames: ["deviceTypeRefs"], reference: indexedRefs }] });
    const inventory = await fixture.api.discover(source(), signal()), value = inventory.targets[0];
    assert.equal(value.reference.kind, expected);
    assert.equal(value.referenceEvidence?.source, "adr-point-get");
    const deviceQueries = fixture.calls.filter(call => call.url.includes("ResourceGraph") &&
      String(call.init?.body).includes("registrydevices"));
    assert.equal(deviceQueries.length, 0);
    assert.equal(fixture.calls.filter(call => call.url.includes("ResourceGraph") &&
      String(call.init?.body).includes("microsoft.devices/iothubs")).length, 1);
    if (expected === "assigned") {
      await assert.rejects(() => fixture.api.patchReference(value, document.id, signal(), source()), /read-only/);
      assert.equal(fixture.calls.some(call => call.init?.method === "PATCH"), false);
    }
  }
  const unavailable = fixtureClient({ completePoint: true, currentProperties: {}, indexedStatus: 503 });
  const inventory = await unavailable.api.discover(source(), signal());
  assert.equal(assignmentBlocker(inventory.targets[0]), undefined);
  assert.equal(inventory.issues.length, 0);
  assert.equal(unavailable.calls.some(call => String(call.init?.body).includes("registrydevices")), false);
});
test("one device discovery and guarded assignment perform zero device ARG queries while preserving all current reads", async () => {
  const fixture = fixtureClient({ completePoint: true, currentProperties: {}, indexedStatus: 503 });
  const inventory = await fixture.api.discover(source(), signal()), selected = inventory.targets[0];
  assert.equal(assignmentBlocker(selected), undefined);
  assert.equal(fixture.calls.length, 7);
  const discoveryCalls = fixture.calls.length;
  const receipts = await new ModelAssignmentOperation().run([selected], model(), true, fixture.api.assignmentServices(source()),
    () => true, signal(), () => undefined);
  assert.equal(receipts[0].state, "confirmed");
  assert.equal(fixture.calls.length - discoveryCalls, 11);
  const counts = {
    hostnameHubArg: fixture.calls.filter(call => call.url.includes("ResourceGraph") &&
      String(call.init?.body).includes("microsoft.devices/iothubs")).length,
    deviceArg: fixture.calls.filter(call => call.url.includes("ResourceGraph") && String(call.init?.body).includes("registrydevices")).length,
    pointGets: fixture.calls.filter(call => new URL(call.url).pathname === deviceId && (!call.init?.method || call.init.method === "GET")).length,
    capabilities: fixture.calls.filter(call => new URL(call.url).pathname === `${deviceId}/capabilities`).length,
    permissions: fixture.calls.filter(call => call.url.includes("Microsoft.Authorization/permissions")).length,
    conditionalPatch: fixture.calls.filter(call => call.init?.method === "PATCH").length,
  };
  assert.deepEqual(counts, { hostnameHubArg: 1, deviceArg: 0, pointGets: 3, capabilities: 3, permissions: 3, conditionalPatch: 1 });
  assert.equal(new Headers(fixture.calls.find(call => call.init?.method === "PATCH")!.init?.headers).get("If-Match"), selected.etag);
});
test("expired dispatched receipts retain customer timeout guidance without resuming operation reads", async () => {
  const selected = pollingTarget(), updates: AssignmentReceipt[] = [];
  const attempt = { operationId: "elapsed", modelId: document.id, targetIds: [selected.id], targets: [selected] };
  const customerMessage = "The assignment wasn't confirmed within 30 seconds. It may still complete in Azure, so check the device later. No assignment will be retried.";
  let requests = 0;
  await pollAssignmentReadback(attempt, [{ operationId: attempt.operationId, modelId: document.id,
    targetId: selected.id, name: selected.name, kind: "device", state: "unknown", writeDispatched: true,
    trackingId: "original-native-operation", readbackDeadline: Date.now() - 1, message: "Earlier timeout guidance." }], {
    current: () => true, read: async () => { requests++; return { confirmed: true }; }, publish: receipt => updates.push(receipt),
  }, signal());
  assert.equal(requests, 0);
  assert.equal(updates.at(-1)?.state, "unknown");
  assert.equal(updates.at(-1)?.message, customerMessage);
  assert.equal(updates.at(-1)?.readbackFailure?.retryable, false);
  await pollAssignmentReadback(attempt, [updates.at(-1)!], {
    current: () => true, read: async () => { requests++; return { confirmed: true }; },
    publish: () => { throw new Error("A settled timeout must not be rewritten again on remount."); },
  }, signal());
  assert.equal(requests, 0);
  const legacy: AssignmentReceipt = { ...updates.at(-1)!,
    message: "The assignment operation tracking budget is missing or expired. Completion is unknown; no assignment will be retried." };
  await pollAssignmentReadback(attempt, [legacy], {
    current: () => true, read: async () => { requests++; return { confirmed: true }; }, publish: receipt => updates.push(receipt),
  }, signal());
  assert.equal(updates.at(-1)?.message, customerMessage);
  assert.equal(requests, 0);
});
test("missing operation timing is tracking uncertainty rather than a claimed elapsed observation timeout", async () => {
  const selected = pollingTarget(), updates: AssignmentReceipt[] = [];
  let requests = 0;
  await pollAssignmentReadback({ operationId: "missing", modelId: document.id, targetIds: [selected.id], targets: [selected] },
    [{ operationId: "missing", modelId: document.id, targetId: selected.id, name: selected.name, kind: "device",
      state: "accepted", writeDispatched: true }], { current: () => true,
      read: async () => { requests++; return { confirmed: true }; }, publish: receipt => updates.push(receipt) }, signal());
  assert.equal(requests, 0); assert.equal(updates.at(-1)?.state, "unknown");
  assert.match(updates.at(-1)!.message!, /tracking information is unavailable/);
  assert.doesNotMatch(updates.at(-1)!.message!, /30 seconds|tracking budget|timed out/);
  assert.equal(updates.at(-1)?.readbackFailure?.kind, "shape");
  assert.equal(updates.at(-1)?.readbackFailure?.retryable, false);
});
test("Hub-only metadata classifies omitted null pending failed and linked states without reading ADR devices", async () => {
  const hostname = { hostName: "fixture-hub.azure-devices.net" };
  for (const [properties, state] of [
    [hostname, "unlinked"], [{ ...hostname, deviceRegistry: null }, "unlinked"],
    [{ ...hostname, deviceRegistry: { namespaceResourceId: nsId, linkingProperties: { state: "Pending" } } }, "pending"],
    [{ ...hostname, deviceRegistry: { namespaceResourceId: nsId, linkingProperties: { state: "Failed" } } }, "unlinked"],
    [{ ...hostname, deviceRegistry: { namespaceResourceId: nsId, linkingProperties: { state: "Succeeded" } } }, "linked"],
  ] as const) {
    const fixture = fixtureClient({ hubProperties: properties });
    const summary = await fixture.api.discoverHubs(source(), signal());
    assert.equal(summary.complete, true); assert.equal(summary.hubs[0].state, state); assert.deepEqual(summary.issues, []);
    assert.equal(fixture.calls.length, 2);
    assert.ok(fixture.calls.every(call => !/registryDevices|capabilities|Authorization\/permissions/.test(call.url)));
    assert.ok(fixture.calls[1].url.includes("api-version=2026-10-01-preview"));
    const inventory = await fixture.api.discover(source(), signal(), undefined, { initial: summary });
    assert.equal(fixture.calls.filter(call => new URL(call.url).pathname === hubId).length, 1);
    assert.equal(inventory.targets.length, state === "linked" ? 1 : 0);
    if (state !== "linked") assert.equal(fixture.calls.length, 2);
  }
});
test("Hub-only failures and incomplete source metadata are not successful all-nonlinked proof", async () => {
  for (const options of [{ hubStatus: 401 }, { hubStatus: 403 }, { hubProperties: { deviceRegistry: [] } },
    { hubProperties: { hostName: "wrong.example", deviceRegistry: null } },
    { hubProperties: { hostName: "fixture-hub.azure-devices.net",
      deviceRegistry: { namespaceResourceId: nsId, linkingProperties: { state: "Unexpected" } } } }]) {
    const summary = await fixtureClient(options).api.discoverHubs(source(), signal());
    assert.equal(summary.complete, false); assert.ok(summary.issues.length);
    assert.equal(allAssignmentHubsUnavailable(summary), false);
  }
  const fixture = fixtureClient();
  const armSource = { ...source(), telemetry: { ...source().telemetry, docstring: `@iot.meta source=${hubId}` },
    twin: { ...source().twin, docstring: `@iot.meta source=${hubId}` } };
  const partialHub = await fixtureClient({ hubProperties: { deviceRegistry: null } }).api.discoverHubs(armSource, signal());
  assert.equal(partialHub.complete, false);
  assert.equal(allAssignmentHubsUnavailable(partialHub), false);
  const noMetadata = { ...source(), telemetry: { ...source().telemetry, docstring: "" }, twin: { ...source().twin, docstring: "" } };
  const summary = await fixture.api.discoverHubs(noMetadata, signal());
  assert.equal(summary.complete, false); assert.equal(summary.hubs.length, 0); assert.equal(fixture.calls.length, 0);
  const truncated = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async () => response({ data: [], resultTruncated: true }) });
  await assert.rejects(() => truncated.discoverHubs(source(), signal()), /incomplete/);
});
test("discovery follows ARM endpoint identity and current canonical reference proof without secret or enrollment reads", async () => {
  const { api, calls } = fixtureClient();
  const result = await api.discover(source(), signal());
  assert.equal(result.targets.length, 1); assert.deepEqual(result.issues, []);
  assert.deepEqual(result.targets[0].hubIds, [hubId]);
  assert.equal(result.targets[0].reference.kind, "unassigned");
  assert.equal(result.targets[0].referenceEvidence?.source, "adr-point-get");
  assert.equal(result.targets[0].referenceEvidence?.currentArmEtag, result.targets[0].etag);
  assert.equal(result.targets[0].updatePermission, "granted"); assert.equal(result.targets[0].writeBlocker, undefined);
  assert.equal((await api.recheck(result.targets[0], source(), signal())).reference.kind, "unassigned");
  assert.ok(calls.every(call => !/listKeys|authenticationProfiles|enrollments|enrollmentGroups|attestation|provisioningServices|azure-devices-provisioning/i.test(call.url)));
  assert.ok(calls.filter(call => new URL(call.url).pathname === deviceId).every(call => call.url.includes("2026-11-02-preview")));
  const unknown = await fixtureClient({ omit: true, completePoint: false }).api.discover(source(), signal());
  assert.equal(unknown.targets[0].reference.kind, "unknown");
  assert.equal((await fixtureClient({ capability: "unrelated-hub-key" }).api.discover(source(), signal())).targets.length, 0);
  const partial = await fixtureClient({ failDevice: true }).api.discover(source(), signal());
  assert.equal(partial.targets.length, 0); assert.equal(partial.issues[0].retryable, false);
  assert.match(partial.issues[0].message, /permission/);
});
test("canonical GET preserves explicit null empty same and other reference states without querying the device index", async () => {
  for (const reference of [[], [document.id], ["urn:fixture:other"], null, ["one", "two"], "unsupported"]) {
    const fixture = fixtureClient({ currentProperties: { deviceTypeRefs: reference },
      indexedRows: [{ id: deviceId, propertyFieldNames: ["deviceTypeRefs"], reference: ["urn:stale"] }] });
    const inventory = await fixture.api.discover(source(), signal()), value = inventory.targets[0];
    assert.deepEqual(value.reference, readReference({ deviceTypeRefs: reference }));
    assert.equal(value.referenceEvidence?.source, "adr-point-get");
    assert.equal(value.referenceEvidence?.rawCarrier, "present");
    assert.ok(Number.isFinite(Date.parse(value.referenceEvidence!.observedAt)));
    assert.equal(!assignmentBlocker(value), value.reference.kind === "unassigned");
    assert.equal(fixture.calls.some(call => call.url.includes("Microsoft.Authorization/permissions")), value.reference.kind === "unassigned");
    assert.equal(fixture.calls.some(call => call.url.includes("ResourceGraph") && String(call.init?.body).includes("registrydevices")), false);
    const point = fixture.calls.filter(call => new URL(call.url).pathname === deviceId);
    assert.equal(point.length, 1);
    assert.equal(new URL(point[0].url).search, "?api-version=2026-11-02-preview");
  }
  for (const indexedRows of [[], [{ id: deviceId, propertyFieldNames: ["externalDeviceId"], reference: null }],
    [{ id: deviceId, propertyFieldNames: ["externalDeviceId"], reference: [document.id] }]]) {
    const value = (await fixtureClient({ indexedRows, completePoint: false }).api.discover(source(), signal())).targets[0];
    assert.equal(value.reference.kind, "unknown");
  }
});
test("stale indexed empty cannot overwrite current assigned malformed or incomplete canonical references", async () => {
  for (const options of [{ currentProperties: {}, completePoint: false }, { currentProperties: { deviceTypeRefs: null } },
    { currentProperties: { deviceTypeRefs: [document.id] } }, { currentProperties: { deviceTypeRefs: ["urn:fixture:other"] } }]) {
    const { api, calls } = fixtureClient(options);
    const value = (await api.discover(source(), signal())).targets[0];
    assert.ok(["assigned", "unknown"].includes(value.reference.kind));
    assert.equal(value.etag, '"fixture-etag-1"');
    assert.equal(referenceIsCurrent(value), value.reference.kind === "assigned");
    assert.equal(value.referenceEvidence?.currentArmEtag, value.reference.kind === "assigned" ? value.etag : undefined);
    assert.equal(value.updatePermission, "unknown");
    await assert.rejects(() => api.patchReference(value, document.id, signal(), source()));
    const forged = { ...value, writeBlocker: undefined };
    await assert.rejects(() => api.patchReference(forged, document.id, signal(), source()));
    assert.ok(!calls.some(call => call.init?.method === "PATCH" || call.url.includes("Microsoft.Authorization/permissions")));
  }
});
test("only a verified full current ADR point GET interprets omitted optional refs as logical unassigned with raw missing provenance", async () => {
  const properties: Record<string, unknown> = {}, fixture = fixtureClient({ completePoint: true, omit: true, currentProperties: properties });
  const inventory = await fixture.api.discover(source(), signal()), value = inventory.targets[0];
  assert.deepEqual(value.reference, { kind: "unassigned", ids: [] });
  assert.deepEqual(value.referenceEvidence, {
    source: "adr-point-get", observedAt: value.referenceEvidence!.observedAt, currentArmEtag: '"fixture-etag-1"',
    apiVersion: "2026-11-02-preview", rawCarrier: "missing",
  });
  assert.deepEqual(value.deviceIdentity, { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: "fixture-device" });
  assert.equal(value.updatePermission, "granted"); assert.equal(assignmentBlocker(value), undefined);
  assert.equal(referenceIsCurrent(value), true);
  assert.deepEqual(properties, {}); assert.equal(Object.hasOwn(properties, "deviceTypeRefs"), false);
  assert.equal(readReference({}).kind, "unknown", "Generic/ARG missing-field classification must not change.");
  assert.ok(fixture.calls.filter(call => new URL(call.url).pathname === deviceId)
    .every(call => new URL(call.url).search === "?api-version=2026-11-02-preview"));
  await fixture.api.patchReference(value, document.id, signal(), source());
  const writes = fixture.calls.filter(call => call.init?.method === "PATCH");
  assert.equal(writes.length, 1); assert.equal(new Headers(writes[0].init?.headers).get("If-Match"), value.etag);
  assert.deepEqual(JSON.parse(String(writes[0].init?.body)), { properties: { deviceTypeRefs: [document.id] } });
});
test("canonical reference authority ignores absent stale and malformed indexed carriers", async () => {
  for (const indexedRows of [[], [{ id: deviceId, propertyFieldNames: [], reference: null }]]) {
    for (const currentProperties of [{}, { deviceTypeRefs: [] }, { deviceTypeRefs: [document.id] },
      { deviceTypeRefs: ["urn:fixture:other"] }, { deviceTypeRefs: null }, { deviceTypeRefs: "invalid" }, { deviceTypeRefs: ["one", "two"] }]) {
      const fixture = fixtureClient({ completePoint: true, currentProperties, indexedRows });
      const value = (await fixture.api.discover(source(), signal())).targets[0];
      const missing = !Object.hasOwn(currentProperties, "deviceTypeRefs");
      assert.deepEqual(value.reference, missing ? { kind: "unassigned", ids: [] } : readReference(currentProperties));
      assert.equal(value.referenceEvidence?.source, "adr-point-get");
      assert.equal(value.referenceEvidence?.rawCarrier, missing ? "missing" : "present");
      assert.equal(!assignmentBlocker(value), missing || Array.isArray(currentProperties.deviceTypeRefs) && !currentProperties.deviceTypeRefs.length);
      assert.equal(fixture.calls.some(call => String(call.init?.body).includes("registrydevices")), false);
    }
  }
  for (const reference of [null, "invalid", ["one", "two"]]) {
    const fixture = fixtureClient({ completePoint: true, currentProperties: {},
      indexedRows: [{ id: deviceId, propertyFieldNames: ["deviceTypeRefs"], reference }] });
    const value = (await fixture.api.discover(source(), signal())).targets[0];
    assert.equal(value.reference.kind, "unassigned"); assert.equal(assignmentBlocker(value), undefined);
    assert.ok(fixture.calls.some(call => call.url.includes("Microsoft.Authorization/permissions")));
    assert.equal(fixture.calls.some(call => String(call.init?.body).includes("registrydevices")), false);
  }
});
test("omission on incomplete, unsupported, failed-state or nonspecific provider reads is not converted to unassigned", async () => {
  for (const options of [
    { currentProperties: { uuid: undefined } }, { currentProperties: { uuid: null } },
    { currentProperties: { uuid: "invalid" } }, { currentProperties: { uuid: "00000000-0000-0000-0000-000000000000" } },
    { currentProperties: { externalDeviceId: undefined } }, { currentProperties: { externalDeviceId: "" } },
    { currentProperties: { externalDeviceId: 42 } }, { currentProperties: { provisioningState: "Pending" } },
    { currentProperties: { provisioningState: "Failed" } }, { currentProperties: { provisioningState: undefined } },
    { pointType: undefined }, { pointType: "Microsoft.DeviceRegistry/namespaces/devices" },
    { pointEtag: undefined }, { pointEtag: "*" }, { pointEtag: '"*"' }, { pointEtag: "bad\nvalue" }, { pointEtag: false },
  ]) {
    const fixture = fixtureClient({ completePoint: true, omit: true, ...options });
    const inventory = await fixture.api.discover(source(), signal()), value = inventory.targets[0];
    assert.equal(value.reference.kind, "unknown"); assert.ok(assignmentBlocker(value));
    assert.equal(value.referenceEvidence?.source, "adr-point-get");
    assert.equal(value.referenceEvidence?.currentArmEtag, undefined);
    assert.equal(value.updatePermission, "unknown");
    assert.ok(!fixture.calls.some(call => call.url.includes("Microsoft.Authorization/permissions") || call.init?.method === "PATCH"));
  }
  for (const options of [{ pointStatus: 206 }, { pointStatus: 400 }, { pointStatus: 500 }, { failDevice: true }, { pointId: `${deviceId}-other` }]) {
    const fixture = fixtureClient({ completePoint: true, omit: true, ...options });
    const inventory = await fixture.api.discover(source(), signal());
    assert.equal(inventory.targets.length, 0); assert.equal(inventory.issues[0].scope, "devices");
    assert.ok(!fixture.calls.some(call => call.url.includes("Microsoft.Authorization/permissions") || call.init?.method === "PATCH"));
  }
});
test("legitimate unset state still requires granted update permission and cannot use an ARG or fabricated provenance bypass", async () => {
  for (const options of [{ denyWrite: true }, { malformedPermissions: true }, { failNamespace: true }]) {
    const fixture = fixtureClient({ completePoint: true, omit: true, ...options });
    const inventory = await fixture.api.discover(source(), signal());
    if (inventory.targets.length) {
      assert.deepEqual(inventory.targets[0].reference, { kind: "unassigned", ids: [] });
      assert.ok(assignmentBlocker(inventory.targets[0]));
    } else assert.ok(inventory.issues.length);
    assert.ok(!fixture.calls.some(call => call.init?.method === "PATCH"));
  }
  const fixture = fixtureClient({ completePoint: true, omit: true });
  const selected = (await fixture.api.discover(source(), signal())).targets[0];
  for (const patch of [{ apiVersion: undefined }, { rawCarrier: undefined }, { currentArmEtag: '"different"' }]) {
    assert.equal(referenceIsCurrent({ ...selected, referenceEvidence: { ...selected.referenceEvidence!, ...patch } }), false);
  }
});
test("fresh unset assignment rechecks reference, ETag, UUID and external identity before the actual adapter can dispatch", async () => {
  for (const change of ["reference", "etag", "uuid", "external-id"] as const) {
    const currentProperties: Record<string, unknown> = {};
    const options = { completePoint: true, omit: true, currentProperties, pointEtag: '"fixture-etag-1"' };
    const fixture = fixtureClient(options), selected = (await fixture.api.discover(source(), signal())).targets[0];
    if (change === "reference") options.currentProperties.deviceTypeRefs = ["urn:fixture:other"];
    if (change === "etag") options.pointEtag = '"changed"';
    if (change === "uuid") options.currentProperties.uuid = "22222222-2222-4222-8222-222222222222";
    if (change === "external-id") options.currentProperties.externalDeviceId = "recreated-device";
    const result = await new ModelAssignmentOperation().run([selected], model(), true, fixture.api.assignmentServices(source()),
      () => true, signal(), () => undefined);
    assert.ok(["failed", "blocked"].includes(result[0].state));
    assert.ok(!fixture.calls.some(call => call.init?.method === "PATCH"));
    await assert.rejects(() => fixture.api.patchReference(selected, document.id, signal(), source()));
    assert.ok(!fixture.calls.some(call => call.init?.method === "PATCH"));
  }
});
test("authoritative not-set eligibility cannot publish across a pending permission read after owner retirement or cancellation", async () => {
  for (const abort of [false, true]) {
    const fixture = fixtureClient({ completePoint: true, omit: true }), entered = deferred<void>(), release = deferred<void>();
    const controller = new AbortController(), publications: AssignmentInventory[] = [];
    let current = true;
    const api = new HubAssignmentApi({ token: async () => "synthetic", isCurrent: () => current, fetcher: async (url, init) => {
      if (String(url).includes("Microsoft.Authorization/permissions")) { entered.resolve(); await release.promise; }
      return fixture.fetcher(url, init);
    } });
    const reader = new HubAssignmentReader(api, source(), value => publications.push(value), () => current);
    const pending = reader.load("device", "view", controller.signal);
    const rejected = assert.rejects(() => pending, { name: "AbortError" });
    await entered.promise;
    if (abort) controller.abort(); else current = false;
    reader.dispose(); release.resolve(); await rejected;
    assert.equal(publications.length, 0);
    assert.ok(!fixture.calls.some(call => call.init?.method === "PATCH"));
  }
});
test("unavailable malformed foreign and duplicate device indexes cannot override canonical references or introduce device queries", async () => {
  for (const options of [{ indexedStatus: 401 }, { indexedStatus: 403 }, { indexedStatus: 500 },
    { indexedRows: [{ id: deviceId, reference: [] }] },
    { indexedRows: [{ id: `${deviceId}-different`, propertyFieldNames: ["deviceTypeRefs"], reference: [] }] },
    { indexedRows: [1, 2].map(() => ({ id: deviceId, propertyFieldNames: ["deviceTypeRefs"], reference: [] })) }]) {
    const fixture = fixtureClient(options), inventory = await fixture.api.discover(source(), signal());
    assert.equal(inventory.targets.length, 1); assert.equal(inventory.issues.length, 0);
    assert.equal(assignmentBlocker(inventory.targets[0]), undefined);
    assert.equal(fixture.calls.some(call => String(call.init?.body).includes("registrydevices")), false);
    assert.deepEqual(assignmentIssuesForView(inventory.issues, "group"), []);
  }
});
test("canonical device-body owner changes and cancellation cannot become a projected row or swallowed discovery issue", async () => {
  for (const abort of [false, true]) {
    const fixture = fixtureClient(), controller = new AbortController();
    let current = true, release!: () => void, entered!: () => void;
    const waiting = new Promise<void>(resolve => { entered = resolve; });
    const api = new HubAssignmentApi({ token: async () => "synthetic", isCurrent: () => current, fetcher: async (url, init) => {
      const result = await fixture.fetcher(url, init);
      if (new URL(String(url)).pathname === deviceId) {
        const body = await result.json();
        Object.defineProperty(result, "json", { value: async () => {
          entered(); await new Promise<void>(resolve => { release = resolve; }); return body;
        } });
      }
      return result;
    } });
    const pending = api.discover(source(), controller.signal);
    await waiting;
    if (abort) controller.abort(); else current = false;
    release();
    await assert.rejects(() => pending, { name: "AbortError" });
    assert.ok(!fixture.calls.some(call => call.url.includes("Microsoft.Authorization/permissions") || call.init?.method === "PATCH"));
  }
});
test("view warning aggregation deduplicates repeated resources and preserves source, permission and legacy uncertainty", () => {
  const issue: AssignmentIssue = { resource: "fixture-service", scope: "enrollments", message: "Metadata could not be checked.",
    retryable: true, kind: "authentication" };
  const issues: AssignmentIssue[] = [issue, { ...issue }, { ...issue, resource: "another-fixture-service" },
    { scope: "source", resource: "Source", message: "Source could not be checked.", retryable: true },
    { resource: "Retained unscoped", message: "Previous uncertainty needs refresh.", retryable: false }];
  const device = assignmentIssuesForView(issues, "device");
  assert.equal(device.length, 2); assert.ok(device.every(group => group.kind !== "authentication"));
  const enrollment = assignmentIssuesForView(issues, "group");
  assert.equal(enrollment.length, 3); assert.equal(enrollment[0].count, 2);
  assert.equal(enrollment[0].kind, "authentication"); assert.equal(enrollment[0].retryable, true);
});
test("initial device collection failures are device-scoped after namespace verification and retain a usable pending page", async () => {
  for (const options of [{ collectionStatus: 401 }, { collectionStatus: 403 }, { collectionStatus: 500 }, { malformedCollection: true }]) {
    const { api, calls } = fixtureClient(options);
    const inventory = await api.discover(source(), signal());
    assert.equal(inventory.targets.length, 0); assert.equal(inventory.hasMore, true);
    assert.equal(inventory.issues.length, 1); assert.equal(inventory.issues[0].scope, "devices");
    const deviceIssues = assignmentIssuesForView(inventory.issues, "device");
    assert.equal(deviceIssues.length, 1);
    assert.equal(deviceIssues.some(issue => issue.kind === "authentication"), "collectionStatus" in options && options.collectionStatus === 401);
    assert.deepEqual(assignmentIssuesForView(inventory.issues, "group"), []);
    assert.deepEqual(assignmentIssuesForView(inventory.issues, "individual"), []);
    assert.equal(calls.filter(call => new URL(call.url).pathname === deviceId).length, 0);
  }
  const options: { collectionStatus?: number } = { collectionStatus: 403 };
  const { api, calls } = fixtureClient(options), failed = await api.discover(source(), signal());
  options.collectionStatus = undefined;
  const recovered = await api.loadMore(failed, signal());
  assert.equal(recovered.targets.length, 1); assert.equal(recovered.hasMore, false);
  assert.equal(calls.filter(call => new URL(call.url).pathname === `${nsId}/registryDevices`).length, 2);
  const namespaceFailure = await fixtureClient({ failNamespace: true }).api.discover(source(), signal());
  assert.equal(namespaceFailure.hasMore, false);
  assert.equal(namespaceFailure.issues[0].scope, "source");
  assert.equal(assignmentIssuesForView(namespaceFailure.issues, "group").length, 1);
});
test("membership needs the IoT Hub discriminator and endpoint join on the same capability entry", async () => {
  assert.equal((await fixtureClient().api.discover(source(), signal())).targets.length, 1);
  for (const options of [{ omitCapabilityType: true }, { capabilityType: "Microsoft.SoftwareUpdate" },
    { capabilityType: "Unsupported" }, { mixedCapabilities: true }]) {
    const { api, calls } = fixtureClient(options);
    assert.equal((await api.discover(source(), signal())).targets.length, 0);
    await assert.rejects(() => api.recheck(target(), source(), signal()), /no longer linked/);
    assert.equal(calls.filter(call => call.init?.method === "PATCH").length, 0);
  }
});
test("default transport preserves a native-style global receiver and ARM arguments", async () => {
  const original = globalThis.fetch, fixture = fixtureClient(), controller = new AbortController();
  let receivers = 0;
  globalThis.fetch = function (this: unknown, input, init) {
    assert.equal(this, globalThis, "The native fetch receiver must be the current global, not the API instance.");
    receivers++;
    assert.equal(init?.signal, controller.signal);
    assert.equal(init?.redirect, "error");
    return fixture.fetcher(input, init);
  };
  try {
    const api = new HubAssignmentApi({ token: async scope => {
      assert.equal(scope, ASSIGNMENT_ARM_SCOPE); return "synthetic-token";
    } });
    const inventory = await api.discover(source(), controller.signal);
    assert.equal(inventory.targets.length, 1);
    assert.equal(inventory.targets[0].reference.kind, "unassigned");
    assert.equal(receivers, fixture.calls.length);
    assert.ok(receivers > 1);
    assert.equal(fixture.calls[0].init?.method, "POST");
    assert.match(fixture.calls[0].url, /Microsoft\.ResourceGraph\/resources\?api-version=2021-03-01$/);
  } finally { globalThis.fetch = original; }
});
test("explicit injected clients retain their binding, results and typed errors without using default fetch", async () => {
  const original = globalThis.fetch, fixture = fixtureClient(), receiver = { calls: 0 };
  globalThis.fetch = function () { throw new Error("An injected client must not fall back to the default transport."); };
  const injected: typeof fetch = function (this: typeof receiver, input, init) {
    assert.equal(this, receiver);
    this.calls++;
    return fixture.fetcher(input, init);
  }.bind(receiver);
  try {
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: injected });
    assert.equal((await api.discover(source(), signal())).targets.length, 1);
    assert.equal(receiver.calls, fixture.calls.length);
    const failure = new AssignmentApiError("permission", "Synthetic access denial.", false, 403);
    const throwing: typeof fetch = function (this: typeof receiver) {
      assert.equal(this, receiver); throw failure;
    }.bind(receiver);
    await assert.rejects(() => new HubAssignmentApi({ token: async () => "synthetic", fetcher: throwing }).discover(source(), signal()),
      error => error === failure);
  } finally { globalThis.fetch = original; }
});
function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>(done => { resolve = done; });
  return { promise, resolve };
}
test("SW1 device PATCH adapter also reports only its actual mutation dispatch, not its repeated token precheck",
  { timeout: 5000 }, async () => {
  const fixture = fixtureClient({ accepted: true }), selected = (await fixture.api.discover(source(), signal())).targets[0];
  const entered = deferred<void>(), release = deferred<void>(), controller = new AbortController();
  let deferToken = false, notifications = 0;
  const api = new HubAssignmentApi({ token: async () => {
    if (deferToken) { entered.resolve(); await release.promise; }
    return "synthetic";
  }, fetcher: fixture.fetcher });
  const pending = new ModelAssignmentOperation().run([selected], model(), true, {
    read: (target, signal) => api.recheck(target, source(), signal),
    write: (target, id, signal, onDispatch) => {
      deferToken = true;
      return api.patchReference(target, id, signal, source(), () => { onDispatch(); notifications++; });
    },
  }, () => true, controller.signal, () => undefined);
  try {
    await entered.promise; controller.abort();
    assert.equal((await pending)[0].state, "cancelled");
    assert.equal(notifications, 0); assert.equal(fixture.calls.filter(call => call.init?.method === "PATCH").length, 0);
  } finally { release.resolve(); }
});
test("ARG pagination and ARM list pagination preserve all pages without exposing arbitrary nextLink", async () => {
  let argPages = 0, devicePages = 0;
  const calls: string[] = [];
  const fetcher: typeof fetch = async (url, init) => {
    const address = String(url); calls.push(address);
    if (address.includes("ResourceGraph")) {
      const body = JSON.parse(String(init?.body));
      if (body.query.includes("provisioningservices")) return response({ data: [], resultTruncated: false });
      argPages++;
      if (!body.options.$skipToken) return response({ data: [], $skipToken: "next-fixture-page" });
      return response({ data: [{ id: hubId, name: "fixture-hub", hostName: "fixture-hub.azure-devices.net" }] });
    }
    const path = new URL(address).pathname;
    if (path === hubId) return response({ id: hubId, name: "fixture-hub", properties: { hostName: "fixture-hub.azure-devices.net",
      deviceRegistry: { namespaceResourceId: nsId, linkingProperties: { state: "Succeeded" } } } });
    if (path === nsId) return response({ id: nsId, properties: { messaging: { endpoints: {} } } });
    if (path === `${nsId}/registryDevices`) {
      devicePages++;
      return response(devicePages === 1 ? { value: [], nextLink: `${address}&$skiptoken=fixture-next` } : { value: [] });
    }
    throw new Error("Unexpected fixture route.");
  };
  const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher });
  const first = await api.discover(source(), signal());
  assert.equal(devicePages, 1); assert.equal(first.hasMore, true);
  assert.equal((await api.loadMore(first, signal())).hasMore, false);
  assert.equal(argPages, 2); assert.equal(devicePages, 2);
  assert.ok(calls.every(url => new URL(url).origin === "https://management.azure.com"));
});
test("ARG malformed, truncated, duplicate continuation and auth failures are not empty success", async () => {
  for (const body of [{}, { data: null }, { data: [], resultTruncated: true }, { data: [], $skipToken: 3 }]) {
    const api = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async () => response(body) });
    await assert.rejects(() => api.discover(source(), signal()), /page|incomplete|continuation/);
  }
  const repeat = new HubAssignmentApi({ token: async () => "synthetic",
    fetcher: async () => response({ data: [], $skipToken: "same-page" }) });
  await assert.rejects(() => repeat.discover(source(), signal()), /continuation/);
  const denied = new HubAssignmentApi({ token: async () => "synthetic", fetcher: async () => response({}, 403) });
  await assert.rejects(() => denied.discover(source(), signal()), /permission/);
});
test("pending token cancellation/current-owner change fences network dispatch; unknown references cannot dispatch", async () => {
  let resolveToken!: (value: string) => void, calls = 0, current = true;
  const token = new Promise<string>(resolve => { resolveToken = resolve; });
  const controller = new AbortController();
  const api = new HubAssignmentApi({ token: () => token, isCurrent: () => current,
    fetcher: async () => { calls++; return response({ data: [] }); } });
  const pending = api.discover(source(), controller.signal); current = false; resolveToken("synthetic");
  await assert.rejects(() => pending, { name: "AbortError" }); assert.equal(calls, 0);
  await assert.rejects(() => api.patchReference({ ...target(), reference: { kind: "unknown", reason: "Reference not confirmed." } },
    document.id, signal(), source()), /not confirmed/);
  assert.equal(calls, 0);
});
test("permission evaluation subtracts within each entry then unions grants; no DataActions shortcut", () => {
  const action = "Microsoft.DeviceRegistry/namespaces/registryDevices/write";
  assert.equal(permissionGrants([{ actions: ["Microsoft.DeviceRegistry/*"], notActions: ["*/write"] },
    { actions: [action.toUpperCase()], notActions: [] }], action), true);
  assert.equal(permissionGrants([{ actions: ["*"], notActions: ["*/write"], dataActions: ["*"] }], action), false);
  assert.equal(permissionGrants([], action), false);
  assert.throws(() => permissionGrants([{ actions: ["*"] }], action), /incomplete/);
  assert.equal(specificEtag("opaque-base64=="), '"opaque-base64=="');
  assert.equal(specificEtag('"quoted"'), '"quoted"');
  for (const etag of ["*", '"*"', "", "bad\nvalue", 'bad"value']) assert.throws(() => specificEtag(etag));
});
test("real ADR transport rechecks membership/reference/permission and sends only refs with specific If-Match", async () => {
  for (const accepted of [false, true]) {
    const { api, calls } = fixtureClient({ accepted });
    const inventory = await api.discover(source(), signal());
    let notifications = 0;
    const result = await api.patchReference(inventory.targets[0], document.id, signal(), source(), () => notifications++);
    assert.equal(result, accepted ? "accepted" : "completed");
    const writes = calls.filter(call => call.init?.method === "PATCH");
    assert.equal(writes.length, 1);
    assert.equal(notifications, 1);
    assert.equal(new Headers(writes[0].init?.headers).get("If-Match"), '"fixture-etag-1"');
    assert.deepEqual(JSON.parse(String(writes[0].init?.body)), { properties: { deviceTypeRefs: [document.id] } });
    assert.ok(calls.filter(call => call.url.includes("Microsoft.Authorization/permissions")).length >= 2);
    assert.ok(calls.findIndex(call => call.init?.method === "PATCH") > calls.findIndex(call => call.url.includes("/capabilities")));
  }
  for (const options of [{ omit: true, completePoint: false }, { denyWrite: true }, { malformedPermissions: true }]) {
    const { api, calls } = fixtureClient(options), inventory = await api.discover(source(), signal());
    const value = inventory.targets[0];
    assert.ok(value.reference.kind === "unknown" || value.writeBlocker);
    await assert.rejects(() => api.patchReference(value, document.id, signal(), source()));
    assert.equal(calls.filter(call => call.init?.method === "PATCH").length, 0);
  }
});
const services = (events: string[], patch: Partial<AssignmentServices> = {}): AssignmentServices => {
  const write: AssignmentServices["write"] = patch.write || (async () => { events.push("write"); return "accepted"; });
  return { read: patch.read || (async value => { events.push("read"); return { ...value }; }),
    write: (target, id, signal, onDispatch) => { onDispatch(); return write(target, id, signal, onDispatch); } };
};
test("first assignment requires explicit confirmation, a fresh empty reference and specific unchanged receipt", async () => {
  const operation = new ModelAssignmentOperation(), events: string[] = [], receipts: AssignmentReceipt[] = [];
  await assert.rejects(() => operation.run([target()], model(), false, services(events), () => true, signal(), () => undefined), /Confirm/);
  const result = await operation.run([target()], model(), true, services(events), () => true, signal(), value => receipts.push(value));
  assert.deepEqual(events, ["read", "write"]); assert.equal(result[0].state, "accepted");
  assert.ok(!result[0].message?.includes("matching data verified"));
  assert.equal(receipts[0].state, "checking");
  for (const changed of [
    { reference: { kind: "assigned" as const, ids: ["urn:other"] as [string] } },
    { etag: '"new-etag"' }, { hubIds: [] }, { namespaceId: `${nsId}-other` },
  ]) {
    let writes = 0;
    const result = await operation.run([target()], model(), true,
      services([], { read: async value => ({ ...value, ...changed }), write: async () => { writes++; return "completed"; } }),
      () => true, signal(), () => undefined);
    assert.equal(writes, 0); assert.ok(["failed", "blocked"].includes(result[0].state));
  }
});
test("same/other/unknown/wildcard/preblocked targets are read-only even when the caller confirms", async () => {
  for (const patch of [
    { reference: { kind: "assigned" as const, ids: [document.id] as [string] } },
    { reference: { kind: "assigned" as const, ids: ["urn:other"] as [string] } },
    { reference: { kind: "unknown" as const, reason: "No reference returned." } }, { etag: "*" }, { writeBlocker: "Contract not verified." },
  ]) {
    const events: string[] = [];
    const result = await new ModelAssignmentOperation().run([{ ...target(), ...patch }], model(), true, services(events), () => true, signal(), () => undefined);
    assert.equal(result[0].state, "blocked"); assert.deepEqual(events, []);
  }
});
test("synchronous completion needs no stale or mutable device readback", async () => {
  const stale: AssignmentTarget = { ...pollingTarget(), referenceEvidence: { source: "adr-point-get",
    observedAt: "2026-11-05T00:00:00Z", apiVersion: "2026-11-02-preview", rawCarrier: "present" } };
  assert.equal(assignmentBlocker(stale), CURRENT_REFERENCE_FRESHNESS);
  let reads = 0;
  const result = await new ModelAssignmentOperation().run([target()], model(), true, services([], {
    read: async value => ++reads === 1 ? value : { ...stale, reference: { kind: "assigned", ids: [document.id] } },
    write: async () => "completed",
  }), () => true, signal(), () => undefined);
  assert.equal(result[0].state, "confirmed");
  assert.equal(reads, 1);
});
test("late preflight cannot dispatch, double submits fail, and postdispatch cancellation stays unknown", async () => {
  let finish!: (value: AssignmentTarget) => void, current = true, writes = 0;
  const operation = new ModelAssignmentOperation();
  const pending = operation.run([target()], model(), true, services([], {
    read: () => new Promise(resolve => { finish = resolve; }), write: async () => { writes++; return "completed"; },
  }), () => current, signal(), () => undefined);
  await assert.rejects(() => operation.run([target()], model(), true, services([]), () => true, signal(), () => undefined), /already running/);
  current = false; finish(target()); assert.equal((await pending)[0].state, "cancelled"); assert.equal(writes, 0);
  const controller = new AbortController();
  const result = await operation.run([target()], model(), true, services([], {
    write: async () => { controller.abort(); return "completed"; },
  }), () => true, controller.signal, () => undefined);
  assert.equal(result[0].state, "unknown");
});
test("cancelling a never-resolving owned read retires checking immediately; late read cannot write", async () => {
  let release!: (value: AssignmentTarget) => void, writes = 0;
  const controller = new AbortController(), operation = new ModelAssignmentOperation();
  const pending = operation.run([target()], model(), true, services([], {
    read: () => new Promise(resolve => { release = resolve; }), write: async () => { writes++; return "completed"; },
  }), () => true, controller.signal, () => undefined);
  controller.abort();
  assert.equal((await pending)[0].state, "cancelled");
  release(target()); await Promise.resolve(); assert.equal(writes, 0);
  assert.equal((await operation.run([target()], model(), true, services([]), () => true, signal(), () => undefined))[0].state, "accepted");
});
test("confirmed readback, rejected writes and uncertain transport outcomes stay distinct", async () => {
  for (const outcome of ["confirmed", "rejected", "uncertain"] as const) {
    let reads = 0;
    const result = await new ModelAssignmentOperation().run([target()], model(), true, services([], {
      read: async value => ++reads === 1 ? value : { ...value, reference: { kind: "assigned", ids: [document.id] } },
      write: async () => {
        if (outcome === "rejected") throw Object.assign(new Error("Permission denied."), { status: 403 });
        if (outcome === "uncertain") throw new Error("Synthetic transport failure.");
        return "completed";
      },
    }), () => true, signal(), () => undefined);
    assert.equal(result[0].state, outcome === "confirmed" ? "confirmed" : outcome === "rejected" ? "failed" : "unknown");
  }
});
