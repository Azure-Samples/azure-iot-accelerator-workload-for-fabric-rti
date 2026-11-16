// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import path from "node:path";
import vm from "node:vm";
import crypto from "node:crypto";
import { fileURLToPath } from "node:url";
import ts from "typescript";
import { HubAssignmentApi, ASSIGNMENT_ARM_SCOPE } from "./HubAssignmentApi.ts";
import { sourceHubs, assignmentBlocker, ModelAssignmentOperation, type AssignmentSource } from "./HubAssignment.ts";
import { buildNativeEventStreamHubUpdate, IOT_HUB_ROUTING_API_VERSION, type NativeEventStreamRoutes } from "./IoTHubRouting.ts";
import { buildEventstreamTopology, validateRawEventstreamTopology, validateRawTableSchema } from "./RawIngestion.ts";
import { mixedHubRoot, mixedNamespace, mixedDtdlRequest } from "../../../tests/mixedSourceDtdlFixtures.ts";

const namespace3 = `${mixedHubRoot}/providers/Microsoft.DeviceRegistry/namespaces/separate-three`;
type HubKey = "regular1" | "linked1" | "linked2" | "regular2" | "linked3";
const hubs: Record<HubKey, { id: string; host: string; namespace: string | undefined }> = {
  regular1: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/regular-hub-1`, host: "regular-hub-1.azure-devices.net", namespace: undefined },
  linked1: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/linked-hub-1`, host: "linked-hub-1.azure-devices.net", namespace: mixedNamespace },
  linked2: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/linked-hub-2`, host: "linked-hub-2.azure-devices.net", namespace: mixedNamespace },
  regular2: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/regular-hub-2`, host: "regular-hub-2.azure-devices.net", namespace: undefined },
  linked3: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/linked-hub-3`, host: "linked-hub-3.azure-devices.net", namespace: namespace3 },
};
const order: HubKey[] = ["regular1", "linked1", "linked2", "regular2", "linked3"];
const outcomes: Record<string, unknown>[] = [];
const sha = (value: string) => crypto.createHash("sha256").update(value).digest("hex");
process.once("exit", () => {
  if (process.env.FIVE_HUB_RESULTS) fs.writeFileSync(process.env.FIVE_HUB_RESULTS,
    JSON.stringify({ provenance: "Agent-authored five-Hub/two-namespace fixtures, not actual provisioned resources",
      cases: outcomes, count: outcomes.length, privateHelperSourceSha256: sha(tableStepSource),
      limits: "Pure/injected ARM only; no browser, cloud, deployment, live ingestion or no-new-table orchestration execution." }, null, 2) + "\n");
});
function source(keys: HubKey[] = order): AssignmentSource {
  const docstring = (role: string) => `Existing raw data meaning.\n@iot.meta tableRole=${role}\n` +
    `@iot.meta source=${keys.map(key => hubs[key].host).join("; ")}\n@iot.meta keep=unrelated`;
  return { key: "five-hub-synthetic", telemetry: { table: "TelemetryRawIoTHub_combined_hubs", docstring: docstring("iothub-telemetry-raw") },
    twin: { table: "PropertiesRawIoTHub_combined_hubs", docstring: docstring("iothub-properties-raw") } };
}
function fixture(absentRegular1 = false) {
  const calls: { url: string; method: string }[] = [], scopes: string[] = [];
  const devices = [
    { namespace: mixedNamespace, name: "hub1-device", endpoints: ["linked1"] },
    { namespace: mixedNamespace, name: "hub2-device", endpoints: ["linked2"] },
    { namespace: mixedNamespace, name: "both-device", endpoints: ["linked1", "linked2"] },
    { namespace: mixedNamespace, name: "cross-entry", endpoints: ["cross-entry"] },
    { namespace: mixedNamespace, name: "foreign-endpoint", endpoints: ["linked3"] },
    { namespace: namespace3, name: "hub3-device", endpoints: ["linked3"] },
    { namespace: namespace3, name: "foreign-endpoint-three", endpoints: ["linked1"] },
  ].map(device => ({ ...device, id: `${device.namespace}/registryDevices/${device.name}` }));
  const reply = (body: unknown) => new Response(JSON.stringify(body), { headers: { "Content-Type": "application/json" } });
  const fetcher: typeof fetch = async (input, init) => {
    const url = String(input), parsed = new URL(url), pathname = parsed.pathname;
    calls.push({ url, method: init?.method || "GET" });
    assert.equal(parsed.origin, "https://management.azure.com");
    assert.ok(["GET", "POST"].includes(init?.method || "GET"), "No fixture writes are allowed.");
    if (url.includes("ResourceGraph")) {
      const query = JSON.parse(String(init?.body)).query as string;
      assert.ok(!query.includes("registrydevices"), "Device references must not use a per-device ARG query.");
      assert.match(query, /microsoft.devices\/iothubs/);
      return reply({ data: Object.values(hubs).filter(hub => query.includes(JSON.stringify(hub.host)))
        .map(hub => ({ id: hub.id, name: hub.host.split(".")[0], hostName: hub.host })) });
    }
    for (const [key, hub] of Object.entries(hubs)) if (pathname === hub.id) return reply({ id: hub.id,
      name: hub.host.split(".")[0], properties: { hostName: hub.host,
        ...(absentRegular1 && key === "regular1" ? {} : { deviceRegistry: hub.namespace
          ? { namespaceResourceId: hub.namespace, linkingProperties: { state: "Succeeded" } } : null }) } });
    if (pathname === mixedNamespace || pathname === namespace3) return reply({ id: pathname, properties: { messaging: { endpoints:
      Object.fromEntries(Object.entries(hubs).filter(([, hub]) => hub.namespace === pathname)
        .map(([key, hub]) => [key, { endpointType: "Microsoft.Devices/IotHubs", resourceId: hub.id }])) } } });
    if ([mixedNamespace, namespace3].some(namespace => pathname === `${namespace}/registryDevices`))
      return reply({ value: devices.filter(device => pathname === `${device.namespace}/registryDevices`).map(device => ({ id: device.id })) });
    const device = devices.find(device => pathname === device.id || pathname.startsWith(`${device.id}/`));
    assert.ok(device, `Unexpected five-Hub fixture path: ${pathname}`);
    if (pathname.endsWith("/capabilities")) return reply({ value: device.name === "cross-entry" ? [
      { id: `${device.id}/capabilities/right-kind-wrong-endpoint`, properties: { capabilityType: "Microsoft.IoTHub", messagingEndpointName: "wrong" } },
      { id: `${device.id}/capabilities/wrong-kind-right-endpoint`, properties: { capabilityType: "Microsoft.SoftwareUpdate", messagingEndpointName: "linked1" } },
    ] : device.endpoints.map(endpoint => ({ id: `${device.id}/capabilities/${endpoint}`,
      properties: { capabilityType: "Microsoft.IoTHub", messagingEndpointName: endpoint } })) });
    if (pathname.endsWith("/permissions")) return reply({ value: [{ actions: ["*"], notActions: [] }] });
    assert.equal(pathname, device.id);
    return reply({ id: device.id, name: device.name, type: "Microsoft.DeviceRegistry/namespaces/registryDevices",
      etag: `"five-${device.name}-current"`, properties: { deviceTypeRefs: [], provisioningState: "Succeeded",
        uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: device.name } });
  };
  return { calls, scopes, fetcher, token: async (scope: string) => {
    scopes.push(scope); assert.equal(scope, ASSIGNMENT_ARM_SCOPE); return "synthetic-only";
  } };
}
test("five source union preserves prior order and deduplicates both stream provenance without a one-Hub restriction", () => {
  const input = source();
  input.twin.docstring = source(["linked3", "regular1", "linked2", "regular2", "linked1", "linked3"]).twin.docstring;
  assert.deepEqual(sourceHubs(input), { sources: order.map(key => hubs[key].host), issues: [] });
  outcomes.push({ name: "five-source union/order/dedup", input, expected: order.map(key => hubs[key].host), actual: sourceHubs(input), passed: true });
});
for (const absentRegular1 of [false, true]) test(`five Hubs use one shared namespace list plus a separate namespace list; absent regular1=${absentRegular1}`, async () => {
  const apiFixture = fixture(absentRegular1), inventory = await new HubAssignmentApi(apiFixture).discover(source(), new AbortController().signal);
  assert.deepEqual(inventory.hubs.map(hub => hub.state), ["unlinked", "linked", "linked", "unlinked", "linked"]);
  assert.deepEqual(inventory.targets.map(target => target.name), ["hub1-device", "hub2-device", "both-device", "hub3-device"]);
  assert.equal(inventory.issues.length, 0);
  assert.equal(new Set(inventory.targets.map(target => target.id)).size, 4);
  assert.equal(new Set(inventory.targets.map(target => target.deviceIdentity!.externalDeviceId)).size, 4, "Fixtures keep the existing unique native identity assumption.");
  for (const namespace of [mixedNamespace, namespace3]) assert.equal(apiFixture.calls
    .filter(call => new URL(call.url).pathname === `${namespace}/registryDevices`).length, 1);
  const expectedHubs = [[hubs.linked1.id], [hubs.linked2.id], [hubs.linked1.id, hubs.linked2.id], [hubs.linked3.id]];
  inventory.targets.forEach((target, index) => {
    assert.deepEqual(target.hubIds, expectedHubs[index]);
    assert.equal(target.namespaceId, index === 3 ? namespace3 : mixedNamespace);
    assert.equal(assignmentBlocker(target), undefined);
  });
  assert.ok(apiFixture.calls.every(call => !/provisioning|enrollment/i.test(call.url)));
  outcomes.push({ name: `five-Hub two-namespace discovery absent=${absentRegular1}`, hubStates: inventory.hubs,
    targets: inventory.targets, expectedHubMemberships: expectedHubs, sourceIssues: inventory.issues,
    namespaceListCounts: [1, 1], dpsRequests: 0, writes: 0, passed: true });
});
for (const selected of [["linked1"], ["linked2"], ["linked3"]] as HubKey[][]) test(`namespace and capability applicability only for selected ${selected[0]}`, async () => {
  const apiFixture = fixture(), inventory = await new HubAssignmentApi(apiFixture).discover(source(selected), new AbortController().signal);
  const expected = selected[0] === "linked1" ? ["hub1-device", "both-device"] : selected[0] === "linked2" ? ["hub2-device", "both-device"] : ["hub3-device"];
  assert.deepEqual(inventory.targets.map(target => target.name), expected);
  assert.ok(inventory.targets.every(target => target.hubIds.length === 1 && target.hubIds[0] === hubs[selected[0]].id));
  const unselectedNamespace = selected[0] === "linked3" ? mixedNamespace : namespace3;
  assert.ok(!apiFixture.calls.some(call => new URL(call.url).pathname.startsWith(unselectedNamespace)));
  outcomes.push({ name: `selected ${selected[0]} only`, expected, actual: inventory.targets, noUnrelatedNamespaceDiscovery: true, passed: true });
});
test("two nonlinked sources alone do not discover either namespace and DTDL cannot invoke assignment adapters", async () => {
  const apiFixture = fixture(true), inventory = await new HubAssignmentApi(apiFixture).discover(source(["regular1", "regular2"]), new AbortController().signal);
  assert.deepEqual(inventory.hubs.map(hub => hub.state), ["unlinked", "unlinked"]); assert.equal(inventory.targets.length, 0);
  assert.ok(!apiFixture.calls.some(call => /namespaces/.test(call.url)));
  let adapters = 0;
  await assert.rejects(() => new ModelAssignmentOperation().run([], mixedDtdlRequest().input.model, true, {
    read: async target => { adapters++; return target; }, write: async () => { adapters++; return "accepted"; },
  }, () => true, new AbortController().signal, () => undefined), /WoT/);
  assert.equal(adapters, 0);
  outcomes.push({ name: "nonlinked nondiscovery/DTDL operation boundary", hubStates: inventory.hubs,
    namespaceReads: 0, dtdlAssignmentAdapters: 0, passed: true });
});

// Compile the exact first-party private helpers, not a reimplementation or the React module.
const tableStepSource = fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), "steps", "TableSetupStep.tsx"), "utf8");
const ast = ts.createSourceFile("TableSetupStep.tsx", tableStepSource, ts.ScriptTarget.ES2022, true, ts.ScriptKind.TSX);
test("raw table reuse retains the original guidance and apply action without a healthy Refresh button", () => {
  const elements: ts.JsxElement[] = [];
  const visit = (node: ts.Node) => {
    if (ts.isJsxElement(node)) elements.push(node);
    ts.forEachChild(node, visit);
  };
  visit(ast);
  const description = elements.find(node => node.openingElement.tagName.getText(ast) === "Text" &&
    node.openingElement.attributes.properties.some(attribute => ts.isJsxAttribute(attribute) &&
      attribute.name.getText(ast) === "className" && attribute.initializer &&
      ts.isStringLiteral(attribute.initializer) && attribute.initializer.text === "iot-solution-step-description"));
  assert.ok(description);
  const guidance = description.children.filter(ts.isJsxText).map(node => node.text).join(" ").replace(/\s+/g, " ").trim();
  assert.ok(guidance.endsWith("You may choose to create new raw data tables for this IoT Hub or combine data from multiple IoT Hubs by reusing existing tables."));
  assert.ok(!description.getText(ast).includes("Existing tables must include"));
  const buttons = elements.filter(node => node.openingElement.tagName.getText(ast) === "Button");
  assert.ok(!buttons.some(node => node.children.some(child => ts.isJsxText(child) && child.text.trim() === "Refresh")));
  const apply = buttons.find(node => node.getText(ast).includes('"Use selected tables"'));
  assert.ok(apply);
  assert.match(apply.openingElement.getText(ast), /onClick=\{applyReuse\}/);
  assert.match(apply.openingElement.getText(ast), /disabled=\{!canApplyReuse\}/);
  const retry = buttons.find(node => node.children.some(child => ts.isJsxText(child) && child.text.trim() === "Retry"));
  assert.ok(retry);
  assert.match(retry.openingElement.getText(ast), /loadReusableTables\(\)/);
});
const helperNames = ["appendIotMetaSource", "toKqlString"];
test("IoT Hub discovery refresh is an accessible icon action with the existing loader and loading gate", () => {
  const hubSource = fs.readFileSync(path.join(path.dirname(fileURLToPath(import.meta.url)), "steps", "IoTHubStep.tsx"), "utf8");
  const hubAst = ts.createSourceFile("IoTHubStep.tsx", hubSource, ts.ScriptTarget.ES2022, true, ts.ScriptKind.TSX);
  let refresh: ts.JsxSelfClosingElement | undefined;
  const visit = (node: ts.Node) => {
    if (ts.isJsxSelfClosingElement(node) && node.tagName.getText(hubAst) === "Button" &&
        node.attributes.properties.some(attribute => ts.isJsxAttribute(attribute) &&
          attribute.name.getText(hubAst) === "aria-label" && attribute.initializer &&
          ts.isStringLiteral(attribute.initializer) && attribute.initializer.text === "Refresh IoT Hubs")) refresh = node;
    ts.forEachChild(node, visit);
  };
  visit(hubAst);
  assert.ok(refresh);
  assert.match(refresh.getText(hubAst), /appearance="transparent"/);
  assert.match(refresh.getText(hubAst), /size="large"/);
  assert.match(refresh.getText(hubAst), /icon=\{<ArrowClockwise24Regular \/>\}/);
  assert.ok(!refresh.attributes.properties.some(attribute => ts.isJsxAttribute(attribute) &&
    attribute.name.getText(hubAst) === "style"));
  assert.match(refresh.getText(hubAst), /onClick=\{\(\) => \{ void loadHubs\(\); \}\}/);
  assert.match(refresh.getText(hubAst), /disabled=\{loadingHubs\}/);
  assert.ok(ts.isJsxElement(refresh.parent));
  assert.equal(refresh.parent.openingElement.getText(hubAst), '<Tooltip content="Refresh IoT Hubs" relationship="label">');
});
const declarations = ast.statements.filter((node): node is ts.FunctionDeclaration => ts.isFunctionDeclaration(node) &&
  !!node.name && helperNames.includes(node.name.text));
assert.equal(declarations.length, 2);
const helperJs = ts.transpileModule(declarations.map(node => node.getText(ast)).join("\n") +
  "\nexport { appendIotMetaSource, toKqlString };", { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 } }).outputText;
const exportsBox: { appendIotMetaSource?: (doc: string, source: string) => string; toKqlString?: (value: string) => string } = {};
vm.runInNewContext(helperJs, { exports: exportsBox }, { timeout: 1000, filename: "first-party-TableSetup-private-helpers.cjs" });
const append = exportsBox.appendIotMetaSource!, escape = exportsBox.toKqlString!;

for (const additions of [["linked2", "regular2", "linked3"], ["linked3", "linked2", "regular2"]] as HubKey[][])
  test(`fresh sequential reuse appends 2 to 3 to 4 to 5 sources in actual order ${additions.join("/")}`, () => {
    const input = source(["regular1", "linked1"]), sequence: Record<string, unknown>[] = [];
    const expectedOrder: HubKey[] = ["regular1", "linked1"];
    for (const addition of additions) {
      expectedOrder.push(addition);
      for (const stream of ["telemetry", "twin"] as const) {
        const before = input[stream].docstring;
        input[stream].docstring = append(before, hubs[addition].host);
        assert.equal(append(input[stream].docstring, hubs[addition].host), input[stream].docstring);
        assert.ok(input[stream].docstring.startsWith("Existing raw data meaning.\n@iot.meta tableRole="));
        assert.ok(input[stream].docstring.endsWith("@iot.meta keep=unrelated"));
        assert.equal(input[stream].docstring.replace(/^@iot\.meta source=.*$/m, ""), before.replace(/^@iot\.meta source=.*$/m, ""));
        assert.equal((input[stream].docstring.match(/^@iot\.meta source=/gm) || []).length, 1);
      }
      assert.deepEqual(sourceHubs(input), { sources: expectedOrder.map(key => hubs[key].host), issues: [] });
      sequence.push({ addition, sourceCount: expectedOrder.length, docstrings: structuredClone(input) });
    }
    assert.ok(escape(input.telemetry.docstring).includes("\\n@iot.meta source="));
    outcomes.push({ name: `actual private append sequence ${additions.join("/")}`, sequence, helperSourceSha256: sha(tableStepSource),
      caveat: "Fresh DocString input per addition. Does not prove concurrent stale-writer preservation or actual UI orchestration.", passed: true });
  });
test("source annotation appends a missing source line without rewriting role or description", () => {
  const before = "Raw meaning.\n@iot.meta tableRole=iothub-telemetry-raw\n@iot.meta other=unchanged";
  assert.equal(append(before, hubs.linked3.host), `${before}\n@iot.meta source=${hubs.linked3.host}`);
  assert.equal(append(`${before}\n@iot.meta source=${hubs.linked3.host}`, hubs.linked3.host), `${before}\n@iot.meta source=${hubs.linked3.host}`);
  outcomes.push({ name: "actual private append missing source/idempotence", before, passed: true });
});

function routingInput(key: HubKey): NativeEventStreamRoutes {
  const endpoint: NativeEventStreamRoutes["telemetryEndpoint"] = { name: `owned-${key}-telemetry`, endpointUri: "sb://synthetic-combined", entityPath: "telemetry",
    workspaceId: "synthetic-workspace", eventStreamId: "existing-telemetry", sourceId: "existing-telemetry-source", userAssignedIdentity: null };
  return { telemetryEndpoint: endpoint, propertiesEndpoint: { ...endpoint, name: `owned-${key}-properties`,
    entityPath: "properties", eventStreamId: "existing-properties", sourceId: "existing-properties-source" },
  telemetryRouteName: `owned-${key}-telemetry-route`, propertiesRouteName: `owned-${key}-properties-route` };
}
test("all three added Hubs route to existing paired Eventstreams without rewriting other endpoints routes auth or Hub state", () => {
  assert.equal(IOT_HUB_ROUTING_API_VERSION, "2026-10-01-preview");
  const updates: Record<string, unknown>[] = [];
  for (const key of ["linked2", "regular2", "linked3"] as HubKey[]) {
    const before = { id: hubs[key].id, etag: `"current-${key}"`, tags: { keep: "original" },
      identity: { type: "UserAssigned", userAssignedIdentities: { "synthetic-existing-identity": {} } },
      properties: { deviceRegistry: hubs[key].namespace ? { namespaceResourceId: hubs[key].namespace } : null,
        disableLocalAuth: true, routing: { fallbackRoute: { isEnabled: false }, enrichments: [{ key: "keep" }],
          endpoints: { eventStreams: [{ name: "existing-nonselected", messagePayloadFormat: "None", identity: { userAssignedIdentity: "retain" } }],
            eventHubs: [{ name: "legacy-not-owned", authenticationType: "identityBased" }] },
          routes: [{ name: "nonselected", source: "DeviceMessages", condition: "false", endpointNames: ["existing-nonselected"] }] } } };
    const original = structuredClone(before), input = routingInput(key), update = buildNativeEventStreamHubUpdate(before, input);
    assert.deepEqual(before, original); assert.equal(update.ifMatch, before.etag);
    assert.deepEqual(update.body.identity, before.identity); assert.deepEqual(update.body.tags, before.tags);
    assert.deepEqual(update.body.properties.deviceRegistry, before.properties.deviceRegistry);
    const routing = update.body.properties.routing!;
    assert.deepEqual(routing.endpoints!.eventStreams![0], before.properties.routing.endpoints.eventStreams[0]);
    assert.deepEqual(routing.endpoints!.eventHubs, before.properties.routing.endpoints.eventHubs);
    assert.deepEqual(routing.routes![0], before.properties.routing.routes[0]);
    assert.deepEqual(routing.fallbackRoute, before.properties.routing.fallbackRoute);
    assert.deepEqual(routing.enrichments, before.properties.routing.enrichments);
    assert.deepEqual(routing.endpoints!.eventStreams!.slice(1).map(endpoint => endpoint.messagePayloadFormat), ["DOObservationV1", "DOObservationV1"]);
    assert.deepEqual(routing.routes!.slice(1).map(route => route.source), ["DeviceMessages", "TwinChangeEvents"]);
    assert.ok(routing.routes!.every(route => !("messagePayloadFormat" in route)));
    assert.deepEqual(buildNativeEventStreamHubUpdate(update.body, input).body, update.body);
    updates.push({ hub: key, ifMatch: update.ifMatch, ownedEndpointCount: 2, existingDestinations: ["existing-telemetry", "existing-properties"] });
  }
  for (const etag of ["", "*", '"*"', "a,b"]) assert.throws(() => buildNativeEventStreamHubUpdate({ etag, properties: {} }, routingInput("linked2")), /ETag/);
  outcomes.push({ name: "three new Hub configs / two existing native Eventstreams / conditional preservation", updates,
    crossHubResourceWrites: 0, actualNetworkCalls: 0, passed: true });
});
test("paired reused combined destinations retain raw three-column processed topology without creating tables", () => {
  const rawSchema = { OrderedColumns: [{ Name: "data", CslType: "string" }, { Name: "headers", CslType: "dynamic" },
    { Name: "user_headers", CslType: "dynamic" }] };
  const before = structuredClone(rawSchema);
  for (const kind of ["Telemetry", "Properties"]) {
    const tableName = `${kind}RawIoTHub_combined_hubs`;
    const topology = buildEventstreamTopology(kind, "existing-workspace", "existing-db", "iot-eh", tableName);
    const original = structuredClone(topology);
    validateRawTableSchema(tableName, rawSchema);
    validateRawEventstreamTopology(kind, topology, { tableName, databaseId: "existing-db", databaseName: "iot-eh" });
    assert.equal(topology.destinations[0].properties.tableName, tableName);
    assert.deepEqual(topology, original); assert.deepEqual(rawSchema, before);
  }
  outcomes.push({ name: "paired existing combined raw contracts", schema: rawSchema, pureValidationOnly: true,
    caveat: "No live/mounted TableSetup orchestration executed; no cloud no-create claim is inferred from helper validation.", passed: true });
});
