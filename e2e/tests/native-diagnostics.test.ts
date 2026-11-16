import test from "node:test";
import assert from "node:assert/strict";
import { assignmentOperationState, assignmentRequestStage, nativeResourceShape } from "../src/recording.js";

test("native response diagnostics retain only structural keys and presence flags, never connection values", () => {
  const result = nativeResourceShape({
    fullyQualifiedNamespace: "sensitive-host-value", eventHubName: "sensitive-eventhub-value",
    primaryKey: "sensitive-key-value", properties: { connectionString: "sensitive-connection-value" },
    sources: [{ id: "sensitive-source-value", type: "CustomEndpoint", properties: { secret: "sensitive-nested-value" } }],
  });

  assert.equal(result.hasNamespace, true);
  assert.equal(result.hasEventHubName, true);
  assert.equal(result.hasCustomEndpoint, true);
  assert.equal(result.hasCustomEndpointId, true);
  assert.doesNotMatch(JSON.stringify(result), /sensitive-/);
  assert.deepEqual(result.childKeys, { properties: ["connectionString"] });
  assert.deepEqual(nativeResourceShape("sensitive-plaintext-value"), { kind: "string" });
});

test("assignment timing classification excludes foreign resources, credentials, bodies and signed URL queries", () => {
  const group = "/subscriptions/00000000-0000-4000-8000-000000000001/resourceGroups/owned";
  const device = `${group}/providers/Microsoft.DeviceRegistry/namespaces/owned/registryDevices/device`;
  const operation = "/subscriptions/00000000-0000-4000-8000-000000000001/providers/microsoft.deviceregistry/locations/australiaeast/operationstatuses/operation";
  const known = new Set([operation]);
  assert.equal(assignmentRequestStage(`https://management.azure.com${device}?sig=sensitive-query`, "PATCH", group, known), "assignment-write");
  assert.equal(assignmentRequestStage(`https://management.azure.com${device}`, "GET", group, known), "device-point-read");
  assert.equal(assignmentRequestStage(`https://management.azure.com${device}/capabilities?sig=sensitive-query`, "GET", group, known), "capability-read");
  assert.equal(assignmentRequestStage(`https://management.azure.com${operation}?sig=sensitive-query`, "GET", group, known), "original-operation-status");
  assert.equal(assignmentRequestStage(`https://management.azure.com${operation}/foreign`, "GET", group, known), undefined);
  assert.equal(assignmentRequestStage(`https://management.azure.com${device.replace("/owned/", "/foreign/")}`, "PATCH", group, known), undefined);
  assert.equal(assignmentRequestStage(`https://untrusted.example${device}`, "PATCH", group, known), undefined);
  assert.equal(assignmentOperationState({ status: "Succeeded", primaryKey: "sensitive-value" }), "Succeeded");
  assert.equal(assignmentOperationState({ status: "sensitive-value" }), "unrecognized");
  assert.equal(assignmentOperationState("sensitive-value"), "unrecognized");
});
