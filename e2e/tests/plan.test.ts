import test from "node:test";
import assert from "node:assert/strict";
import { runConfig } from "../src/config.js";
import { assertOwnedPlan, createPlan } from "../src/plan.js";

const env = {
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "Australia East",
  E2E_RETENTION: "retain",
  E2E_DEPLOYMENT: "development",
};

test("plan owns exactly five Hubs with the requested 2/3 namespace split", () => {
  const plan = createPlan(runConfig(env));
  assert.equal(plan.azure.hubs.length, 5);
  assert.equal(plan.azure.hubs.filter(hub => hub.kind === "nonlinked").length, 2);
  assert.deepEqual(plan.azure.hubs.filter(hub => hub.kind === "future-linked").map(hub => hub.namespaceKey), ["a", "a", "b"]);
  assert.equal(plan.azure.namespaces.length, 2);
  assert.equal(plan.requiredSupport.dpsInstances, 0);
  assert.equal(plan.deferredSupport.dpsInstances, 2);
  assert.equal(plan.requiredSupport.userAssignedIdentities, 1);
  assert.ok(plan.azure.hubs.every(hub => hub.sku === "S1"));
  assert.equal(plan.provisioningContract, "s1-hubs-linked-deferred");
  assert.equal(plan.azure.tags["e2e-run-id"], plan.runId);
  assert.equal(plan.target.location, "australiaeast");
  assert.equal(plan.resourcesCreated, false);
  assert.equal(plan.verification.rawData, "not-run");
});

test("every planned Azure resource is inside one fresh resource group", () => {
  const plan = createPlan(runConfig(env));
  const resources = [
    ...plan.azure.hubs,
    ...plan.azure.namespaces,
    ...plan.azure.namespaces.map(namespace => namespace.dps),
    plan.azure.regularIdentity,
  ];
  assert.equal(new Set(resources.map(resource => resource.id)).size, 10);
  for (const resource of resources) {
    assert.ok(resource.id.startsWith(`${plan.azure.resourceGroupId}/providers/`));
    assert.ok(resource.name.length <= 64);
    assert.match(resource.name, /^[a-z0-9][a-z0-9-]+[a-z0-9]$/);
  }
});

test("runs get distinct ownership and globally collision-resistant names", () => {
  const config = runConfig(env);
  const first = createPlan(config);
  const second = createPlan(config);
  assert.notEqual(first.runId, second.runId);
  assert.notEqual(first.azure.resourceGroup, second.azure.resourceGroup);
  const firstNames = new Set(first.azure.hubs.map(hub => hub.name));
  assert.ok(second.azure.hubs.every(hub => !firstNames.has(hub.name)));
});

test("live configuration validates targets and defaults to full-suite cleanup", () => {
  assert.throws(() => runConfig({ ...env, E2E_LOCATION: "" }), /E2E_LOCATION/);
  assert.equal(runConfig({ ...env, E2E_LOCATION: "westus" }).location, "westus");
  assert.throws(() => runConfig({ ...env, E2E_SUBSCRIPTION_ID: "" }), /SUBSCRIPTION_ID/);
  assert.throws(() => runConfig({ ...env, E2E_RETENTION: "delete" }), /E2E_RETENTION/);
  assert.equal(runConfig({ ...env, E2E_RETENTION: undefined }).retention, "cleanup");
  assert.equal(runConfig({ ...env, E2E_RETENTION: "retain" }).retention, "retain");
  assert.throws(() => runConfig({ ...env, E2E_DEPLOYMENT: "" }), /E2E_DEPLOYMENT/);
});

test("continuation accepts only the exact owned topology and supplied target", () => {
  const config = runConfig(env);
  const plan = createPlan(config);
  assert.doesNotThrow(() => assertOwnedPlan(plan, config, plan.runId));
  const foreignHub = structuredClone(plan);
  const hub = foreignHub.azure.hubs[0];
  assert.ok(hub);
  hub.id = "/subscriptions/foreign/resourceGroups/existing/providers/Microsoft.Devices/IotHubs/existing";
  assert.throws(() => assertOwnedPlan(foreignHub, config, plan.runId), /ownership/);
  const wrongNamespace = structuredClone(plan);
  const linked = wrongNamespace.azure.hubs[2];
  assert.ok(linked);
  linked.namespaceKey = "b";
  assert.throws(() => assertOwnedPlan(wrongNamespace, config, plan.runId), /ownership/);
  assert.throws(() => assertOwnedPlan(plan, runConfig({ ...env, E2E_LOCATION: "japaneast" }), plan.runId), /ownership/);
  assert.throws(() => assertOwnedPlan(plan, config, "00000000-0000-4000-8000-000000000099"), /ownership/);
});
