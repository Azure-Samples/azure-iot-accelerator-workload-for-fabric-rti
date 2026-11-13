import test from "node:test";
import assert from "node:assert/strict";
import { readdir, readFile, rmdir, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import { Cloud, object, redactSensitiveText } from "../src/cloud.js";
import { runConfig } from "../src/config.js";
import { armOperationUrl, associateNamespace, grantLinkingRole, linkingRoleDelay } from "../src/linking.js";
import { createPlan, writePlan, type RunPlan } from "../src/plan.js";

const config = runConfig({
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development",
});

test("linking grants require exact owned-scope principal and role readback before propagation", async () => {
  const plan = createPlan(config), scope = plan.azure.hubs[2]!.id;
  const principal = "00000000-0000-4000-8000-000000000004";
  class RoleCloud extends Cloud {
    writes = 0;
    changes: Record<string, unknown> = {};
    assignment: Record<string, unknown> = {};
    override async arm(path: string, version: string, method = "GET", body?: unknown): Promise<unknown> {
      assert.equal(version, "2022-04-01");
      if (method === "PUT") {
        this.writes++;
        this.assignment = { id: path, properties: { ...object(object(body, "grant body").properties, "grant properties"), scope } };
        return this.assignment;
      }
      assert.equal(method, "GET");
      return { ...this.assignment, properties: { ...object(this.assignment.properties, "role properties"), ...this.changes } };
    }
  }
  const cloud = new RoleCloud(config);
  const grant = await grantLinkingRole(cloud, plan, scope, principal, "b24988ac-6180-42a0-ab88-20f7382dd24c");
  assert.equal(cloud.writes, 1);
  assert.equal(grant.scope, scope);
  assert.equal(grant.principalId, principal);
  assert.equal(linkingRoleDelay([grant], grant.verifiedAt), 600_000);
  assert.equal(linkingRoleDelay([grant], grant.verifiedAt + 599_999), 1);
  assert.equal(linkingRoleDelay([grant], grant.verifiedAt + 600_000), 0);
  assert.equal(linkingRoleDelay([grant, { ...grant, verifiedAt: grant.verifiedAt + 5000 }], grant.verifiedAt), 605_000);
  assert.throws(() => linkingRoleDelay([]), /timestamps/);
  assert.throws(() => linkingRoleDelay([{ ...grant, verifiedAt: NaN }]), /timestamps/);
  for (const changes of [
    { scope: "/subscriptions/foreign" }, { principalId: "foreign" }, { roleDefinitionId: "foreign" },
    { principalType: "User" }, { condition: "conditional-access" },
  ]) {
    cloud.changes = changes;
    await assert.rejects(() => grantLinkingRole(cloud, plan, scope, principal, "b24988ac-6180-42a0-ab88-20f7382dd24c"), /readback/);
  }
  const before = cloud.writes;
  await assert.rejects(() => grantLinkingRole(cloud, plan, "/subscriptions/foreign/providers/Microsoft.Devices/IotHubs/foreign", principal, "role"), /outside/);
  assert.equal(cloud.writes, before, "Foreign resource scopes must fail before any mutation.");
});

class FixtureCloud extends Cloud {
  failed = false;
  endpointFailed = false;
  writes = 0;
  operation = `https://management.azure.com/subscriptions/${config.subscriptionId}/providers/Microsoft.DeviceRegistry/locations/australiaeast/operationStatuses/owned?api-version=2026-11-02-preview`;
  constructor(readonly plan: RunPlan) { super(config); }
  private namespace() {
    const namespace = this.plan.azure.namespaces[0]!, hub = this.plan.azure.hubs[2]!;
    return { id: namespace.id, tags: this.plan.azure.tags, properties: {
      provisioningState: this.failed ? "Failed" : "Succeeded",
      messaging: { endpoints: { [hub.name]: {
        resourceId: hub.id, linkingState: this.endpointFailed ? "Failed" : "Succeeded",
      } } },
    } };
  }
  override async send(_url: string, _scope: string, method = "GET"): Promise<Awaited<ReturnType<Cloud["send"]>>> {
    if (method === "PATCH") {
      this.writes++;
      return { status: 202, headers: new Headers({ "Azure-AsyncOperation": this.operation, "x-ms-request-id": "owned-request" }), body: {} };
    }
    return { status: 200, headers: new Headers({ ETag: '"specific"' }), body: {
      id: this.plan.azure.namespaces[0]!.id, tags: this.plan.azure.tags, properties: {},
    } };
  }
  override async arm(): Promise<unknown> { return this.namespace(); }
  override async request(url: string): Promise<unknown> {
    assert.equal(url, this.operation);
    return { status: "Failed", error: { code: "EndpointLinkFailed", message: "Denied SharedAccessKey=fake-test-only" } };
  }
}

test("native operation readback stays in the exact ARM subscription", () => {
  const plan = createPlan(config), cloud = new FixtureCloud(plan);
  assert.equal(armOperationUrl(cloud.operation, plan), cloud.operation);
  assert.throws(() => armOperationUrl("https://external.invalid/operations/id", plan), /outside/);
  assert.throws(() => armOperationUrl("https://management.azure.com/subscriptions/foreign/operations/id", plan), /outside/);
  assert.throws(() => armOperationUrl("https://user:password@management.azure.com/subscriptions/foreign", plan), /outside/);
});

for (const state of ["success", "namespace-failed", "endpoint-failed"] as const) {
  test(`namespace association ${state} preserves native operation evidence and dispatches once`, async context => {
    const plan = createPlan(config), manifest = await writePlan(plan), cloud = new FixtureCloud(plan);
    cloud.failed = state === "namespace-failed";
    cloud.endpointFailed = state === "endpoint-failed";
    if (cloud.failed) cloud.operation += "&s=synthetic-operation-signature&c=synthetic-operation-certificate";
    context.after(async () => {
      for (const name of await readdir(dirname(manifest))) await unlink(join(dirname(manifest), name));
      await rmdir(dirname(manifest));
    });
    const namespace = plan.azure.namespaces[0]!, hub = plan.azure.hubs[2]!;
    const action = associateNamespace(cloud, plan, namespace.id, "messaging", {
      [hub.name]: { resourceId: hub.id, endpointType: "Microsoft.Devices/IotHubs" },
    });
    if (state === "success") await action;
    else await assert.rejects(action, /Native association evidence/);
    assert.equal(cloud.writes, 1, "Failed writes must not be blindly retried.");
    const evidence = (await readdir(dirname(manifest))).find(name => name.startsWith("linking-"));
    assert.ok(evidence);
    const raw = await readFile(join(dirname(manifest), evidence), "utf8"), result = JSON.parse(raw);
    assert.equal(result.response.status, 202);
    assert.equal(result.response.operationUrl, redactSensitiveText(cloud.operation));
    assert.doesNotMatch(raw, /synthetic-operation-(signature|certificate)/);
    if (state !== "success") {
      assert.equal(result.operation.error.code, "EndpointLinkFailed");
      assert.doesNotMatch(raw, /fake-test-only/);
      assert.match(raw, /REDACTED/);
    }
  });
}

  test("a successful parent namespace cannot hide endpoint links still in progress", async context => {
    const plan = createPlan(config), manifest = await writePlan(plan);
    context.after(async () => {
      for (const name of await readdir(dirname(manifest))) await unlink(join(dirname(manifest), name));
      await rmdir(dirname(manifest));
    });
    class PendingCloud extends FixtureCloud {
      reads = 0;
      override async arm(): Promise<unknown> {
        const resource = object(await super.arm(), "namespace");
        const endpoints = object(object(object(resource.properties, "properties").messaging, "messaging").endpoints, "endpoints");
        object(Object.values(endpoints)[0], "endpoint").linkingState = ++this.reads === 1 ? "InProgress" : "Succeeded";
        return resource;
      }
    }
    const cloud = new PendingCloud(plan), hub = plan.azure.hubs[2]!;
    await associateNamespace(cloud, plan, plan.azure.namespaces[0]!.id, "messaging", { [hub.name]: { resourceId: hub.id } });
    assert.equal(cloud.reads, 2);
    assert.equal(cloud.writes, 1);
  });
