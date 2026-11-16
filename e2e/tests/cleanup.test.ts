import test, { type TestContext } from "node:test";
import assert from "node:assert/strict";
import { readdir, rmdir, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import { Cloud } from "../src/cloud.js";
import { cleanupRun, verifyCleanupInventory } from "../src/cleanup.js";
import { runConfig } from "../src/config.js";
import { createPlan, readPlan, writePlan } from "../src/plan.js";
import { ingestionPath } from "../src/paths.js";
import { captureProcessOwner } from "../src/process-owner.js";
import { spawn } from "node:child_process";
import { stopOwnedChild } from "../src/owned-child.js";

const config = runConfig({
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development",
});
class FixtureCloud extends Cloud {
  readonly deletes: string[] = [];
  foreignGroup = false;
  override async arm(id: string, _version: string, method = "GET"): Promise<unknown> {
    if (method === "DELETE") { this.deletes.push(id); return undefined; }
    return this.foreignGroup && /\/resourceGroups\/[^/]+$/.test(id) ? { tags: { "e2e-run-id": "foreign-run" } } : undefined;
  }
  override async fabric(path: string, method = "GET"): Promise<unknown> {
    if (method === "DELETE") { this.deletes.push(path); return undefined; }
    return { value: [] };
  }
}

async function fixturePlan(context: TestContext) {
  const plan = createPlan(config), manifest = await writePlan(plan);
  context.after(async () => {
    for (const file of await readdir(dirname(manifest))) await unlink(join(dirname(manifest), file));
    await rmdir(dirname(manifest));
  });
  return plan;
}

test("Fabric inventory failure is recorded while independent Azure deletion still runs", async context => {
  const plan = await fixturePlan(context);
  class InventoryFailure extends FixtureCloud {
    groupPresent = true;
    listings = 0;
    override async arm(id: string, version: string, method = "GET"): Promise<unknown> {
      if (id === plan.azure.resourceGroupId) {
        if (method === "DELETE") {
          this.deletes.push(id);
          this.groupPresent = false;
          return undefined;
        }
        return this.groupPresent ? { tags: plan.azure.tags } : undefined;
      }
      if (id === `${plan.azure.resourceGroupId}/resources`) return { value: [] };
      return super.arm(id, version, method);
    }
    override async fabric(path: string, method = "GET"): Promise<unknown> {
      if (path.endsWith("/items") && this.listings++ === 0) throw new Error("Synthetic Fabric listing failure");
      return super.fabric(path, method);
    }
  }
  const cloud = new InventoryFailure(config);
  await assert.rejects(cleanupRun(cloud, plan, async () => {}), /Run-owned resources remain/);
  assert.ok(cloud.deletes.includes(plan.azure.resourceGroupId));
  assert.equal(cloud.groupPresent, false);
  assert.equal(plan.phase, "cleanup-failed");
  assert.equal(plan.observed.cleanupState, "failed");
});

test("workspace role recovery uses the recorded owned principal after its Hub is gone", async context => {
  const plan = await fixturePlan(context), hub = plan.azure.hubs[0]!;
  const principal = "00000000-0000-4000-8000-000000000009";
  const role = "00000000-0000-4000-8000-000000000010";
  plan.observed[hub.id] = "native-hub-provisioning-verified";
  plan.observed.regular1IdentityResourceId = hub.id;
  plan.observed.regular1IdentityPrincipalId = principal;
  class RoleRecovery extends FixtureCloud {
    rolePresent = true;
    override async fabric(path: string, method = "GET"): Promise<unknown> {
      if (path.endsWith(`/roleAssignments/${role}`) && method === "DELETE") {
        this.deletes.push(path);
        this.rolePresent = false;
        return undefined;
      }
      if (path.endsWith("/roleAssignments")) return { value: this.rolePresent ? [{ id: role, principal: { id: principal } }] : [] };
      return super.fabric(path, method);
    }
  }
  const cloud = new RoleRecovery(config);
  await cleanupRun(cloud, plan, async () => {});
  assert.equal(cloud.rolePresent, false);
  assert.equal(cloud.deletes.length, 1);
  await cleanupRun(cloud, plan, async () => {});
  assert.equal(cloud.deletes.length, 1);
});

test("a recorded principal cannot be rebound to a different resource", async context => {
  const plan = await fixturePlan(context), cloud = new FixtureCloud(config);
  plan.observed.regular1IdentityResourceId = plan.azure.hubs[1]!.id;
  plan.observed.regular1IdentityPrincipalId = "00000000-0000-4000-8000-000000000009";
  await assert.rejects(cleanupRun(cloud, plan, async () => {}), /Run-owned resources remain/);
  assert.deepEqual(cloud.deletes, []);
  assert.equal(plan.observed.cleanupState, "failed");
});

test("principal proof survives a failed role removal and resource-group deletion", async context => {
  const plan = await fixturePlan(context), hub = plan.azure.hubs[0]!;
  const principal = "00000000-0000-4000-8000-000000000009", roleId = "00000000-0000-4000-8000-000000000010";
  class RetryCloud extends FixtureCloud {
    groupPresent = true;
    rolePresent = true;
    rejectRoles = true;
    override async arm(id: string, version: string, method = "GET"): Promise<unknown> {
      if (id === hub.id) return this.groupPresent ? { tags: plan.azure.tags, identity: { principalId: principal } } : undefined;
      if (id === plan.azure.resourceGroupId) {
        if (method === "DELETE") { this.deletes.push(id); this.groupPresent = false; return undefined; }
        return this.groupPresent ? { tags: plan.azure.tags } : undefined;
      }
      if (id === `${plan.azure.resourceGroupId}/resources`) return { value: [{ id: hub.id, tags: plan.azure.tags }] };
      return super.arm(id, version, method);
    }
    override async fabric(path: string, method = "GET"): Promise<unknown> {
      if (path.endsWith("/roleAssignments")) {
        if (this.rejectRoles) throw new Error("Synthetic role-removal failure");
        return { value: this.rolePresent ? [{ id: roleId, principal: { id: principal } }] : [] };
      }
      if (path.endsWith(`/roleAssignments/${roleId}`) && method === "DELETE") {
        this.deletes.push(path);
        this.rolePresent = false;
        return undefined;
      }
      return super.fabric(path, method);
    }
  }
  const cloud = new RetryCloud(config);
  await assert.rejects(cleanupRun(cloud, plan, async () => {}), /Run-owned resources remain/);
  assert.equal(cloud.groupPresent, false);
  const recovered = await readPlan(config, plan.runId);
  assert.equal(recovered.observed.regular1IdentityPrincipalId, principal);
  cloud.rejectRoles = false;
  await cleanupRun(cloud, recovered, async () => {});
  assert.equal(cloud.rolePresent, false);
  assert.equal(recovered.observed.cleanupState, "confirmed");
});

test("cleanup refuses foreign tags even on an exact planned Azure resource ID", async () => {
  const plan = createPlan(config), cloud = new FixtureCloud(config);
  await verifyCleanupInventory(cloud, plan, [{ id: plan.azure.hubs[0]!.id, tags: plan.azure.tags }]);
  await assert.rejects(verifyCleanupInventory(cloud, plan, [{
    id: plan.azure.hubs[0]!.id, tags: { ...plan.azure.tags, "e2e-run-id": "foreign-run" },
  }]), /foreign ownership tags/);
  assert.deepEqual(cloud.deletes, []);
});

for (const state of ["Running", "Succeeded"] as const) {
  test(`cleanup reconciles an unlisted original Fabric create that is ${state}`, async context => {
    const plan = await fixturePlan(context), operation = "/operations/00000000-0000-4000-8000-000000000009";
    const id = "00000000-0000-4000-8000-000000000010";
    plan.observed.eventhouseIdOperation = operation;
    class CreatingCloud extends FixtureCloud {
      itemPresent = true;
      override async fabric(path: string, method = "GET"): Promise<unknown> {
        if (path === operation) return { status: state };
        if (path === `${operation}/result`) return {
          id, displayName: plan.fabric.eventhouseName, type: "Eventhouse", workspaceId: plan.target.workspaceId,
        };
        return super.fabric(path, method);
      }
      override async request(url: string, _resource: string, method = "GET"): Promise<unknown> {
        assert.ok(url.endsWith(`/items/${id}`));
        if (method === "DELETE") { this.deletes.push(url); this.itemPresent = false; return undefined; }
        return this.itemPresent ? { id, displayName: plan.fabric.eventhouseName, type: "Eventhouse" } : undefined;
      }
    }
    const cloud = new CreatingCloud(config);
    if (state === "Running") {
      await assert.rejects(cleanupRun(cloud, plan, async () => {}), /Run-owned resources remain/);
      assert.equal(plan.observed.eventhouseIdOperation, operation);
      assert.equal(plan.observed.cleanupState, "failed");
      assert.deepEqual(cloud.deletes, []);
    } else {
      await cleanupRun(cloud, plan, async () => {});
      assert.equal(cloud.itemPresent, false);
      assert.equal(plan.observed.eventhouseIdOperation, undefined);
      assert.equal(plan.observed.eventhouseIdCreationOperation, operation);
      assert.equal(plan.observed.cleanupState, "confirmed");
    }
  });
}

for (const state of ["owned-root", "foreign-database", "foreign-folder"] as const) {
  test(`interrupted raw Eventstream cleanup ${state}`, async context => {
    const plan = createPlan(config), manifest = await writePlan(plan), path = ingestionPath(plan, "combined");
    plan.observed.folderId = "00000000-0000-4000-8000-000000000008";
    plan.observed.databaseId = "00000000-0000-4000-8000-000000000009";
    const id = "00000000-0000-4000-8000-000000000010";
    context.after(async () => {
      for (const file of await readdir(dirname(manifest))) await unlink(join(dirname(manifest), file));
      await rmdir(dirname(manifest));
    });
    const item = { id, displayName: path.propertiesStream, type: "Eventstream",
      folderId: state === "foreign-folder" ? "00000000-0000-4000-8000-000000000011" : null };
    class RootStreamCloud extends FixtureCloud {
      removed = false;
      override async fabric(resource: string, method = "GET"): Promise<unknown> {
        if (resource.includes("/eventstreams/")) return { destinations: [{ type: "Eventhouse", properties: {
          workspaceId: config.workspaceId, itemId: state === "foreign-database" ? "foreign" : plan.observed.databaseId,
          tableName: path.propertiesTable, dataIngestionMode: "ProcessedIngestion",
        } }] };
        if (resource.endsWith("/items")) return { value: this.removed ? [] : [item] };
        return super.fabric(resource, method);
      }
      override async request(url: string, _resource: string, method = "GET"): Promise<unknown> {
        if (url.endsWith(`/items/${id}`)) {
          if (method === "DELETE") { this.deletes.push(url); this.removed = true; return undefined; }
          return this.removed ? undefined : item;
        }
        if (url.includes("/folders/")) return undefined;
        if (url.endsWith(`/items/${plan.observed.databaseId}`)) return undefined;
        throw new Error("Unplanned resource read.");
      }
    }
    const cloud = new RootStreamCloud(config);
    if (state === "owned-root") await cleanupRun(cloud, plan, async () => {});
    else await assert.rejects(() => cleanupRun(cloud, plan, async () => {}), /Run-owned resources remain/);
    assert.equal(cloud.deletes.length, state === "owned-root" ? 1 : 0);
    assert.equal(plan.observed.cleanupState, state === "owned-root" ? "confirmed" : "failed");
  });
}

for (const foreignGroup of [false, true]) {
  test(`full cleanup ${foreignGroup ? "refuses a foreign-owned group and records failure" : "certifies absent owned resources without deleting the workspace"}`, async context => {
    const plan = createPlan(config), manifest = await writePlan(plan), cloud = new FixtureCloud(config);
    cloud.foreignGroup = foreignGroup;
    context.after(async () => {
      for (const file of await readdir(dirname(manifest))) await unlink(join(dirname(manifest), file));
      await rmdir(dirname(manifest));
    });
    let reconciled = false;
    const action = cleanupRun(cloud, plan, async () => { reconciled = true; });
    if (foreignGroup) await assert.rejects(action, /Run-owned resources remain/);
    else await action;
    assert.equal(reconciled, true);
    assert.deepEqual(cloud.deletes, []);
    assert.equal(plan.observed.cleanupState, foreignGroup ? "failed" : "confirmed");
    const receipts = (await readdir(dirname(manifest))).filter(file => /^cleanup-\d+\.json$/.test(file));
    assert.equal(receipts.length, 1);
  });
}

test("full cleanup refuses running devices before any native resource access", async context => {
  const plan = await fixturePlan(context), cloud = new FixtureCloud(config);
  plan.observed.simulationState = "running-in-process";
  plan.observed.cleanupState = "confirmed";
  await assert.rejects(cleanupRun(cloud, plan, async () => assert.fail("Must not start cleanup")), /Stop the run's local devices/);
  assert.deepEqual(cloud.deletes, []);
  assert.equal(plan.observed.cleanupState, "failed");
});

test("cleanup recovers a dead simulator generation but never clears a live owner", async context => {
  const plan = await fixturePlan(context), cloud = new FixtureCloud(config);
  const owner = await captureProcessOwner(plan.runId);
  assert.ok(owner);
  plan.observed.simulationState = "running-in-process";
  plan.observed.simulationOwner = JSON.stringify(owner);
  await assert.rejects(cleanupRun(cloud, plan, async () => {}), /still alive/);
  assert.equal(plan.observed.simulationState, "running-in-process");
  plan.observed.simulationOwner = JSON.stringify({ ...owner, stamp: `${owner.stamp}-previous-generation` });
  await cleanupRun(cloud, plan, async () => {});
  assert.equal(plan.observed.simulationState, "stopped");
  assert.equal(plan.observed.simulationStopReason, "recorded-owner-process-exited");
});

test("manual cleanup succeeds after the actual recorded simulator worker is terminated", async context => {
  const plan = await fixturePlan(context), cloud = new FixtureCloud(config);
  const worker = spawn(process.execPath, ["-e", "setInterval(()=>{},1000)"], {
    detached: true, windowsHide: true, stdio: "ignore",
  });
  assert.ok(worker.pid);
  const owner = await captureProcessOwner(plan.runId, worker.pid);
  assert.ok(owner);
  let workerStopped = false;
  context.after(async () => { if (!workerStopped) await stopOwnedChild(worker, owner); });
  plan.observed.simulationState = "running-in-process";
  plan.observed.simulationOwner = JSON.stringify(owner);
  await assert.rejects(cleanupRun(cloud, plan, async () => {}), /still alive/);
  await stopOwnedChild(worker, owner);
  workerStopped = true;
  await cleanupRun(cloud, plan, async () => {});
  assert.equal(plan.observed.simulationState, "stopped");
  assert.equal(plan.observed.cleanupState, "confirmed");
});

test("cleanup accepts only the exact manifest-derived ADR devices with native identity and Hub binding proof", async () => {
  const plan = createPlan(config), hub = plan.azure.hubs.find(value => value.key === "linked1");
  assert.ok(hub);
  const namespace = plan.azure.namespaces.find(value => value.key === hub.namespaceKey);
  assert.ok(namespace);
  const id = `${namespace.id}/registryDevices/${hub.name}-device`;
  plan.observed["linked1.adrDeviceUuid"] = "owned-device-uuid";
  class RegistryCloud extends FixtureCloud {
    wrongIdentity = false;
    wrongHub = false;
    override async arm(resource: string): Promise<unknown> {
      if (resource === namespace!.id) return { tags: plan.azure.tags };
      if (resource === id) return { properties: {
        uuid: this.wrongIdentity ? "foreign-uuid" : "owned-device-uuid", externalDeviceId: `${hub!.name}-device`,
      } };
      if (resource === `${id}/capabilities/${hub!.name}`) return { properties: {
        messagingEndpointName: this.wrongHub ? "foreign-hub" : hub!.name,
        capabilityType: "Microsoft.IoTHub", provisioningState: "Succeeded",
      } };
      throw new Error("Unplanned resource read.");
    }
  }
  const cloud = new RegistryCloud(config);
  await verifyCleanupInventory(cloud, plan, [{ id }]);
  await assert.rejects(verifyCleanupInventory(cloud, plan, [{ id: `${namespace.id}/registryDevices/foreign-device` }]), /Unexpected resource/);
  cloud.wrongIdentity = true;
  await assert.rejects(verifyCleanupInventory(cloud, plan, [{ id }]), /exact simulated identity/);
  cloud.wrongIdentity = false;
  cloud.wrongHub = true;
  await assert.rejects(verifyCleanupInventory(cloud, plan, [{ id }]), /planned Hub capability/);
  assert.deepEqual(cloud.deletes, []);
});
