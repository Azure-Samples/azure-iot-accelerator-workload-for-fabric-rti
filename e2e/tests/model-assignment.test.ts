import test from "node:test";
import assert from "node:assert/strict";
import { rmdir, unlink } from "node:fs/promises";
import { dirname } from "node:path";
import { Cloud, object } from "../src/cloud.js";
import { runConfig } from "../src/config.js";
import { expectedReference, initializeModelReference, otherModelReference, referenceRowsReady, verifyModelReferences } from "../src/model-assignment.js";
import { modelFixture } from "../src/models.js";
import { createPlan, writePlan, type RunPlan } from "../src/plan.js";

const config = runConfig({
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development",
});

class FixtureCloud extends Cloud {
  readonly devices = new Map<string, Record<string, unknown>>();
  readonly patches: { id: string; refs: unknown; etag: string | undefined }[] = [];
  constructor(readonly plan: RunPlan) {
    super(config);
    for (const hub of plan.azure.hubs.filter(value => value.namespaceKey)) {
      const namespace = plan.azure.namespaces.find(value => value.key === hub.namespaceKey);
      assert.ok(namespace);
      const id = `${namespace.id}/registryDevices/${hub.name}-device`;
      plan.observed[`${hub.key}.adrDeviceUuid`] = `uuid-${hub.key}`;
      this.devices.set(id, { id, etag: '"specific-etag"', properties: {
        externalDeviceId: `${hub.name}-device`, uuid: `uuid-${hub.key}`, deviceTypeRefs: [],
      } });
    }
  }
  override async arm(id: string, _version: string, method = "GET", body?: unknown, _missing = false, etag?: string): Promise<unknown> {
    const resource = this.devices.get(id);
    assert.ok(resource, "A test cannot touch an unplanned resource.");
    if (method === "PATCH") {
      const refs = object(object(body, "body").properties, "patch properties").deviceTypeRefs;
      this.patches.push({ id, refs, etag });
      object(resource.properties, "properties").deviceTypeRefs = refs;
    } else assert.equal(method, "GET");
    return resource;
  }
  override async request(_url: string, _scope: string, method = "GET", body?: unknown): Promise<unknown> {
    assert.equal(method, "POST");
    const query = object(body, "ARG request").query;
    assert.equal(typeof query, "string");
    return { data: [...this.devices.values()].filter(resource => String(query).includes(String(resource.id).toLowerCase()))
      .map(resource => ({ id: resource.id, refs: object(resource.properties, "properties").deviceTypeRefs })) };
  }
}

test("linked fixture starts unassigned/current/other and preserves the other model after assignment", () => {
  assert.deepEqual(expectedReference("linked1", false), []);
  assert.deepEqual(expectedReference("linked2", false), [modelFixture("combined").id]);
  assert.deepEqual(expectedReference("linked3", false), [otherModelReference]);
  assert.deepEqual(expectedReference("linked1", true), [modelFixture("combined").id]);
  assert.deepEqual(expectedReference("linked3", true), [otherModelReference]);
  assert.throws(() => expectedReference("regular1", false), /No linked model-reference fixture/);
});

test("model-reference scoping waits for actual observations, not just a successful ARM write", () => {
  const devices = ["linked1", "linked2", "linked3"].map(hubKey => ({ id: `${hubKey}-device`, hubKey }));
  const rows = devices.map(device => ({ deviceId: device.id, modelReference: expectedReference(device.hubKey, true)[0] }));
  assert.equal(referenceRowsReady(rows, devices), true);
  assert.equal(referenceRowsReady(rows.slice(0, 2), devices), false);
  assert.equal(referenceRowsReady([{ ...rows[0], modelReference: "" }, ...rows.slice(1)], devices), false);
  assert.equal(referenceRowsReady(rows, []), false);
});

test("initialize each owned reference once with specific ETags, then never reset the UI assignment", async context => {
  const plan = createPlan(config), manifest = await writePlan(plan), cloud = new FixtureCloud(plan);
  context.after(async () => { await unlink(manifest); await rmdir(dirname(manifest)); });
  for (const key of ["linked1", "linked2", "linked3"]) await initializeModelReference(cloud, plan, key);
  assert.equal(cloud.patches.length, 2);
  assert.ok(cloud.patches.every(value => value.etag === '"specific-etag"'));
  await verifyModelReferences(cloud, plan, false);
  for (const [id, resource] of cloud.devices) {
    if (id.includes("-linked1-device")) object(resource.properties, "properties").deviceTypeRefs = [modelFixture("combined").id];
  }
  await verifyModelReferences(cloud, plan, true);
  for (const key of ["linked1", "linked2", "linked3"]) await initializeModelReference(cloud, plan, key);
  assert.equal(cloud.patches.length, 2, "Subsequent phases must not reseed or reset model references.");
});

test("foreign references, uncertain writes and wildcard ETags block initialization before any mutation", async context => {
  const plan = createPlan(config), manifest = await writePlan(plan), cloud = new FixtureCloud(plan);
  context.after(async () => { await unlink(manifest); await rmdir(dirname(manifest)); });
  const resource = [...cloud.devices.values()][1];
  assert.ok(resource);
  object(resource.properties, "properties").deviceTypeRefs = ["urn:foreign"];
  await assert.rejects(initializeModelReference(cloud, plan, "linked2"), /Refusing to overwrite/);
  object(resource.properties, "properties").deviceTypeRefs = [];
  plan.observed["linked2.initialModelReference"] = "initializing";
  await assert.rejects(initializeModelReference(cloud, plan, "linked2"), /earlier reference initialization is uncertain/);
  delete plan.observed["linked2.initialModelReference"];
  resource.etag = "*";
  await assert.rejects(initializeModelReference(cloud, plan, "linked2"), /specific ETag/);
  assert.equal(cloud.patches.length, 0);
});
