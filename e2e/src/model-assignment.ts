import assert from "node:assert/strict";
import { setTimeout as pause } from "node:timers/promises";
import { Cloud, object, text } from "./cloud.js";
import { modelFixture } from "./models.js";
import { updatePlan, type RunPlan } from "./plan.js";

const API = "2026-11-02-preview";
export const otherModelReference = "urn:example:E2E:OtherThermoProX200";

export function expectedReference(hubKey: string, assigned: boolean): string[] {
  if (hubKey === "linked1") return assigned ? [modelFixture("combined").id] : [];
  if (hubKey === "linked2") return [modelFixture("combined").id];
  if (hubKey === "linked3") return [otherModelReference];
  throw new Error(`[fixture] No linked model-reference fixture for ${hubKey}.`);
}

export function referenceRowsReady(rows: Record<string, unknown>[], devices: readonly { id: string; hubKey: string }[]): boolean {
  return devices.length === 3 && new Set(devices.map(device => device.id)).size === 3 &&
    ["linked1", "linked2", "linked3"].every(key => devices.some(device => device.hubKey === key)) &&
    devices.every(device => rows.some(row =>
    row.deviceId === device.id && row.modelReference === expectedReference(device.hubKey, true)[0]));
}

function deviceResource(plan: RunPlan, hubKey: string) {
  const hub = plan.azure.hubs.find(value => value.key === hubKey);
  const namespace = plan.azure.namespaces.find(value => value.key === hub?.namespaceKey);
  if (!hub || !namespace) throw new Error("[ownership] Model assignment requires an owned linked Hub and namespace.");
  return { hub, id: `${namespace.id}/registryDevices/${hub.name}-device` };
}

async function readReference(cloud: Cloud, plan: RunPlan, hubKey: string) {
  const { hub, id } = deviceResource(plan, hubKey);
  const resource = object(await cloud.arm(id, API), "owned registry device");
  assert.equal(text(resource.id, "registry device ID").toLowerCase(), id.toLowerCase());
  const properties = object(resource.properties, "registry device properties");
  assert.equal(properties.externalDeviceId, `${hub.name}-device`);
  assert.equal(properties.uuid, text(plan.observed[`${hubKey}.adrDeviceUuid`], "owned ADR device UUID"));
  const references = properties.deviceTypeRefs ?? [];
  assert.ok(Array.isArray(references) && references.every(value => typeof value === "string"), "[fixture] Invalid deviceTypeRefs.");
  return { id, resource, references };
}

export async function initializeModelReference(cloud: Cloud, plan: RunPlan, hubKey: string): Promise<void> {
  const marker = `${hubKey}.initialModelReference`;
  const expected = expectedReference(hubKey, plan.verification.modelAssignment === "confirmed");
  const current = await readReference(cloud, plan, hubKey);
  if (plan.observed[marker] === "initialized") {
    assert.deepEqual(current.references, expected, "[fixture] An initialized model reference changed; it will not be reset.");
    return;
  }
  if (plan.observed[marker]) throw new Error("[fixture] An earlier reference initialization is uncertain. Inspect the owned device before continuing.");
  assert.deepEqual(current.references, [], "[ownership] Refusing to overwrite an existing model reference while initializing the fixture.");
  plan.observed[marker] = "initializing";
  await updatePlan(plan);
  if (expected.length) {
    const etag = text(current.resource.etag, "registry device ETag");
    assert.notEqual(etag, "*", "[ownership] Reference initialization requires a specific ETag.");
    await cloud.arm(current.id, API, "PATCH", { properties: { deviceTypeRefs: expected } }, false, etag);
  }
  await waitForReferences(cloud, plan, [hubKey], false);
  plan.observed[marker] = "initialized";
  await updatePlan(plan);
}

async function waitForReferences(cloud: Cloud, plan: RunPlan, keys: string[], assigned: boolean): Promise<void> {
  const deadline = Date.now() + 180_000;
  while (Date.now() < deadline) {
    const current = await Promise.all(keys.map(key => readReference(cloud, plan, key)));
    if (current.every((value, i) => JSON.stringify(value.references) === JSON.stringify(expectedReference(keys[i]!, assigned)))) {
      const ids = current.map(value => value.id.toLowerCase());
      const inventory = object(await cloud.request("https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2022-10-01",
        "https://management.azure.com", "POST", {
          subscriptions: [plan.target.subscriptionId],
          query: `resources | where id in~ (${ids.map(id => JSON.stringify(id)).join(",")}) | project id, refs=properties.deviceTypeRefs`,
          options: { resultFormat: "objectArray", $top: 100 },
        }), "indexed model-reference readback");
      assert.ok(Array.isArray(inventory.data), "[service-shape] Resource Graph did not return reference rows.");
      const rows = inventory.data.map(value => object(value, "indexed device"));
      if (rows.length === keys.length && ids.every((id, i) => {
        const row = rows.find(value => typeof value.id === "string" && value.id.toLowerCase() === id);
        return row && JSON.stringify(row.refs ?? []) === JSON.stringify(expectedReference(keys[i]!, assigned));
      })) return;
    }
    await pause(3000);
  }
  throw new Error("[model-assignment] Native ADR model references did not reach the expected state.");
}

export async function verifyModelReferences(cloud: Cloud, plan: RunPlan, assigned: boolean): Promise<void> {
  await waitForReferences(cloud, plan, ["linked1", "linked2", "linked3"], assigned);
  if (assigned) {
    plan.verification.modelAssignment = "confirmed";
    await updatePlan(plan);
  }
}
