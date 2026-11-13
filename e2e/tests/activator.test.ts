import test from "node:test";
import assert from "node:assert/strict";
import { readdir, rmdir, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import { ActivatorAttempt, activatorInputSchema } from "../src/activator.js";
import { Cloud } from "../src/cloud.js";
import { runConfig } from "../src/config.js";
import { createPlan, writePlan } from "../src/plan.js";

const destination = { id: "owned-destination", type: "Activator", properties: { workspaceId: "owned-workspace", itemId: "owned-activator" } };
const input = { name: "owned-stream-stream", schema: { columns: [{ name: "deviceId", type: "Nvarchar(max)" }] } };
const definition = (overrides: Record<string, unknown> = {}) => [
  { path: "eventstream.json", value: { destinations: [{ ...destination, inputSchemas: [input], ...overrides }] } },
];

test("Activator schemas come from the native authoring definition, not the simplified runtime topology", () => {
  assert.equal(activatorInputSchema(definition(), destination), input);
  assert.ok(!("inputSchemas" in destination));
});

test("Activator authoring schema must match the exact runtime destination and target", () => {
  for (const overrides of [
    { id: "another-destination" }, { type: "Eventhouse" },
    { properties: { workspaceId: "another-workspace", itemId: "owned-activator" } },
    { properties: { workspaceId: "owned-workspace", itemId: "another-activator" } },
    { inputSchemas: [] }, { inputSchemas: [input, input] },
  ]) assert.throws(() => activatorInputSchema(definition(overrides), destination));
  assert.throws(() => activatorInputSchema([], destination));
});

test("a confirmed deletion is terminal even when the original create response was lost", async context => {
  const config = runConfig({
    E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
    E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
    E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
    E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development",
  });
  const plan = createPlan(config), manifest = await writePlan(plan);
  context.after(async () => {
    for (const file of await readdir(dirname(manifest))) await unlink(join(dirname(manifest), file));
    await rmdir(dirname(manifest));
  });
  class DeletedCloud extends Cloud {
    items: Record<string, unknown>[] = [];
    override async fabric(): Promise<unknown> { return { value: this.items }; }
    override async request(url: string): Promise<unknown> {
      return url.endsWith("/v1/connections") ? { value: [] } : undefined;
    }
  }
  const cloud = new DeletedCloud(config), attempt = await ActivatorAttempt.begin(cloud, plan, "regular");
  const resource = attempt.journal.resources[0]!;
  resource.requested = true;
  resource.id = "00000000-0000-4000-8000-000000000009";
  resource.deleted = true;
  await attempt.cleanup();
  await ActivatorAttempt.cleanupJournals(cloud, plan);
  assert.equal(plan.verification["regular.activatorCleanup"], "all-attempt-resources-confirmed-absent");
  resource.deleted = false;
  await assert.rejects(attempt.reconcile(), /no native response or resource readback/);
  resource.deleted = true;
  cloud.items = [{ id: "00000000-0000-4000-8000-000000000010", displayName: resource.name, type: resource.kind }];
  await assert.rejects(attempt.reconcile(), /reappeared/);
});
