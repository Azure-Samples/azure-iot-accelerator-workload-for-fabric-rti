import test from "node:test";
import assert from "node:assert/strict";
import { readFile, rmdir, unlink } from "node:fs/promises";
import { dirname } from "node:path";
import { ActivatorAttempt } from "../src/activator.js";
import { Cloud } from "../src/cloud.js";
import { runConfig } from "../src/config.js";
import { assertOwnedPlan, createPlan, writePlan } from "../src/plan.js";

const env = {
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development", E2E_RETENTION: "cleanup",
};

class FixtureCloud extends Cloud {
  calls = 0;
  override async fabric(): Promise<unknown> { this.calls++; return { value: [] }; }
  override async request(): Promise<unknown> { this.calls++; return { value: [] }; }
}

test("keeping resources is an invocation policy, not a change to recorded ownership", () => {
  const config = runConfig(env), plan = createPlan(config);
  assert.equal(config.keepResources, false);
  const kept = runConfig({ ...env, E2E_KEEP_RESOURCES: "1" });
  assert.equal(kept.keepResources, true);
  assert.equal(kept.retention, "cleanup");
  assert.doesNotThrow(() => assertOwnedPlan(plan, kept, plan.runId));
  assert.equal(runConfig({ ...env, E2E_RETENTION: "retain" }).keepResources, true);
  assert.equal(runConfig({ ...env, E2E_KEEP_RESOURCES: "0" }).keepResources, false);
  assert.throws(() => runConfig({ ...env, E2E_KEEP_RESOURCES: "invalid" }), /E2E_KEEP_RESOURCES/);
});

for (const keep of [true, false]) {
  test(`Activator teardown ${keep ? "does not delete retained resources" : "still cleans up by default"}`, async context => {
    const config = runConfig({ ...env, E2E_KEEP_RESOURCES: keep ? "1" : "0" });
    const plan = createPlan(config), manifest = await writePlan(plan);
    const cloud = new FixtureCloud(config);
    let journal: string | undefined;
    context.after(async () => {
      if (journal) await unlink(journal);
      await unlink(manifest);
      await rmdir(dirname(manifest));
    });
    const attempt = await ActivatorAttempt.begin(cloud, plan, "regular");
    journal = attempt.file;
    cloud.calls = 0;
    const cleanup = context.mock.method(attempt, "cleanup", async () => {});
    await attempt.cleanupAfterScenario();
    assert.equal(cleanup.mock.callCount(), keep ? 0 : 1);
    assert.equal(cloud.calls, 0, "Retention must not issue native read/delete calls.");
    if (keep) {
      const saved = JSON.parse(await readFile(manifest, "utf8"));
      assert.equal(saved.verification["regular.activatorCleanup"], "retained-by-request");
      assert.ok(await readFile(journal, "utf8"), "The ownership journal must remain available for explicit cleanup.");
    }
    // Explicit cleanup remains available even when automatic cleanup is disabled.
    await attempt.cleanup();
    assert.equal(cleanup.mock.callCount(), keep ? 1 : 2);
  });
}
