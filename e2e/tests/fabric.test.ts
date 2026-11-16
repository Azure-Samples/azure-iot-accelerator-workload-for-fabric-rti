import test from "node:test";
import assert from "node:assert/strict";
import { operationFromHeaders } from "../src/fabric.js";
import { reconcileFabricCreates } from "../src/fabric.js";
import { Cloud } from "../src/cloud.js";
import { runConfig } from "../src/config.js";
import { createPlan, writePlan } from "../src/plan.js";
import { readdir, rmdir, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";

test("Fabric LRO uses the public operation ID, not the regional redirect host", () => {
  const id = "00000000-0000-4000-8000-000000000001";
  const headers = new Headers({
    "x-ms-operation-id": id,
    Location: `https://wabi-us-central-b-primary-redirect.analysis.windows.net/v1/operations/${id}`,
  });
  assert.equal(operationFromHeaders(headers), `/operations/${id}`);
});

test("missing or malformed operation IDs fail instead of following arbitrary locations", () => {
  assert.throws(() => operationFromHeaders(new Headers({ Location: "https://example.com/operation" })), /operation ID/);
  assert.throws(() => operationFromHeaders(new Headers({ "x-ms-operation-id": "../../other" })), /operation ID/);
});

for (const state of ["Failed", "wrong-workspace", "wrong-name", "wrong-type", "changed-id"]) {
  test(`Fabric cleanup reconciles original operation ${state} without replaying creation`, async context => {
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
    const operation = "/operations/00000000-0000-4000-8000-000000000009";
    plan.observed.eventhouseIdOperation = operation;
    if (state === "changed-id") plan.observed.eventhouseId = "00000000-0000-4000-8000-000000000011";
    class OperationCloud extends Cloud {
      override async fabric(path: string, method = "GET"): Promise<unknown> {
        assert.equal(method, "GET");
        if (path === operation) return { status: state === "Failed" ? "Failed" : "Succeeded" };
        assert.equal(path, `${operation}/result`);
        return {
          id: "00000000-0000-4000-8000-000000000010",
          displayName: state === "wrong-name" ? "foreign" : plan.fabric.eventhouseName,
          type: state === "wrong-type" ? "Lakehouse" : "Eventhouse",
          workspaceId: state === "wrong-workspace" ? "foreign" : plan.target.workspaceId,
        };
      }
    }
    const action = reconcileFabricCreates(new OperationCloud(config), plan);
    if (state === "Failed") {
      await action;
      assert.equal(plan.observed.eventhouseIdOperation, undefined);
      assert.equal(plan.observed.eventhouseIdCreationOperationState, "Failed");
    } else {
      await assert.rejects(action, /changed its item identity/);
      assert.equal(plan.observed.eventhouseIdOperation, operation);
    }
  });
}
