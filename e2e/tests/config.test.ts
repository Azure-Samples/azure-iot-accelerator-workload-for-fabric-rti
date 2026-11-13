import test from "node:test";
import assert from "node:assert/strict";
import { authConfig } from "../src/config.js";

const tenant = "00000000-0000-4000-8000-000000000001";
const workspace = "00000000-0000-4000-8000-000000000002";
const args = ["--tenant", tenant, "--workspace", workspace];

test("configuration takes explicit IDs and adds only a tenant hint", () => {
  const config = authConfig(args, {});
  assert.equal(config.tenantId, tenant);
  assert.equal(config.workspaceUrl, `https://app.fabric.microsoft.com/groups/${workspace}?ctid=${tenant}`);
});

test("configuration accepts the user-selected Power BI workspace host", () => {
  const config = authConfig([...args, "--fabric-url", `https://app.powerbi.com/groups/${workspace}`], {});
  assert.equal(new URL(config.workspaceUrl).hostname, "app.powerbi.com");
});

test("missing and malformed IDs fail before opening a browser", () => {
  assert.throws(() => authConfig([], {}), /tenant.*GUID/);
  assert.throws(() => authConfig(["--tenant", tenant, "--workspace", "not-an-id"], {}), /workspace.*GUID/);
});

test("authentication can opt into a valid deployment target without a silent default", () => {
  assert.equal(authConfig(args, {}).deployment, undefined);
  assert.equal(authConfig([...args, "--deployment", "development"], {}).deployment, "development");
  assert.equal(authConfig(args, { E2E_DEPLOYMENT: "published" }).deployment, "published");
  assert.throws(() => authConfig(args, { E2E_DEPLOYMENT: "invalid" }), /E2E_DEPLOYMENT/);
});

test("configuration rejects another workspace, tenant, untrusted host and injected URL parameters", () => {
  for (const url of [
    `https://app.powerbi.com/groups/${tenant}`,
    `https://app.powerbi.com/groups/${workspace}?ctid=${workspace}`,
    `https://example.com/groups/${workspace}`,
    `https://app.powerbi.com/groups/${workspace}?redirect=https://example.com`,
    `https://user:password@app.powerbi.com/groups/${workspace}`,
  ]) assert.throws(() => authConfig([...args, "--fabric-url", url], {}));
});
