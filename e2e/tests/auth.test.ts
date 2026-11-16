import test from "node:test";
import assert from "node:assert/strict";
import type { BrowserContext } from "@playwright/test";
import { grantDevelopmentAccess } from "../src/auth.js";
import { authConfig } from "../src/config.js";

const args = [
  "--tenant", "00000000-0000-4000-8000-000000000001",
  "--workspace", "00000000-0000-4000-8000-000000000002",
];

function permissionRecorder() {
  const grants: { permissions: readonly string[]; origin: string | undefined }[] = [];
  const context: Pick<BrowserContext, "grantPermissions"> = {
    async grantPermissions(permissions, options) {
      grants.push({ permissions, origin: options?.origin });
    },
  };
  return { context, grants };
}

test("development grants local access only to the configured Fabric origin", async () => {
  for (const host of ["app.powerbi.com", "app.fabric.microsoft.com"]) {
    const config = authConfig([
      ...args, "--fabric-url", `https://${host}/groups/00000000-0000-4000-8000-000000000002`,
    ], { E2E_DEPLOYMENT: "development" });
    const { context, grants } = permissionRecorder();
    await grantDevelopmentAccess(context, config);
    assert.deepEqual(grants, [{
      permissions: ["local-network-access"],
      origin: `https://${host}`,
    }]);
  }
});

test("published and unspecified targets do not grant local-network permissions", async () => {
  for (const env of [{ E2E_DEPLOYMENT: "published" }, {}]) {
    const { context, grants } = permissionRecorder();
    await grantDevelopmentAccess(context, authConfig(args, env));
    assert.deepEqual(grants, []);
  }
});

test("unsupported permission errors propagate instead of silently disabling browser security", async () => {
  const failure = new Error("Permission is not supported");
  const context: Pick<BrowserContext, "grantPermissions"> = {
    async grantPermissions() { throw failure; },
  };
  await assert.rejects(
    grantDevelopmentAccess(context, authConfig(args, { E2E_DEPLOYMENT: "development" })),
    error => error === failure,
  );
});
