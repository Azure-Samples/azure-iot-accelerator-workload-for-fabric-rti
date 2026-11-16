import test from "node:test";
import assert from "node:assert/strict";
import { AzureCliCredential } from "@azure/identity";
import { Cloud, redactSensitiveText, safeServiceMessage } from "../src/cloud.js";
import { runConfig } from "../src/config.js";

test("credential failures retain the original error and record only safe same-process availability diagnostics", async context => {
  const failure = new Error("Synthetic credential failure"), diagnostics: string[] = [];
  context.mock.method(AzureCliCredential.prototype, "getToken", async () => { throw failure; });
  context.mock.method(console, "error", (message: string) => { diagnostics.push(message); });
  const config = runConfig({
    E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
    E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
    E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
    E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "published",
  });
  await assert.rejects(() => new Cloud(config).token("https://management.azure.com"), error => error === failure);
  assert.equal(diagnostics.length, 1);
  const record = JSON.parse(diagnostics[0]!.slice("[authentication-process] ".length));
  assert.equal(typeof record.azureCliFileOnPath, "boolean");
  assert.equal(typeof record.cwdExists, "boolean");
  assert.ok(Array.isArray(record.pathKeys));
  assert.ok(!("PATH" in record) && !("Path" in record) && !("token" in record));
});

test("native service validation messages remain actionable without credentials", () => {
  const message = safeServiceMessage('Invalid sku GEN2; primaryKey="fake-key-value"; SharedAccessKey=fake-connection-key; Bearer fake-bearer-value');
  assert.match(message, /Invalid sku GEN2/);
  for (const secret of ["fake-key-value", "fake-connection-key", "fake-bearer-value"]) assert.ok(!message.includes(secret));
  assert.match(message, /REDACTED/);
});

test("native operation diagnostics redact signed query parameters without losing the API version or operation path", () => {
  const url = "https://management.azure.com/subscriptions/owned/asyncOperationStatuses/example?api-version=2026-11-02-preview&t=synthetic-time&c=synthetic-certificate&s=synthetic-signature&h=synthetic-hash";
  const result = redactSensitiveText(JSON.stringify({ operationUrl: url }));
  assert.ok(result.includes("/asyncOperationStatuses/example?api-version=2026-11-02-preview"));
  for (const value of ["synthetic-time", "synthetic-certificate", "synthetic-signature", "synthetic-hash"]) {
    assert.ok(!result.includes(value));
  }
  assert.equal((result.match(/\[REDACTED\]/g) ?? []).length, 4);
});
