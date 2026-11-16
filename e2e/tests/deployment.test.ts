import test from "node:test";
import assert from "node:assert/strict";
import { requireWorkloadOrigin } from "../src/deployment.js";

test("published workloads reject local development frames without a cross-target fallback", () => {
  for (const url of ["http://127.0.0.1:60006/item", "https://127.0.0.1/item", "https://localhost/item", "https://[::1]/item", "http://published.example/item"]) {
    assert.throws(() => requireWorkloadOrigin(url, { deployment: "published" }));
  }
  assert.doesNotThrow(() => requireWorkloadOrigin("https://published.example/item", { deployment: "published" }));
  assert.doesNotThrow(() => requireWorkloadOrigin("http://127.0.0.1:60006/item", { deployment: "development" }));
});
