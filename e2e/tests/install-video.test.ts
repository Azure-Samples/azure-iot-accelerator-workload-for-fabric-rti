import test from "node:test";
import assert from "node:assert/strict";
import childProcess from "node:child_process";
import { syncBuiltinESMExports } from "node:module";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { root } from "../src/config.js";

for (const configured of [undefined, "custom-browser-cache", "0"]) {
  test(`video installation respects browser cache ${configured ?? "default"}`, async context => {
    const previous = process.env.PLAYWRIGHT_BROWSERS_PATH;
    const exitCode = process.exitCode;
    if (configured === undefined) delete process.env.PLAYWRIGHT_BROWSERS_PATH;
    else process.env.PLAYWRIGHT_BROWSERS_PATH = configured;
    const spawn = context.mock.method(childProcess, "spawnSync", () => ({
      pid: 0, output: [], stdout: Buffer.alloc(0), stderr: Buffer.alloc(0), status: 0, signal: null,
    }));
    syncBuiltinESMExports();
    context.after(() => {
      context.mock.restoreAll();
      syncBuiltinESMExports();
      if (previous === undefined) delete process.env.PLAYWRIGHT_BROWSERS_PATH;
      else process.env.PLAYWRIGHT_BROWSERS_PATH = previous;
      process.exitCode = exitCode;
    });
    await import(`${pathToFileURL(join(root, "scripts", "install-video.mjs")).href}?cache=${configured ?? "default"}`);
    assert.equal(spawn.mock.callCount(), 1);
    const options = spawn.mock.calls[0]?.arguments[2];
    assert.equal(options?.env?.PLAYWRIGHT_BROWSERS_PATH, configured ?? join(root, "artifacts", "browsers"));
  });
}
