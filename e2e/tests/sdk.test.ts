import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { root } from "../src/config.js";
import { loadSimulationSdk, withSimulation } from "../src/data.js";

test("CommonJS IoT SDK loads through native ESM without opening a device connection", () => {
  const sdk = loadSimulationSdk();
  const message = new sdk.Message('{"temperature":23.75}');
  assert.equal(message.getData().toString(), '{"temperature":23.75}');
  assert.equal(typeof sdk.Client.fromConnectionString, "function");
  assert.equal(typeof sdk.MqttWs, "function");
  assert.equal(typeof withSimulation, "function");
});

test("collecting simulation helpers does not eagerly load either CommonJS IoT SDK", () => {
  const result = spawnSync(process.execPath, ["--import", "tsx", "--input-type=module", "-e", `
    import assert from "node:assert/strict";
    import { createRequire } from "node:module";
    const require = createRequire(import.meta.url);
    await import("./src/data.ts");
    for (const name of ["azure-iot-device", "azure-iot-device-mqtt"]) {
      assert.equal(require.cache[require.resolve(name)], undefined, name);
    }
  `], { cwd: root, encoding: "utf8", timeout: 30_000 });
  assert.ifError(result.error);
  assert.equal(result.status, 0, result.stderr);
});
