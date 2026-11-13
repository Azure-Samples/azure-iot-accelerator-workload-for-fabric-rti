import test from "node:test";
import assert from "node:assert/strict";
import { readdir, readFile, rmdir, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import { Cloud, object, text } from "../src/cloud.js";
import { runConfig } from "../src/config.js";
import type { DeviceFleet } from "../src/data.js";
import { modeledNames, verifyScopedData } from "../src/modeling.js";
import { modelFixture } from "../src/models.js";
import { ingestionPath } from "../src/paths.js";
import { createPlan, writePlan } from "../src/plan.js";

test("scope windows close on raw native observation time after both cohorts arrive, not on local publish acknowledgements", async context => {
  const config = runConfig({
    E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
    E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
    E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
    E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "published",
  });
  const plan = createPlan(config), manifest = await writePlan(plan), path = ingestionPath(plan, "linked");
  plan.observed.queryServiceUri = "https://owned.kusto.fabric.microsoft.com";
  context.after(async () => {
    for (const file of await readdir(dirname(manifest))) await unlink(join(dirname(manifest), file));
    await rmdir(dirname(manifest));
  });
  const clock = Date.now() + 60_000, nativeEnd = new Date(clock + 3000).toISOString().replace(/\.\d{3}Z$/, ".9876543Z");
  const reference = modelFixture("linked").id;
  const positive = { deviceId: "linked1", event_time: new Date(clock).toISOString(), reference, sequence: 13 };
  const negativeEven = { deviceId: "linked3", event_time: new Date(clock + 1000).toISOString(), reference: "other-model", sequence: 12 };
  const lateNegativeOdd = { deviceId: "linked3", event_time: nativeEnd, reference: "other-model", sequence: 13 };
  const telemetry = [positive, negativeEven], twin = [positive, negativeEven, lateNegativeOdd];
  const output = (rows: Record<string, unknown>[]) => rows.map(row => ({ deviceId: row.deviceId, event_time: row.event_time, rows: 1 }));
  const result = (rows: Record<string, unknown>[]) => {
    const keys = Object.keys(rows[0]!);
    return { Tables: [{ Columns: keys.map(ColumnName => ({ ColumnName })), Rows: rows.map(row => keys.map(key => row[key])) }] };
  };
  const calls: string[] = [];
  class ScopeCloud extends Cloud {
    override async request(_url: string, _resource: string, method = "GET", body?: unknown): Promise<unknown> {
      assert.equal(method, "POST");
      const csl = text(object(body, "query").csl, "query text");
      calls.push(csl);
      if (csl.startsWith("union ")) {
        assert.ok(!calls.slice(0, -1).some(call => call.startsWith(modeledNames(path).modeled) || call.startsWith(modeledNames(path).normalized)),
          "Never choose the window from modeled output values.");
        return result([{ nativeEnd }]);
      }
      if (csl.includes("between")) assert.ok(csl.includes(`datetime(${nativeEnd})`), "Do not truncate the late negative cohort at the host ACK time.");
      if (csl.startsWith(path.telemetryTable)) return result(telemetry);
      if (csl.startsWith(path.propertiesTable)) return result(twin);
      if (csl.startsWith(modeledNames(path).modeled)) return result(output([positive]));
      if (csl.startsWith(modeledNames(path).normalized)) return result(output([positive, negativeEven]));
      throw new Error("Unplanned Kusto query.");
    }
  }
  let rounds = 0;
  const fleet: DeviceFleet = {
    attempt: "current-attempt", devices: [], assertHealthy() {},
    async sendRound() { rounds++; }, async reportProperties() {},
  };
  await verifyScopedData(new ScopeCloud(config), plan, path, fleet);
  assert.equal(rounds, 2);
  const evidence = JSON.parse(await readFile(join(dirname(manifest), "scope-evidence-linked-current-attempt.json"), "utf8"));
  assert.equal(evidence.windowClock, "native-observation");
  assert.equal(evidence.end, nativeEnd);
  assert.ok(Date.parse(evidence.end) > Date.now(), "The fixture's observation time intentionally follows the immediate local ACK.");
  assert.equal(evidence.twin.raw.length, 3);
  assert.equal(evidence.twin.output.length, 2);
});
