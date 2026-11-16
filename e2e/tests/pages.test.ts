import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { root } from "../src/config.js";
import * as ingestion from "./pages/ingestion.js";
import * as modeling from "./pages/modeling.js";
import * as dashboards from "./pages/dashboards.js";
import * as agents from "./pages/data-agents.js";

test("each wizard exposes one named unit per page, without a bulk workflow runner", () => {
  for (const [pages, names] of [
    [ingestion, ["overview", "iotHub", "eventhouse", "rawDataTables", "eventstreams", "iotHubRouting", "readyToUse"]],
    [modeling, ["overview", "iotDataSource", "deviceModel", "deviceModelAssignment", "modelCoverage", "modeledData", "readyToUse"]],
    [dashboards, ["overview", "dataSource", "generateDashboard", "readyToUse"]],
    [agents, ["overview", "modeledDataSource", "configureDataAgent", "readyToUse"]],
  ] as const) {
    assert.deepEqual(Object.keys(pages).sort(), [...names].sort());
    for (const unit of Object.values(pages)) assert.equal(typeof unit, "function");
  }
});

test("scenario specs show page steps directly instead of hiding UI in scenario runners", async () => {
  for (const [folder, files] of [
    ["ingestion", ["nonlinked", "combined", "linked-only"]],
    ["modeling", ["nonlinked-dtdl", "combined-wot", "linked-only-wot"]],
    ["dashboards", ["nonlinked", "combined", "linked-only"]],
    ["data-agents", ["nonlinked", "combined", "linked-only"]],
  ] as const) {
    for (const file of files) {
      const source = await readFile(join(root, "tests", "scenarios", folder, `${file}.spec.ts`), "utf8");
      assert.match(source, /await test\.step\(/);
      assert.doesNotMatch(source, /runModelingScenario|runExperienceScenario|ingestRegularHubs/);
      assert.doesNotMatch(source, /\.getByRole\(|\.locator\(|\.setInputFiles\(/);
    }
  }
});
