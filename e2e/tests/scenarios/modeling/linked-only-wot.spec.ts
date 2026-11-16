import { test } from "@playwright/test";
import { Cloud } from "../../../src/cloud.js";
import { requireGuid, runConfig } from "../../../src/config.js";
import { withSimulation } from "../../../src/data.js";
import { verifyModelReferences } from "../../../src/model-assignment.js";
import { modelFixture } from "../../../src/models.js";
import { modeledNames, verifyIncomingReferences, verifyModeled } from "../../../src/modeling.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan, updatePlan } from "../../../src/plan.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as modeling from "../../pages/modeling.js";

// Upload ThermoProX200-wot-tm.json against only the three linked Hubs' raw tables.
test(scenarioTitle("linkedModel"), async ({}, info) => {
  test.setTimeout(30 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  if (!plan.verification["linked.rawData"]?.startsWith("native-kusto")) throw new Error("[prerequisite] Complete linked-only ingestion first.");
  if (plan.verification.modelAssignment !== "confirmed") throw new Error("[prerequisite] Complete combined model assignment before checking the linked-only already-assigned path.");
  const cloud = new Cloud(config), path = ingestionPath(plan, "linked"), fixture = modelFixture("linked");
  const names = modeledNames(path);
  await recordedWorkflow(config, info, plan.runId, "linked-modeling", async page => {
    await withSimulation(cloud, plan, async fleet => {
      await test.step("3.3 Native references retained from combined assignment", () => verifyModelReferences(cloud, plan, true));
      const workload = await test.step("3.3.0 Solution home", () => solutionHome(page, plan, "Model your device data"));
      await test.step("3.3.1 Overview", () => modeling.overview(workload));
      await test.step("3.3.2 IoT Data Source", () => modeling.iotDataSource(workload, plan, path));
      await test.step("3.3.3 Device Model", () => modeling.deviceModel(workload, fixture));
      await test.step("3.3.4 All devices already assigned", () => modeling.deviceModelAssignment(workload, fixture, path.id, plan));
      await test.step("3.3 Fresh referenced telemetry and properties", () => verifyIncomingReferences(cloud, plan, path, fleet));
      await test.step("3.3.5 Model reference OR property condition", () => modeling.modelCoverage(workload, fixture, path.id));
      await test.step("3.3.6 Modeled Data", () => modeling.modeledData(workload, names));
      await test.step("3.3.7 Ready to Use", () => modeling.readyToUse(workload, names));
      plan.observed["linked.modelCreatedAt"] = new Date().toISOString();
      await updatePlan(plan);
      await test.step("3.3.8 Native outputs, enrichment and exclusion", () => verifyModeled(cloud, plan, path, fleet));
    }, true);
  });
});
