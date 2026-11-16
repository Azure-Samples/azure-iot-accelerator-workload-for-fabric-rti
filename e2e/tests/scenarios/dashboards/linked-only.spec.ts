import { test } from "@playwright/test";
import { Cloud } from "../../../src/cloud.js";
import { requireGuid, runConfig } from "../../../src/config.js";
import { verifyDashboard } from "../../../src/experience-verification.js";
import { verifyArrivingData, withSimulation } from "../../../src/data.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan } from "../../../src/plan.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as dashboard from "../../pages/dashboards.js";
import * as native from "../../pages/native-experiences.js";

// Bind generated visuals only to the independently verified linked-only raw table pair.
test(scenarioTitle("linkedDashboard"), async ({}, info) => {
  test.setTimeout(15 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  if (!plan.verification["linked.modeling"]?.startsWith("native-model")) throw new Error("[prerequisite] Complete WoT Linked-only modeling first.");
  const cloud = new Cloud(config), path = ingestionPath(plan, "linked");
  const name = `${plan.azure.resourceGroup.replace(/^fabric-/, "")}-linked-dashboard`;
  await recordedWorkflow(config, info, plan.runId, "linked-dashboard", async page => {
    await withSimulation(cloud, plan, async fleet => {
      await test.step("4.3 Fresh source data from the same owned devices", () => verifyArrivingData(cloud, plan, fleet, path));
      const workload = await test.step("4.3.0 Solution home", () => solutionHome(page, plan, "Generate Real-Time Dashboards"));
      await test.step("4.3.1 Overview", () => dashboard.overview(workload));
      await test.step("4.3.2 Data Source", () => dashboard.dataSource(workload, plan, path));
      await test.step("4.3.3 Generate Dashboard", () => dashboard.generateDashboard(workload, name));
      await test.step("4.3.4 Ready to Use", () => dashboard.readyToUse(workload, path));
      await test.step("4.3.5 Native definition and source data", () => verifyDashboard(cloud, plan, path, name));
    }, true);
    await test.step("4.3.6 Open native dashboard and verify rendered device data", () => native.dashboard(page, plan, path));
  });
});
