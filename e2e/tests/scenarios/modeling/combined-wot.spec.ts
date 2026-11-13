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

// Upload ThermoProX200-wot-tm.json against the combined five-Hub raw tables.
test(scenarioTitle("combinedModel"), async ({}, info) => {
  test.setTimeout(30 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  if (!plan.verification["combined.rawData"]?.startsWith("native-kusto")) throw new Error("[prerequisite] Complete combined ingestion first.");
  const cloud = new Cloud(config), path = ingestionPath(plan, "combined"), fixture = modelFixture("combined");
  const names = modeledNames(path);
  await recordedWorkflow(config, info, plan.runId, "combined-modeling", async page => {
    await withSimulation(cloud, plan, async fleet => {
      await test.step("3.2 Native initial references: unassigned, this model, another model", () => verifyModelReferences(cloud, plan, false));
      const workload = await test.step("3.2.0 Solution home", () => solutionHome(page, plan, "Model your device data"));
      await test.step("3.2.1 Overview", () => modeling.overview(workload));
      await test.step("3.2.2 IoT Data Source", () => modeling.iotDataSource(workload, plan, path));
      await test.step("3.2.3 Device Model", () => modeling.deviceModel(workload, fixture));
      await test.step("3.2.4 Device Model Assignment", () => modeling.deviceModelAssignment(workload, fixture, path.id, plan));
      await test.step("3.2 Native reference readback; existing assignments unchanged", () => verifyModelReferences(cloud, plan, true));
      await test.step("3.2 Fresh referenced telemetry and properties", () => verifyIncomingReferences(cloud, plan, path, fleet));
      await test.step("3.2.5 Model reference AND telemetry condition", () => modeling.modelCoverage(workload, fixture, path.id));
      await test.step("3.2.6 Modeled Data", () => modeling.modeledData(workload, names));
      await test.step("3.2.7 Ready to Use", () => modeling.readyToUse(workload, names));
      plan.observed["combined.modelCreatedAt"] = new Date().toISOString();
      await updatePlan(plan);
      await test.step("3.2.8 Native outputs and fresh enrichment", () => verifyModeled(cloud, plan, path, fleet));
    }, true);
  });
});
