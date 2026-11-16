import { test } from "@playwright/test";
import { Cloud } from "../../../src/cloud.js";
import { requireGuid, runConfig } from "../../../src/config.js";
import { withSimulation } from "../../../src/data.js";
import { modelFixture } from "../../../src/models.js";
import { modeledNames, verifyModeled } from "../../../src/modeling.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan, updatePlan } from "../../../src/plan.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as modeling from "../../pages/modeling.js";

// Upload ThermoProX200-dtdl.json against the two nonlinked Hubs' raw tables.
test(scenarioTitle("regularModel"), async ({}, info) => {
  test.setTimeout(30 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  if (!plan.verification.rawData?.startsWith("native-kusto")) throw new Error("[prerequisite] Complete nonlinked ingestion first.");
  const cloud = new Cloud(config), path = ingestionPath(plan, "regular"), fixture = modelFixture("regular");
  const names = modeledNames(path);
  await recordedWorkflow(config, info, plan.runId, "regular-modeling", async page => {
    await withSimulation(cloud, plan, async fleet => {
      const workload = await test.step("3.1.0 Solution home", () => solutionHome(page, plan, "Model your device data"));
      await test.step("3.1.1 Overview", () => modeling.overview(workload));
      await test.step("3.1.2 IoT Data Source", () => modeling.iotDataSource(workload, plan, path));
      await test.step("3.1.3a WoT Device Model", () => modeling.deviceModel(workload, modelFixture("combined")));
      await test.step("3.1.4a WoT assignment unavailable for nonlinked Hubs; return to Device Model",
        () => modeling.deviceModelAssignment(workload, modelFixture("combined"), path.id, plan));
      await test.step("3.1.3b Replace WoT with DTDL", () => modeling.deviceModel(workload, fixture, true));
      await test.step("3.1.4b Assignment unavailable for DTDL", () => modeling.deviceModelAssignment(workload, fixture, path.id, plan));
      await test.step("3.1.5 Model Coverage with telemetry and property conditions", () => modeling.modelCoverage(workload, fixture, path.id));
      await test.step("3.1.6 Modeled Data", () => modeling.modeledData(workload, names));
      await test.step("3.1.7 Ready to Use", () => modeling.readyToUse(workload, names));
      plan.observed["regular.modelCreatedAt"] = new Date().toISOString();
      await updatePlan(plan);
      await test.step("3.1.8 Native outputs and fresh enrichment", () => verifyModeled(cloud, plan, path, fleet));
    });
  });
});
