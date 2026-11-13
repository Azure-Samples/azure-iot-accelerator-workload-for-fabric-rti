import { test } from "@playwright/test";
import { ActivatorAttempt } from "../../../src/activator.js";
import { Cloud } from "../../../src/cloud.js";
import { requireGuid, runConfig } from "../../../src/config.js";
import { withSimulation } from "../../../src/data.js";
import { modeledNames } from "../../../src/modeling.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan } from "../../../src/plan.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as activator from "../../pages/activator.js";
import * as native from "../../pages/native-experiences.js";

let resources: ActivatorAttempt | undefined;
test.afterEach(async () => {
  const current = resources;
  if (current) await test.step(
    `7.2.8 ${current.cloud.config.keepResources ? "Retain" : "Cleanup"} temporary stream, Activator and connection`,
    () => current.cleanupAfterScenario());
});

// All five simulated Hubs feed the combined modeled source.
test(scenarioTitle("combinedActivator"), async ({}, info) => {
  test.setTimeout(20 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  if (!plan.verification["combined.modeling"]?.startsWith("native-model")) throw new Error("[prerequisite] Complete WoT Combined modeling first.");
  const cloud = new Cloud(config), path = ingestionPath(plan, "combined"), table = modeledNames(path).modeled;
  const attempt = resources = await ActivatorAttempt.begin(cloud, plan, "combined");
  await recordedWorkflow(config, info, plan.runId, "combined-activator", async page => {
    await attempt.observe(page);
    await withSimulation(cloud, plan, async fleet => {
      const workload = await test.step("7.2.0 Solution home", () => solutionHome(page, plan, "Set up Activator rules"));
      await test.step("7.2.1 Overview", () => activator.overview(workload));
      await test.step("7.2.2 Monitoring Data", () => activator.monitoringData(workload, plan, table));
      await test.step("7.2.3 Workspace Identity", () => activator.workspaceIdentity(workload));
      await test.step("7.2.4 Activator Resources", () => activator.activatorResources(workload, attempt.activatorName, attempt.streamName));
      await test.step("7.2.5 Ready to Configure", () => activator.readyToConfigure(workload, plan, table, attempt.activatorName, attempt.streamName));
      await test.step("7.2.6 Native bindings, schema and fresh source data", () => attempt.verify(path, fleet));
      await test.step("7.2.7 Open Activator and observe incoming device data", () =>
        native.activator(page, plan, attempt.activatorId, attempt.incomingStreamName, plan.observed[`${path.id}.activatorDeviceId`]));
    }, true);
  });
});
