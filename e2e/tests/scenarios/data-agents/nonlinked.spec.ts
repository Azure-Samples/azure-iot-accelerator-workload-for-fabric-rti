import { test } from "@playwright/test";
import { Cloud } from "../../../src/cloud.js";
import { requireGuid, runConfig } from "../../../src/config.js";
import { verifyDataAgent } from "../../../src/experience-verification.js";
import { modelFixture } from "../../../src/models.js";
import { modeledNames } from "../../../src/modeling.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan } from "../../../src/plan.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as agent from "../../pages/data-agents.js";
import * as native from "../../pages/native-experiences.js";

// Ground the agent in the DTDL-modeled table; inspect actual draft and published definitions.
test(scenarioTitle("regularAgent"), async ({}, info) => {
  test.setTimeout(15 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  if (!plan.verification["regular.modeling"]?.startsWith("native-model")) throw new Error("[prerequisite] Complete DTDL Nonlinked modeling first.");
  const cloud = new Cloud(config), path = ingestionPath(plan, "regular"), fixture = modelFixture("regular");
  const table = modeledNames(path).modeled;
  const name = `${plan.azure.resourceGroup.replace(/^fabric-/, "")}-regular-data-agent`;
  await recordedWorkflow(config, info, plan.runId, "regular-data-agent", async page => {
    const workload = await test.step("5.1.0 Solution home", () => solutionHome(page, plan, "Set up Data Agent"));
    await test.step("5.1.1 Overview", () => agent.overview(workload));
    await test.step("5.1.2 Modeled Data Source", () => agent.modeledDataSource(workload, plan, table, fixture));
    await test.step("5.1.3 Configure Data Agent", () => agent.configureDataAgent(workload, name, table, fixture));
    await test.step("5.1.4 Ready to Use", () => agent.readyToUse(workload, name, table));
    await test.step("5.1.5 Native definition and grounding", () => verifyDataAgent(cloud, plan, path, name));
    await test.step("5.1.6 Ask the agent for a known fixture firmware version", () => native.dataAgent(page, plan, path, table, plan.observed[`${path.id}.agentDeviceId`]));
  });
});
