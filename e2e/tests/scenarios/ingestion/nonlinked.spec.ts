import { test } from "@playwright/test";
import { join } from "node:path";
import { Cloud } from "../../../src/cloud.js";
import { artifacts, requireGuid, runConfig } from "../../../src/config.js";
import { verifyArrivingData, withSimulation } from "../../../src/data.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { prepareFabric } from "../../../src/fabric.js";
import { recordIngestionStreams, routingBefore, verifyRouting } from "../../../src/routing.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan, updatePlan } from "../../../src/plan.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as ingestion from "../../pages/ingestion.js";

// Two nonlinked Hubs: one system identity and one dedicated user-assigned identity.
test(scenarioTitle("ingestion"), async ({}, info) => {
  test.setTimeout(45 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  const path = ingestionPath(plan, "regular");
  if (!plan.observed.hubsProvisionedAt) throw new Error("[prerequisite] Provision the five owned test Hubs before ingestion. ADR/DPS setup is deferred and not required.");
  await recordedWorkflow(config, info, plan.runId, "regular", async page => {
    const cloud = new Cloud(config);
    // Keep local devices sending throughout setup; record the workflow after sign-in.
    await withSimulation(cloud, plan, async fleet => {
      await test.step("Create and verify the owned Fabric destination", () => prepareFabric(cloud, plan));
      // Hub 1 creates the raw tables/streams; Hub 2 reuses them through the UI.
      for (const [index, hub] of plan.azure.hubs.filter(hub => path.hubKeys.includes(hub.key)).entries()) {
        const id = `2.1.${index + 1}`;
        const previous = await routingBefore(cloud, plan, hub.id);
        const workload = await test.step(`${id}.0 ${hub.key}: Solution home`, () => solutionHome(page, plan, "Ingest device data from your IoT Hub"));
        await test.step(`${id}.1 Overview`, () => ingestion.overview(workload));
        await test.step(`${id}.2 IoT Hub`, () => ingestion.iotHub(workload, hub));
        await test.step(`${id}.3 Eventhouse`, () => ingestion.eventhouse(workload, plan));
        await test.step(`${id}.4 Raw Data Tables`, () => ingestion.rawDataTables(workload, path, index === 0));
        await test.step(`${id}.5 Eventstreams`, () => ingestion.eventstreams(workload, path, index === 0));
        await test.step(`${id}.5a Record native output identity and run folder`, () => recordIngestionStreams(cloud, plan, path));
        await test.step(`${id}.6 IoT Hub Routing`, () => ingestion.iotHubRouting(workload, plan, hub));
        await test.step(`${id}.7 Ready to Use`, () => ingestion.readyToUse(workload, plan, path, hub));
        await test.step(`${id}.8 Native routing and preservation`, () => verifyRouting(cloud, plan, hub.id, path, previous));
        plan.observed[`${path.id}.${hub.key}.ingestionCompletedAt`] = new Date().toISOString();
        await updatePlan(plan);
      }
      plan.phase = "regular-hub-ingestion-ui-and-routing-verified";
      plan.verification.ui = "2-selected-hubs-completed-ingestion";
      await updatePlan(plan);
      // Require fresh data from both devices, not just a successful wizard.
      await test.step("Require device acknowledgements and exact arriving data in native Kusto", () => verifyArrivingData(cloud, plan, fleet));
    });
    await info.attach("run-manifest", { path: join(artifacts, "runs", plan.runId, "manifest.json"), contentType: "application/json" });
  });
});
