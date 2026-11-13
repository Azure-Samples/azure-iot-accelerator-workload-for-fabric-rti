import { test } from "@playwright/test";
import { join } from "node:path";
import { artifacts, requireGuid, runConfig } from "../../../src/config.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { Cloud } from "../../../src/cloud.js";
import { verifyArrivingData, withSimulation } from "../../../src/data.js";
import { prepareFabric } from "../../../src/fabric.js";
import { recordIngestionStreams, routingBefore, verifyRouting } from "../../../src/routing.js";
import { verifyLinks } from "../../../src/linking.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan, updatePlan } from "../../../src/plan.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as ingestion from "../../pages/ingestion.js";

// Combine two nonlinked and three ADR-linked Hubs into a separate raw output pair.
test(scenarioTitle("combined"), async ({}, info) => {
  test.setTimeout(45 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  const path = ingestionPath(plan, "combined");
  await recordedWorkflow(config, info, plan.runId, "combined", async page => {
    const cloud = new Cloud(config);
    // Verify real namespace links: two Hubs share A; the third uses B.
    await test.step("Require native linked-Hub topology A/A/B", () => verifyLinks(cloud, plan));
    await withSimulation(cloud, plan, async fleet => {
      await prepareFabric(cloud, plan);
      // Create outputs with the first Hub, then add the other four through UI reuse.
      for (const [index, hub] of plan.azure.hubs.filter(hub => path.hubKeys.includes(hub.key)).entries()) {
        const id = `2.2.${index + 1}`;
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
      plan.verification["combined.ui"] = "5-selected-hubs-completed-ingestion";
      await updatePlan(plan);
      // Match fresh data to every device and Hub, including components and whole values.
      await test.step("Require exact native root/component telemetry and properties from every Hub", () => verifyArrivingData(cloud, plan, fleet, path));
    }, true);
    await info.attach("combined-run-manifest", { path: join(artifacts, "runs", plan.runId, "manifest.json"), contentType: "application/json" });
  });
});
