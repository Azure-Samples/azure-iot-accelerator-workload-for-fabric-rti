import { test } from "@playwright/test";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import { artifacts, requireGuid, runConfig } from "../../../src/config.js";
import { recordedWorkflow } from "../../../src/recording.js";
import { Cloud } from "../../../src/cloud.js";
import { verifyArrivingData, withSimulation } from "../../../src/data.js";
import { recordIngestionStreams, routingBefore, verifyRouting } from "../../../src/routing.js";
import { verifyLinks } from "../../../src/linking.js";
import { ingestionPath } from "../../../src/paths.js";
import { readPlan, updatePlan } from "../../../src/plan.js";
import { scenarioTitle } from "../../../src/scenarios.js";
import { solutionHome } from "../../pages/solution-home.js";
import * as ingestion from "../../pages/ingestion.js";

// New outputs must receive the three linked Hubs and exclude both nonlinked Hubs.
test(scenarioTitle("linked"), async ({}, info) => {
  test.setTimeout(45 * 60_000);
  const config = runConfig();
  const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
  if (!plan.verification["combined.rawData"]?.startsWith("native-kusto")) throw new Error("[prerequisite] Complete combined ingestion first.");
  const combined = ingestionPath(plan, "combined");
  const linked = ingestionPath(plan, "linked");
  await recordedWorkflow(config, info, plan.runId, "linked-only", async page => {
    const cloud = new Cloud(config);
    await test.step("Require actual three-Hub Succeeded ADR topology", () => verifyLinks(cloud, plan));
    // Run all five devices so exclusion cannot pass simply because regular devices are idle.
    await withSimulation(cloud, plan, async fleet => {
      const negative = fleet.devices.filter(device => !linked.hubKeys.includes(device.hubKey));
      for (const [index, hub] of plan.azure.hubs.filter(hub => linked.hubKeys.includes(hub.key)).entries()) {
        const id = `2.3.${index + 1}`;
        const previous = await routingBefore(cloud, plan, hub.id);
        const workload = await test.step(`${id}.0 ${hub.key}: Solution home`, () => solutionHome(page, plan, "Ingest device data from your IoT Hub"));
        await test.step(`${id}.1 Overview`, () => ingestion.overview(workload));
        await test.step(`${id}.2 IoT Hub`, () => ingestion.iotHub(workload, hub));
        await test.step(`${id}.3 Eventhouse`, () => ingestion.eventhouse(workload, plan));
        await test.step(`${id}.4 Raw Data Tables`, () => ingestion.rawDataTables(workload, linked, index === 0));
        await test.step(`${id}.5 Eventstreams`, () => ingestion.eventstreams(workload, linked, index === 0));
        await test.step(`${id}.5a Record native output identity and run folder`, () => recordIngestionStreams(cloud, plan, linked));
        await test.step(`${id}.6 IoT Hub Routing`, () => ingestion.iotHubRouting(workload, plan, hub));
        await test.step(`${id}.7 Ready to Use`, () => ingestion.readyToUse(workload, plan, linked, hub));
        await test.step(`${id}.8 Native routing and preservation`, () => verifyRouting(cloud, plan, hub.id, linked, previous));
        plan.observed[`${linked.id}.${hub.key}.ingestionCompletedAt`] = new Date().toISOString();
        await updatePlan(plan);
      }
      plan.phase = "regular-hub-ingestion-ui-and-routing-verified";
      plan.verification["linked.ui"] = "3-selected-hubs-completed-ingestion";
      await updatePlan(plan);
      await test.step("Prove all five currently sending devices still arrive in combined outputs", () => verifyArrivingData(cloud, plan, fleet, combined));
      await test.step("Prove linked positives and regular exclusion in the same window", async () => {
        await verifyArrivingData(cloud, plan, fleet, linked, negative);
        // Continue sending for a full minute, then recheck delivery and exclusion.
        await pause(60_000);
        await verifyArrivingData(cloud, plan, fleet, combined);
        await verifyArrivingData(cloud, plan, fleet, linked, negative);
        plan.observed.linkedExclusionWindowSeconds = "60";
        await updatePlan(plan);
      });
    }, true);
    await info.attach("linked-only-run-manifest", { path: join(artifacts, "runs", plan.runId, "manifest.json"), contentType: "application/json" });
  });
});
