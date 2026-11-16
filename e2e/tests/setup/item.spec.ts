import { test } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import { join } from "node:path";
import { artifacts, requireGuid, runConfig } from "../../src/config.js";
import { recordedWorkflow } from "../../src/recording.js";
import { creationDialog } from "../../src/item.js";
import { createPlan, readPlan, updatePlan, writePlan } from "../../src/plan.js";
import { Cloud } from "../../src/cloud.js";
import { ensureRunFolder, moveOwnedItem, folderName } from "../../src/folders.js";
import { scenarioTitle } from "../../src/scenarios.js";
import { requireWorkloadOrigin } from "../../src/deployment.js";
import { overview } from "../pages/ingestion.js";

test(scenarioTitle("item"), async ({}, info) => {
  const config = runConfig();
  const requested = process.env.E2E_RUN_ID;
  const plan = requested ? await readPlan(config, requireGuid(requested, "E2E_RUN_ID")) : createPlan(config);
  if (plan.phase !== "planned" || plan.observed.itemId) {
    throw new Error("[ownership] This run already attempted or completed item creation. Continue its provisioning/ingestion instead of creating another item.");
  }
  const manifestPath = requested ? join(artifacts, "runs", plan.runId, "manifest.json") : await writePlan(plan);
  console.log(`Run ownership manifest: ${manifestPath}`);
  await recordedWorkflow(config, info, plan.runId, "item", async page => {
    await page.goto(config.workspaceUrl, { waitUntil: "domcontentloaded" });
    await page.getByRole("button", { name: "New item", exact: true }).waitFor();
    const cloud = new Cloud(config);
    await ensureRunFolder(cloud, plan);
    const dialog = await creationDialog(page, config);
    await dialog.getByRole("button", { name: "Location *", exact: true }).click();
    await page.getByRole("button", { name: folderName(plan), exact: true }).click();
    await expect(dialog.getByRole("button", { name: "Location *", exact: true })).toContainText(folderName(plan));
    const name = dialog.getByRole("textbox", { name: "Name *", exact: true });
    await name.click();
    await name.pressSequentially(plan.fabric.itemName);
    await name.press("Tab");
    try {
      await expect(name).toHaveValue(plan.fabric.itemName);
      await expect(dialog.getByRole("button", { name: "Create", exact: true })).toBeEnabled();
    } finally {
      await info.attach("filled-creation-dialog", {
        body: await dialog.ariaSnapshot(), contentType: "text/plain",
      });
    }
    plan.phase = "item-create-requested";
    await updatePlan(plan);
    await dialog.getByRole("button", { name: "Create", exact: true }).click();
    await expect(dialog).not.toBeVisible({ timeout: 90_000 });
    plan.observed.itemUrl = page.url();
    plan.phase = "item-create-submitted";
    await updatePlan(plan);
    const workload = page.getByTestId("iframe-page-Org.IoTSolutionAccelerator").contentFrame();
    await expect(workload.getByRole("button", { name: "Ingest device data from your IoT Hub", exact: true }))
      .toBeVisible({ timeout: 90_000 });
    plan.resourcesCreated = true;
    plan.phase = "item-created";
    const frame = page.frames().find(candidate => /\/IoTSolutionItem-editor\/[0-9a-f-]{36}/i.test(candidate.url()));
    if (!frame) {
      throw new Error("[deployment] The created item loaded but its expected item editor route was not found. The manifest retains its top-level URL.");
    }
    requireWorkloadOrigin(frame.url(), config);
    const match = /\/IoTSolutionItem-editor\/([0-9a-f-]{36})/i.exec(frame.url());
    if (!match?.[1]) throw new Error("[harness] The item editor did not expose a valid item ID.");
    plan.observed.itemFrameUrl = frame.url();
    plan.observed.itemId = match[1];
    plan.observed.itemUrl = page.url();
    await updatePlan(plan);
    await moveOwnedItem(cloud, plan, plan.observed.itemId, plan.fabric.itemName);
    await workload.getByRole("button", { name: "Ingest device data from your IoT Hub", exact: true }).click();
    await overview(workload);
    await expect(workload.getByRole("heading", { name: "Select your IoT Hub", exact: true })).toBeVisible();
    await expect(workload.getByRole("combobox", { name: "IoT Hub", exact: true })).toBeVisible({ timeout: 60_000 });
    plan.verification.ui = "fresh-item-and-ingestion-entrypoint-verified";
    plan.phase = "ingestion-entrypoint-verified";
    await updatePlan(plan);
    await info.attach("run-manifest", { path: manifestPath, contentType: "application/json" });
    console.log("New item retained for this run. Azure topology, simulator and native raw-data verification have not run.");
  });
});
