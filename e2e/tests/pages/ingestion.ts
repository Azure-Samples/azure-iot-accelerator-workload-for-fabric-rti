import { type FrameLocator } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import type { RunPlan } from "../../src/plan.js";
import type { IngestionPath } from "../../src/paths.js";
import { redactSensitiveText } from "../../src/cloud.js";
import { ingestionAdrPrerequisite, documentationLinks, expectExternalDocumentationLink } from "./solution-home.js";

type Hub = RunPlan["azure"]["hubs"][number];

export async function overview(workload: FrameLocator): Promise<void> {
  await expect(workload.locator(".wizard-control__step-title")).toHaveText([
    "Overview", "IoT Hub", "Eventhouse", "Raw data tables", "Eventstreams", "IoT Hub routing", "Ready to use",
  ]);
  await expect(workload.getByRole("heading", { name: "Connect your IoT data to Microsoft Fabric", exact: true })).toBeVisible();
  const prerequisites = workload.locator(".iot-solution-prerequisites > ul > li");
  await expect(prerequisites).toHaveText([
    "An existing Azure IoT Hub with a system-assigned or user-assigned managed identity enabled. An identity is needed for IoT Hub to securely route data to your Fabric solution.",
    ingestionAdrPrerequisite,
    "An existing Eventhouse and KQL Database in this Fabric workspace to store the ingested device data used for analytics and operational intelligence.",
    "Permission to read and configure the IoT Hub, create Fabric items, and manage workspace access for the IoT Hub identity. These permissions allow the accelerator to securely configure the data path on your behalf.",
  ]);
  await expectExternalDocumentationLink(prerequisites.nth(1).getByRole("link", { name: "linked to your IoT Hub", exact: true }),
    documentationLinks.adr);
  await workload.getByRole("button", { name: "Next", exact: true }).click();
}

export async function iotHub(workload: FrameLocator, hub: Hub): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Select your IoT Hub", exact: true })).toBeVisible();
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeDisabled();
  const picker = workload.locator("#hub-combobox");
  await expect(picker).toBeVisible({ timeout: 180_000 });
  await expect(picker).toBeEnabled({ timeout: 180_000 });
  await expect(workload.getByRole("button", { name: "Refresh IoT Hubs", exact: true })).toBeVisible();
  await expect(workload.getByRole("button", { name: "Refresh", exact: true })).toHaveCount(0);
  await picker.click();
  await picker.fill(hub.name);
  await picker.press("ArrowDown");
  await workload.getByRole("option").filter({ hasText: hub.name }).click();
  await expect(workload.getByText("IoT Hub validated successfully", { exact: true })).toBeVisible({ timeout: 60_000 });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function eventhouse(workload: FrameLocator, plan: RunPlan): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Select Eventhouse & database", exact: true })).toBeVisible();
  for (const [selector, value] of [["#eventhouse-select", plan.fabric.eventhouseName], ["#database-select", plan.fabric.databaseName]] as const) {
    await expect(workload.locator(selector)).toBeVisible({ timeout: 180_000 });
    await expect(workload.locator(selector)).toBeEnabled({ timeout: 180_000 });
    await workload.locator(selector).click();
    await workload.getByRole("option", { name: value, exact: true }).click();
  }
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function rawDataTables(workload: FrameLocator, path: IngestionPath, create: boolean): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Create or reuse raw data tables", exact: true })).toBeVisible();
  if (create) {
    await workload.getByRole("radio", { name: "Create new tables", exact: true }).check();
    for (const [selector, value] of [["#telemetry-table", path.telemetryTable], ["#properties-table", path.propertiesTable]] as const) {
      await workload.locator(selector).fill(value);
      await expect(workload.locator(selector)).toHaveValue(value);
    }
    await workload.getByRole("button", { name: "Create Tables", exact: true }).click();
    await expect(workload.getByText("Tables created successfully", { exact: true })).toBeVisible({ timeout: 180_000 });
  } else {
    await workload.getByRole("radio", { name: "Reuse existing tables", exact: true }).check();
    for (const [selector, value, stream] of [
      ["#reuse-telemetry", path.telemetryTable, path.telemetryStream],
      ["#reuse-properties", path.propertiesTable, path.propertiesStream],
    ] as const) {
      await expect(workload.locator(selector).describe("Reusable-table discovery finished")).toBeVisible({ timeout: 180_000 });
      await expect(workload.locator(selector)).toBeEnabled({ timeout: 180_000 });
      await workload.locator(selector).click();
      const option = workload.getByRole("option", { name: `${value} via ${stream}`, exact: true });
      await expect(option).toHaveCount(1);
      await option.click();
      await expect(workload.locator(selector)).toContainText(value);
    }
    await expect(workload.getByRole("button", { name: "Refresh", exact: true })).toHaveCount(0);
    await workload.getByRole("button", { name: "Use selected tables", exact: true }).click();
    await expect(workload.getByText("Existing tables selected", { exact: true })).toBeVisible({ timeout: 120_000 });
  }
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function eventstreams(workload: FrameLocator, path: IngestionPath, create: boolean): Promise<void> {
  if (create) {
    await expect(workload.getByRole("heading", { name: "Create Eventstreams", exact: true })).toBeVisible();
    for (const [selector, value] of [["#telemetry-stream", path.telemetryStream], ["#properties-stream", path.propertiesStream]] as const) {
      await workload.locator(selector).fill(value);
      await expect(workload.locator(selector)).toHaveValue(value);
    }
    await workload.getByRole("button", { name: "Create Eventstreams", exact: true }).click();
    const success = workload.getByText("Eventstreams created successfully", { exact: true });
    const errors = workload.locator(".iot-solution-field-error, .iot-solution-error-text");
    await expect(success.or(errors.first())).toBeVisible({ timeout: 300_000 });
    if (await errors.count()) throw new Error(`[eventstreams] ${redactSensitiveText((await errors.allTextContents()).join("\n"))}`);
    await expect(success).toBeVisible();
  } else {
    await expect(workload.getByText("Using existing Eventstreams", { exact: true })).toBeVisible();
  }
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function iotHubRouting(workload: FrameLocator, plan: RunPlan, hub: Hub): Promise<void> {
  await expect(workload.getByRole("heading", { name: "IoT Hub routing", exact: true, level: 2 })).toBeVisible();
  const identity = workload.getByRole("combobox").last();
  const name = hub.key === "regular2" ? plan.azure.regularIdentity.name : "System-assigned";
  await identity.click();
  await workload.getByRole("option").filter({ hasText: name }).click();
  await expect(identity).toContainText(name);
  await workload.getByRole("button", { name: /^(Check Access|Re-check)$/ }).click();
  const add = workload.getByRole("button", { name: "Add as Contributor", exact: true });
  const configure = workload.getByRole("button", { name: "Configure Routing", exact: true });
  await expect(add.or(configure)).toBeVisible({ timeout: 60_000 });
  if (await add.isVisible()) await add.click();
  await expect(configure).toBeEnabled({ timeout: 90_000 });
  await configure.click();
  await expect(workload.getByText("IoT Hub routing configured successfully", { exact: true })).toBeVisible({ timeout: 180_000 });
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function readyToUse(workload: FrameLocator, plan: RunPlan, path: IngestionPath, hub: Hub): Promise<void> {
  await expect(workload.getByRole("heading", {
    name: "Your IoT data is ready for analytics and operational intelligence", exact: true,
  })).toBeVisible();
  const architecture = workload.getByLabel("Deployed ingestion architecture", { exact: true });
  await expect(architecture).toContainText(hub.name);
  await expect(architecture).toContainText(plan.fabric.databaseName);
  await expect(workload.locator(".iot-solution-summary")).toContainText(path.telemetryTable);
  await expect(workload.locator(".iot-solution-summary")).toContainText(path.propertiesTable);
  await expect(workload.locator(".iot-solution-error, .iot-solution-field-error")).toHaveCount(0);
  await workload.getByRole("button", { name: "Finish and Close", exact: true }).click();
  await expect(workload.getByRole("button", { name: "Ingest device data from your IoT Hub", exact: true })).toBeVisible({ timeout: 90_000 });
}
