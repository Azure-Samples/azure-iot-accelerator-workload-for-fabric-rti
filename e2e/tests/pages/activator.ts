import { type FrameLocator } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import { redactSensitiveText } from "../../src/cloud.js";
import type { RunPlan } from "../../src/plan.js";

export async function overview(workload: FrameLocator): Promise<void> {
  await expect(workload.locator(".wizard-control__step-title")).toHaveText([
    "Overview", "Monitoring data", "Workspace identity", "Activator resources", "Ready to configure",
  ]);
  await expect(workload.getByRole("heading", { name: "Generate alerts and automated actions based on live IoT data", exact: true })).toBeVisible();
  await expect(workload.getByLabel("Activator monitoring flow", { exact: true })).toContainText("Modeled device data");
  await workload.getByRole("button", { name: "Next", exact: true }).click();
}

export async function monitoringData(workload: FrameLocator, plan: RunPlan, table: string): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Select data for monitoring", exact: true })).toBeVisible();
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeDisabled();
  for (const [selector, value] of [
    ["#act-eh-dropdown", plan.fabric.eventhouseName], ["#act-db-dropdown", plan.fabric.databaseName], ["#act-table-dropdown", table],
  ] as const) {
    await expect(workload.locator(selector)).toBeVisible({ timeout: 180_000 });
    await expect(workload.locator(selector)).toBeEnabled({ timeout: 180_000 });
    await workload.locator(selector).click();
    await workload.getByRole("option", { name: value, exact: true }).click();
    await expect(workload.locator(selector)).toContainText(value);
  }
  await expect(next).toBeEnabled();
  await next.click();
}

export async function workspaceIdentity(workload: FrameLocator): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Workspace identity", level: 2, exact: true })).toBeVisible();
  const ready = workload.getByText("Workspace identity is ready.", { exact: false });
  const enable = workload.getByRole("button", { name: "Enable workspace identity", exact: true });
  const errors = workload.locator(".fui-MessageBar");
  await expect.poll(async () => await ready.isVisible() || await enable.isVisible() || await errors.count() > 0, { timeout: 90_000 }).toBe(true);
  if (await errors.count()) throw new Error(`[workspace-identity] ${redactSensitiveText((await errors.allTextContents()).join("\n"))}`);
  if (await enable.isVisible()) await enable.click();
  await expect(ready).toBeVisible({ timeout: 180_000 });
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function activatorResources(workload: FrameLocator, activator: string, stream: string): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Create Activator resources", exact: true })).toBeVisible();
  const next = workload.getByRole("button", { name: "Next", exact: true });
  const create = workload.getByRole("button", { name: "Set Up Activator", exact: true });
  await expect(next).toBeDisabled();
  const name = workload.getByRole("textbox", { name: "Activator name", exact: true });
  await name.fill("");
  await expect(create).toBeDisabled();
  await name.fill(activator);
  await expect(name).toHaveValue(activator);
  const streamName = workload.getByRole("textbox", { name: "Eventstream name", exact: true });
  await streamName.fill("");
  await expect(create).toBeDisabled();
  await streamName.fill(stream);
  await expect(streamName).toHaveValue(stream);
  await expect(create).toBeEnabled();
  await create.click();
  const errors = workload.locator(".iot-solution-field-error, .fui-MessageBar");
  await expect.poll(async () => await next.isEnabled() || await errors.count() > 0, { timeout: 240_000 }).toBe(true);
  if (await errors.count()) throw new Error(`[activator-resources] ${redactSensitiveText((await errors.allTextContents()).join("\n"))}`);
  await expect(workload.locator(".iot-solution-creation-step-label.done")).toHaveCount(4);
  await expect(workload.getByText("Activator data feed created.", { exact: false })).toBeVisible();
  await expect(next).toBeEnabled();
  await next.click();
}

export async function readyToConfigure(workload: FrameLocator, plan: RunPlan, table: string, activator: string, stream: string): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Your device data is ready for Activator monitoring", exact: true })).toBeVisible();
  for (const value of [plan.fabric.databaseName, table, activator, stream]) {
    await expect(workload.locator(".iot-solution-validation")).toContainText(value);
  }
  await expect(workload.getByLabel("Activator resource architecture", { exact: true })).toContainText("Eventstream");
  await expect(workload.getByRole("button", { name: "Open Activator and create rules", exact: true })).toBeEnabled();
  await workload.getByRole("button", { name: "Finish and Close", exact: true }).click();
  await expect(workload.getByRole("button", { name: "Set up Activator rules", exact: true })).toBeVisible();
}
