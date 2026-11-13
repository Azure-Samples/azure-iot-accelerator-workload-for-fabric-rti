import { type FrameLocator } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import type { RunPlan } from "../../src/plan.js";
import type { ModelFixture } from "../../src/models.js";

export async function overview(workload: FrameLocator): Promise<void> {
  await expect(workload.locator(".wizard-control__step-title")).toHaveText(["Overview", "Modeled data source", "Configure Data Agent", "Ready to use"]);
  await workload.getByRole("button", { name: "Next", exact: true }).click();
}

export async function modeledDataSource(workload: FrameLocator, plan: RunPlan, table: string, fixture: ModelFixture): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Select modeled data source", level: 2 })).toBeVisible();
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeDisabled();
  for (const [selector, value] of [
    ["#agent-eh", plan.fabric.eventhouseName], ["#agent-db", plan.fabric.databaseName], ["#agent-table", table],
  ] as const) {
    await expect(workload.locator(selector)).toBeVisible({ timeout: 180_000 });
    await expect(workload.locator(selector)).toBeEnabled({ timeout: 180_000 });
    await workload.locator(selector).click();
    await workload.getByRole("option", { name: value, exact: true }).click();
    await expect(workload.locator(selector)).toContainText(value);
  }
  await expect(workload.locator(".iot-solution-validation")).toContainText(`${fixture.fields.length + 2} typed columns available to the agent`);
  await expect(next).toBeEnabled();
  await next.click();
}

export async function configureDataAgent(workload: FrameLocator, name: string, table: string, fixture: ModelFixture): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Configure Data Agent", level: 2 })).toBeVisible();
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeDisabled();
  const create = workload.getByRole("button", { name: "Create Data Agent", exact: true });
  const input = workload.getByRole("textbox", { name: "Data Agent name", exact: true });
  await input.fill("");
  await expect(create).toBeDisabled();
  await input.fill(name);
  await expect(input).toHaveValue(name);
  await expect(workload.locator("pre")).toContainText(table);
  await expect(workload.locator("pre")).toContainText("event_time");
  await expect(workload.locator("pre")).toContainText(fixture.id);
  await expect(create).toBeEnabled();
  await create.click();
  await expect(next).toBeEnabled({ timeout: 180_000 });
  await expect(workload.locator(".iot-solution-error-text")).toHaveCount(0);
  await next.click();
}

export async function readyToUse(workload: FrameLocator, name: string, table: string): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Your IoT Data Agent is ready", level: 2 })).toBeVisible();
  await expect(workload.locator(".iot-solution-validation")).toContainText(name);
  await expect(workload.locator(".iot-solution-validation")).toContainText(table);
  await expect(workload.getByRole("button", { name: "Open Data Agent in Fabric", exact: true })).toBeEnabled();
  await workload.getByRole("button", { name: "Finish and Close", exact: true }).click();
  await expect(workload.getByRole("button", { name: "Set up Data Agent", exact: true })).toBeVisible();
}
