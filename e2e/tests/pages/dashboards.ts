import { type FrameLocator } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import type { RunPlan } from "../../src/plan.js";
import type { IngestionPath } from "../../src/paths.js";
import { dashboardSelection } from "../../src/dashboard.js";

export async function overview(workload: FrameLocator): Promise<void> {
  await expect(workload.locator(".wizard-control__step-title")).toHaveText(["Overview", "Data source", "Generate dashboard", "Ready to use"]);
  await workload.getByRole("button", { name: "Next", exact: true }).click();
}

export async function dataSource(workload: FrameLocator, plan: RunPlan, path: IngestionPath): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Select data source and fields", level: 2 })).toBeVisible();
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeDisabled();
  for (const [selector, value] of [
    ["#dash-eh-select", plan.fabric.eventhouseName], ["#dash-db-select", plan.fabric.databaseName],
    ["#dash-telemetry-table", path.telemetryTable], ["#dash-properties-table", path.propertiesTable],
  ] as const) {
    await expect(workload.locator(selector)).toBeVisible({ timeout: 180_000 });
    await expect(workload.locator(selector)).toBeEnabled({ timeout: 180_000 });
    await workload.locator(selector).click();
    await workload.getByRole("option", { name: value, exact: true }).click();
    await expect(workload.locator(selector)).toContainText(value);
  }
  await workload.getByRole("button", { name: "Inspect Tables", exact: true }).click();
  await expect(workload.getByRole("checkbox", { name: "temperature", exact: true })).toBeVisible({ timeout: 90_000 });
  for (const checkbox of await workload.getByRole("checkbox").all()) await checkbox.uncheck();
  await expect(workload.getByRole("checkbox", { checked: true })).toHaveCount(0);
  await expect(workload.locator(".iot-solution-config-card").filter({ hasText: "Telemetry Fields" })).toContainText("0 selected");
  await expect(workload.locator(".iot-solution-config-card").filter({ hasText: "Property Fields" })).toContainText("0 selected");
  const fields = dashboardSelection(path.id);
  for (const field of fields) {
    const checkbox = workload.getByRole("checkbox", { name: field.name, exact: true });
    await expect(checkbox).toBeVisible();
    if (field.component && field.visual === "line") {
      await expect(checkbox.locator("xpath=ancestor::div[contains(@class, 'iot-solution-field-list-item')][1]")).toContainText("Time-series");
    }
    await checkbox.check();
  }
  await expect(workload.locator(".iot-solution-config-card").filter({ hasText: "Telemetry Fields" }))
    .toContainText(`${fields.filter(field => field.visual !== "card").length} selected`);
  await expect(workload.locator(".iot-solution-config-card").filter({ hasText: "Property Fields" }))
    .toContainText(`${fields.filter(field => field.visual === "card").length} selected`);
  await expect(next).toBeEnabled();
  await next.click();
}

export async function generateDashboard(workload: FrameLocator, name: string): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Generate dashboard", level: 2 })).toBeVisible();
  const next = workload.getByRole("button", { name: "Next", exact: true });
  await expect(next).toBeDisabled();
  const input = workload.getByLabel("Dashboard name", { exact: true });
  const create = workload.getByRole("button", { name: "Generate Dashboard", exact: true });
  await input.fill("");
  await expect(create).toBeDisabled();
  await input.fill(name);
  await expect(create).toBeEnabled();
  await create.click();
  await expect(workload.getByText(`Dashboard "${name}" created successfully.`, { exact: true })).toBeVisible({ timeout: 180_000 });
  await expect(workload.locator(".iot-solution-error-text")).toHaveCount(0);
  await next.click();
}

export async function readyToUse(workload: FrameLocator, path: IngestionPath): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Your real-time device dashboard is ready", level: 2 })).toBeVisible();
  const summary = workload.getByLabel("Generated dashboard composition", { exact: true });
  await expect(summary).toContainText(path.telemetryTable);
  await expect(summary).toContainText(path.propertiesTable);
  const fields = dashboardSelection(path.id);
  await expect(summary).toContainText(`${fields.filter(field => field.visual === "line").length} charts, 1 tables, ${fields.filter(field => field.visual === "card").length} value cards`);
  await expect(workload.getByRole("button", { name: "Open dashboard in Fabric", exact: true })).toBeEnabled();
  await workload.getByRole("button", { name: "Finish and Close", exact: true }).click();
  await expect(workload.getByRole("button", { name: "Generate Real-Time Dashboards", exact: true })).toBeVisible();
}
