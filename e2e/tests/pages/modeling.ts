import { type FrameLocator } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { redactSensitiveText } from "../../src/cloud.js";
import type { RunPlan } from "../../src/plan.js";
import type { IngestionPath } from "../../src/paths.js";
import type { ModelFixture } from "../../src/models.js";
import type { modeledNames } from "../../src/modeling.js";
import { scenarioScope } from "../../src/scoping.js";
import type { PathId } from "../../src/paths.js";
import { otherModelReference } from "../../src/model-assignment.js";
import { modelingAdrPrerequisite, documentationLinks, expectExternalDocumentationLink } from "./solution-home.js";

type Names = ReturnType<typeof modeledNames>;

export async function overview(workload: FrameLocator): Promise<void> {
  await expect(workload.locator(".wizard-control__step-title")).toHaveText([
    "Get started", "Select data source", "Add device model", "Assign model", "Check model coverage", "Create modeled data", "Review results",
  ]);
  await expect(workload.getByRole("heading", { name: "Turn raw IoT data into modeled data", exact: true })).toBeVisible();
  await expect(workload.getByRole("heading", { name: "What modeling unlocks", exact: true }).locator("..").getByRole("listitem")).toHaveText([
    "Compare measurements with device state, such as temperature by operating mode or vibration by firmware version.",
    "Query typed columns instead of parsing raw JSON in every query.",
    "Set more precise alerts that account for device configuration and status.",
    "Build dashboards, Activator rules, reports, and AI experiences on the same modeled data.",
  ]);
  const prerequisites = workload.locator(".iot-solution-prerequisites > ul > li");
  await expect(prerequisites).toHaveText([
    "Raw telemetry and properties tables from the Ingest device data from your IoT Hub experience.",
    modelingAdrPrerequisite,
    "A device model: a Web of Things (WoT) Thing Model (recommended), a Digital Twins Definition Language (DTDL) model, or an IoT Central device template export.",
    "At least one device event in your raw tables (more events give better suggestions).",
    "Permission to read raw data and to create tables, update policies, and a materialized view in the KQL database.",
  ]);
  await expect(prerequisites.first().locator("strong")).toHaveText("Ingest device data from your IoT Hub");
  await expectExternalDocumentationLink(prerequisites.nth(1).getByRole("link", { name: "linked to your IoT Hub", exact: true }),
    documentationLinks.adr);
  await workload.getByRole("button", { name: "Continue", exact: true }).click();
}

export async function iotDataSource(workload: FrameLocator, plan: RunPlan, path: IngestionPath): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Select your data source", level: 2 })).toBeVisible();
  const next = workload.getByRole("button", { name: "Continue", exact: true });
  await expect(next).toBeDisabled();
  for (const [selector, placeholder] of [
    ["#hub-model-eventhouse", "Select an Eventhouse"], ["#hub-model-database", "Select a database"],
    ["#hub-model-telemetry", "Select a table"], ["#hub-model-twin", "Select a table"],
  ] as const) await expect(workload.locator(selector)).toHaveText(placeholder, { timeout: 180_000 });
  await expect(workload.getByText("Select an Eventhouse first", { exact: true })).toHaveCount(0);
  await expect(workload.getByText("Select a database first", { exact: true })).toHaveCount(0);
  for (const [selector, value] of [
    ["#hub-model-eventhouse", plan.fabric.eventhouseName], ["#hub-model-database", plan.fabric.databaseName],
    ["#hub-model-telemetry", path.telemetryTable], ["#hub-model-twin", path.propertiesTable],
  ] as const) {
    await expect(workload.locator(selector)).toBeVisible({ timeout: 180_000 });
    await expect(workload.locator(selector)).toBeEnabled({ timeout: 180_000 });
    await workload.locator(selector).click();
    await workload.getByRole("option", { name: value, exact: true }).click();
    await expect(workload.locator(selector)).toContainText(value);
  }
  await workload.getByRole("button", { name: "Validate source tables", exact: true }).click();
  await expect(workload.getByText("Source tables validated", { exact: true })).toBeVisible({ timeout: 120_000 });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function deviceModel(workload: FrameLocator, fixture: ModelFixture, replacing = false): Promise<void> {
  if (createHash("sha256").update(await readFile(fixture.file)).digest("hex") !== fixture.hash) {
    throw new Error("[fixture] The supplied model file differs from its recorded hash.");
  }
  await expect(workload.getByRole("heading", { name: "Add a device model", level: 2 })).toBeVisible();
  const next = workload.getByRole("button", { name: "Continue", exact: true });
  if (!replacing) await expect(next).toBeDisabled();
  await workload.getByRole("radio", { name: fixture.format === "dtdl" ? "Digital Twin Definition Language (DTDL)" : "WoT Thing Model (recommended)", exact: true }).check();
  if (fixture.format === "wot") {
    const supported = workload.getByRole("list", { name: "Supported WoT models", exact: true });
    await expect(supported.getByRole("listitem")).toHaveText([
      "Supported model files: a self-contained WoT 1.1 Thing Model object or an array of related Thing Models. Each model needs its own supplied absolute ID.",
      "Supported capabilities: properties and event data using primitive, enum, and complex schemas.",
      "Not included: actions, submodels, and external definitions.",
    ]);
    await expect(supported.locator("strong")).toHaveText(["Supported model files:", "Supported capabilities:", "Not included:"]);
  }
  await workload.getByLabel("Device model file", { exact: true }).setInputFiles(fixture.file);
  await expect(workload.locator("#model-upload-reference")).toHaveText(fixture.id);
  await expect(workload.locator("#model-upload-error")).toHaveCount(0);
  if (fixture.format === "wot") {
    // Ignored submodels must not be presented as modeled component coverage.
    await expect(workload.locator("#model-upload-warnings")).toContainText("thermostat");
    await expect(workload.locator("#model-upload-warnings")).toContainText("powerModule");
  }
  await workload.getByRole("button", { name: "Deselect all", exact: true }).click();
  await expect(next).toBeDisabled();
  for (const field of fixture.fields) {
    const row = workload.locator(`[data-field-id="${field.id}"]`);
    await expect(row).toHaveCount(1);
    await row.getByRole("checkbox").check();
    if (fixture.format === "wot") {
      await row.getByRole("combobox").click();
      await workload.getByRole("option", { name: field.role === "telemetry" ? "Telemetry" : "Twin Property", exact: true }).click();
    } else {
      await expect(row).toContainText(field.role === "telemetry" ? "Telemetry" : "Twin Property");
      await expect(row.getByRole("combobox")).toHaveCount(0);
    }
  }
  await expect(workload.locator(".iot-solution-model-fields-inline").getByRole("checkbox", { checked: true })).toHaveCount(fixture.fields.length);
  await expect(next).toBeEnabled();
  await next.click();
}

export async function deviceModelAssignment(workload: FrameLocator, fixture: ModelFixture, path: PathId, plan: RunPlan): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Assign the model to devices", level: 2 })).toBeVisible();
  if (fixture.format === "dtdl") {
    await expect(workload.getByText("Model assignment isn't available for DTDL models.", { exact: false })).toBeVisible();
    await expect(workload.getByRole("button", { name: "Assign model to devices" })).toHaveCount(0);
    await workload.getByRole("button", { name: "Continue", exact: true }).click();
  } else {
    await workload.getByRole("button", { name: "Assign model to devices", exact: true }).click();
    if (path === "regular") {
      await expect(workload.getByLabel("Source IoT Hubs", { exact: true }))
        .toContainText("Your IoT hubs aren't linked to Azure Device Registry");
      await expectExternalDocumentationLink(workload.getByRole("link", { name: "Learn how to link an IoT hub to ADR", exact: true }),
        documentationLinks.adr);
      await expect(workload.locator("[data-assignment-unavailable-guidance]")).toBeVisible();
      await expect(workload.getByRole("table", { name: "Device model assignments", exact: true })).toHaveCount(0);
      await workload.getByTestId("wizard-previous-btn").click();
      return;
    }
    if (path === "linked") {
      await expect(workload.locator("[data-all-devices-assigned]")).toHaveText("All devices already have a model assigned", { timeout: 180_000 });
      await expect(workload.getByRole("button", { name: "Assign to selected devices", exact: true })).toHaveCount(0);
    } else {
      const showAll = workload.getByRole("button", { name: "Show all devices", exact: true });
      await expect(showAll.describe("Native device inventory finished loading")).toBeEnabled({ timeout: 180_000 });
      await showAll.click();
      await expectExternalDocumentationLink(workload.getByRole("link", { name: "Learn how to link an IoT hub to ADR", exact: true }),
        documentationLinks.adr);
      const table = workload.getByRole("table", { name: "Device model assignments", exact: true });
      await expect(table.locator("tbody tr")).toHaveCount(3, { timeout: 180_000 });
      for (const [hubKey, status, reference] of [
        ["linked1", "None", ""],
        ["linked2", "Assigned to this model", fixture.id],
        ["linked3", "Assigned to another model", otherModelReference],
      ] as const) {
        const hub = plan.azure.hubs.find(value => value.key === hubKey);
        if (!hub) throw new Error("[fixture] Missing linked assignment Hub.");
        const row = table.getByRole("row").filter({ hasText: `${hub.name}-device` });
        await expect(row).toHaveCount(1);
        await expect(row.getByText(status, { exact: true })).toBeVisible();
        if (reference) {
          await expect(row.getByText(reference, { exact: true })).toBeVisible();
          await expect(row.getByRole("checkbox")).toHaveCount(0);
        } else {
          await row.getByRole("checkbox", { name: `Select ${hub.name}-device`, exact: true }).check();
        }
      }
      await expect(table.getByRole("checkbox", { checked: true })).toHaveCount(1);
      await workload.getByRole("button", { name: "Assign to selected devices", exact: true }).click();
      const dialog = workload.getByRole("dialog");
      await expect(dialog).toContainText(`Assign ${fixture.title} to 1 device?`);
      await expect(dialog).toContainText(`Model ID: ${fixture.id}`);
      const selectedHub = plan.azure.hubs.find(hub => hub.key === "linked1");
      if (!selectedHub) throw new Error("[fixture] Missing unassigned linked device.");
      await expect(dialog).toContainText(`${selectedHub.name}-device`);
      console.log(`[model-assignment] Explicit Assign UI click at ${new Date().toISOString()}.`);
      await dialog.getByRole("button", { name: "Assign", exact: true }).click();
      const progress = workload.getByRole("region", { name: "Device assignment progress", exact: true });
      const result = progress.locator("[data-assignment-state]");
      await expect(result).toHaveCount(1);
      await expect.poll(async () => await progress.getAttribute("aria-busy") === "false" &&
        ["confirmed", "failed", "blocked", "cancelled", "unknown"].includes(await result.getAttribute("data-assignment-state") ?? ""),
      { timeout: 180_000 }).toBe(true);
      const state = await result.getAttribute("data-assignment-state");
      if (state !== "confirmed") {
        throw new Error(`[model-assignment] Assignment operation settled as ${state}: ${redactSensitiveText(await progress.innerText())}`);
      }
      await expect(result).toHaveAttribute("data-assignment-state", "confirmed");
      await expect(result.getByLabel("Assignment confirmed", { exact: true })).toBeVisible();
      await expect(result.locator("[data-assignment-warning]")).toHaveCount(0);
      const summary = progress.locator('[data-assignment-summary="confirmed"]');
      await expect(summary).toBeVisible();
      await expect(summary).toHaveText("Model assigned to 1 device");
    }
    await expectExternalDocumentationLink(workload.getByRole("link", { name: "Learn how to update DPS enrollments", exact: true }),
      documentationLinks.dps);
    await expect(workload.locator("#dps-guidance-unavailable")).toHaveCount(0);
    const next = workload.getByRole("button", { name: "Continue", exact: true });
    await expect(next).toBeEnabled();
    await next.click();
  }
}

export async function modelCoverage(workload: FrameLocator, fixture: ModelFixture, path: PathId): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Filter events and check model coverage", level: 2 })).toBeVisible();
  if (fixture.format === "wot") {
    const choice = path === "combined" ? "Require the model ID" : "Model ID or additional conditions";
    await workload.getByRole("radio", { name: choice, exact: true }).check();
    await expect(workload.getByRole("radio", { name: choice, exact: true })).toBeChecked();
  }
  const scope = scenarioScope(path);
  await workload.getByRole("radio", { name: "Add conditions", exact: true }).check();
  await workload.locator("#hub-telemetry-scope").fill(scope.telemetry);
  await workload.locator("#hub-twin-scope").fill(scope.twin);
  await expect(workload.locator("#hub-telemetry-scope")).toHaveValue(scope.telemetry);
  await expect(workload.locator("#hub-twin-scope")).toHaveValue(scope.twin);
  await workload.getByRole("button", { name: "Check coverage", exact: true }).click();
  for (const [stream, role] of [["telemetry", "telemetry"], ["twin", "twin"]] as const) {
    const section = workload.locator(`[data-coverage-stream="${stream}"]`);
    const count = fixture.fields.filter(field => field.role === role).length;
    await expect(section.getByRole("status")).toHaveText("All fields found", { timeout: 120_000 });
    await expect(section.locator("[data-observed-fields]")).toHaveAttribute("aria-label", `Selected model fields found: ${count} of ${count} in this sample.`);
    const samples = Number(await section.locator("[data-sample-count] .iot-solution-coverage-number").textContent());
    expect(samples).toBeGreaterThan(0);
    expect(samples).toBeLessThanOrEqual(100);
  }
  await expect(workload.getByRole("alert")).toHaveCount(0);
  const next = workload.getByRole("button", { name: "Continue", exact: true });
  await expect(next).toBeEnabled();
  await next.click();
}

export async function modeledData(workload: FrameLocator, names: Names): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Create modeled data", level: 2 })).toBeVisible();
  const next = workload.getByRole("button", { name: "Continue", exact: true });
  await expect(next).toBeDisabled();
  for (const [selector, value] of [
    ["#twin-changes-name", names.normalized], ["#twin-state-name", names.lkv], ["#modeled-output-name", names.modeled],
  ] as const) {
    await workload.locator(selector).fill(value);
    await expect(workload.locator(selector)).toHaveValue(value);
  }
  const create = workload.getByRole("button", { name: "Create modeled data", exact: true });
  await expect(create).toBeEnabled();
  await create.click();
  await expect(workload.locator("#modeled-data-success").or(workload.getByRole("alert").first())).toBeVisible({ timeout: 180_000 });
  const alert = workload.getByRole("alert");
  if (await alert.count()) throw new Error(`[modeled-data] ${redactSensitiveText((await alert.allTextContents()).join("\n"))}`);
  await expect(workload.locator("#modeled-data-success")).toBeVisible();
  await next.click();
}

export async function readyToUse(workload: FrameLocator, names: Names): Promise<void> {
  await expect(workload.getByRole("heading", { name: "Your modeled data is ready", level: 2 })).toBeVisible();
  await expect(workload.getByRole("heading", { name: "What was created", exact: true })).toBeVisible();
  const resources = workload.locator(".iot-solution-ready-resources");
  for (const value of Object.values(names)) await expect(resources).toContainText(value);
  for (const label of ["Modeled data table", "Normalized properties table", "Latest property values (materialized view)"]) {
    await expect(resources).toContainText(label);
  }
  await expectExternalDocumentationLink(workload.getByRole("link", { name: "Learn about Activator", exact: true }),
    documentationLinks.activator);
  await expectExternalDocumentationLink(workload.getByRole("link", { name: "Learn about Fabric data agents", exact: true }),
    documentationLinks.dataAgent);
  await workload.getByRole("button", { name: "Finish and Close", exact: true }).click();
  await expect(workload.getByRole("button", { name: "Model your device data", exact: true })).toBeVisible();
}
