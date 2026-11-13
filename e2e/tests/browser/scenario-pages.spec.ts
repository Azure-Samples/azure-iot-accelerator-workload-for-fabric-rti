import { test, expect } from "@playwright/test";
import { readFile, readdir, rmdir, unlink } from "node:fs/promises";
import { dirname, join } from "node:path";
import { runConfig } from "../../src/config.js";
import { dashboardSelection } from "../../src/dashboard.js";
import { otherModelReference } from "../../src/model-assignment.js";
import { modelFixture } from "../../src/models.js";
import { modeledNames } from "../../src/modeling.js";
import { ingestionPath } from "../../src/paths.js";
import { authenticatedAppSessions } from "../../src/recording.js";
import { creationDialog } from "../../src/item.js";
import { closeOwnedItemTabs, ownedItemTabs } from "../../src/browser-tabs.js";
import { createPlan, writePlan } from "../../src/plan.js";
import { scenarioScope } from "../../src/scoping.js";
import { scenarioTitle } from "../../src/scenarios.js";
import * as modeling from "../pages/modeling.js";
import * as native from "../pages/native-experiences.js";
import * as ingestion from "../pages/ingestion.js";
import * as dashboardPages from "../pages/dashboards.js";
import { enableVisualChecks, expect as visualExpect } from "../../src/visual-expect.js";
import { ingestionAdrPrerequisite, documentationLinks, expectExternalDocumentationLink } from "../pages/solution-home.js";

test.use({ channel: "msedge", headless: true });
const config = runConfig({
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development",
});
test.beforeEach(async ({ context, page }) => {
  await context.route("**/*", route => route.abort("blockedbyclient"));
  enableVisualChecks(page, 0);
});

test(scenarioTitle("pageCatalog"), async ({ page }) => {
  const label = "Azure IoT Solution Accelerator (preview)";
  const description = "\u2022 Optionally enrich raw device data with semantic context from WoT Thing Models or DTDL device models and generate query-ready modeled tables.";
  await page.setContent(`<button onclick="document.querySelector('[role=region]').hidden=false">New item</button>
    <section role="region" aria-label="New item" hidden>
      <input placeholder="Filter by keyword" oninput="document.querySelector('#workload').hidden=!('Azure IoT Solution Accelerator'.includes(this.value))">
      <button id="workload" aria-label="${label}" aria-describedby="capability" onclick="document.querySelector('[role=dialog]').hidden=false">
        ${label}<span id="capability">${description}</span></button>
      <button aria-label="${label} other">Wrong workload<span>${description}</span></button>
    </section>
    <div role="dialog" aria-label="New Azure IoT Solution Accelerator" hidden><button disabled>Create</button></div>`);
  const dialog = await creationDialog(page, { ...config, itemLabel: label });
  await expect(page.getByPlaceholder("Filter by keyword")).toHaveValue("Azure IoT Solution Accelerator");
  await expect(page.locator("#workload")).toHaveAccessibleName(label);
  await expect(page.locator("#workload")).toContainText(description);
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole("button", { name: "Create", exact: true })).toBeDisabled();
});

test(scenarioTitle("pageModelSource"), async ({ page }) => {
  const plan = createPlan(config), path = ingestionPath(plan, "regular");
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated modeling frame.");
  await frame.setContent(`
    <aside>${["Get started", "Select data source", "Add device model", "Assign model", "Check model coverage", "Create modeled data", "Review results"]
      .map(title => `<h3 class="wizard-control__step-title">${title}</h3>`).join("")}</aside>
    <section id="overview"><h2>Turn raw IoT data into modeled data</h2>
      <section><h3>What modeling unlocks</h3><ul>
        <li>Compare measurements with device state, such as temperature by operating mode or vibration by firmware version.</li>
        <li>Query typed columns instead of parsing raw JSON in every query.</li>
        <li>Set more precise alerts that account for device configuration and status.</li>
        <li>Build dashboards, Activator rules, reports, and AI experiences on the same modeled data.</li>
      </ul></section>
      <section class="iot-solution-prerequisites"><h3>What you'll need</h3><ul>
        <li>Raw telemetry and properties tables from the <strong>Ingest device data from your IoT Hub</strong> experience.</li>
        <li>Optionally, an Azure Device Registry (ADR) namespace <a href="${documentationLinks.adr}" target="_blank" rel="noopener noreferrer">linked to your IoT Hub</a>. Linking an ADR namespace to your IoT Hub allows devices to reference your model.</li>
        <li>A device model: a Web of Things (WoT) Thing Model (recommended), a Digital Twins Definition Language (DTDL) model, or an IoT Central device template export.</li>
        <li>At least one device event in your raw tables (more events give better suggestions).</li>
        <li>Permission to read raw data and to create tables, update policies, and a materialized view in the KQL database.</li>
      </ul></section>
      <button onclick="this.parentElement.remove();document.querySelector('#source').hidden=false">Continue</button></section>
    <section id="source" hidden><h2>Select your data source</h2>
      ${[
        ["hub-model-eventhouse", "Select an Eventhouse", plan.fabric.eventhouseName],
        ["hub-model-database", "Select a database", plan.fabric.databaseName],
        ["hub-model-telemetry", "Select a table", path.telemetryTable],
        ["hub-model-twin", "Select a table", path.propertiesTable],
      ].map(([id, placeholder, value]) => `<button role="combobox" id="${id}" onclick="this.nextElementSibling.hidden=false">${placeholder}</button>
        <div role="listbox" hidden><div role="option" onclick="document.getElementById('${id}').textContent='${value}';this.parentElement.hidden=true">${value}</div></div>`).join("")}
      <button onclick="document.querySelector('#validation').textContent='Source tables validated';document.querySelector('#next').disabled=false">Validate source tables</button>
      <p id="validation"></p><button id="next" disabled onclick="document.querySelector('#finished').hidden=false">Continue</button>
    </section><h2 id="finished" hidden>Add a device model</h2>`);
  const workload = page.frameLocator("iframe");
  await modeling.overview(workload);
  await modeling.iotDataSource(workload, plan, path);
  await expect(frame.getByRole("heading", { name: "Add a device model", exact: true })).toBeVisible();
});

test(scenarioTitle("pageOwnedTabs"), async ({ page }) => {
  const plan = createPlan(config), prefix = `${plan.azure.resourceGroup.slice("fabric-".length)}-`;
  plan.observed.itemId = "00000000-0000-4000-8000-000000000008";
  plan.observed["regular.dashboardId"] = "00000000-0000-4000-8000-000000000009";
  const targets = ownedItemTabs(plan);
  expect(targets.map(tab => tab.name)).toEqual([plan.fabric.itemName, `${prefix}regular-dashboard`]);
  await page.route(plan.target.workspaceUrl, route => route.fulfill({ contentType: "text/html", body: "<main>Navigation is loading</main>" }));
  await page.goto(plan.target.workspaceUrl);
  const closing = closeOwnedItemTabs(page, plan);
  await expect(page.getByRole("menu", { name: "Sidebar", exact: true })).toHaveCount(0);
  await page.evaluate(html => document.body.insertAdjacentHTML("beforeend", html), `<nav role="menu" aria-label="Sidebar">
    ${targets.map(({ name }) => `<div role="menuitem" aria-label="${name} - Workspace">${name}
      <button data-testid="tab-close" aria-label="Close tab" onclick="this.parentElement.remove()"></button></div>`).join("")}
    <div role="menuitem" aria-label="Customer solution" id="foreign">Customer solution
      <button data-testid="tab-close" onclick="this.parentElement.remove()">Close</button></div>
    <div role="menuitem" aria-label="e2e-fedcba9876543210-not-owned" id="unknown">e2e-fedcba9876543210-not-owned
      <button data-testid="tab-close" onclick="this.parentElement.remove()">Close</button></div>
    <div role="menuitem" aria-label="e2e-823476c61570bba9-solution" id="manual">e2e-823476c61570bba9-solution
      <button data-testid="tab-close" onclick="this.parentElement.remove()">Close</button></div>
    <div role="menuitem" aria-label="${prefix}unrecorded" id="unrecorded">${prefix}unrecorded
      <button data-testid="tab-close" onclick="this.parentElement.remove()">Close</button></div>
    </nav>`);
  expect(await closing).toEqual(targets);
  await expect(page.locator("#foreign")).toBeVisible();
  await expect(page.locator("#unknown")).toBeVisible();
  await expect(page.locator("#manual")).toBeVisible();
  await expect(page.locator("#unrecorded")).toBeVisible();
  plan.azure.tags["e2e-run-id"] = "00000000-0000-4000-8000-000000000099";
  await expect(closeOwnedItemTabs(page, plan)).rejects.toThrow("ownership");
});

test(scenarioTitle("pageStreamFailure"), async ({ page }) => {
  const path = ingestionPath(createPlan(config), "regular");
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated Eventstream frame.");
  await frame.setContent(`<h2>Create Eventstreams</h2><input id="telemetry-stream"><input id="properties-stream">
    <button id="create" data-clicks="0" onclick="this.dataset.clicks=String(Number(this.dataset.clicks)+1);document.querySelector('#failure').hidden=false">Create Eventstreams</button>
    <p id="failure" class="iot-solution-field-error" hidden>Eventstream creation is still in progress. primaryKey=fake-key-must-not-appear</p>
    <button disabled>Next</button>`);
  let message = "";
  try { await ingestion.eventstreams(page.frameLocator("iframe"), path, true); } catch (failure) {
    if (!(failure instanceof Error)) throw failure;
    message = failure.message;
  }
  expect(message).toContain("[eventstreams] Eventstream creation is still in progress");
  expect(message).not.toContain("fake-key-must-not-appear");
  await expect(frame.locator("#create")).toHaveAttribute("data-clicks", "1");
  await expect(frame.getByRole("button", { name: "Next", exact: true })).toBeDisabled();
});

test(scenarioTitle("pageVisual"), async ({ page }, info) => {
  await page.setContent('<h1>Verification preview</h1><p id="status">Ready</p><p id="late">Waiting</p>');
  enableVisualChecks(page, 0);
  const status = page.locator("#status").describe("Device model assignment status");
  await visualExpect(status).toHaveText("Ready");
  await visualExpect(status).not.toHaveText("Not ready");
  await visualExpect(page.locator(".absent")).toHaveCount(0);
  const failed = visualExpect(page.locator("#late").describe("Expected status: All devices assigned"))
    .toHaveText("All devices assigned", { timeout: 500 });
  await expect(page.locator("#fabric-e2e-verification-overlay")).toBeAttached();
  await expect(page.getByText("All devices assigned", { exact: true })).toHaveCount(0);
  await expect(page.locator('#fabric-e2e-verification-overlay[data-check-state="failed"]')).toBeAttached();
  await page.screenshot({ path: info.outputPath("verification-overlay.png") });
  let message = "";
  try { await failed; } catch (error) {
    if (!(error instanceof Error)) throw error;
    message = error.message;
  }
  expect(message).toContain("All devices assigned");
  expect(message).toContain("Waiting");
  await expect(page.locator("#fabric-e2e-verification-overlay")).toHaveCount(0);
  await visualExpect(status).toHaveAttribute("id", "status");
  await expect(page.locator("#fabric-e2e-verification-overlay")).toHaveCount(0);
});

for (const [path, scenario] of [
  ["regular", "pageScopeRegular"], ["combined", "pageScopeCombined"], ["linked", "pageScopeLinked"],
] as const) {
  test(scenarioTitle(scenario), async ({ page }) => {
    const fixture = modelFixture(path);
    await page.setContent("<iframe></iframe>");
    const frame = page.frames()[1];
    if (!frame) throw new Error("Missing isolated test frame.");
    await frame.setContent(`
      <h2>Filter events and check model coverage</h2>
      <label><input type="radio" name="reference">Require the model ID</label>
      <label><input type="radio" name="reference">Model ID or additional conditions</label>
      <label><input type="radio">Add conditions</label>
      <textarea id="hub-telemetry-scope"></textarea><textarea id="hub-twin-scope"></textarea>
      ${(["telemetry", "twin"] as const).map(stream => `
        <section data-coverage-stream="${stream}"><span role="status">All fields found</span>
          <div data-observed-fields aria-label="Selected model fields found: ${fixture.fields.filter(field => field.role === stream).length} of ${fixture.fields.filter(field => field.role === stream).length} in this sample."></div>
          <div data-sample-count><span class="iot-solution-coverage-number">2</span></div>
        </section>`).join("")}
      <button>Check coverage</button><button>Continue</button>`);
    await modeling.modelCoverage(page.frameLocator("iframe"), fixture, path);
    await expect(frame.locator("#hub-telemetry-scope")).toHaveValue(scenarioScope(path).telemetry);
    await expect(frame.locator("#hub-twin-scope")).toHaveValue(scenarioScope(path).twin);
  });
}

test(scenarioTitle("pageAssignCombined"), async ({ page, context }) => {
  const plan = createPlan(config), fixture = modelFixture("combined");
  const deviceId = `${plan.azure.hubs.find(hub => hub.key === "linked1")!.name}-device`;
  page.setDefaultTimeout(1000);
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated test frame.");
  await frame.setContent(`
    <h2>Assign the model to devices</h2>
    <button onclick="setTimeout(()=>document.querySelector('#show-all').disabled=false,1500)">Assign model to devices</button>
    <button id="show-all" disabled>Show all devices</button>
    <table aria-label="Device model assignments"><tbody>
      ${plan.azure.hubs.filter(hub => hub.namespaceKey).map((hub, index) => `<tr><td>${hub.name}-device</td><td>
        <span>${["None", "Assigned to this model", "Assigned to another model"][index]}</span>
        ${index ? `<span>${index === 1 ? fixture.id : otherModelReference}</span>` : `<input type="checkbox" aria-label="Select ${hub.name}-device">`}
      </td></tr>`).join("")}
    </tbody></table>
    <button onclick="document.querySelector('[role=dialog]').hidden=false">Assign to selected devices</button>
    <div role="dialog" hidden><h2>Assign ${fixture.title} to 1 device?</h2>
      <p>Model ID: ${fixture.id}</p><p>${deviceId}</p>
      <button id="confirm" data-clicks="0" onclick="this.dataset.clicks=String(Number(this.dataset.clicks)+1);this.parentElement.hidden=true;
        const progress=document.querySelector('[aria-label=\\'Device assignment progress\\']');
        const result=document.querySelector('[data-assignment-state]');
        result.setAttribute('data-assignment-state','unknown');
        result.innerHTML=document.querySelector('#rechecking-indicator').innerHTML">Assign</button>
    </div><section aria-label="Device assignment progress" aria-busy="true"><div data-assignment-state="checking">
      <svg aria-label="Assignment not confirmed" data-assignment-warning width="24" height="24"></svg>
      <span class="iot-solution-creation-step-label warning">${deviceId}</span><p>Previous status read failed</p></div>
      <div id="summary" data-assignment-summary="confirmed" hidden>Model assigned to 1 device</div></section>
    <template id="rechecking-indicator"><span role="progressbar" aria-label="Checking assignment for ${deviceId}" style="display:inline-block;width:12px;height:12px"></span>
      <span class="iot-solution-creation-step-label running">${deviceId}</span></template>
    <a href="${documentationLinks.adr}" target="_blank" rel="noopener noreferrer">Learn how to link an IoT hub to ADR</a>
    <a href="${documentationLinks.dps}" target="_blank" rel="noopener noreferrer">Learn how to update DPS enrollments</a>
    <button id="next" disabled onclick="this.dataset.finished=String(document.querySelector('[data-assignment-state]').getAttribute('data-assignment-state')==='confirmed'&&document.querySelector('[aria-label=\\'Device assignment progress\\']').getAttribute('aria-busy')==='false')">Continue</button>`);
  const before = page.url();
  const workflow = modeling.deviceModelAssignment(page.frameLocator("iframe"), fixture, "combined", plan);
  const progress = frame.getByRole("region", { name: "Device assignment progress", exact: true });
  await expect(progress.getByLabel(`Checking assignment for ${deviceId}`, { exact: true })).toBeVisible();
  await expect(progress.locator("[data-assignment-state]")).toHaveAttribute("data-assignment-state", "unknown");
  await expect(progress).toHaveAttribute("aria-busy", "true");
  await expect(progress.locator("[data-assignment-warning]")).toHaveCount(0);
  await expect(progress.getByText("Previous status read failed", { exact: true })).toHaveCount(0);
  await expect(frame.locator("#next")).toBeDisabled();
  await progress.evaluate((element, deviceId) => {
    const row = element.querySelector<HTMLElement>("[data-assignment-state]");
    const summary = element.querySelector<HTMLElement>("#summary");
    const next = element.ownerDocument.querySelector<HTMLButtonElement>("#next");
    if (!row || !summary || !next) throw new Error("Missing mocked assignment completion controls.");
    row.setAttribute("data-assignment-state", "confirmed");
    row.innerHTML = `<svg aria-label="Assignment confirmed" width="24" height="24"></svg>
      <span class="iot-solution-creation-step-label success">${deviceId}</span>`;
    element.setAttribute("aria-busy", "false");
    summary.hidden = false;
    next.disabled = false;
  }, deviceId);
  await workflow;
  await expect(frame.getByRole("checkbox", { checked: true })).toHaveCount(1);
  await expect(frame.locator("#confirm")).toHaveAttribute("data-clicks", "1");
  await expect(frame.locator("#next")).toHaveAttribute("data-finished", "true");
  expect(context.pages()).toHaveLength(1);
  expect(page.url()).toBe(before);
});

test(scenarioTitle("pageAssignmentOutcome"), async ({ page }) => {
  const plan = createPlan(config), fixture = modelFixture("combined");
  const hub = plan.azure.hubs.find(hub => hub.key === "linked1")!;
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated assignment-outcome frame.");
  for (const state of ["failed", "blocked", "cancelled", "unknown"]) {
    await frame.setContent(`<h2>Assign the model to devices</h2>
      <button>Assign model to devices</button><button>Show all devices</button>
      <a href="${documentationLinks.adr}" target="_blank" rel="noopener noreferrer">Learn how to link an IoT hub to ADR</a>
      <table aria-label="Device model assignments"><tbody>
      ${plan.azure.hubs.filter(hub => hub.namespaceKey).map((hub, index) => `<tr><td>${hub.name}-device</td><td>
        <span>${["None", "Assigned to this model", "Assigned to another model"][index]}</span>
        ${index ? `<span>${index === 1 ? fixture.id : otherModelReference}</span>` : `<input type="checkbox" aria-label="Select ${hub.name}-device">`}
      </td></tr>`).join("")}</tbody></table>
      <button onclick="document.querySelector('[role=dialog]').hidden=false">Assign to selected devices</button>
      <div role="dialog" hidden><h2>Assign ${fixture.title} to 1 device?</h2>
        <p>Model ID: ${fixture.id}</p><p>${hub.name}-device</p>
        <button data-clicks="0" id="confirm" onclick="this.dataset.clicks=String(Number(this.dataset.clicks)+1);this.parentElement.hidden=true">Assign</button></div>
      <section aria-label="Device assignment progress" aria-busy="false"><div data-assignment-state="${state}">
        <svg aria-label="Assignment not confirmed" ${state === "unknown" ? "data-assignment-warning" : ""} width="24" height="24"></svg>
        <span class="iot-solution-creation-step-label ${state === "unknown" ? "warning" : "error"}">${hub.name}-device</span>
        Completion ${state}; primaryKey=fake-assignment-key</div></section>
      <button id="refresh" data-clicks="0" onclick="this.dataset.clicks=String(Number(this.dataset.clicks)+1)">Refresh results</button>
      <button id="next" data-clicks="0" onclick="this.dataset.clicks=String(Number(this.dataset.clicks)+1)">Continue</button>`);
    let message = "";
    try { await modeling.deviceModelAssignment(page.frameLocator("iframe"), fixture, "combined", plan); } catch (failure) {
      if (!(failure instanceof Error)) throw failure;
      message = failure.message;
    }
    expect(message).toContain(`[model-assignment] Assignment operation settled as ${state}`);
    expect(message).not.toContain("fake-assignment-key");
    await expect(frame.getByLabel("Assignment not confirmed", { exact: true })).toBeVisible();
    await expect(frame.locator("[data-assignment-warning]")).toHaveCount(state === "unknown" ? 1 : 0);
    await expect(frame.locator("#confirm")).toHaveAttribute("data-clicks", "1");
    await expect(frame.locator("#refresh")).toHaveAttribute("data-clicks", "0");
    await expect(frame.locator("#next")).toHaveAttribute("data-clicks", "0");
  }
});

test(scenarioTitle("pageHubSelector"), async ({ page }) => {
  const hub = createPlan(config).azure.hubs[0]!;
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated Hub-picker frame.");
  await frame.setContent(`<h2>Select your IoT Hub</h2>
    <input id="hub-combobox" role="combobox" onkeydown="if(event.key==='ArrowDown')document.querySelector('[role=listbox]').hidden=false">
    <button id="refresh" aria-label="Refresh IoT Hubs" data-clicks="0" onclick="this.dataset.clicks=String(Number(this.dataset.clicks)+1)"><svg aria-hidden="true"></svg></button>
    <div role="listbox" hidden><div role="option" onclick="this.parentElement.hidden=true;document.querySelector('#status').textContent='IoT Hub validated successfully';document.querySelector('#next').disabled=false">${hub.name}</div></div>
    <p id="status"></p><button id="next" disabled onclick="this.dataset.finished='true'">Next</button>`);
  await ingestion.iotHub(page.frameLocator("iframe"), hub);
  await expect(frame.locator("#hub-combobox")).toHaveValue(hub.name);
  await expect(frame.locator("#refresh")).toHaveAttribute("data-clicks", "0");
  await expect(frame.locator("#next")).toHaveAttribute("data-finished", "true");
});

test(scenarioTitle("pageReuse"), async ({ page }) => {
  const path = ingestionPath(createPlan(config), "regular");
  page.setDefaultTimeout(1000);
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated test frame.");
  await frame.setContent(`<h2>Create or reuse raw data tables</h2>
    <label><input type="radio" onchange="setTimeout(()=>document.querySelector('#pickers').hidden=false,1500)">Reuse existing tables</label>
    <section id="pickers" hidden>
    ${[["reuse-telemetry", path.telemetryTable, path.telemetryStream], ["reuse-properties", path.propertiesTable, path.propertiesStream]]
      .map(([id, table, stream]) => `<button id="${id}" role="combobox" onclick="this.nextElementSibling.hidden=false">Choose table</button>
      <div role="listbox" hidden>
        <div role="option"><span>${table}_other</span> <span>via ${stream}</span></div>
        <div role="option" onclick="document.getElementById('${id}').textContent='${table}';this.parentElement.hidden=true"><span>${table}</span> <span>via ${stream}</span></div>
      </div>`).join("")}</section>
    <button onclick="document.getElementById('status').textContent='Existing tables selected';document.getElementById('next').disabled=false">Use selected tables</button>
    <span id="status"></span><button id="next" disabled>Next</button>`);
  await ingestion.rawDataTables(page.frameLocator("iframe"), path, false);
  await expect(frame.locator("#reuse-telemetry")).toHaveText(path.telemetryTable);
  await expect(frame.locator("#reuse-properties")).toHaveText(path.propertiesTable);
});

test(scenarioTitle("pageComponents"), async ({ page }) => {
  for (const id of ["regular", "combined"] as const) {
    const plan = createPlan(config), path = ingestionPath(plan, id);
    await page.setContent("<iframe></iframe>");
    const frame = page.frames()[1];
    if (!frame) throw new Error("Missing isolated test frame.");
    const fields = dashboardSelection(id);
    await frame.setContent(`<h2>Select data source and fields</h2>
      ${[["dash-eh-select", plan.fabric.eventhouseName], ["dash-db-select", plan.fabric.databaseName],
        ["dash-telemetry-table", path.telemetryTable], ["dash-properties-table", path.propertiesTable]]
        .map(([selector, value]) => `<button id="${selector}" role="combobox" onclick="this.nextElementSibling.hidden=false">Choose</button>
          <div role="listbox" hidden><div role="option" onclick="document.getElementById('${selector}').textContent='${value}';this.parentElement.hidden=true">${value}</div></div>`).join("")}
      <button onclick="document.getElementById('fields').hidden=false">Inspect Tables</button>
      <section id="fields" hidden>
        ${["Telemetry Fields", "Property Fields"].map((title, index) => `<div class="iot-solution-config-card">${title}<span data-count="${index}"></span>
          ${fields.filter(field => (field.visual === "card") === (index === 1)).map(field => `<div class="iot-solution-field-list-item">
            <label><input type="checkbox" checked data-group="${index}" onchange="updateCounts()">${field.name}</label>
            ${field.visual === "line" ? "<span>Time-series</span>" : ""}</div>`).join("")}</div>`).join("")}
      </section><button id="next" disabled>Next</button>
      <script>function updateCounts() {
        for(let i=0;i<2;i++) document.querySelector('[data-count="'+i+'"]').textContent=document.querySelectorAll('input[data-group="'+i+'"]:checked').length+' selected';
        document.getElementById('next').disabled=document.getElementById('fields').hidden||!document.querySelector('input:checked');
      } updateCounts();</script>`);
    await dashboardPages.dataSource(page.frameLocator("iframe"), plan, path);
    await expect(frame.getByRole("checkbox", { checked: true })).toHaveCount(7);
    await frame.setContent(`<h2>Generate dashboard</h2>
      <label>Dashboard name<input oninput="document.querySelector('#create').disabled=!this.value.trim()"></label>
      <button id="create" disabled onclick="document.querySelector('#result').textContent='Dashboard &quot;owned-dashboard&quot; created successfully.';document.querySelector('#next').disabled=false">Generate Dashboard</button>
      <p id="result"></p><button id="next" disabled>Next</button>`);
    await dashboardPages.generateDashboard(page.frameLocator("iframe"), "owned-dashboard");
  }
});

for (const [path, scenario] of [["regular", "pageAssignRegular"], ["linked", "pageAssignLinked"]] as const) {
  test(scenarioTitle(scenario), async ({ page, context }) => {
    await page.setContent("<iframe></iframe>");
    const frame = page.frames()[1];
    if (!frame) throw new Error("Missing isolated test frame.");
    await frame.setContent(`<h2>Assign the model to devices</h2><button>Assign model to devices</button>
      ${path === "regular" ? `<section aria-label="Source IoT Hubs">Your IoT hubs aren't linked to Azure Device Registry</section>
        <a href="${documentationLinks.adr}" target="_blank" rel="noopener noreferrer">Learn how to link an IoT hub to ADR</a>
        <p data-assignment-unavailable-guidance>Continue</p><button data-testid="wizard-previous-btn">Previous</button>`
        : `<p data-all-devices-assigned>All devices already have a model assigned</p>
        <a href="${documentationLinks.dps}" target="_blank" rel="noopener noreferrer">Learn how to update DPS enrollments</a>
        <button>Continue</button>`}`);
    const before = page.url();
    await modeling.deviceModelAssignment(page.frameLocator("iframe"), modelFixture("combined"), path, createPlan(config));
    expect(context.pages()).toHaveLength(1);
    expect(page.url()).toBe(before);
  });
}

test(scenarioTitle("pageModelFields"), async ({ page }) => {
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated model-fields frame.");
  for (const path of ["regular", "combined"] as const) {
    const fixture = modelFixture(path);
    await frame.setContent(`<h2>Add a device model</h2>
      <label><input type="radio" name="format" checked>WoT Thing Model (recommended)</label>
      <label><input type="radio" name="format">Digital Twin Definition Language (DTDL)</label>
      ${fixture.format === "wot" ? `<ul aria-label="Supported WoT models">
        <li><strong>Supported model files:</strong> a self-contained WoT 1.1 Thing Model object or an array of related Thing Models. Each model needs its own supplied absolute ID.</li>
        <li><strong>Supported capabilities:</strong> properties and event data using primitive, enum, and complex schemas.</li>
        <li><strong>Not included:</strong> actions, submodels, and external definitions.</li></ul>` : ""}
      <input type="file" aria-label="Device model file" onchange="document.querySelector('#model-upload-reference').textContent='${fixture.id}';document.querySelector('.iot-solution-model-fields-inline').hidden=false;updateNext()">
      <span id="model-upload-reference"></span>
      ${fixture.format === "wot" ? '<p id="model-upload-warnings">Submodels: thermostat, powerModule</p>' : ""}
      <section class="iot-solution-model-fields-inline" hidden>
        <h3>3. Select fields</h3>
        <button onclick="for(const checkbox of document.querySelectorAll('[data-field-id] input'))checkbox.checked=false;updateNext()">Deselect all</button>
        ${fixture.fields.map(field => `<div data-field-id="${field.id}" data-role="${fixture.format === "dtdl" ? field.role : ""}">
          <label><input type="checkbox" checked onchange="updateNext()">${field.column}</label>
          ${fixture.format === "dtdl" ? `<span>${field.role === "telemetry" ? "Telemetry" : "Twin Property"}</span>` :
            `<button role="combobox" onclick="this.nextElementSibling.hidden=false">Choose a Role</button>
             <div role="listbox" hidden>${[["telemetry", "Telemetry"], ["twin", "Twin Property"]].map(([value, label]) =>
               `<div role="option" data-role="${value}" onclick="this.closest('[data-field-id]').dataset.role=this.dataset.role;this.parentElement.previousElementSibling.textContent=this.textContent;this.parentElement.hidden=true;updateNext()">${label}</div>`).join("")}</div>`}
        </div>`).join("")}
        <div data-field-id="not-selected" data-role="telemetry"><label><input type="checkbox" checked onchange="updateNext()">Unselected field</label></div>
      </section><button id="next" disabled onclick="this.dataset.finished='true'">Continue</button>
      <script>function updateNext(){
        const selected=[...document.querySelectorAll('[data-field-id]')].filter(row=>row.querySelector('input').checked);
        document.querySelector('#next').disabled=!selected.length||selected.some(row=>!row.dataset.role);
      }</script>`);
    await modeling.deviceModel(page.frameLocator("iframe"), fixture);
    await expect(frame.locator("#next")).toHaveAttribute("data-finished", "true");
    await expect(frame.locator('[data-field-id="not-selected"] input')).not.toBeChecked();
    await expect(frame.getByRole("checkbox", { checked: true })).toHaveCount(fixture.fields.length);
  }
});

test(scenarioTitle("pageModelCreation"), async ({ page }) => {
  const names = modeledNames(ingestionPath(createPlan(config), "regular"));
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated modeled-creation frame.");
  await frame.setContent(`<h2>Create modeled data</h2>
    ${[["modeled-output-name", "Modeled data table"], ["twin-changes-name", "Normalized properties table"],
      ["twin-state-name", "Latest property values (materialized view)"]].map(([id, label]) =>
      `<label>${label}<input id="${id}" oninput="document.querySelector('#create').disabled=[...document.querySelectorAll('input')].some(input=>!input.value.trim())"></label>`).join("")}
    <button id="create" disabled data-clicks="0" onclick="this.dataset.clicks=String(Number(this.dataset.clicks)+1);document.querySelector('#modeled-data-success').hidden=false;document.querySelector('#next').disabled=false">Create modeled data</button>
    <p id="modeled-data-success" role="status" hidden>Modeled data created. Select Continue to see what was created and what to do next.</p>
    <button id="next" disabled onclick="this.dataset.finished='true'">Continue</button>`);
  await modeling.modeledData(page.frameLocator("iframe"), names);
  await expect(frame.locator("#modeled-output-name")).toHaveValue(names.modeled);
  await expect(frame.locator("#twin-changes-name")).toHaveValue(names.normalized);
  await expect(frame.locator("#twin-state-name")).toHaveValue(names.lkv);
  await expect(frame.locator("#create")).toHaveAttribute("data-clicks", "1");
  await expect(frame.locator("#next")).toHaveAttribute("data-finished", "true");
});

test(scenarioTitle("pageModelReady"), async ({ page }) => {
  const names = modeledNames(ingestionPath(createPlan(config), "regular"));
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated modeled-results frame.");
  await frame.setContent(`<section id="ready"><h2>Your modeled data is ready</h2><h3>What was created</h3>
    <div class="iot-solution-validation iot-solution-ready-resources">
      <p>Modeled data table: ${names.modeled}</p><p>Normalized properties table: ${names.normalized}</p>
      <p>Latest property values (materialized view): ${names.lkv}</p>
    </div>
    <a href="${documentationLinks.activator}" target="_blank" rel="noopener noreferrer">Learn about Activator</a>
    <a href="${documentationLinks.dataAgent}" target="_blank" rel="noopener noreferrer">Learn about Fabric data agents</a>
    <button onclick="document.querySelector('#ready').hidden=true;document.querySelector('#home').hidden=false">Finish and Close</button></section>
    <button id="home" hidden>Model your device data</button>`);
  await modeling.readyToUse(page.frameLocator("iframe"), names);
  await expect(frame.locator("#home")).toBeVisible();
});

test(scenarioTitle("pageIngestionOverview"), async ({ page, context }) => {
  await page.setContent("<iframe></iframe>");
  const frame = page.frames()[1];
  if (!frame) throw new Error("Missing isolated ingestion-overview frame.");
  await frame.setContent(`
    <aside>${["Overview", "IoT Hub", "Eventhouse", "Raw data tables", "Eventstreams", "IoT Hub routing", "Ready to use"]
      .map(title => `<h3 class="wizard-control__step-title">${title}</h3>`).join("")}</aside>
    <h2>Connect your IoT data to Microsoft Fabric</h2>
    <section class="iot-solution-prerequisites"><h3>What you will need</h3><ul>
      <li>An existing Azure IoT Hub with a system-assigned or user-assigned managed identity enabled. An identity is needed for IoT Hub to securely route data to your Fabric solution.</li>
      <li>Optionally, an Azure Device Registry (ADR) namespace <a href="${documentationLinks.adr}" target="_blank" rel="noopener noreferrer">linked to your IoT Hub</a>. Linking an ADR namespace to your IoT Hub lets you assign a model to your devices later, when you model your data.</li>
      <li>An existing Eventhouse and KQL Database in this Fabric workspace to store the ingested device data used for analytics and operational intelligence.</li>
      <li>Permission to read and configure the IoT Hub, create Fabric items, and manage workspace access for the IoT Hub identity. These permissions allow the accelerator to securely configure the data path on your behalf.</li>
    </ul></section>
    <button onclick="this.dataset.finished='true'">Next</button>
    <section aria-label="DPS enrollment guidance"><a href="${documentationLinks.dps}" target="_blank" rel="noopener noreferrer">Learn how to update DPS enrollments</a></section>`);
  const before = page.url();
  await ingestion.overview(page.frameLocator("iframe"));
  await expect(frame.locator(".iot-solution-prerequisites li").nth(1)).toHaveText(ingestionAdrPrerequisite);
  await expect(frame.getByRole("button", { name: "Next", exact: true })).toHaveAttribute("data-finished", "true");
  await expectExternalDocumentationLink(frame.getByRole("link", { name: "Learn how to update DPS enrollments", exact: true }),
    documentationLinks.dps);
  await expect(frame.locator("#dps-guidance-unavailable")).toHaveCount(0);
  expect(context.pages()).toHaveLength(1);
  expect(page.url()).toBe(before);
});

for (const [kind, scenario] of [
  ["dashboard", "pageDashboard"], ["agent", "pageAgent"], ["activator", "pageActivator"],
] as const) {
  test(scenarioTitle(scenario), async ({ page, context }) => {
    const plan = createPlan(config), manifest = await writePlan(plan);
    const directory = dirname(manifest);
    try {
      const path = ingestionPath(plan, "regular"), deviceId = `${plan.azure.hubs[0]!.name}-device`;
      const id = "00000000-0000-4000-8000-000000000009";
      const prefix = `${plan.azure.resourceGroup.slice("fabric-".length)}-`;
      const nativeName = `${prefix}regular-${kind === "agent" ? "data-agent" : kind === "activator" ? "activator-00000000" : "dashboard"}`;
      plan.observed.itemId = "00000000-0000-4000-8000-000000000008";
      if (kind === "dashboard") plan.observed["regular.dashboardId"] = id;
      if (kind === "agent") plan.observed["regular.dataAgentId"] = id;
      if (kind === "activator") {
        plan.observed["regular.activatorId"] = id;
        plan.observed["regular.activatorName"] = nativeName;
      }
      let openOwnedName = plan.fabric.itemName;
      const closedNames: string[] = [], nativeNavigationHadCapacity: boolean[] = [];
      await page.exposeFunction("recordOwnedTabClose", (name: string) => {
        closedNames.push(name);
        openOwnedName = "";
      });
      const sidebar = () => `<nav role="menu" aria-label="Sidebar">
        ${Array.from({ length: 19 }, (_, index) => `<div role="menuitem" aria-label="Foreign item ${index}" data-foreign-tab>${index}</div>`).join("")}
        ${openOwnedName ? `<div role="menuitem" aria-label="${openOwnedName} - Workspace">${openOwnedName}
          <button data-testid="tab-close" onclick="window.recordOwnedTabClose(${JSON.stringify(openOwnedName).replace(/"/g, "&quot;")}).then(()=>this.parentElement.remove())"></button></div>` : ""}
        </nav>`;
      let html = "", agentReply = "Firmware: 1.0.0";
      if (kind === "dashboard") html = `
        <button aria-label="Device ID">Choose device</button>
        <button onclick="document.querySelector('#data').hidden=false;document.querySelector('#data').setAttribute('aria-hidden','true');document.querySelector('#tip').hidden=false">${deviceId}</button>
        <div id="tip" role="dialog" hidden><strong>New &middot; </strong><span id="tip-title">Search this table</span><button onclick="dismissTip()">Dismiss</button></div>
        <div id="data" hidden>${dashboardSelection("regular").map(field => `<section role="region" aria-label="${field.name}"><h3>${field.name}</h3>
          ${field.visual === "line"
            ? `<section role="region" aria-label="Chart. Highcharts interactive chart."><p>Line chart with 2 data points. Data ranges from ${field.expected} to ${field.expected}.</p><svg role="img" aria-label="Interactive chart" style="visibility:hidden"></svg></section>`
            : `<span>${field.expected}</span>`}</section>`).join("")}</div>
          <script>function dismissTip() {
            const title=document.querySelector('#tip-title'), data=document.querySelector('#data'), tip=document.querySelector('#tip');
            if(title.textContent==='Search this table'){title.textContent='Maximize this visual';return;}
            tip.hidden=true;data.removeAttribute('aria-hidden');
            if(!document.body.dataset.lateTip) {
              document.body.dataset.lateTip='true';
              setTimeout(()=>{title.textContent='Maximize this visual';tip.hidden=false;data.setAttribute('aria-hidden','true');},100);
              setTimeout(()=>document.querySelectorAll('svg').forEach(image=>image.style.visibility='visible'),400);
            }
          }</script>`;
      if (kind === "agent") html = `<p>Example firmware: 1.0.0</p>
        <textarea onkeydown="if(event.key==='Enter'){event.preventDefault();document.querySelector('#answer').textContent='Firmware: 1.0.0';}"></textarea><p id="answer"></p>`;
      if (kind === "activator") html = `<div role="tree" aria-label="Artifact document tree">
          <div role="treeitem" aria-label="Main"><div role="group"><div role="treeitem" onclick="document.querySelector('table').hidden=false">owned-</div></div></div>
        </div><button role="tab">Live feed</button>
        <table hidden><tr><th>deviceId</th><th>firmwareVersion</th><th>___subject</th></tr><tr><td>${deviceId}</td><td>1.0.0</td><td>/streams/owned-stream</td></tr></table>
        <div role="dialog" aria-label="Start monitoring your real-time data"><button onclick="this.parentElement.remove()">Close dialog</button></div>`;
      await context.route("**/*", route => {
        const url = new URL(route.request().url()), host = url.hostname;
        const main = route.request().isNavigationRequest() && route.request().frame() === page.mainFrame();
        const workspace = new URL(plan.target.workspaceUrl);
        if (main && host === workspace.hostname && url.pathname === workspace.pathname) {
          return route.fulfill({ contentType: "text/html", body: sidebar() });
        }
        if (main && host === workspace.hostname) {
          nativeNavigationHadCapacity.push(openOwnedName === "");
          openOwnedName = nativeName;
        }
        let body = host === "dataexplorer.azure.com" || host === "login.microsoftonline.com"
          ? '<body><script>sessionStorage.setItem("synthetic-frame", "local-only")</script></body>'
          : kind === "dashboard" && host === "native-check.powerbi.com"
            ? html + `<script>if(!location.search) document.querySelector('[aria-label="Device ID"]').onclick=()=>parent.postMessage('remount-native-check',${JSON.stringify(new URL(plan.target.workspaceUrl).origin)});</script>`
            : kind === "dashboard"
              ? `<iframe id="native" src="https://native-check.powerbi.com/dashboard/${id}"></iframe>
                <script>addEventListener('message', event => {
                  if(event.origin==='https://native-check.powerbi.com' && event.data==='remount-native-check')
                    document.querySelector('#native').src='https://native-check.powerbi.com/dashboard/${id}?ready=1';
                });</script>`
              : kind === "agent" ? html.replace("Firmware: 1.0.0", agentReply) : html;
        if (main && host === workspace.hostname) body = sidebar() + body;
        return route.fulfill({ contentType: "text/html", body });
      });
      if (kind === "dashboard") {
        await native.prepareDashboard(page, plan, path, 5000);
        await page.evaluate(() => {
          sessionStorage.setItem("synthetic-workspace", "local-only");
          for (const host of ["dataexplorer.azure.com", "login.microsoftonline.com"]) {
            const frame = document.createElement("iframe");
            frame.src = `https://${host}/local-auth-state-check`;
            document.body.append(frame);
          }
        });
        await expect.poll(() => page.frames().filter(frame => frame.url().includes("local-auth-state-check")).length).toBe(2);
        for (const frame of page.frames().slice(1)) await frame.locator("body").waitFor();
        const sessions = await authenticatedAppSessions(page, plan.target.workspaceUrl);
        expect(sessions[new URL(plan.target.workspaceUrl).origin]?.["synthetic-workspace"]).toBe("local-only");
        expect(sessions["https://dataexplorer.azure.com"]?.["synthetic-frame"]).toBe("local-only");
        expect(sessions).not.toHaveProperty("https://login.microsoftonline.com");
      }
      if (kind === "dashboard") {
        await native.dashboard(page, plan, path);
        expect(page.frames().some(frame => frame.url() === `https://native-check.powerbi.com/dashboard/${id}?ready=1`)).toBe(true);
        await expect(page.frameLocator("#native").locator("body")).toHaveAttribute("data-late-tip", "true");
        await expect(page.frameLocator("#native").getByRole("dialog")).toHaveCount(0);
      }
      if (kind === "agent") {
        await native.dataAgent(page, plan, path, "Modeled_owned", deviceId);
        const evidence = JSON.parse(await readFile(join(directory, "data-agent-prompt-regular.json"), "utf8"));
        expect(evidence.prompt).not.toContain("1.0.0");
        expect(evidence.answer.length).toBeGreaterThan(1);
        expect(evidence.newestAnswer).toBe("Firmware: 1.0.0");
        agentReply = "Firmware: 11.0.0";
        await expect(native.dataAgent(page, plan, path, "Modeled_owned", deviceId)).rejects.toThrow("exact fixture firmware");
      }
      if (kind === "activator") await native.activator(page, plan, id, "owned-stream", deviceId);
      expect(new URL(page.url()).searchParams.get("ctid")).toBe(config.tenantId);
      expect(nativeNavigationHadCapacity).toEqual(kind === "activator" ? [true] : [true, true]);
      expect(closedNames).toEqual(kind === "activator" ? [plan.fabric.itemName] : [plan.fabric.itemName, nativeName]);
      await expect(page.locator("[data-foreign-tab]")).toHaveCount(19);
    } finally {
      for (const file of await readdir(directory)) await unlink(join(directory, file));
      await rmdir(directory);
    }
  });
}
