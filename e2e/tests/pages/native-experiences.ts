import { type Frame, type FrameLocator, type Locator, type Page } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { artifacts } from "../../src/config.js";
import { dashboardSelection, dashboardTipTitle, dashboardTipTitles } from "../../src/dashboard.js";
import { nativeItemUrl, type IngestionPath } from "../../src/paths.js";
import { updatePlan, type RunPlan } from "../../src/plan.js";
import { closeOwnedItemTabs } from "../../src/browser-tabs.js";
import { assertFirmwareAnswer } from "../../src/experience-verification.js";
import { redactSensitiveText } from "../../src/cloud.js";

type NativeSurface = Page | FrameLocator;
type NativeControl = (frame: Frame | NativeSurface) => Locator;

async function surface(page: Page, control: NativeControl, timeout = 90_000): Promise<NativeSurface> {
  let matches: Frame[] = [];
  await expect.poll(async () => {
    matches = [];
    for (const frame of page.frames()) if (await control(frame).filter({ visible: true }).count()) matches.push(frame);
    return matches.length;
  }, { timeout, message: "Expected exactly one native Fabric surface with its ready control." }).toBe(1);
  const frame = matches[0];
  if (!frame) throw new Error("[native-ui] Native Fabric surface is unavailable.");
  const selectors: string[] = [];
  for (let current = frame; current.parentFrame(); current = current.parentFrame()!) {
    const source = await (await current.frameElement()).getAttribute("src");
    if (!source) throw new Error("[native-ui] Native iframe has no navigable source.");
    const url = new URL(source, current.parentFrame()!.url());
    selectors.unshift(`iframe[src^=${JSON.stringify(url.origin + url.pathname)}]`);
  }
  let native: NativeSurface = page;
  for (const selector of selectors) native = native.locator(selector).filter({ visible: true }).contentFrame();
  await expect(control(native).filter({ visible: true })).toHaveCount(1, { timeout });
  return native;
}

const dashboardParameter: NativeControl = frame => frame.getByRole("combobox", { name: /Device ID/i })
  .or(frame.getByRole("button", { name: /^Device ID\b/i }));

async function dismissDashboardTips(frame: NativeSurface): Promise<void> {
  const dismissed = new Set<string>();
  const dialogs = frame.getByRole("dialog").filter({ visible: true });
  for (let index = 0; index < dashboardTipTitles.length; index++) {
    if (await dialogs.count() !== 1) break;
    const title = dashboardTipTitle(await dialogs.innerText());
    if (!title || dismissed.has(title)) break;
    const tip = dialogs.filter({ hasText: title });
    await tip.getByRole("button", { name: "Dismiss", exact: true }).describe(`Dismiss native ${title} first-use tip`).click();
    await expect(tip).toHaveCount(0);
    dismissed.add(title);
  }
  await expect(dialogs).toHaveCount(0);
}

export async function prepareDashboard(page: Page, plan: RunPlan, path: IngestionPath, timeout: number): Promise<void> {
  const id = plan.observed[`${path.id}.dashboardId`];
  if (!id) throw new Error("[prerequisite] Native dashboard identity has not been verified.");
  await closeOwnedItemTabs(page, plan);
  await page.goto(nativeItemUrl(plan, "kustodashboards", id), { waitUntil: "domcontentloaded" });
  await surface(page, dashboardParameter, timeout);
}

export async function dashboard(page: Page, plan: RunPlan, path: IngestionPath): Promise<void> {
  await prepareDashboard(page, plan, path, 90_000);
  const id = plan.observed[`${path.id}.dashboardId`]!;
  const frame = await surface(page, dashboardParameter);
  await dashboardParameter(frame).filter({ visible: true }).click();
  const hub = plan.azure.hubs.find(value => value.key === path.hubKeys[0]);
  if (!hub) throw new Error("[fixture] Missing dashboard device.");
  const deviceId = `${hub.name}-device`;
  await frame.getByText(deviceId, { exact: true }).click();
  for (const field of dashboardSelection(path.id)) {
    const tile = frame.getByRole("region", { name: field.name, exact: true, includeHidden: true })
      .describe(`${field.name}: rendered fixture value ${field.expected}`);
    await tile.scrollIntoViewIfNeeded();
    await dismissDashboardTips(frame);
    await expect(tile).toBeVisible();
    if (field.visual === "line") {
      const chart = tile.getByRole("region", { name: "Chart. Highcharts interactive chart.", exact: true, includeHidden: true });
      const image = chart.getByRole("img", { name: "Interactive chart", exact: true, includeHidden: true });
      await expect(image).toBeVisible({ timeout: 120_000 });
      await dismissDashboardTips(frame);
      await expect(image).toBeVisible();
      await expect(chart).toContainText(/Line chart with [1-9]\d* data points/);
      await expect(chart).toContainText(`Data ranges from ${field.expected} to ${field.expected}.`);
    } else {
      await expect(tile.getByText(String(field.expected), { exact: true }).first()).toBeVisible({ timeout: 120_000 });
      await dismissDashboardTips(frame);
      await expect(tile.getByText(String(field.expected), { exact: true }).first()).toBeVisible();
    }
    await dismissDashboardTips(frame);
    await expect(tile.getByRole("progressbar")).toHaveCount(0);
  }
  await expect(frame.getByText(/No data to display|Failed to execute query|Query execution failed/i)).toHaveCount(0);
  await writeFile(join(artifacts, "runs", plan.runId, `dashboard-rendered-${path.id}.json`),
    JSON.stringify({ id, deviceId, renderedFields: dashboardSelection(path.id).map(field => field.name),
      renderedValues: dashboardSelection(path.id).map(field => field.expected) }, null, 2));
  plan.verification[`${path.id}.dashboardRendered`] = "native-dashboard-selected-device-and-rendered-data-verified";
  await updatePlan(plan);
}

export async function dataAgent(page: Page, plan: RunPlan, path: IngestionPath, table: string, deviceId: string | undefined): Promise<void> {
  const id = plan.observed[`${path.id}.dataAgentId`];
  if (!id) throw new Error("[prerequisite] Native Data Agent identity has not been verified.");
  if (!deviceId) throw new Error("[prerequisite] No known fixture device exists in the agent grounding table.");
  await closeOwnedItemTabs(page, plan);
  await page.goto(nativeItemUrl(plan, "aiskills", id), { waitUntil: "domcontentloaded" });
  const composer: NativeControl = frame => frame.locator("textarea").filter({ visible: true });
  const frame = await surface(page, composer);
  const prompt = `What is the latest firmwareVersion for deviceId '${deviceId}' in table '${table}'? Read the data and return only the firmware version.`;
  const answer = frame.getByText(/\d+\.\d+\.\d+/).filter({ visible: true });
  const before = await answer.count();
  await composer(frame).fill(prompt);
  await composer(frame).press("Enter");
  await expect.poll(() => answer.count(), { timeout: 180_000, message: "The agent must return a new version response, not a pre-existing example." }).toBeGreaterThan(before);
  await expect(answer.last()).toBeVisible();
  const newestAnswer = await answer.last().innerText();
  assertFirmwareAnswer(newestAnswer);
  await expect(frame.getByText(/Unable to generate a response|Something went wrong|You don't have permission/i)).toHaveCount(0);
  await writeFile(join(artifacts, "runs", plan.runId, `data-agent-prompt-${path.id}.json`),
    JSON.stringify({ id, table, deviceId, prompt, answer: (await answer.allTextContents()).map(redactSensitiveText),
      newestAnswer: redactSensitiveText(newestAnswer), expected: "1.0.0" }, null, 2));
  plan.verification[`${path.id}.dataAgentPrompt`] = "native-prompt-returned-known-fixture-firmware";
  await updatePlan(plan);
}

export async function activator(page: Page, plan: RunPlan, id: string, streamName: string, deviceId: string | undefined): Promise<void> {
  if (!deviceId) throw new Error("[prerequisite] No fresh source device was verified for Activator.");
  await closeOwnedItemTabs(page, plan);
  await page.goto(nativeItemUrl(plan, "reflexes", id), { waitUntil: "domcontentloaded" });
  const explorer: NativeControl = frame => frame.getByRole("tree", { name: "Artifact document tree", exact: true });
  const introduction: NativeControl = frame => frame.getByRole("dialog", { name: "Start monitoring your real-time data", exact: true });
  const frame = await surface(page, value => introduction(value).or(explorer(value)).first());
  const dialog = introduction(frame);
  if (await dialog.isVisible()) {
    await dialog.getByRole("button", { name: "Close dialog", exact: true }).describe("Close Activator introduction without creating rules").click();
    await expect(dialog).toHaveCount(0);
  }
  const entries = explorer(frame).getByRole("treeitem");
  await expect(entries).toHaveCount(2);
  await expect(explorer(frame).getByRole("treeitem", { name: "Main", exact: true })).toHaveCount(1);
  await entries.last().describe("Open the single owned incoming event stream (native display name may be shortened)").click();
  await frame.getByRole("tab", { name: "Live feed", exact: true }).click();
  await expect(frame.getByText("deviceId", { exact: true }).first()).toBeVisible({ timeout: 120_000 });
  const row = frame.getByRole("row").filter({ hasText: deviceId }).filter({ hasText: "1.0.0" })
    .filter({ hasText: `/streams/${streamName}` });
  await expect(row.first()).toBeVisible({ timeout: 180_000 });
  await writeFile(join(artifacts, "runs", plan.runId, `activator-arrival-${id}.json`),
    JSON.stringify({ id, streamName, deviceId, incomingRows: await row.allTextContents(),
      scope: "Actual incoming Activator events, not rule firing or external actions" }, null, 2));
  plan.verification[`activator.${id}.arrival`] = "actual-incoming-activator-events-observed";
  await updatePlan(plan);
}
