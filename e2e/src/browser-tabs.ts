import { expect, type Page } from "@playwright/test";
import { requireGuid } from "./config.js";
import type { RunPlan } from "./plan.js";

export interface OwnedItemTab {
  id: string;
  name: string;
}

export function ownedItemTabs(plan: RunPlan): OwnedItemTab[] {
  requireGuid(plan.runId, "tab ownership run ID");
  if (plan.azure.tags["e2e-run-id"] !== plan.runId || plan.azure.tags["e2e-suite"] !== "fabric-iot-release" ||
      !/^fabric-e2e-[0-9a-f]{16}$/.test(plan.azure.resourceGroup) ||
      plan.azure.resourceGroupId !== `/subscriptions/${plan.target.subscriptionId}/resourceGroups/${plan.azure.resourceGroup}`) {
    throw new Error("[ownership] Cannot establish the current run's Fabric tab ownership.");
  }
  const prefix = `${plan.azure.resourceGroup.replace(/^fabric-/, "")}-`;
  const tabs: OwnedItemTab[] = [];
  const add = (name: string, id: string | undefined) => {
    if (!id) return;
    requireGuid(id, "recorded Fabric tab item ID");
    if (!name.startsWith(prefix)) throw new Error("[ownership] Recorded Fabric tab name is outside the current run.");
    const previous = tabs.find(tab => tab.id === id || tab.name === name);
    if (previous) {
      if (previous.id === id && previous.name === name) return;
      throw new Error("[ownership] Ambiguous recorded Fabric tab name/item ID.");
    }
    tabs.push({ id, name });
  };
  add(plan.fabric.itemName, plan.observed.itemId);
  add(plan.fabric.eventhouseName, plan.observed.eventhouseId);
  add(plan.fabric.databaseName, plan.observed.databaseId);
  for (const path of ["regular", "combined", "linked"]) {
    add(`${prefix}${path}-dashboard`, plan.observed[`${path}.dashboardId`]);
    add(`${prefix}${path}-data-agent`, plan.observed[`${path}.dataAgentId`]);
    add(plan.observed[`${path}.activatorName`] ?? "", plan.observed[`${path}.activatorId`]);
  }
  return tabs;
}

export async function closeOwnedItemTabs(page: Page, plan: RunPlan): Promise<OwnedItemTab[]> {
  const tabs = ownedItemTabs(plan);
  if (!tabs.length) return [];
  const workspace = new URL(plan.target.workspaceUrl);
  const workspacePath = `/groups/${requireGuid(plan.target.workspaceId, "tab ownership workspace ID")}`;
  if (workspace.protocol !== "https:" || !["app.powerbi.com", "app.fabric.microsoft.com"].includes(workspace.hostname) ||
      workspace.pathname !== workspacePath || workspace.searchParams.get("ctid") !== requireGuid(plan.target.tenantId, "tab ownership tenant ID")) {
    throw new Error("[ownership] Invalid workspace URL for closing current-run Fabric tabs.");
  }
  const current = new URL(page.url());
  if (current.origin !== workspace.origin || ![workspacePath, `${workspacePath}/list`].includes(current.pathname)) {
    await page.goto(workspace.href, { waitUntil: "domcontentloaded" });
  }
  const sidebar = page.getByRole("menu", { name: "Sidebar", exact: true });
  const escaped = (name: string) => name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const owned = sidebar.getByRole("menuitem", { name: new RegExp(`^(?:${tabs.map(tab => escaped(tab.name)).join("|")})(?: - |$)`) });
  const closed: OwnedItemTab[] = [];
  while (true) {
    await expect(sidebar).toBeVisible();
    if (!await owned.count()) break;
    if (closed.length >= 50) throw new Error("[browser-tabs] Too many owned tabs; refusing an unbounded UI cleanup.");
    const tab = owned.first(), element = await tab.elementHandle();
    if (!element) throw new Error("[browser-tabs] Owned tab disappeared before close.");
    try {
      const name = (await tab.innerText()).trim();
      const target = tabs.find(tab => tab.name === name);
      if (!target) throw new Error("[ownership] Fabric tab text differs from its recorded item name.");
      await expect(sidebar.getByRole("menuitem", { name: new RegExp(`^${escaped(name)}(?: - |$)`) })).toHaveCount(1);
      await tab.hover();
      const close = tab.getByTestId("tab-close");
      await expect(close).toHaveCount(1);
      await close.click();
      await element.waitForElementState("hidden", { timeout: 30_000 });
      closed.push(target);
    } finally { await element.dispose(); }
  }
  if (closed.length) console.log(`[browser-tabs] Closed ${closed.length} current-run Fabric item tabs: ${closed.map(tab =>
    `${tab.name} (${tab.id})`).join(", ")}. No resources or unrelated items changed.`);
  return closed;
}
