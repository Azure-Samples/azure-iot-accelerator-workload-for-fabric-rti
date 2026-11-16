import { Cloud, object, text } from "./cloud.js";
import { requireGuid } from "./config.js";
import { updatePlan, type RunPlan } from "./plan.js";

export function folderName(plan: RunPlan): string {
  return `${plan.azure.resourceGroup.replace(/^fabric-/, "")}-items`;
}

export async function ensureRunFolder(cloud: Cloud, plan: RunPlan): Promise<string> {
  const root = `/workspaces/${plan.target.workspaceId}`;
  const recorded = plan.observed.folderId;
  const folder = object(await cloud.fabric(recorded
    ? `${root}/folders/${requireGuid(recorded, "run folder ID")}` : `${root}/folders`,
  recorded ? "GET" : "POST", recorded ? undefined : { displayName: folderName(plan) }), "run folder");
  if (folder.displayName !== folderName(plan) || folder.workspaceId !== plan.target.workspaceId || folder.parentFolderId) {
    throw new Error("[ownership] Folder does not match the dedicated per-run folder.");
  }
  const id = requireGuid(text(folder.id, "folder ID"), "folder ID");
  plan.observed.folderId = id;
  await updatePlan(plan);
  return id;
}

export async function moveOwnedItem(cloud: Cloud, plan: RunPlan, id: string, name: string): Promise<void> {
  const root = `/workspaces/${plan.target.workspaceId}`;
  const folderId = requireGuid(plan.observed.folderId, "run folder ID");
  const item = object(await cloud.fabric(`${root}/items/${requireGuid(id, "owned item ID")}`), "owned item");
  if (item.displayName !== name || !name.startsWith(plan.azure.resourceGroup.replace(/^fabric-/, "") + "-")) {
    throw new Error("[ownership] Refusing to move an item not named and identified by this run.");
  }
  if (item.folderId !== folderId) await cloud.fabric(`${root}/items/${id}/move`, "POST", { targetFolderId: folderId });
  const final = object(await cloud.fabric(`${root}/items/${id}`), "moved item");
  if (final.folderId !== folderId) throw new Error("[folder] Owned Fabric item is not inside the run folder.");
}
