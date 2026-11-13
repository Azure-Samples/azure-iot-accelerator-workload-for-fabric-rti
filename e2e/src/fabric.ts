import { setTimeout as pause } from "node:timers/promises";
import { Cloud, object, text } from "./cloud.js";
import { requireGuid } from "./config.js";
import { updatePlan, type RunPlan } from "./plan.js";
import { ensureRunFolder, moveOwnedItem } from "./folders.js";

export function operationFromHeaders(headers: Headers): string {
  const id = requireGuid(text(headers.get("x-ms-operation-id"), "Fabric operation ID"), "Fabric operation ID");
  return `/operations/${id}`;
}

export async function reconcileFabricCreates(cloud: Cloud, plan: RunPlan): Promise<void> {
  for (const [key, name, type] of [
    ["eventhouseId", plan.fabric.eventhouseName, "Eventhouse"],
    ["databaseId", plan.fabric.databaseName, "KQLDatabase"],
  ] as const) {
    const operation = plan.observed[`${key}Operation`];
    if (!operation) continue;
    if (!/^\/operations\/[0-9a-f-]{36}$/i.test(operation)) {
      throw new Error("[ownership] Invalid recorded Fabric creation operation.");
    }
    const state = object(await cloud.fabric(operation), "recorded Fabric creation operation");
    if (!["Succeeded", "Failed", "Canceled", "Cancelled"].includes(String(state.status))) {
      throw new Error(`[cleanup] ${type} creation is not terminal; retain its recorded operation and folder.`);
    }
    if (state.status === "Succeeded") {
      const result = object(await cloud.fabric(`${operation}/result`), "recorded Fabric creation result");
      const id = requireGuid(text(result.id, "created item ID"), "created item ID");
      if (result.displayName !== name || result.type !== type || result.workspaceId !== plan.target.workspaceId ||
          plan.observed[key] && plan.observed[key] !== id) {
        throw new Error("[ownership] Recorded Fabric creation result changed its item identity.");
      }
      plan.observed[key] = id;
    }
    plan.observed[`${key}CreationOperation`] = operation;
    plan.observed[`${key}CreationOperationState`] = text(state.status, "terminal creation state");
    delete plan.observed[`${key}Operation`];
    await updatePlan(plan);
  }
}

export async function prepareFabric(cloud: Cloud, plan: RunPlan): Promise<void> {
  const workspace = `/workspaces/${plan.target.workspaceId}`;
  const marker = `Fabric IoT release E2E run ${plan.runId}`;
  const folderId = await ensureRunFolder(cloud, plan);

  async function create(
    collection: "eventhouses" | "kqlDatabases",
    type: "Eventhouse" | "KQLDatabase",
    name: string,
    key: string,
    creationPayload?: Record<string, unknown>,
  ): Promise<string> {
    const recorded = plan.observed[key];
    if (recorded) {
      const current = object(await cloud.fabric(`${workspace}/${collection}/${requireGuid(recorded, key)}`), type);
      if (current.displayName !== name || current.description !== marker || current.workspaceId !== plan.target.workspaceId) {
        throw new Error(`[ownership] Recorded ${type} does not match this run.`);
      }
      return recorded;
    }
    let result: Record<string, unknown>;
    const response = await cloud.send(`https://api.fabric.microsoft.com/v1${workspace}/${collection}`,
      "https://api.fabric.microsoft.com", "POST", {
        displayName: name, description: marker,
        folderId,
        ...(creationPayload ? { creationPayload } : {}),
      });
    if (response.status === 202) {
      const operation = operationFromHeaders(response.headers);
      plan.observed[`${key}Operation`] = operation;
      await updatePlan(plan);
      const deadline = Date.now() + 15 * 60_000;
      let succeeded = false;
      while (Date.now() < deadline) {
        const current = object(await cloud.fabric(operation), "Fabric operation");
        if (current.status === "Succeeded") { succeeded = true; break; }
        if (current.status === "Failed" || current.status === "Canceled" || current.status === "Cancelled") {
          const details = current.error === undefined ? {} : object(current.error, "Fabric operation error");
          throw new Error(`[service] ${type} creation failed: ${String(details.errorCode ?? details.code ?? current.status)}`);
        }
        if (current.status !== "Running" && current.status !== "NotStarted") throw new Error("[service-shape] Unknown Fabric operation status.");
        await pause(10_000);
      }
      if (!succeeded) throw new Error(`[service] ${type} creation exceeded 15 minutes; the operation remains recorded in the manifest.`);
      result = object(await cloud.fabric(`${operation}/result`), `${type} creation result`);
      delete plan.observed[`${key}Operation`];
    } else {
      result = object(response.body, `${type} creation result`);
    }
    const id = requireGuid(text(result.id, `${type} ID`), `${type} ID`);
    if (result.displayName !== name || result.type !== type || result.workspaceId !== plan.target.workspaceId) {
      throw new Error(`[service-shape] Created ${type} does not match the requested name/type/workspace.`);
    }
    plan.observed[key] = id;
    await updatePlan(plan);
    if (type === "Eventhouse") await moveOwnedItem(cloud, plan, id, name);
    return id;
  }

  const eventhouseId = await create("eventhouses", "Eventhouse", plan.fabric.eventhouseName, "eventhouseId");
  await create("kqlDatabases", "KQLDatabase", plan.fabric.databaseName, "databaseId", {
    databaseType: "ReadWrite", parentEventhouseItemId: eventhouseId,
  });
  const eventhouse = object(await cloud.fabric(`${workspace}/eventhouses/${eventhouseId}`), "Eventhouse metadata");
  const properties = object(eventhouse.properties, "Eventhouse properties");
  const query = new URL(text(properties.queryServiceUri, "Kusto query URI"));
  if (query.protocol !== "https:" || query.username || query.password || query.port ||
      ![".kusto.fabric.microsoft.com", ".kusto.windows.net"].some(suffix => query.hostname.endsWith(suffix))) {
    throw new Error("[service-shape] Unexpected native Eventhouse Kusto endpoint.");
  }
  if (!Array.isArray(properties.databasesItemIds) || !properties.databasesItemIds.includes(plan.observed.databaseId)) {
    throw new Error("[ownership] The new database is not attached to this run's Eventhouse.");
  }
  plan.observed.queryServiceUri = query.origin;
  plan.phase = "fabric-destination-native-verified";
  await updatePlan(plan);
}
