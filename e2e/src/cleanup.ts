import { readdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import { Cloud, object, text, redactSensitiveText } from "./cloud.js";
import { artifacts, requireGuid } from "./config.js";
import { folderName } from "./folders.js";
import { ingestionPath } from "./paths.js";
import { assertRawStreamDestination } from "./routing.js";
import { assertOwnedPlan, updatePlan, type RunPlan } from "./plan.js";
import { reconcileFabricCreates } from "./fabric.js";
import { assertSimulationStopped } from "./process-owner.js";

export const FABRIC = "https://api.fabric.microsoft.com";

export async function workspaceItems(cloud: Cloud, plan: RunPlan): Promise<Record<string, unknown>[]> {
  const items: Record<string, unknown>[] = [];
  let continuation: string | undefined;
  do {
    const page = object(await cloud.fabric(`/workspaces/${plan.target.workspaceId}/items${continuation ? `?continuationToken=${encodeURIComponent(continuation)}` : ""}`), "workspace items");
    if (!Array.isArray(page.value)) throw new Error("[service-shape] Missing workspace item inventory.");
    items.push(...page.value.map(value => object(value, "Fabric item")));
    continuation = typeof page.continuationToken === "string" ? page.continuationToken : undefined;
  } while (continuation);
  return items;
}

export async function connections(cloud: Cloud): Promise<Record<string, unknown>[]> {
  const items: Record<string, unknown>[] = [];
  let continuation: string | undefined;
  do {
    const page = object(await cloud.request(`${FABRIC}/v1/connections${continuation ? `?continuationToken=${encodeURIComponent(continuation)}` : ""}`, FABRIC), "connections");
    if (!Array.isArray(page.value)) throw new Error("[service-shape] Missing connection inventory.");
    items.push(...page.value.map(value => object(value, "connection")));
    continuation = typeof page.continuationToken === "string" ? page.continuationToken : undefined;
  } while (continuation);
  return items;
}

export async function deleteFabricResource(cloud: Cloud, path: string, name: string, type?: string): Promise<void> {
  const workspace = `/workspaces/${cloud.config.workspaceId}/`;
  if (!path.startsWith(`${workspace}items/`) && !path.startsWith(`${workspace}folders/`) && !path.startsWith("/connections/")) {
    throw new Error("[ownership] Cleanup path is outside the configured workspace.");
  }
  requireGuid(text(path.split("/").at(-1), "cleanup ID"), "cleanup ID");
  const url = `${FABRIC}/v1${path}`;
  const current = await cloud.request(url, FABRIC, "GET", undefined, true);
  if (current === undefined) return;
  const resource = object(current, "resource before cleanup");
  if (resource.displayName !== name || type !== undefined && resource.type !== type) {
    throw new Error("[ownership] Cleanup resource identity/name/type changed; refusing deletion.");
  }
  await cloud.request(url, FABRIC, "DELETE");
  const deadline = Date.now() + 180_000;
  while (Date.now() < deadline) {
    if (await cloud.request(url, FABRIC, "GET", undefined, true) === undefined) return;
    await pause(2000);
  }
  throw new Error(`[cleanup] Deletion was not confirmed for ${name}.`);
}

interface CleanupReceipt {
  resource: string;
  status: "confirmed-absent" | "completed" | "failed";
  detail?: string;
}

export async function cleanupRun(
  cloud: Cloud, plan: RunPlan,
  cleanupActivatorAttempts: () => Promise<void>,
): Promise<void> {
  assertOwnedPlan(plan, cloud.config, plan.runId);
  try {
    await assertSimulationStopped(plan);
  } catch (error) {
    plan.phase = "cleanup-blocked";
    plan.observed.cleanupState = "failed";
    await updatePlan(plan);
    console.error("[cleanup] Local device shutdown has not been confirmed; no resource access was started.");
    throw error;
  }
  const prefix = plan.azure.resourceGroup.replace(/^fabric-/, "");
  const root = `/workspaces/${plan.target.workspaceId}`;
  const errors: Error[] = [];
  const receipts: CleanupReceipt[] = [];
  const attempt = async <T>(
    label: string, action: () => Promise<T>, success: "confirmed-absent" | "completed" = "confirmed-absent",
  ): Promise<T | undefined> => {
    try {
      const result = await action();
      receipts.push({ resource: label, status: success });
      console.log(`[cleanup] ${success === "completed" ? "Completed" : "Confirmed absent"}: ${label}`);
      return result;
    } catch (error) {
      const failure = new Error(redactSensitiveText(error instanceof Error ? error.message : String(error)));
      errors.push(failure);
      receipts.push({ resource: label, status: "failed", detail: failure.message });
      console.error(`[cleanup] ${label}: ${failure.message}`);
      return undefined;
    }
  };
  plan.phase = "cleanup-in-progress";
  await updatePlan(plan);
  try {
    // Collect only principals from native, tagged resources before their group disappears.
    const principals = new Set<string>();
    let foreignIdentity = false;
    const identitySources = [
      ...plan.azure.hubs.filter(hub => hub.key !== "regular2").map(hub => ({ key: hub.key, id: hub.id })),
      { key: "regular2", id: plan.azure.regularIdentity.id },
    ];
    await attempt("owned Hub workspace roles", async () => {
      for (const source of identitySources) {
        const principal = plan.observed[`${source.key}IdentityPrincipalId`];
        const resource = plan.observed[`${source.key}IdentityResourceId`];
        if (principal === undefined && resource === undefined) continue;
        if (resource !== source.id) throw new Error("[ownership] Recorded workspace principal is bound to another resource.");
        principals.add(requireGuid(principal, "recorded owned identity principal"));
      }
      for (const resource of [...plan.azure.hubs, plan.azure.regularIdentity]) {
        const native = await cloud.arm(resource.id, resource.type === "Microsoft.ManagedIdentity/userAssignedIdentities" ? "2023-01-31" : "2026-10-01-preview", "GET", undefined, true);
        if (native === undefined) {
          const source = identitySources.find(source => source.id === resource.id);
          if (source && plan.observed[resource.id] === "native-hub-provisioning-verified" &&
              !plan.observed[`${source.key}IdentityPrincipalId`]) {
            throw new Error("[cleanup] A deleted Hub has no recorded principal; workspace role cleanup cannot be certified.");
          }
          continue;
        }
        const value = object(native, "owned identity resource");
        const tags = object(value.tags, "identity ownership");
        if (tags["e2e-run-id"] !== plan.runId || tags["e2e-suite"] !== "fabric-iot-release") {
          foreignIdentity = true;
          throw new Error("[ownership] Identity resource has foreign ownership tags.");
        }
        const identity = resource.type === "Microsoft.ManagedIdentity/userAssignedIdentities"
          ? object(value.properties, "UAMI properties") : object(value.identity, "Hub identity");
        if (typeof identity.principalId === "string") {
          const principal = requireGuid(identity.principalId, "owned principal");
          const source = identitySources.find(source => source.id === resource.id);
          if (!source) throw new Error("[ownership] Unexpected identity on a planned Hub.");
          const previous = plan.observed[`${source.key}IdentityPrincipalId`];
          if (previous && requireGuid(previous, "recorded owned principal") !== principal) {
            foreignIdentity = true;
            throw new Error("[ownership] An owned identity changed its recorded principal.");
          }
          principals.add(principal);
          plan.observed[`${source.key}IdentityResourceId`] = source.id;
          plan.observed[`${source.key}IdentityPrincipalId`] = principal;
          await updatePlan(plan);
        }
      }
      const assignments = object(await cloud.fabric(`${root}/roleAssignments`), "workspace role assignments");
      if (!Array.isArray(assignments.value) || assignments.continuationToken) throw new Error("[service-shape] Ambiguous workspace role inventory.");
      for (const value of assignments.value) {
        const role = object(value, "workspace role");
        const principal = object(role.principal, "workspace principal");
        if (typeof principal.id === "string" && principals.has(principal.id.toLowerCase())) {
          await cloud.fabric(`${root}/roleAssignments/${requireGuid(text(role.id, "role ID"), "role ID")}`, "DELETE");
        }
      }
      const after = object(await cloud.fabric(`${root}/roleAssignments`), "remaining workspace roles");
      if (!Array.isArray(after.value) || after.continuationToken) throw new Error("[service-shape] Missing remaining role inventory.");
      if (after.value.some(value => principals.has(String(object(object(value, "role").principal, "principal").id).toLowerCase()))) {
        throw new Error("[cleanup] An owned Hub identity still has workspace access.");
      }
    });

    await attempt("Activator items, streams and new connections", cleanupActivatorAttempts);
    const createsSettled = await attempt("recorded Fabric creation operations", async () => {
      await reconcileFabricCreates(cloud, plan);
      return true;
    }, "completed");
    await attempt("owned Fabric items", async () => {
      const names = new Map<string, Set<string>>();
      const rawStreamTables = new Map<string, string>();
      const allow = (name: string, type: string) => {
        if (!name.startsWith(`${prefix}-`)) throw new Error("[ownership] Planned cleanup item is outside the run prefix.");
        const types = names.get(name) ?? new Set<string>();
        types.add(type);
        names.set(name, types);
      };
      allow(plan.fabric.itemName, "Org.IoTSolutionAccelerator.IoTSolution");
      allow(plan.fabric.eventhouseName, "Eventhouse");
      allow(plan.fabric.eventhouseName, "KQLDatabase");
      allow(plan.fabric.databaseName, "KQLDatabase");
      for (const path of ["regular", "combined", "linked"] as const) {
        const source = ingestionPath(plan, path);
        allow(source.telemetryStream, "Eventstream");
        allow(source.propertiesStream, "Eventstream");
        rawStreamTables.set(source.telemetryStream, source.telemetryTable);
        rawStreamTables.set(source.propertiesStream, source.propertiesTable);
        allow(`${prefix}-${path}-dashboard`, "KQLDashboard");
        allow(`${prefix}-${path}-data-agent`, "DataAgent");
      }
      const listed = createsSettled
        ? await attempt("owned Fabric item inventory", () => workspaceItems(cloud, plan), "completed") : undefined;
      const recorded = new Map<string, { id: string; displayName: string; type: string }>();
      const record = (key: string, displayName: string, type: string) => {
        const id = plan.observed[key];
        if (id) recorded.set(`${type}:${displayName}`, { id: requireGuid(id, "recorded cleanup item"), displayName, type });
      };
      record("itemId", plan.fabric.itemName, "Org.IoTSolutionAccelerator.IoTSolution");
      record("eventhouseId", plan.fabric.eventhouseName, "Eventhouse");
      record("databaseId", plan.fabric.databaseName, "KQLDatabase");
      for (const path of ["regular", "combined", "linked"]) {
        record(`${path}.dashboardId`, `${prefix}-${path}-dashboard`, "KQLDashboard");
        record(`${path}.dataAgentId`, `${prefix}-${path}-data-agent`, "DataAgent");
      }
      const owned = createsSettled ? new Map<string, Record<string, unknown>>(
        [...recorded.values()].map(item => [item.id, item]),
      ) : new Map<string, Record<string, unknown>>();
      for (const item of listed ?? []) {
        if (typeof item.displayName === "string" && item.displayName.startsWith(`${prefix}-`)) {
          owned.set(requireGuid(text(item.id, "cleanup item ID"), "cleanup item ID"), item);
        }
      }
      const order = ["Eventstream", "Reflex", "DataAgent", "KQLDashboard", "Org.IoTSolutionAccelerator.IoTSolution", "KQLDatabase", "Eventhouse"];
      for (const item of [...owned.values()].sort((a, b) => order.indexOf(String(a.type)) - order.indexOf(String(b.type)))) {
        await attempt(text(item.displayName, "owned item name"), async () => {
          const name = text(item.displayName, "item name"), type = text(item.type, "item type");
          if (!names.get(name)?.has(type)) throw new Error("[ownership] Unrecognized item shares the run prefix; not deleting it.");
          const id = requireGuid(text(item.id, "item ID"), "item ID");
          const expected = recorded.get(`${type}:${name}`);
          if (expected && expected.id !== id) throw new Error("[ownership] Planned cleanup item name is bound to another ID.");
          const native = await cloud.request(`${FABRIC}/v1${root}/items/${id}`, FABRIC, "GET", undefined, true);
          if (native !== undefined) {
            const current = object(native, "owned item");
            if (plan.observed.folderId && current.folderId !== plan.observed.folderId) {
              const table = rawStreamTables.get(name);
              if (type !== "Eventstream" || !table || current.folderId !== undefined && current.folderId !== null) {
                throw new Error("[ownership] Item left the dedicated run folder; refusing deletion.");
              }
              const topology = object(await cloud.fabric(`${root}/eventstreams/${id}/topology`), "root raw Eventstream topology");
              assertRawStreamDestination(topology, plan, table);
              console.log(`[ownership] Root raw Eventstream identity and exclusive manifest database/table binding verified: ${name}`);
            }
          }
          await deleteFabricResource(cloud, `${root}/items/${id}`, name, type);
        });
      }
    }, "completed");
    await attempt("run Fabric folder", async () => {
      if (!createsSettled) throw new Error("[cleanup] Pending Fabric creation prevents folder cleanup.");
      if (!plan.observed.folderId) return;
      const id = requireGuid(plan.observed.folderId, "run folder ID");
      const current = await cloud.request(`${FABRIC}/v1${root}/folders/${id}`, FABRIC, "GET", undefined, true);
      if (current === undefined) return;
      const folder = object(current, "run folder");
      if (folder.displayName !== folderName(plan) || folder.workspaceId !== plan.target.workspaceId || folder.parentFolderId) {
        throw new Error("[ownership] Folder no longer matches the dedicated run folder.");
      }
      if ((await workspaceItems(cloud, plan)).some(item => item.folderId === id)) throw new Error("[ownership] Folder is not empty; refusing deletion.");
      let continuation: string | undefined;
      do {
        const folders = object(await cloud.fabric(`${root}/folders${continuation ? `?continuationToken=${encodeURIComponent(continuation)}` : ""}`), "remaining folders");
        if (!Array.isArray(folders.value)) throw new Error("[service-shape] Missing folder inventory.");
        if (folders.value.some(value => object(value, "folder").parentFolderId === id)) throw new Error("[ownership] Run folder contains an unexpected child folder; refusing deletion.");
        continuation = typeof folders.continuationToken === "string" ? folders.continuationToken : undefined;
      } while (continuation);
      await deleteFabricResource(cloud, `${root}/folders/${id}`, folderName(plan));
    });
    await attempt(plan.azure.resourceGroup, async () => {
      if (foreignIdentity) throw new Error("[ownership] A foreign identity was observed; refusing group deletion.");
      const group = await cloud.arm(plan.azure.resourceGroupId, "2024-03-01", "GET", undefined, true);
      if (group === undefined) return;
      const tags = object(object(group, "resource group").tags, "group tags");
      if (tags["e2e-run-id"] !== plan.runId || tags["e2e-suite"] !== "fabric-iot-release") throw new Error("[ownership] Resource group tags changed; refusing deletion.");
      const inventory = object(await cloud.arm(`${plan.azure.resourceGroupId}/resources`, "2021-04-01"), "group resources");
      if (!Array.isArray(inventory.value) || inventory.nextLink) throw new Error("[ownership] Ambiguous group inventory; refusing deletion.");
      await verifyCleanupInventory(cloud, plan, inventory.value);
      await cloud.arm(plan.azure.resourceGroupId, "2024-03-01", "DELETE");
      const deadline = Date.now() + 15 * 60_000;
      while (Date.now() < deadline) {
        if (await cloud.arm(plan.azure.resourceGroupId, "2024-03-01", "GET", undefined, true) === undefined) return;
        await pause(10_000);
      }
      throw new Error("[cleanup] Azure resource-group deletion is still pending; resources may still incur charges.");
    });
    await attempt("remaining run-owned Fabric items", async () => {
      if (!createsSettled) throw new Error("[cleanup] Recorded Fabric creation is unresolved; absence cannot be certified.");
      const remaining = (await workspaceItems(cloud, plan)).filter(item => typeof item.displayName === "string" && item.displayName.startsWith(`${prefix}-`));
      if (remaining.length) throw new Error(`[cleanup] ${remaining.length} run-prefixed Fabric items remain.`);
    });
  } catch (error) {
    const failure = new Error(redactSensitiveText(error instanceof Error ? error.message : String(error)));
    errors.push(failure);
    receipts.push({ resource: "unexpected cleanup failure", status: "failed", detail: failure.message });
    console.error(`[cleanup] ${failure.message}`);
  } finally {
    plan.observed.cleanupState = errors.length ? "failed" : "confirmed";
    plan.observed.cleanupCompletedAt = new Date().toISOString();
    plan.phase = errors.length ? "cleanup-failed" : "cleaned";
    await updatePlan(plan);
    await writeFile(join(artifacts, "runs", plan.runId, `cleanup-${Date.now()}.json`), JSON.stringify({ runId: plan.runId, receipts }, null, 2));
  }
  if (errors.length) throw new AggregateError(errors, "[cleanup] Run-owned resources remain or cleanup could not be verified. See the cleanup receipt.");
}

export async function verifyCleanupInventory(cloud: Cloud, plan: RunPlan, resources: unknown[]): Promise<void> {
  const expected = new Set([
    ...plan.azure.hubs, ...plan.azure.namespaces, ...plan.azure.namespaces.map(namespace => namespace.dps), plan.azure.regularIdentity,
  ].map(resource => resource.id.toLowerCase()));
  const devices = plan.azure.hubs.flatMap(hub => {
    const namespace = plan.azure.namespaces.find(value => value.key === hub.namespaceKey);
    return namespace ? [{ hub, namespace, id: `${namespace.id}/registryDevices/${hub.name}-device` }] : [];
  });
  for (const value of resources) {
    const resource = object(value, "group resource");
    const id = text(resource.id, "resource ID");
    if (expected.has(id.toLowerCase())) {
      const tags = object(resource.tags, "planned resource ownership tags");
      if (tags["e2e-run-id"] !== plan.runId || tags["e2e-suite"] !== "fabric-iot-release") {
        throw new Error("[ownership] Planned resource has foreign ownership tags; refusing group deletion.");
      }
      continue;
    }
    const device = devices.find(value => value.id.toLowerCase() === id.toLowerCase());
    if (!device) throw new Error(`[ownership] Unexpected resource in the run group: ${id}; refusing group deletion.`);
    const namespace = object(await cloud.arm(device.namespace.id, "2026-11-02-preview"), "owned namespace");
    if (object(namespace.tags, "namespace tags")["e2e-run-id"] !== plan.runId) throw new Error("[ownership] Registry device namespace is not owned by this run.");
    const registry = object(await cloud.arm(device.id, "2026-11-02-preview"), "owned registry device");
    const properties = object(registry.properties, "registry device properties");
    if (properties.externalDeviceId !== `${device.hub.name}-device` || typeof properties.uuid !== "string" || !properties.uuid ||
        plan.observed[`${device.hub.key}.adrDeviceUuid`] && properties.uuid !== plan.observed[`${device.hub.key}.adrDeviceUuid`]) {
      throw new Error("[ownership] Registry device does not match the exact simulated identity.");
    }
    const capability = object(object(await cloud.arm(`${device.id}/capabilities/${device.hub.name}`, "2026-11-02-preview"),
      "registry capability").properties, "capability properties");
    if (capability.messagingEndpointName !== device.hub.name || capability.capabilityType !== "Microsoft.IoTHub" ||
        capability.provisioningState !== "Succeeded") throw new Error("[ownership] Registry device is not bound to the planned Hub capability.");
  }
}
