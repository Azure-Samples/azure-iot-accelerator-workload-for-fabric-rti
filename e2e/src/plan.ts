import { randomBytes, randomUUID } from "node:crypto";
import { mkdir, readFile, rename, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { artifacts, requireGuid, type RunConfig } from "./config.js";

export function createPlan(config: RunConfig) {
  const runId = randomUUID();
  const suffix = randomBytes(8).toString("hex");
  const prefix = `e2e-${suffix}`;
  const observed: Record<string, string> = { itemUrl: "", itemFrameUrl: "", itemId: "" };
  const verification: Record<string, string> = { ui: "not-run", rawData: "not-run", topology: "planned-hubs-only;adr-dps-deferred" };
  const resourceGroup = `fabric-${prefix}`;
  const resourceGroupId = `/subscriptions/${config.subscriptionId}/resourceGroups/${resourceGroup}`;
  const resource = (type: string, name: string) => ({
    type, name, id: `${resourceGroupId}/providers/${type}/${name}`,
  });
  const namespaces = ["a", "b"].map(key => ({
    key,
    ...resource("Microsoft.DeviceRegistry/namespaces", `${prefix}-adr-${key}`),
    dps: resource("Microsoft.Devices/provisioningServices", `${prefix}-dps-${key}`),
  }));
  const hubs = [
    { key: "regular1", kind: "nonlinked", namespaceKey: null },
    { key: "regular2", kind: "nonlinked", namespaceKey: null },
    { key: "linked1", kind: "future-linked", namespaceKey: "a" },
    { key: "linked2", kind: "future-linked", namespaceKey: "a" },
    { key: "linked3", kind: "future-linked", namespaceKey: "b" },
  ].map(hub => ({
    ...hub,
    ...resource("Microsoft.Devices/IotHubs", `${prefix}-${hub.key}`),
    sku: "S1",
    capacity: 1,
  }));
  return {
    schemaVersion: 4,
    provisioningContract: "s1-hubs-linked-deferred",
    runId,
    createdAt: new Date().toISOString(),
    phase: "planned",
    resourcesCreated: false,
    target: {
      tenantId: config.tenantId,
      subscriptionId: config.subscriptionId,
      workspaceId: config.workspaceId,
      workspaceUrl: config.workspaceUrl,
      location: config.location,
      itemLabel: config.itemLabel,
      deployment: config.deployment,
      retention: config.retention,
    },
    azure: {
      regularIdentity: resource("Microsoft.ManagedIdentity/userAssignedIdentities", `${prefix}-regular2-identity`),
      resourceGroup,
      resourceGroupId,
      tags: { "e2e-suite": "fabric-iot-release", "e2e-run-id": runId },
      namespaces,
      hubs,
    },
    fabric: {
      itemName: `${prefix}-solution`,
      eventhouseName: `${prefix}-eventhouse`,
      databaseName: `${prefix}-database`,
      telemetryTable: `TelemetryRaw_${suffix}`,
      propertiesTable: `PropertiesRaw_${suffix}`,
      telemetryStream: `${prefix}-telemetry`,
      propertiesStream: `${prefix}-properties`,
    },
    requiredSupport: {
      dpsInstances: 0,
      userAssignedIdentities: 1,
      systemAssignedIdentities: 4,
      roleAssignmentScope: resourceGroupId,
      certificateManagement: false,
    },
    deferredSupport: { adrNamespaces: 2, dpsInstances: 2 },
    verification,
    observed,
  };
}

export type RunPlan = ReturnType<typeof createPlan>;

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function strings(value: unknown): value is Record<string, string> {
  return record(value) && Object.values(value).every(entry => typeof entry === "string");
}

export function assertOwnedPlan(value: unknown, config: RunConfig, runId: string): asserts value is RunPlan {
  const invalid = () => { throw new Error("[ownership] Invalid or mismatched run manifest; refusing resource access."); };
  if (!record(value) || value.schemaVersion !== 4 || value.provisioningContract !== "s1-hubs-linked-deferred" || value.runId !== runId ||
      typeof value.createdAt !== "string" || !Number.isFinite(Date.parse(value.createdAt)) ||
      typeof value.phase !== "string" || typeof value.resourcesCreated !== "boolean" ||
      !strings(value.target) || !record(value.azure) || !strings(value.fabric) ||
      !strings(value.observed) || !strings(value.verification) || !record(value.requiredSupport)) return invalid();
  for (const key of ["tenantId", "subscriptionId", "workspaceId", "workspaceUrl", "location", "itemLabel", "deployment", "retention"] as const) {
    if (value.target[key] !== config[key]) return invalid();
  }
  const azure = value.azure;
  if (typeof azure.resourceGroup !== "string" || !/^fabric-e2e-[a-f0-9]{16}$/.test(azure.resourceGroup)) return invalid();
  const groupId = `/subscriptions/${config.subscriptionId}/resourceGroups/${azure.resourceGroup}`;
  const prefix = azure.resourceGroup.slice("fabric-".length);
  if (azure.resourceGroupId !== groupId || !strings(azure.tags) ||
      azure.tags["e2e-run-id"] !== runId || azure.tags["e2e-suite"] !== "fabric-iot-release" ||
      !Array.isArray(azure.hubs) || azure.hubs.length !== 5 ||
      !Array.isArray(azure.namespaces) || azure.namespaces.length !== 2) return invalid();
  const resourceMatches = (resource: unknown, type: string, name: string) =>
    record(resource) && resource.type === type && resource.name === name &&
    resource.id === `${groupId}/providers/${type}/${name}`;
  if (!resourceMatches(azure.regularIdentity, "Microsoft.ManagedIdentity/userAssignedIdentities", `${prefix}-regular2-identity`)) return invalid();
  for (const [index, key] of ["regular1", "regular2", "linked1", "linked2", "linked3"].entries()) {
    const hub: unknown = azure.hubs[index];
    if (!record(hub) || !resourceMatches(hub, "Microsoft.Devices/IotHubs", `${prefix}-${key}`) ||
        hub.key !== key || hub.kind !== (index < 2 ? "nonlinked" : "future-linked") ||
        hub.namespaceKey !== (index < 2 ? null : index < 4 ? "a" : "b") ||
        hub.sku !== "S1" || hub.capacity !== 1) return invalid();
  }
  for (const [index, key] of ["a", "b"].entries()) {
    const namespace: unknown = azure.namespaces[index];
    if (!record(namespace) || namespace.key !== key ||
        !resourceMatches(namespace, "Microsoft.DeviceRegistry/namespaces", `${prefix}-adr-${key}`) ||
        !resourceMatches(namespace.dps, "Microsoft.Devices/provisioningServices", `${prefix}-dps-${key}`)) return invalid();
  }
  const fabric = value.fabric;
  if (fabric.itemName !== `${prefix}-solution` ||
      !["itemName", "eventhouseName", "databaseName", "telemetryTable", "propertiesTable", "telemetryStream", "propertiesStream"]
        .every(key => typeof fabric[key] === "string" && fabric[key] !== "") ||
      !Object.values(fabric).every(name => /^[a-zA-Z0-9_-]+$/.test(name)) ||
      value.requiredSupport.dpsInstances !== 0 || value.requiredSupport.userAssignedIdentities !== 1 ||
      value.requiredSupport.systemAssignedIdentities !== 4 ||
      !record(value.deferredSupport) || value.deferredSupport.adrNamespaces !== 2 || value.deferredSupport.dpsInstances !== 2 ||
      value.requiredSupport.roleAssignmentScope !== groupId || value.requiredSupport.certificateManagement !== false ||
      typeof value.verification.ui !== "string" || typeof value.verification.rawData !== "string") return invalid();
  if (value.observed.itemId) requireGuid(value.observed.itemId, "owned item ID");
}

export async function readPlan(config: RunConfig, runId: string): Promise<RunPlan> {
  requireGuid(runId, "E2E_RUN_ID");
  const value: unknown = JSON.parse(await readFile(join(artifacts, "runs", runId, "manifest.json"), "utf8"));
  assertOwnedPlan(value, config, runId);
  return value;
}

export async function writePlan(plan: ReturnType<typeof createPlan>): Promise<string> {
  const directory = join(artifacts, "runs", plan.runId);
  await mkdir(directory, { recursive: true, mode: 0o700 });
  const path = join(directory, "manifest.json");
  await writeFile(path, JSON.stringify(plan, null, 2), { flag: "wx", mode: 0o600 });
  return path;
}

export async function updatePlan(plan: ReturnType<typeof createPlan>): Promise<void> {
  const directory = join(artifacts, "runs", requireGuid(plan.runId, "run ID"));
  const path = join(directory, "manifest.json");
  const temporary = join(directory, `manifest-${randomUUID()}.tmp`);
  await writeFile(temporary, JSON.stringify(plan, null, 2), { flag: "wx", mode: 0o600 });
  await rename(temporary, path);
}
