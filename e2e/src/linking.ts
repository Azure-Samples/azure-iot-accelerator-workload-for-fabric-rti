import { createHash } from "node:crypto";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import { Cloud, CloudError, object, text, redactSensitiveText } from "./cloud.js";
import { artifacts } from "./config.js";
import { updatePlan, type RunPlan } from "./plan.js";

export const HUB_API = "2026-10-01-preview";
const ADR_API = "2026-11-02-preview";
export const DPS_API = "2026-06-01-preview";
const roles = {
  contributor: "b24988ac-6180-42a0-ab88-20f7382dd24c",
  hubData: "4fc6c259-987e-4a07-842e-c321cc9d413f",
  hubRegistry: "4ea46cd5-c1b2-4a8e-910b-273211f9ce47",
  registryAdmin: "12675fd7-7f59-493f-9201-f7944860a2f1",
  dpsData: "dfce44e4-17b7-4bd1-a6d1-04996ec95633",
};
export const sameId = (value: unknown, expected: string) =>
  typeof value === "string" && value.toLowerCase() === expected.toLowerCase();

function owned(resource: Record<string, unknown>, plan: RunPlan): void {
  if (!text(resource.id, "resource ID").toLowerCase().startsWith(`${plan.azure.resourceGroupId}/providers/`.toLowerCase()) ||
      object(resource.tags, "ownership tags")["e2e-run-id"] !== plan.runId) {
    throw new Error("[ownership] Resource is not owned by this exact test run.");
  }
}

interface LinkingRoleGrant {
  id: string;
  scope: string;
  principalId: string;
  roleDefinitionId: string;
  verifiedAt: number;
}

export async function grantLinkingRole(cloud: Cloud, plan: RunPlan, scope: string, principal: string, roleId: string, type = "ServicePrincipal"): Promise<LinkingRoleGrant> {
  if (!scope.toLowerCase().startsWith(`${plan.azure.resourceGroupId}/providers/`.toLowerCase())) throw new Error("[ownership] Role scope is outside the owned test resources.");
  const definitionId = `/subscriptions/${plan.target.subscriptionId}/providers/Microsoft.Authorization/roleDefinitions/${roleId}`;
  const hash = createHash("sha256").update(`${scope}|${principal}|${definitionId}`).digest("hex").slice(0, 32);
  const id = `${hash.slice(0, 8)}-${hash.slice(8, 12)}-${hash.slice(12, 16)}-${hash.slice(16, 20)}-${hash.slice(20)}`;
  const assignmentId = `${scope}/providers/Microsoft.Authorization/roleAssignments/${id}`;
  await cloud.arm(assignmentId, "2022-04-01", "PUT", {
    properties: { principalId: principal, principalType: type, roleDefinitionId: definitionId },
  });
  const assignment = object(await cloud.arm(assignmentId, "2022-04-01"), "linking role assignment");
  const properties = object(assignment.properties, "linking role properties");
  if (!sameId(assignment.id, assignmentId) || !sameId(properties.scope, scope) ||
      !sameId(properties.principalId, principal) || !sameId(properties.roleDefinitionId, definitionId) ||
      properties.principalType !== type || properties.condition) {
    throw new Error("[permissions] Linking role readback does not match the exact owned resource/principal/role.");
  }
  return { id: assignmentId, scope, principalId: principal, roleDefinitionId: definitionId, verifiedAt: Date.now() };
}

export function linkingRoleDelay(grants: readonly LinkingRoleGrant[], now = Date.now()): number {
  if (!grants.length || grants.some(grant => !Number.isFinite(grant.verifiedAt))) {
    throw new Error("[permissions] Missing verified linking role timestamps.");
  }
  // ARM caches direct role changes for up to ten minutes; readback alone is not effective-access proof.
  return Math.max(0, Math.max(...grants.map(grant => grant.verifiedAt)) + 600_000 - now);
}

class NativeLinkingFailure extends Error {
  constructor(message: string, readonly resource: Record<string, unknown>) { super(message); }
}

async function wait(cloud: Cloud, id: string, api: string, plan: RunPlan, ready = (_resource: Record<string, unknown>) => true) {
  const deadline = Date.now() + 15 * 60_000;
  while (Date.now() < deadline) {
    const resource = object(await cloud.arm(id, api), "created resource");
    owned(resource, plan);
    const props = object(resource.properties, "provisioning properties");
    const state = props.provisioningState ?? props.state;
    if (["Failed", "Canceled", "ActivationFailed"].includes(String(state))) throw new NativeLinkingFailure(`[service] ${id} failed: ${String(state)}`, resource);
    if ((state === "Succeeded" || state === "Active") && ready(resource)) return resource;
    await pause(5000);
  }
  throw new Error(`[service] Timed out provisioning ${id}.`);
}

export function armOperationUrl(value: string, plan: RunPlan): string {
  const url = new URL(value, "https://management.azure.com");
  if (url.origin !== "https://management.azure.com" || url.username || url.password || url.hash ||
      !url.pathname.toLowerCase().startsWith(`/subscriptions/${plan.target.subscriptionId}/`.toLowerCase())) {
    throw new Error("[ownership] Namespace operation URL is outside the configured ARM subscription.");
  }
  return url.href;
}

export async function associateNamespace(cloud: Cloud, plan: RunPlan, nsId: string, section: "messaging" | "provisioning", additions: Record<string, unknown>) {
  const response = await cloud.send(`https://management.azure.com${nsId}?api-version=${ADR_API}`, "https://management.azure.com");
  const ns = object(response.body, "namespace preimage");
  owned(ns, plan);
  const props = object(ns.properties, "namespace properties");
  const group = props[section] === undefined || props[section] === null ? {} : object(props[section], section);
  const existing = group.endpoints === undefined || group.endpoints === null ? {} : object(group.endpoints, "endpoints");
  if (Object.entries(additions).every(([key, value]) => {
    const observed = existing[key];
    return observed !== undefined && sameId(object(observed, "existing endpoint").resourceId, text(object(value, "target endpoint").resourceId, "target ID")) &&
      object(observed, "existing endpoint").linkingState === "Succeeded";
  })) {
    console.log(`[linking] Native ${section} bindings already succeeded; preserving them.`);
    return;
  }
  const endpoints: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(existing)) {
    const endpoint = object(value, "existing endpoint");
    if (!text(endpoint.resourceId, "endpoint resource").toLowerCase().startsWith(`${plan.azure.resourceGroupId}/providers/`.toLowerCase())) throw new Error("[ownership] Namespace contains an external endpoint.");
    endpoints[key] = Object.fromEntries(["endpointType", "resourceId", "inboundCallerIdentity", "provisioning", "address"]
      .filter(field => endpoint[field] !== undefined && endpoint[field] !== null).map(field => [field, endpoint[field]]));
  }
  for (const [key, value] of Object.entries(additions)) {
    const expected = object(value, "new endpoint");
    if (endpoints[key] && !sameId(object(endpoints[key], "existing endpoint").resourceId, text(expected.resourceId, "target resource"))) throw new Error("[ownership] Endpoint name is already assigned.");
    endpoints[key] = value;
  }
  console.log(`[linking] Associate namespace ${section}: ${Object.keys(additions).join(", ")}`);
  const mutation = await cloud.send(`https://management.azure.com${nsId}?api-version=${ADR_API}`, "https://management.azure.com", "PATCH", {
    properties: { outboundIdentity: { type: "SystemAssigned" }, [section]: { endpoints } },
  }, false, response.headers.get("ETag") ?? (typeof ns.etag === "string" ? ns.etag : undefined));
  const evidenceFile = join(artifacts, "runs", plan.runId, `linking-${nsId.split("/").at(-1)}-${section}-${Date.now()}.json`);
  const operation = mutation.headers.get("Azure-AsyncOperation");
  const evidence: Record<string, unknown> = {
    namespaceId: nsId, section, requestedEndpoints: endpoints, requestedAt: new Date().toISOString(),
    response: { status: mutation.status, body: mutation.body,
      operationUrl: operation, location: mutation.headers.get("Location"),
      requestId: mutation.headers.get("x-ms-request-id"), correlationId: mutation.headers.get("x-ms-correlation-request-id") },
  };
  const persist = () => writeFile(evidenceFile, redactSensitiveText(JSON.stringify(evidence, null, 2)), { mode: 0o600 });
  await persist();
  try {
    const terminal = await wait(cloud, nsId, ADR_API, plan, resource => {
      const current = object(resource.properties, "namespace properties")[section];
      if (!current) return false;
      const observed = object(object(current, section).endpoints ?? {}, "terminal endpoints");
      return Object.entries(endpoints).every(([key, value]) => {
        if (!observed[key]) return false;
        const endpoint = object(observed[key], "terminal endpoint");
        if (!sameId(endpoint.resourceId, text(object(value, "expected endpoint").resourceId, "resource ID"))) {
          throw new NativeLinkingFailure("[topology] Namespace association changed an endpoint identity.", resource);
        }
        if (endpoint.linkingState === "Failed") throw new NativeLinkingFailure(`[service] Endpoint ${key} linking failed.`, resource);
        return endpoint.linkingState === "Succeeded";
      });
    });
    evidence.terminalResource = terminal;
    evidence.completedAt = new Date().toISOString();
    await persist();
  } catch (error) {
    if (!(error instanceof Error)) throw error;
    evidence.failure = error.message;
    if (error instanceof NativeLinkingFailure) evidence.terminalResource = error.resource;
    if (operation) {
      try {
        evidence.operation = await cloud.request(armOperationUrl(operation, plan), "https://management.azure.com");
      } catch (readbackError) {
        if (!(readbackError instanceof CloudError)) throw readbackError;
        evidence.operationReadbackFailure = readbackError.message;
        console.error(`[linking] Operation error readback: ${readbackError.message}`);
      }
    }
    await persist();
    throw new Error(`${redactSensitiveText(error.message)} Native association evidence: ${evidenceFile}`, { cause: error });
  }
}

export async function verifyLinks(cloud: Cloud, plan: RunPlan): Promise<void> {
  const receipts: unknown[] = [];
  for (const ns of plan.azure.namespaces) {
    const namespace = object(await cloud.arm(ns.id, ADR_API), "namespace");
    owned(namespace, plan);
    const nsProps = object(namespace.properties, "namespace properties");
    const endpoints = object(object(nsProps.messaging, "namespace messaging").endpoints, "Hub endpoints");
    const dpsEndpoints = object(object(nsProps.provisioning, "namespace provisioning").endpoints, "DPS endpoints");
    const dpsEndpoint = object(dpsEndpoints[ns.dps.name], "DPS endpoint");
    if (!sameId(dpsEndpoint.resourceId, ns.dps.id) || dpsEndpoint.linkingState !== "Succeeded") {
      throw new Error("[topology] Wrong or nonterminal namespace DPS association.");
    }
    const dps = object(await cloud.arm(ns.dps.id, DPS_API), "native linked DPS");
    owned(dps, plan);
    const bindings = object(dps.properties, "DPS properties").deviceRegistryNamespaces;
    if (!Array.isArray(bindings) || !bindings.some(value => {
      const binding = object(value, "DPS namespace binding");
      return sameId(binding.resourceId, ns.id) && binding.linkingState === "Succeeded";
    })) throw new Error("[topology] Native DPS readback did not confirm Succeeded namespace binding.");
    for (const hub of plan.azure.hubs.filter(hub => hub.namespaceKey === ns.key)) {
      const current = object(await cloud.arm(hub.id, HUB_API), "Hub");
      owned(current, plan);
      const registry = object(object(current.properties, "Hub properties").deviceRegistry, "Hub registry");
      if (!sameId(registry.namespaceResourceId, ns.id) ||
          object(registry.linkingProperties, "link state").state !== "Succeeded" ||
          !sameId(object(endpoints[hub.name], "namespace Hub endpoint").resourceId, hub.id) ||
          object(endpoints[hub.name], "namespace Hub endpoint").linkingState !== "Succeeded") {
        throw new Error(`[topology] ${hub.name} does not have the expected Succeeded native namespace link.`);
      }
      receipts.push({ hubId: hub.id, namespaceId: ns.id, linkingState: "Succeeded", principalId: object(current.identity, "Hub identity").principalId });
    }
  }
  await writeFile(join(artifacts, "runs", plan.runId, "linking-proof.json"), JSON.stringify({ checkedAt: new Date().toISOString(), hubs: receipts }, null, 2));
}

export async function provisionLinks(cloud: Cloud, plan: RunPlan): Promise<void> {
  for (const ns of plan.azure.namespaces) {
    const existing = await cloud.arm(ns.id, ADR_API, "GET", undefined, true);
    if (existing !== undefined) owned(object(existing, "namespace"), plan);
    else await cloud.arm(ns.id, ADR_API, "PUT", { location: plan.target.location, tags: plan.azure.tags, identity: { type: "SystemAssigned" }, properties: { outboundIdentity: { type: "SystemAssigned" } } });
    const namespace = existing === undefined ? await wait(cloud, ns.id, ADR_API, plan) : object(existing, "owned namespace");
    let dps = await cloud.arm(ns.dps.id, DPS_API, "GET", undefined, true);
    if (dps === undefined) {
      console.log(`[linking] Create owned DPS ${ns.dps.name} through ARM deployment`);
      const deployment = `${plan.azure.resourceGroupId}/providers/Microsoft.Resources/deployments/e2e-dps-${ns.key}`;
      await cloud.arm(deployment, "2022-09-01", "PUT", { properties: {
        mode: "Incremental",
        template: {
          "$schema": "https://schema.management.azure.com/schemas/2019-04-01/deploymentTemplate.json#",
          contentVersion: "1.0.0.0",
          resources: [{ type: "Microsoft.Devices/provisioningServices", apiVersion: DPS_API, name: ns.dps.name, location: plan.target.location,
            tags: plan.azure.tags, identity: { type: "SystemAssigned" }, sku: { name: "S1", capacity: 1 },
            properties: { disableLocalAuth: true, publicNetworkAccess: "Enabled", iotHubs: [] } }],
        },
      } });
      const deadline = Date.now() + 15 * 60_000;
      let succeeded = false;
      while (Date.now() < deadline) {
        const state = object(object(await cloud.arm(deployment, "2022-09-01"), "DPS deployment").properties, "deployment properties");
        if (state.provisioningState === "Succeeded") { succeeded = true; break; }
        if (state.provisioningState === "Failed" || state.provisioningState === "Canceled") throw new Error(`[service] DPS ARM deployment failed: ${JSON.stringify(state.error)}`);
        await pause(5000);
      }
      if (!succeeded) throw new Error("[service] DPS ARM deployment timed out.");
      dps = await wait(cloud, ns.dps.id, DPS_API, plan);
    }
    const dpsResource = object(dps, "DPS");
    owned(dpsResource, plan);
    const nsPrincipal = text(object(namespace.identity, "ADR identity").principalId, "ADR principal ID");
    const dpsPrincipal = text(object(dpsResource.identity, "DPS identity").principalId, "DPS principal ID");
    const token = await cloud.token("https://management.azure.com");
    const claims = object(JSON.parse(Buffer.from(text(token.split(".")[1], "token payload"), "base64url").toString()), "claims");
    const user = text(claims.oid, "current user");
    const grants = [
      await grantLinkingRole(cloud, plan, ns.id, nsPrincipal, roles.registryAdmin),
      await grantLinkingRole(cloud, plan, ns.dps.id, nsPrincipal, roles.contributor),
      await grantLinkingRole(cloud, plan, ns.id, dpsPrincipal, roles.contributor),
      await grantLinkingRole(cloud, plan, ns.dps.id, user, roles.dpsData, "User"),
    ];
    const additions: Record<string, unknown> = {};
    for (const hub of plan.azure.hubs.filter(hub => hub.namespaceKey === ns.key)) {
      const current = object(await cloud.arm(hub.id, HUB_API), "owned Hub");
      owned(current, plan);
      const registry = object(current.properties, "Hub properties").deviceRegistry;
      if (registry !== undefined && registry !== null && !sameId(object(registry, "registry").namespaceResourceId, ns.id)) throw new Error("[ownership] Hub is already linked elsewhere.");
      const principal = text(object(current.identity, "Hub identity").principalId, "Hub principal");
      grants.push(
        await grantLinkingRole(cloud, plan, hub.id, nsPrincipal, roles.contributor),
        await grantLinkingRole(cloud, plan, hub.id, nsPrincipal, roles.hubData),
        await grantLinkingRole(cloud, plan, ns.id, principal, roles.contributor),
        await grantLinkingRole(cloud, plan, hub.id, user, roles.hubRegistry, "User"),
      );
      additions[hub.name] = { endpointType: "Microsoft.Devices/IotHubs", resourceId: hub.id,
        inboundCallerIdentity: { type: "SystemAssigned" }, provisioning: { availability: "Available", allocationWeight: 1 } };
    }
    const delay = linkingRoleDelay(grants);
    const grantFile = join(artifacts, "runs", plan.runId, `linking-role-grants-${ns.key}.json`);
    const propagation = { grants, waitMs: delay, startedAt: new Date().toISOString(), completedAt: "" };
    await writeFile(grantFile, JSON.stringify(propagation, null, 2));
    console.log(`[linking] Exact resource/principal/role readbacks verified. Allow ${Math.ceil(delay / 1000)} seconds for ARM RBAC propagation before the first association; no request retries.`);
    await pause(delay);
    propagation.completedAt = new Date().toISOString();
    await writeFile(grantFile, JSON.stringify(propagation, null, 2));
    await associateNamespace(cloud, plan, ns.id, "provisioning", {
      [ns.dps.name]: { endpointType: "Microsoft.Devices/provisioningServices", resourceId: ns.dps.id, inboundCallerIdentity: { type: "SystemAssigned" } },
    });
    await associateNamespace(cloud, plan, ns.id, "messaging", additions);
  }
  // Association operations complete asynchronously. These are native-state waits, not request retries.
  for (const ns of plan.azure.namespaces) for (const hub of plan.azure.hubs.filter(hub => hub.namespaceKey === ns.key)) {
    const deadline = Date.now() + 10 * 60_000;
    let ready = false;
    while (Date.now() < deadline) {
      const registry = object(object(object(await cloud.arm(hub.id, HUB_API), "Hub").properties, "properties").deviceRegistry, "registry");
      const state = object(registry.linkingProperties, "link status").state;
      if (state === "Succeeded") { ready = true; break; }
      if (state === "Failed") throw new Error(`[service] Hub ${hub.name} linking failed.`);
      await pause(5000);
    }
    if (!ready) throw new Error(`[service] ${hub.name} link timed out.`);
  }
  await verifyLinks(cloud, plan);
  plan.observed.linkedTopologyState = "native-succeeded-A-A-B";
  plan.observed.linkedTopologyVerifiedAt = new Date().toISOString();
  await updatePlan(plan);
}
