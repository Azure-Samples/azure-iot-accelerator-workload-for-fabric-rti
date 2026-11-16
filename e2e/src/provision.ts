import { createHash } from "node:crypto";
import { setTimeout as pause } from "node:timers/promises";
import { Cloud, object, text } from "./cloud.js";
import { updatePlan, type RunPlan } from "./plan.js";
import { requireGuid } from "./config.js";

const hubVersion = "2026-10-01-preview";
const registryContributor = "4ea46cd5-c1b2-4a8e-910b-273211f9ce47";

export async function provision(cloud: Cloud, plan: RunPlan): Promise<void> {
  const groupId = plan.azure.resourceGroupId;
  const tags = plan.azure.tags;

  function owned(resource: Record<string, unknown>): void {
    const actual = object(resource.tags, "resource ownership");
    if (actual["e2e-run-id"] !== plan.runId || actual["e2e-suite"] !== tags["e2e-suite"]) {
      throw new Error("[ownership] Refusing to reuse or change a resource outside this exact run.");
    }
  }

  async function ensure(id: string, version: string, body: Record<string, unknown>) {
    if (id !== groupId && !id.startsWith(`${groupId}/providers/`)) throw new Error("[ownership] Resource is outside the run resource group.");
    let result = await cloud.arm(id, version, "GET", undefined, true);
    if (result === undefined) {
      console.log(`[azure] CREATE ${id.split("/").at(-1)} (API ${version})`);
      await cloud.arm(id, version, "PUT", { ...body, location: plan.target.location, tags });
    } else {
      owned(object(result, id));
      console.log(`[azure] REUSE  ${id.split("/").at(-1)} (owned by this run)`);
    }
    const deadline = Date.now() + 15 * 60_000;
    let lastState = "";
    while (Date.now() < deadline) {
      const current = object(result ?? await cloud.arm(id, version), id);
      result = undefined;
      owned(current);
      const properties = object(current.properties, "provisioning properties");
      if (id === plan.azure.regularIdentity.id && properties.principalId && properties.clientId) return current;
      const state = String(properties.provisioningState ?? properties.state);
      if (state !== lastState) console.log(`[azure] ${id.split("/").at(-1)}: ${state}`);
      lastState = state;
      if (state === "Succeeded" || state === "Active") return current;
      if (["Failed", "Canceled", "Cancelled", "ActivationFailed"].includes(state)) throw new Error(`[service] ${id} provisioning failed: ${state}`);
      await pause(10_000);
    }
    throw new Error(`[service] ${id} did not finish provisioning within 15 minutes.`);
  }

  const itemId = text(plan.observed.itemId, "owned workload item ID");
  const item = object(await cloud.fabric(`/workspaces/${plan.target.workspaceId}/items/${itemId}`), "workload item");
  if (item.displayName !== plan.fabric.itemName || item.type !== "Org.IoTSolutionAccelerator.IoTSolution") {
    throw new Error("[ownership] Recorded workload item does not match this run.");
  }
  const token = await cloud.token("https://management.azure.com");
  const payload = token.split(".")[1];
  if (!payload) throw new Error("[authentication] Missing ARM token claims.");
  const userId = text(object(JSON.parse(Buffer.from(payload, "base64url").toString("utf8")), "ARM claims").oid, "current principal ID");

  plan.phase = "provisioning-five-s1-hubs";
  await updatePlan(plan);
  await ensure(groupId, "2021-04-01", {});
  const uami = await ensure(plan.azure.regularIdentity.id, "2023-01-31", {});
  const uamiProps = object(uami.properties, "regular UAMI properties");
  plan.observed.regular2IdentityResourceId = plan.azure.regularIdentity.id;
  plan.observed.regular2IdentityPrincipalId = requireGuid(text(uamiProps.principalId, "UAMI principal"), "UAMI principal");
  plan.observed.regular2IdentityClientId = text(uamiProps.clientId, "UAMI client");
  await updatePlan(plan);
  const results = new Map(await Promise.all(plan.azure.hubs.map(async hub => {
    let current = await ensure(hub.id, hubVersion, {
      identity: hub.key === "regular2"
        ? { type: "UserAssigned", userAssignedIdentities: { [plan.azure.regularIdentity.id]: {} } }
        : { type: "SystemAssigned" }, sku: { name: "S1", capacity: 1 },
      properties: { minTlsVersion: "1.2", publicNetworkAccess: "Enabled", disableLocalAuth: true },
    });
    if (hub.key === "regular2" && object(current.identity, "Hub identity").type !== "UserAssigned") {
      console.log(`[azure] Set owned ${hub.name} to UserAssigned-only`);
      await cloud.arm(hub.id, hubVersion, "PUT", {
        location: current.location, tags: { ...object(current.tags, "Hub tags"), ...plan.azure.tags }, sku: { name: "S1", capacity: 1 },
        properties: current.properties,
        identity: { type: "UserAssigned", userAssignedIdentities: { [plan.azure.regularIdentity.id]: {} } },
      },
        false, typeof current.etag === "string" ? current.etag : undefined);
      current = await ensure(hub.id, hubVersion, {});
    }
    const identity = object(current.identity, "Hub identity");
    if (hub.key === "regular2") {
      const attached = object(identity.userAssignedIdentities, "attached UAMIs");
      if (identity.type !== "UserAssigned" || Object.keys(attached).length !== 1 ||
          !Object.keys(attached).some(key => key.toLowerCase() === plan.azure.regularIdentity.id.toLowerCase()) || identity.principalId) {
        throw new Error("[topology] regular2 must have exactly its dedicated UAMI and no system identity.");
      }
    } else {
      if (identity.type !== "SystemAssigned") throw new Error("[topology] Expected a system-assigned-only test Hub.");
      plan.observed[`${hub.key}IdentityResourceId`] = hub.id;
      plan.observed[`${hub.key}IdentityPrincipalId`] = requireGuid(text(identity.principalId, "Hub identity principal"), "Hub identity principal");
    }
    if (object(current.sku, "Hub SKU").name !== "S1") throw new Error("[topology] Expected an S1 Hub.");
    plan.observed[hub.id] = "native-hub-provisioning-verified";
    return [hub.key, current] as const;
  })));
  const definitionId = `/subscriptions/${plan.target.subscriptionId}/providers/Microsoft.Authorization/roleDefinitions/${registryContributor}`;
  for (const hub of plan.azure.hubs.filter(value => value.kind === "nonlinked")) {
    const properties = object(results.get(hub.key)?.properties, "nonlinked Hub properties");
    if (properties.deviceRegistry !== undefined && properties.deviceRegistry !== null) throw new Error("[topology] An intended nonlinked Hub is unexpectedly ADR-linked.");
    const hash = createHash("sha256").update(`${hub.id}|${userId}|${definitionId}`).digest("hex").slice(0, 32);
    const assignment = `${hash.slice(0, 8)}-${hash.slice(8, 12)}-${hash.slice(12, 16)}-${hash.slice(16, 20)}-${hash.slice(20)}`;
    await cloud.arm(`${hub.id}/providers/Microsoft.Authorization/roleAssignments/${assignment}`, "2022-04-01", "PUT", {
      properties: { principalId: userId, principalType: "User", roleDefinitionId: definitionId },
    });
  }
  const inventory = object(await cloud.arm(`${groupId}/resources`, "2021-04-01"), "owned resource inventory");
  if (!Array.isArray(inventory.value) || inventory.nextLink ||
      inventory.value.filter(value => String(object(value, "owned resource").type).toLowerCase() === "microsoft.devices/iothubs").length !== 5) {
    throw new Error("[topology] The run resource group must contain exactly five IoT Hubs.");
  }
  plan.phase = "five-s1-hubs-ready-linked-setup-deferred";
  plan.observed.hubsProvisionedAt = new Date().toISOString();
  plan.observed.regularIdentityCoverage = "regular1:SystemAssigned;regular2:UserAssigned-only";
  plan.verification.topology = "five-s1-hubs-provisioned;adr-dps-not-verified";
  await updatePlan(plan);
  console.log("Five owned S1 Hubs verified. Using the two intended nonlinked Hubs; ADR/DPS associations and linked scenarios are deferred.");
}
