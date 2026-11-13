import { Cloud, object, text } from "./cloud.js";
import { requireGuid } from "./config.js";
import type { RunPlan } from "./plan.js";
import type { IngestionPath } from "./paths.js";
import { moveOwnedItem } from "./folders.js";
import { artifacts } from "./config.js";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import { updatePlan } from "./plan.js";

export function assertRawStreamDestination(topology: Record<string, unknown>, plan: RunPlan, table: string): void {
  if (!Array.isArray(topology.destinations) || topology.destinations.length !== 1) {
    throw new Error("[ownership] Expected one exact run-owned raw Eventstream destination.");
  }
  const destination = object(topology.destinations[0], "stream destination");
  const props = object(destination.properties, "stream destination properties");
  if (destination.type !== "Eventhouse" || props.workspaceId !== plan.target.workspaceId ||
      !plan.observed.databaseId || props.itemId !== plan.observed.databaseId ||
      props.tableName !== table || props.dataIngestionMode !== "ProcessedIngestion") {
    throw new Error("[ownership] Raw Eventstream does not bind exclusively to the manifest's exact database/table.");
  }
}

export async function recordIngestionStreams(cloud: Cloud, plan: RunPlan, path: IngestionPath): Promise<void> {
  const workspace = `/workspaces/${plan.target.workspaceId}`;
  const listing = object(await cloud.fabric(`${workspace}/items?type=Eventstream`), "workspace streams");
  if (!Array.isArray(listing.value) || listing.continuationToken) throw new Error("[service-shape] Stream inventory is missing or unexpectedly paginated.");
  for (const [kind, name, table] of [
    ["telemetry", path.telemetryStream, path.telemetryTable], ["properties", path.propertiesStream, path.propertiesTable],
  ] as const) {
    const matches = listing.value.map(value => object(value, "stream")).filter(stream => stream.displayName === name && stream.type === "Eventstream");
    if (matches.length !== 1) throw new Error(`[product-or-service] Expected exactly one per-run Eventstream named ${name}.`);
    const id = requireGuid(text(matches[0]?.id, "Eventstream ID"), "Eventstream ID");
    const topology = object(await cloud.fabric(`${workspace}/eventstreams/${id}/topology`), "Eventstream topology");
    assertRawStreamDestination(topology, plan, table);
    plan.observed[`${path.id}.${kind}StreamId`] = id;
    await updatePlan(plan);
    await moveOwnedItem(cloud, plan, id, name);
  }
}

export function routingTargetsReady(routing: Record<string, unknown>, workspaceId: string, streamIds: readonly string[]): boolean {
  if (streamIds.length !== 2 || new Set(streamIds).size !== 2) throw new Error("[prerequisite] Require two distinct recorded raw Eventstream IDs.");
  const endpoints = object(routing.endpoints ?? {}, "routing endpoints").eventStreams;
  if (endpoints === undefined) return false;
  if (!Array.isArray(endpoints)) throw new Error("[service-shape] Native Fabric endpoints are not an array.");
  return streamIds.every(id => {
    const matches = endpoints.map(value => object(value, "endpoint")).filter(endpoint => endpoint.eventStreamId === id && endpoint.workspaceId === workspaceId);
    return matches.length === 1 && matches[0]?.messagePayloadFormat === "DOObservationV1";
  });
}

export async function routingBefore(cloud: Cloud, plan: RunPlan, hubId: string): Promise<Record<string, unknown>> {
  const hub = object(await cloud.arm(hubId, "2026-10-01-preview"), "owned Hub");
  if (object(hub.tags, "Hub ownership")["e2e-run-id"] !== plan.runId) throw new Error("[ownership] Routing snapshot reached an unowned Hub.");
  return object(object(hub.properties, "Hub properties").routing ?? {}, "prior routing");
}

export async function verifyRouting(
  cloud: Cloud, plan: RunPlan, hubId: string, path: IngestionPath, previous: Record<string, unknown>,
): Promise<void> {
  const streamIds = ["telemetry", "properties"].map(kind => requireGuid(text(plan.observed[`${path.id}.${kind}StreamId`], "recorded raw stream"), "recorded raw stream"));
  const deadline = Date.now() + 180_000;
  let hub: Record<string, unknown>, routing: Record<string, unknown>;
  const evidenceFile = join(artifacts, "runs", plan.runId, `routing-${path.id}-${hubId.split("/").at(-1)}.json`);
  while (true) {
    hub = object(await cloud.arm(hubId, "2026-10-01-preview"), "routed Hub");
    if (object(hub.tags, "Hub ownership")["e2e-run-id"] !== plan.runId) throw new Error("[ownership] Routing verification reached an unowned Hub.");
    routing = object(object(hub.properties, "Hub properties").routing ?? {}, "Hub routing");
    const streams = object(routing.endpoints ?? {}, "routing endpoints").eventStreams;
    await writeFile(evidenceFile, JSON.stringify({
      checkedAt: new Date().toISOString(), hubId, path: path.id, expectedStreamIds: streamIds,
      endpoints: Array.isArray(streams) ? streams.map(value => {
        const endpoint = object(value, "native endpoint");
        return Object.fromEntries(["name", "eventStreamId", "workspaceId", "messagePayloadFormat", "authenticationType", "identity"]
          .map(key => [key, endpoint[key]]));
      }) : [],
      routes: routing.routes,
    }, null, 2));
    if (routingTargetsReady(routing, plan.target.workspaceId, streamIds)) break;
    if (Date.now() >= deadline) {
      throw new Error(`[product-or-deployed-contract] Native routing did not expose both exact DOObservationV1 endpoints. Evidence: ${evidenceFile}`);
    }
    console.log(`[routing] Waiting for native ${path.id} endpoint readback on ${hubId.split("/").at(-1)}; no routing mutation or retry.`);
    await pause(5000);
  }
  if (object(hub.tags, "Hub ownership")["e2e-run-id"] !== plan.runId) throw new Error("[ownership] Routing verification reached an unowned Hub.");
  const endpoints = object(routing.endpoints, "routing endpoints").eventStreams;
  if (!Array.isArray(endpoints) || !Array.isArray(routing.routes)) throw new Error("[service-shape] Native Fabric endpoints/routes are missing.");
  for (const value of Array.isArray(previous.routes) ? previous.routes : []) {
    const old = object(value, "prior route");
    const current = routing.routes.map(value => object(value, "route")).find(value => value.name === old.name);
    if (!current || current.source !== old.source || current.condition !== old.condition || current.isEnabled !== old.isEnabled ||
        JSON.stringify(current.endpointNames) !== JSON.stringify(old.endpointNames)) throw new Error("[additive-ingestion] A prior scenario route was changed.");
  }
  const previousEndpoints = object(previous.endpoints ?? {}, "prior endpoints").eventStreams;
  for (const value of Array.isArray(previousEndpoints) ? previousEndpoints : []) {
    const old = object(value, "prior endpoint");
    const current = endpoints.map(value => object(value, "endpoint")).find(value => value.name === old.name);
    if (!current || current.eventStreamId !== old.eventStreamId || current.workspaceId !== old.workspaceId ||
        JSON.stringify(current.identity) !== JSON.stringify(old.identity)) throw new Error("[additive-ingestion] A prior scenario endpoint was changed.");
  }
  const workspace = `/workspaces/${plan.target.workspaceId}`;
  const identity = object(hub.identity, "native Hub identity");
  const regular2 = plan.azure.hubs.find(value => value.key === "regular2")?.id === hubId;
  const identityResourceId = regular2 ? plan.azure.regularIdentity.id : null;
  const attached = regular2 ? object(identity.userAssignedIdentities, "attached identities") : {};
  const attachedKey = Object.keys(attached).find(key => key.toLowerCase() === plan.azure.regularIdentity.id.toLowerCase());
  const principal = regular2 ? object(attached[attachedKey ?? ""], "attached UAMI").principalId : identity.principalId;
  if (identity.type !== (regular2 ? "UserAssigned" : "SystemAssigned")) throw new Error("[topology] Wrong Hub identity type.");
  const assignments = object(await cloud.fabric(`${workspace}/roleAssignments`), "workspace access");
  if (!Array.isArray(assignments.value) || !assignments.value.some(value => {
    const assignment = object(value, "workspace role");
    return object(assignment.principal, "workspace principal").id === principal && assignment.role === "Contributor";
  })) throw new Error("[native-identity] The selected Hub identity does not have native workspace Contributor membership.");
  for (const [kind, source] of [
    ["telemetry", "DeviceMessages"],
    ["properties", "TwinChangeEvents"],
  ] as const) {
    const streamId = requireGuid(text(plan.observed[`${path.id}.${kind}StreamId`], "recorded raw Eventstream"), "Eventstream ID");
    const endpoint = endpoints.map(value => object(value, "native Eventstream endpoint"))
      .find(value => value.eventStreamId === streamId && value.workspaceId === plan.target.workspaceId);
    if (!endpoint || endpoint.messagePayloadFormat !== "DOObservationV1") throw new Error("[product-or-deployed-contract] Native routing must use the intended DOObservationV1 Eventstream endpoint.");
    const endpointIdentity = object(endpoint.identity, "routing identity").userAssignedIdentity;
    if (endpoint.authenticationType !== "identityBased" ||
        (identityResourceId ? typeof endpointIdentity !== "string" || endpointIdentity.toLowerCase() !== identityResourceId.toLowerCase() : endpointIdentity !== null && endpointIdentity !== undefined)) {
      throw new Error("[native-identity] Eventstream endpoint did not use the exact selected managed identity.");
    }
    if (!routing.routes.some(value => {
      const route = object(value, "Hub route");
      return route.source === source && route.isEnabled === true && route.condition === "true" &&
        Array.isArray(route.endpointNames) && route.endpointNames.includes(endpoint.name);
    })) throw new Error(`[product-or-service] Native ${source} route does not feed the expected Eventstream.`);
    plan.observed[`${path.id}.${source === "DeviceMessages" ? "telemetryStreamId" : "propertiesStreamId"}`] = streamId;
  }
}
