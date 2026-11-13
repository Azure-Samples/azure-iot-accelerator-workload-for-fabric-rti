import type { RunPlan } from "./plan.js";
import { requireGuid } from "./config.js";

export function nativeItemUrl(plan: RunPlan, kind: "kustodashboards" | "aiskills" | "reflexes", id: string): string {
  const url = new URL(`/groups/${requireGuid(plan.target.workspaceId, "workspace ID")}/${kind}/${requireGuid(id, "owned native item")}`, plan.target.workspaceUrl);
  url.searchParams.set("ctid", requireGuid(plan.target.tenantId, "tenant ID"));
  return url.href;
}

export type PathId = "regular" | "combined" | "linked";
export interface IngestionPath {
  id: PathId;
  hubKeys: string[];
  telemetryTable: string;
  propertiesTable: string;
  telemetryStream: string;
  propertiesStream: string;
}

export function ingestionPath(plan: RunPlan, id: PathId): IngestionPath {
  const suffix = id === "regular" ? "" : `_${id}`;
  return {
    id,
    hubKeys: plan.azure.hubs.filter(hub => id === "combined" || (id === "regular" ? hub.namespaceKey === null : hub.namespaceKey !== null)).map(hub => hub.key),
    telemetryTable: `${plan.fabric.telemetryTable}${suffix}`,
    propertiesTable: `${plan.fabric.propertiesTable}${suffix}`,
    telemetryStream: `${plan.fabric.telemetryStream}${suffix}`,
    propertiesStream: `${plan.fabric.propertiesStream}${suffix}`,
  };
}
