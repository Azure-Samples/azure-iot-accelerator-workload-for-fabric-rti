import { parseArgs } from "node:util";
import { fileURLToPath } from "node:url";
import { join } from "node:path";

export const root = fileURLToPath(new URL("../", import.meta.url));
export const artifacts = join(root, "artifacts");
export const profile = join(artifacts, "browser-profile");
process.env.PLAYWRIGHT_BROWSERS_PATH ??= join(artifacts, "browsers");

const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function requireGuid(value: string | undefined, name: string): string {
  if (!value || !guid.test(value)) throw new Error(`[configuration] ${name} must be a GUID.`);
  return value.toLowerCase();
}

export interface AuthConfig {
  tenantId: string;
  workspaceId: string;
  workspaceUrl: string;
  browserChannel: "msedge" | "chromium";
  loginTimeoutMs: number;
  deployment?: "development" | "published";
}

export function authConfig(
  args = process.argv.slice(2),
  env: NodeJS.ProcessEnv = process.env,
): AuthConfig {
  const { values } = parseArgs({
    args,
    options: {
      tenant: { type: "string" },
      workspace: { type: "string" },
      "fabric-url": { type: "string" },
      browser: { type: "string" },
      deployment: { type: "string" },
    },
  });
  const tenantId = requireGuid(values.tenant ?? env.E2E_TENANT_ID, "--tenant / E2E_TENANT_ID");
  const workspaceId = requireGuid(values.workspace ?? env.E2E_WORKSPACE_ID, "--workspace / E2E_WORKSPACE_ID");
  const url = new URL(values["fabric-url"] ?? env.E2E_FABRIC_URL ?? `https://app.fabric.microsoft.com/groups/${workspaceId}`);
  if (url.protocol !== "https:" || !["app.powerbi.com", "app.fabric.microsoft.com"].includes(url.hostname) ||
      url.username || url.password || url.port ||
      !new RegExp(`^/groups/${workspaceId}/?$`, "i").test(url.pathname)) {
    throw new Error("[configuration] Fabric URL must be the HTTPS /groups/<workspace ID> page on app.powerbi.com or app.fabric.microsoft.com.");
  }
  if ([...url.searchParams.keys()].some(key => key !== "ctid")) {
    throw new Error("[configuration] Only the ctid tenant hint is supported in the workspace URL.");
  }
  if (url.searchParams.has("ctid") && url.searchParams.get("ctid")?.toLowerCase() !== tenantId) {
    throw new Error("[configuration] The workspace URL ctid differs from the requested tenant.");
  }
  url.searchParams.set("ctid", tenantId);
  url.hash = "";
  const browserChannel = values.browser ?? env.E2E_BROWSER ?? "msedge";
  if (browserChannel !== "msedge" && browserChannel !== "chromium") {
    throw new Error("[configuration] Browser must be msedge or chromium.");
  }
  const deployment = values.deployment ?? env.E2E_DEPLOYMENT;
  if (deployment !== undefined && deployment !== "development" && deployment !== "published") {
    throw new Error("[configuration] E2E_DEPLOYMENT / --deployment must be development or published.");
  }
  return {
    tenantId, workspaceId, workspaceUrl: url.href, browserChannel, loginTimeoutMs: 600_000,
    ...(deployment ? { deployment } : {}),
  };
}

export interface RunConfig extends AuthConfig {
  subscriptionId: string;
  location: string;
  retention: "cleanup" | "retain";
  keepResources: boolean;
  itemLabel: string;
  deployment: "development" | "published";
}

export function runConfig(env: NodeJS.ProcessEnv = process.env): RunConfig {
  const auth = authConfig([], env);
  const subscriptionId = requireGuid(env.E2E_SUBSCRIPTION_ID, "E2E_SUBSCRIPTION_ID");
  const location = env.E2E_LOCATION?.trim().toLowerCase().replace(/\s+/g, "");
  if (!location || !/^[a-z][a-z0-9]+$/.test(location)) {
    throw new Error("[configuration] E2E_LOCATION is required (for example australiaeast).");
  }
  const retention = env.E2E_RETENTION ?? "cleanup";
  if (retention !== "cleanup" && retention !== "retain") {
    throw new Error("[configuration] E2E_RETENTION must be cleanup or retain.");
  }
  const keepResources = env.E2E_KEEP_RESOURCES;
  if (keepResources !== undefined && keepResources !== "0" && keepResources !== "1") {
    throw new Error("[configuration] E2E_KEEP_RESOURCES must be 0 or 1.");
  }
  const itemLabel = env.E2E_ITEM_LABEL?.trim() || "Azure IoT Solution Accelerator";
  if (itemLabel.length > 200 || /[\r\n]/.test(itemLabel)) {
    throw new Error("[configuration] E2E_ITEM_LABEL must be a single UI label of at most 200 characters.");
  }
  const deployment = auth.deployment;
  if (!deployment) {
    throw new Error("[configuration] Set E2E_DEPLOYMENT to development or published explicitly. No cross-target fallback is allowed.");
  }
  return { ...auth, subscriptionId, location, retention,
    keepResources: keepResources === "1" || retention === "retain", itemLabel, deployment };
}
