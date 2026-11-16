// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { canonicalJson, isJsonObject } from "./ModelAuthoring.ts";
import {
  ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE, ASSIGNMENT_UNTRACKED_DISPATCH_MESSAGE, ASSIGNMENT_OPERATION_TIMEOUT_MS, ASSIGNMENT_POLL_TIMEOUT_MS, DPS_INFORMATIONAL_ONLY, UPDATE_PERMISSION_UNKNOWN, armIdentity, assignmentBlocker, awaitAssignment,
  permissionGrants, readReference,
  sameResource, sourceHubs, specificEtag,
  type AssignmentAttempt, type AssignmentReceipt, type AssignmentHubSummary, type AssignmentInventory, type AssignmentIssue, type AssignmentIssueScope, type AssignmentServices, type AssignmentSource, type AssignmentTarget,
} from "./HubAssignment.ts";

export const ASSIGNMENT_ARM_SCOPE = "https://management.azure.com/.default";
const ARM = "https://management.azure.com";
const PREVIEW = "2026-11-02-preview";
type RecordValue = Record<string, unknown>;
const object = (value: unknown, where: string): RecordValue => {
  if (!isJsonObject(value)) throw new AssignmentApiError("shape", `The service returned incomplete ${where}.`);
  return value;
};
const text = (value: unknown, where: string): string => {
  if (typeof value !== "string" || !value) throw new AssignmentApiError("shape", `The service returned incomplete ${where}.`);
  return value;
};
export class AssignmentApiError extends Error {
  readonly kind: "permission" | "authentication" | "shape" | "request" | "unsupported";
  readonly retryable: boolean;
  readonly status?: number;
  readonly authScope?: "arm" | "dps";
  readonly retryAfterMs?: number;
  constructor(kind: AssignmentApiError["kind"], message: string, retryable = false, status?: number,
    authScope?: "arm" | "dps", retryAfterMs?: number) {
    super(message);
    this.kind = kind;
    this.retryable = retryable;
    this.status = status;
    this.authScope = authScope;
    this.retryAfterMs = retryAfterMs;
  }
}
const ASSIGNMENT_TRACKING_UNAVAILABLE_MESSAGE =
  "The original assignment operation tracking information is unavailable. Completion is unknown; no assignment will be retried.";
class AssignmentTrackingUnavailableError extends AssignmentApiError {
  constructor() {
    super("shape", ASSIGNMENT_TRACKING_UNAVAILABLE_MESSAGE);
  }
}
function retryAfter(response: Response): number | undefined {
  const value = response.headers.get("Retry-After");
  if (!value) return undefined;
  const milliseconds = /^\d+$/.test(value) ? Number(value) * 1000 : Date.parse(value) - Date.now();
  return Number.isFinite(milliseconds) && milliseconds > 0 ? Math.min(milliseconds, 2147483647) : undefined;
}
export interface AssignmentReadback { confirmed: boolean; terminalState?: "failed" | "cancelled" }
interface NativeAssignmentOperation {
  url: string; targetId: string; modelId: string; deadline: number;
  terminal?: AssignmentReadback;
}
// Operation URLs can carry opaque credentials. Keep them out of wizard context and receipts.
const nativeAssignmentOperations = new Map<string, NativeAssignmentOperation>();
function operationUrl(value: string, targetId: string): string {
  let url: URL;
  try { url = new URL(value); } catch {
    throw new AssignmentApiError("shape", "The assignment response did not contain a valid operation address.");
  }
  const subscription = targetId.split("/")[2];
  const operationPath = /^\/subscriptions\/([^/]+)\/providers\/Microsoft\.DeviceRegistry\/locations\/[^/]+\/(?:asyncOperationStatuses|operationStatuses)\/[^/]+$/i.exec(url.pathname);
  if (url.origin !== ARM || url.username || url.password || url.hash || !operationPath ||
      operationPath[1].toLowerCase() !== subscription.toLowerCase() ||
      /%|\\/.test(value.split(/[?#]/, 1)[0]) || /\/\.{1,2}(?:\/|$)/.test(value.split(/[?#]/, 1)[0])) {
    throw new AssignmentApiError("shape", "The assignment response contained an unsafe or unsupported operation address.");
  }
  return url.href;
}
export function assignmentTokenFailure(failure: unknown): AssignmentApiError {
  if (failure instanceof AssignmentApiError) return failure;
  const authScope = "arm";
  const resource = "Azure resource";
  const code = isJsonObject(failure) && typeof failure.error === "number" ? failure.error : undefined;
  // WorkloadAuthError distinguishes unsupported/configuration/scope failures from incomplete interaction.
  if (code === 0 || code === 2 || code === 4) {
    return new AssignmentApiError("unsupported", code === 0 ? "Fabric authentication is unavailable in this environment."
      : `The workload's ${resource} authorization is not configured for this request. Ask the workload owner to check its application permissions.`,
    false, undefined, authScope);
  }
  return new AssignmentApiError("authentication",
    `Fabric couldn't complete ${resource} authorization. Use Sign in to retry. If consent is still required, ask the workload owner to check its application permissions.`,
    false, undefined, authScope);
}
function cancelled(signal: AbortSignal, isCurrent?: () => boolean): void {
  if (signal.aborted || isCurrent?.() === false) throw new DOMException("Assignment request cancelled.", "AbortError");
}
interface ClientOptions {
  token: (scope: string) => Promise<string>;
  fetcher?: typeof fetch;
  isCurrent?: () => boolean;
}
interface DevicePage {
  namespaceId: string;
  hubIds: string[];
  endpoints: RecordValue;
  next: string;
  seen: Set<string>;
  ids: Set<string>;
}
interface ReadPublication {
  isCurrent: () => boolean;
  publish: (inventory: AssignmentInventory) => void;
}
interface HubSummaryOptions {
  initial?: AssignmentHubSummary;
  publish?: (summary: AssignmentHubSummary) => void;
}
export class HubAssignmentApi {
  private readonly options: ClientOptions;
  private readonly fetcher: typeof fetch;
  private pendingPages: DevicePage[] = [];
  private readonly operationIds = new Map<string, string>();
  constructor(options: ClientOptions) {
    this.options = options;
    this.fetcher = options.fetcher || globalThis.fetch.bind(globalThis);
  }

  private async request(url: string, signal: AbortSignal, init: RequestInit = {},
    onDispatch?: () => void): Promise<{ body: unknown; status: number; retryAfterMs?: number; operationAddress?: string }> {
    const parsed = new URL(url, ARM);
    if (parsed.origin !== ARM || parsed.username || parsed.password || parsed.hash) {
      throw new AssignmentApiError("shape", "The service returned an unsupported resource address.");
    }
    cancelled(signal);
    if (this.options.isCurrent?.() === false) throw new DOMException("The assignment context changed.", "AbortError");
    const token = await awaitAssignment(this.options.token(ASSIGNMENT_ARM_SCOPE), signal);
    cancelled(signal);
    if (this.options.isCurrent?.() === false) throw new DOMException("The assignment context changed.", "AbortError");
    let response: Response;
    try {
      const request: RequestInit = { ...init, signal, redirect: "error",
        headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...init.headers } };
      cancelled(signal, this.options.isCurrent);
      onDispatch?.();
      response = await awaitAssignment(this.fetcher(parsed.href, request), signal);
    } catch (error: unknown) {
      cancelled(signal, this.options.isCurrent);
      if (error instanceof AssignmentApiError || error instanceof DOMException && error.name === "AbortError") throw error;
      throw new AssignmentApiError("request", "The Azure request could not be completed. Try again.", true);
    }
    cancelled(signal);
    if (!response.ok) {
      const status = response.status;
      throw new AssignmentApiError(status === 401 ? "authentication" : status === 403 ? "permission" : "request",
        status === 401 ? "Connect to Azure again to check these resources."
          : status === 403 ? "You do not have permission to read or update this Azure resource."
            : status === 412 ? "The assignment changed. Refresh and review the device again."
              : `The Azure request failed (${status}).`,
        status === 408 || status === 429 || status >= 500, status, status === 401 ? "arm" : undefined, retryAfter(response));
    }
    if (response.status === 204 || init.method === "PATCH") return {
      body: undefined, status: response.status, retryAfterMs: retryAfter(response),
      operationAddress: response.headers.get("Azure-AsyncOperation") || response.headers.get("Location") || undefined,
    };
    let body: unknown;
    try { body = await response.json(); } catch {
      throw new AssignmentApiError("shape", "The Azure response could not be read.");
    }
    cancelled(signal);
    return { body, status: response.status, retryAfterMs: retryAfter(response) };
  }
  private async get(id: string, version: string, signal: AbortSignal): Promise<RecordValue> {
    return object((await this.request(`${ARM}${id}?api-version=${version}`, signal)).body, "resource information");
  }
  private async list(path: string, signal: AbortSignal): Promise<RecordValue[]> {
    let next = `${ARM}${path}`;
    const seen = new Set<string>(), ids = new Set<string>(), result: RecordValue[] = [];
    const collection = new URL(next).pathname.toLowerCase();
    while (next) {
      const url = new URL(next, ARM);
      if (url.origin !== ARM || url.pathname.toLowerCase() !== collection || seen.has(url.href) || seen.size >= 100) {
        throw new AssignmentApiError("shape", "The resource list returned incomplete or repeated pages.");
      }
      seen.add(url.href);
      const page = object((await this.request(url.href, signal)).body, "resource list");
      if (!Array.isArray(page.value)) throw new AssignmentApiError("shape", "The resource list did not contain a valid page.");
      for (const raw of page.value) {
        const item = object(raw, "resource list entry"), id = text(item.id, "resource identity").toLowerCase();
        if (ids.has(id) || result.length >= 10000) throw new AssignmentApiError("shape", "The resource list is duplicated or too large. Narrow the source selection.");
        ids.add(id); result.push(item);
      }
      if (page.nextLink !== undefined && page.nextLink !== null && typeof page.nextLink !== "string") {
        throw new AssignmentApiError("shape", "The resource list returned an invalid continuation.");
      }
      next = typeof page.nextLink === "string" ? page.nextLink : "";
    }
    return result;
  }
  private async resources(type: "microsoft.devices/iothubs",
    signal: AbortSignal, sources: string[] = []): Promise<RecordValue[]> {
    const query = `resources | where type =~ '${type}'` + (sources.length
      ? ` | where ${sources.map(source => `${armIdentity(source, "hub") ? "id" : "tostring(properties.hostName)"} =~ ${JSON.stringify(source)}`).join(" or ")}` : "") +
      " | project id, name, hostName=tostring(properties.hostName) | order by id asc";
    return this.argRows(query, signal);
  }
  private async argRows(query: string, signal: AbortSignal): Promise<RecordValue[]> {
    const result: RecordValue[] = [], seen = new Set<string>(), ids = new Set<string>();
    let skip: string | undefined;
    do {
      const page = object((await this.request(`${ARM}/providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01`, signal, {
        method: "POST", body: JSON.stringify({
          query,
          options: { resultFormat: "objectArray", $top: 1000, ...(skip ? { $skipToken: skip } : {}) },
        }),
      })).body, "discovery page");
      if (!Array.isArray(page.data)) throw new AssignmentApiError("shape", "Azure discovery returned an invalid resource page.");
      for (const raw of page.data) {
        const item = object(raw, "discovered resource"), id = text(item.id, "resource identity").toLowerCase();
        if (ids.has(id) || result.length >= 10000) throw new AssignmentApiError("shape", "Azure discovery returned duplicated or too many resources.");
        ids.add(id); result.push(item);
      }
      const next = page.$skipToken;
      if (next !== undefined && (typeof next !== "string" || !next || seen.has(next) || seen.size >= 100)) {
        throw new AssignmentApiError("shape", "Azure discovery returned an invalid continuation.");
      }
      skip = typeof next === "string" ? next : undefined;
      if (!skip && (page.resultTruncated === true || page.resultTruncated === "true")) {
        throw new AssignmentApiError("shape", "Azure discovery is incomplete. No complete assignment inventory is available.");
      }
      if (skip) seen.add(skip);
    } while (skip);
    return result;
  }
  async discoverHubs(source: AssignmentSource, signal: AbortSignal,
    progress: (message: string) => void = () => undefined): Promise<AssignmentHubSummary> {
    const inventory: AssignmentInventory = { hubs: [], targets: [], issues: [], hasMore: false };
    const provenance = sourceHubs(source);
    inventory.issues.push(...provenance.issues.map<AssignmentIssue>(message => ({ resource: "Source tables", message, retryable: false, scope: "source" })));
    if (!provenance.sources.length) {
      return { sourceSignature: canonicalJson(source), hubs: [], issues: inventory.issues, complete: false };
    }
    progress("Finding source Hubs");
    const hostnames = provenance.sources.filter(value => !armIdentity(value, "hub"));
    const candidates = hostnames.length ? await this.resources("microsoft.devices/iothubs", signal, hostnames) : [];
    for (const recorded of provenance.sources) {
      cancelled(signal);
      const matches = candidates.filter(item => armIdentity(recorded, "hub") ? typeof item.id === "string" && sameResource(item.id, recorded)
        : typeof item.hostName === "string" && item.hostName.toLowerCase() === recorded);
      const hubId = armIdentity(recorded, "hub") ? recorded : matches.length === 1 ? matches[0].id : undefined;
      if (typeof hubId !== "string" || !armIdentity(hubId, "hub") || !armIdentity(recorded, "hub") && matches.length !== 1) {
        inventory.issues.push({ resource: recorded, message: matches.length > 1
          ? "More than one accessible Hub matches this source. Devices for this source could not be verified."
          : "The source Hub could not be resolved. Check Azure access and the table's source information.", retryable: true, scope: "source" });
        continue;
      }
      if (inventory.hubs.some(hub => sameResource(hub.id, hubId))) continue;
      try {
        const hub = await this.get(hubId, "2026-10-01-preview", signal);
        if (typeof hub.id !== "string" || !sameResource(hub.id, hubId)) throw new AssignmentApiError("shape", "The returned Hub did not match the source.");
        const properties = object(hub.properties, "Hub properties");
        const hostname = text(properties.hostName, "Hub hostname");
        if (!armIdentity(recorded, "hub") && hostname.toLowerCase() !== recorded) {
          throw new AssignmentApiError("shape", "The current Hub hostname does not match the selected source.");
        }
        const name = text(hub.name, "Hub name");
        if (!Object.prototype.hasOwnProperty.call(properties, "deviceRegistry") || properties.deviceRegistry === null) {
          inventory.hubs.push({ id: hubId, name, state: "unlinked" }); continue;
        }
        const registry = object(properties.deviceRegistry, "Hub registry link");
        if (!armIdentity(registry.namespaceResourceId, "namespace")) throw new AssignmentApiError("shape", "The Hub returned an invalid namespace link.");
        const link = object(registry.linkingProperties, "Hub linking state");
        if (link.state !== "Succeeded") {
          if (!["Pending", "InProgress", "Failed", "Canceled"].includes(String(link.state))) {
            throw new AssignmentApiError("shape", "The Hub returned an unsupported registry-link state.");
          }
          inventory.hubs.push({ id: hubId, name, namespaceId: registry.namespaceResourceId,
            state: link.state === "Pending" || link.state === "InProgress" ? "pending" : "unlinked" });
          continue;
        }
        inventory.hubs.push({ id: hubId, name, namespaceId: registry.namespaceResourceId, state: "linked" });
      } catch (error: unknown) {
        this.issue(inventory, recorded, error, signal, "source");
        if (!inventory.hubs.some(hub => sameResource(hub.id, hubId))) {
          inventory.hubs.push({ id: hubId, name: recorded, state: "unverified" });
        }
      }
    }
    cancelled(signal, this.options.isCurrent);
    return { sourceSignature: canonicalJson(source), hubs: inventory.hubs, issues: inventory.issues,
      complete: inventory.hubs.length > 0 && inventory.issues.length === 0 };
  }
  async discover(source: AssignmentSource, signal: AbortSignal,
    progress: (message: string) => void = () => undefined, hubs: HubSummaryOptions = {}): Promise<AssignmentInventory> {
    const summary = hubs.initial || await this.discoverHubs(source, signal, progress);
    cancelled(signal, this.options.isCurrent);
    if (summary.sourceSignature !== canonicalJson(source)) {
      throw new AssignmentApiError("shape", "The Hub summary does not match the selected source.");
    }
    hubs.publish?.(summary);
    const inventory: AssignmentInventory = { hubs: [...summary.hubs], targets: [], issues: [...summary.issues], hasMore: false };
    const pendingPages: DevicePage[] = [];
    const namespaces = new Map<string, { id: string; hubIds: string[] }>();
    for (const hub of summary.hubs) {
      if (hub.state !== "linked") continue;
      if (!armIdentity(hub.id, "hub") || !armIdentity(hub.namespaceId, "namespace")) {
        throw new AssignmentApiError("shape", "The Hub summary contains an invalid namespace link.");
      }
      const key = hub.namespaceId.toLowerCase();
      const group = namespaces.get(key) || { id: hub.namespaceId, hubIds: [] };
      if (!group.hubIds.some(id => sameResource(id, hub.id))) group.hubIds.push(hub.id);
      namespaces.set(key, group);
    }
    for (const namespace of namespaces.values()) {
      progress("Checking devices linked to your source Hubs");
      try {
        const ns = await this.get(namespace.id, "2026-04-01", signal);
        if (typeof ns.id !== "string" || !sameResource(ns.id, namespace.id)) throw new AssignmentApiError("shape", "The returned namespace did not match the Hub link.");
        const messaging = object(object(ns.properties, "namespace properties").messaging, "namespace messaging");
        const endpoints = object(messaging.endpoints, "namespace endpoints");
        const page: DevicePage = { namespaceId: namespace.id, hubIds: namespace.hubIds, endpoints,
          next: `${ARM}${namespace.id}/registryDevices?api-version=${PREVIEW}`, seen: new Set(), ids: new Set() };
        pendingPages.push(page);
        try { await this.devicePage(page, inventory, signal); }
        catch (error: unknown) { this.issue(inventory, `${namespace.id}/registryDevices`, error, signal, "devices"); }
      } catch (error: unknown) { this.issue(inventory, namespace.id, error, signal, "source"); }
    }
    cancelled(signal);
    this.pendingPages = pendingPages.filter(page => !!page.next);
    inventory.hasMore = this.pendingPages.length > 0;
    return inventory;
  }
  async loadMore(previous: AssignmentInventory, signal: AbortSignal, publication?: ReadPublication): Promise<AssignmentInventory> {
    const inventory: AssignmentInventory = { ...previous, targets: previous.targets.filter(target => target.kind === "device"),
      issues: previous.issues.filter(issue => issue.scope !== "enrollments" && issue.authScope !== "dps") };
    const pages = this.pendingPages.map(page => ({ ...page, seen: new Set(page.seen), ids: new Set(page.ids) }));
    const page = pages[0];
    if (page) {
      try { await this.devicePage(page, inventory, signal); }
      catch (error: unknown) { this.issue(inventory, "More devices", error, signal, "devices"); }
    }
    cancelled(signal, () => this.options.isCurrent?.() !== false && publication?.isCurrent() !== false);
    this.pendingPages = pages.filter(value => !!value.next);
    inventory.hasMore = this.pendingPages.length > 0;
    publication?.publish(inventory);
    return inventory;
  }
  private async devicePage(page: DevicePage, inventory: AssignmentInventory, signal: AbortSignal): Promise<void> {
    cancelled(signal);
    const url = new URL(page.next, ARM);
    const expected = `${page.namespaceId}/registryDevices`.toLowerCase();
    if (url.origin !== ARM || url.pathname.toLowerCase() !== expected || page.seen.has(url.href) || page.seen.size >= 100) {
      throw new AssignmentApiError("shape", "The device list returned incomplete or repeated pages.");
    }
    const body = object((await this.request(url.href, signal)).body, "device page");
    if (!Array.isArray(body.value)) throw new AssignmentApiError("shape", "The device list did not contain a valid page.");
    const devices = body.value.map(raw => object(raw, "device list entry"));
    const incoming = new Set<string>();
    for (const device of devices) {
      if (!armIdentity(device.id, "device") || !device.id.toLowerCase().startsWith(`${page.namespaceId.toLowerCase()}/registrydevices/`) ||
          page.ids.has(device.id.toLowerCase()) || incoming.has(device.id.toLowerCase()) || page.ids.size + devices.length > 10000) {
        throw new AssignmentApiError("shape", "The device page contains an invalid or duplicate identity.");
      }
      incoming.add(device.id.toLowerCase());
    }
    let next = "";
    if (body.nextLink !== undefined && body.nextLink !== null && body.nextLink !== "") {
      if (typeof body.nextLink !== "string") throw new AssignmentApiError("shape", "The device page returned an invalid continuation.");
      const link = new URL(body.nextLink, ARM);
      if (link.origin !== ARM || link.pathname.toLowerCase() !== expected || page.seen.has(link.href) || link.href === url.href) {
        throw new AssignmentApiError("shape", "The device page returned an unsupported continuation.");
      }
      next = link.href;
    }
    for (const device of devices) {
      try {
        const deviceId = text(device.id, "device identity");
        const target = await this.deviceTarget(deviceId, page, signal);
        if (target) inventory.targets.push(target);
      } catch (error: unknown) { this.issue(inventory, typeof device.id === "string" ? device.id : "Device", error, signal, "devices"); }
    }
    page.seen.add(url.href);
    for (const id of incoming) page.ids.add(id);
    page.next = next;
  }
  private async deviceTarget(deviceId: string, context: Pick<DevicePage, "namespaceId" | "hubIds" | "endpoints">,
    signal: AbortSignal): Promise<AssignmentTarget | undefined> {
    if (!armIdentity(deviceId, "device") || !deviceId.toLowerCase().startsWith(`${context.namespaceId.toLowerCase()}/registrydevices/`)) {
      throw new AssignmentApiError("shape", "This device does not belong to the verified namespace.");
    }
    const capabilities = await this.list(`${deviceId}/capabilities?api-version=${PREVIEW}`, signal);
    const hubIds = new Set<string>();
    for (const capability of capabilities) {
      if (typeof capability.id !== "string" || !capability.id.toLowerCase().startsWith(`${deviceId.toLowerCase()}/capabilities/`)) {
        throw new AssignmentApiError("shape", "The capability does not belong to this device.");
      }
      const capabilityProperties = object(capability.properties, "device capability");
      if (capabilityProperties.capabilityType !== "Microsoft.IoTHub") continue;
      const endpointName = capabilityProperties.messagingEndpointName;
      if (typeof endpointName !== "string" || !Object.prototype.hasOwnProperty.call(context.endpoints, endpointName)) continue;
      const endpoint = object(context.endpoints[endpointName], "messaging endpoint");
      if (endpoint.endpointType !== "Microsoft.Devices/IotHubs" || !armIdentity(endpoint.resourceId, "hub")) continue;
      for (const hubId of context.hubIds) if (sameResource(hubId, endpoint.resourceId)) hubIds.add(hubId);
    }
    if (!hubIds.size) return undefined;
    const pointRead = await this.request(`${ARM}${deviceId}?api-version=${PREVIEW}`, signal);
    if (pointRead.status !== 200) throw new AssignmentApiError("shape", "The current device response is not a complete ordinary resource read.");
    const current = object(pointRead.body, "device information");
    if (typeof current.id !== "string" || !sameResource(current.id, deviceId)) throw new AssignmentApiError("shape", "The returned device identity changed.");
    const currentProperties = object(current.properties, "device properties");
    cancelled(signal, this.options.isCurrent);
    const rawMissing = !Object.prototype.hasOwnProperty.call(currentProperties, "deviceTypeRefs");
    const uuid = currentProperties.uuid, externalDeviceId = currentProperties.externalDeviceId;
    let pointIssue: string | undefined;
    if (typeof current.type !== "string" || current.type.toLowerCase() !== "microsoft.deviceregistry/namespaces/registrydevices" ||
        currentProperties.provisioningState !== "Succeeded" || typeof uuid !== "string" ||
        !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(uuid) ||
        uuid === "00000000-0000-0000-0000-000000000000" ||
        typeof externalDeviceId !== "string" || !externalDeviceId.trim()) {
      pointIssue = "The current device read did not verify its native identity and successful provider state.";
    } else {
      try { specificEtag(current.etag); }
      catch { pointIssue = "The current device read did not return a valid specific ETag."; }
    }
    const pointVerified = pointIssue === undefined;
    // In this provider version, an omitted optional field on a full GET means stored not-set, not an observed empty array.
    const reference: AssignmentTarget["reference"] = pointIssue ? { kind: "unknown", reason: pointIssue }
      : rawMissing ? { kind: "unassigned", ids: [] } : readReference(currentProperties);
    const deviceIdentity = pointVerified && typeof uuid === "string" && typeof externalDeviceId === "string"
      ? { uuid, externalDeviceId } : undefined;
    const currentArmEtag = pointVerified && reference.kind !== "unknown" && typeof current.etag === "string"
      ? current.etag : undefined;
    let updatePermission: AssignmentTarget["updatePermission"] = "unknown", writeBlocker: string | undefined;
    if (reference.kind === "unassigned") {
      try {
        updatePermission = await this.updatePermission(deviceId, signal) ? "granted" : "not-granted";
        if (updatePermission === "not-granted") writeBlocker = "You can view these references, but can't change them with your current access.";
      } catch (error: unknown) {
        cancelled(signal);
        writeBlocker = `${UPDATE_PERMISSION_UNKNOWN} ${error instanceof Error ? error.message : "Refresh to check access again."}`;
      }
      cancelled(signal, this.options.isCurrent);
    }
    return { id: deviceId, name: text(current.name, "device name"), kind: "device",
      namespaceId: context.namespaceId, hubIds: [...hubIds], reference, updatePermission,
      referenceEvidence: { source: "adr-point-get", observedAt: new Date().toISOString(),
        ...(currentArmEtag ? { currentArmEtag } : {}),
        apiVersion: PREVIEW, rawCarrier: rawMissing ? "missing" : "present" },
      ...(deviceIdentity ? { deviceIdentity } : {}),
      ...(typeof current.etag === "string" ? { etag: current.etag } : {}), ...(writeBlocker ? { writeBlocker } : {}) };
  }
  private async updatePermission(id: string, signal: AbortSignal): Promise<boolean> {
    return permissionGrants(await this.permissions(id, signal), "Microsoft.DeviceRegistry/namespaces/registryDevices/write");
  }
  private async permissions(id: string, signal: AbortSignal): Promise<unknown[]> {
    const path = `${id}/providers/Microsoft.Authorization/permissions`;
    let next = `${ARM}${path}?api-version=2022-04-01`;
    const seen = new Set<string>(), entries: unknown[] = [];
    while (next) {
      const url = new URL(next, ARM);
      if (url.origin !== ARM || url.pathname.toLowerCase() !== path.toLowerCase() || seen.has(url.href) || seen.size >= 100) {
        throw new AssignmentApiError("shape", "Update-permission pages could not be verified.");
      }
      seen.add(url.href);
      const page = object((await this.request(url.href, signal)).body, "update permissions");
      if (!Array.isArray(page.value)) throw new AssignmentApiError("shape", "Update permissions were not returned.");
      entries.push(...page.value);
      if (page.nextLink !== undefined && page.nextLink !== null && typeof page.nextLink !== "string") {
        throw new AssignmentApiError("shape", "The update-permission continuation is invalid.");
      }
      next = typeof page.nextLink === "string" ? page.nextLink : "";
    }
    cancelled(signal, this.options.isCurrent);
    return entries;
  }
  async recheck(target: AssignmentTarget, source: AssignmentSource, signal: AbortSignal): Promise<AssignmentTarget> {
    if (target.kind !== "device" || !armIdentity(target.namespaceId, "namespace") || !target.hubIds.length) {
      throw new AssignmentApiError("unsupported", "This device cannot be checked safely with this connection.");
    }
    const recorded = sourceHubs(source).sources;
    for (const id of target.hubIds) {
      if (!armIdentity(id, "hub")) throw new AssignmentApiError("shape", "The device's source Hub could not be verified.");
      const hub = await this.get(id, "2026-10-01-preview", signal);
      const properties = object(hub.properties, "Hub properties");
      const registry = object(properties.deviceRegistry, "current Hub registry link");
      if (typeof hub.id !== "string" || !sameResource(hub.id, id) ||
          !recorded.some(value => armIdentity(value, "hub") ? sameResource(value, id)
            : typeof properties.hostName === "string" && properties.hostName.toLowerCase() === value) ||
          typeof registry.namespaceResourceId !== "string" || !sameResource(registry.namespaceResourceId, target.namespaceId) ||
          object(registry.linkingProperties, "Hub linking state").state !== "Succeeded") {
        throw new AssignmentApiError("shape", "The source Hub link changed. Refresh devices before assigning.");
      }
    }
    const namespace = await this.get(target.namespaceId, "2026-04-01", signal);
    if (typeof namespace.id !== "string" || !sameResource(namespace.id, target.namespaceId)) {
      throw new AssignmentApiError("shape", "The current namespace identity could not be verified.");
    }
    const endpoints = object(object(object(namespace.properties, "namespace properties").messaging, "namespace messaging").endpoints, "namespace endpoints");
    const current = await this.deviceTarget(target.id, { namespaceId: target.namespaceId, hubIds: target.hubIds, endpoints }, signal);
    if (!current) throw new AssignmentApiError("shape", "This device is no longer linked to a selected source Hub.");
    return current;
  }
  async readAssignment(target: AssignmentTarget, modelId: string, signal: AbortSignal, trackingId?: string): Promise<AssignmentReadback> {
    cancelled(signal, this.options.isCurrent);
    const operation = trackingId ? nativeAssignmentOperations.get(trackingId) : undefined;
    if (!operation || !sameResource(operation.targetId, target.id) || operation.modelId !== modelId) {
      throw new AssignmentTrackingUnavailableError();
    }
    if (operation.terminal) return operation.terminal;
    const remaining = operation.deadline - Date.now();
    if (remaining <= 0) throw new AssignmentApiError("request", ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
    const url = operationUrl(operation.url, target.id);
    const controller = new AbortController();
    const abort = () => controller.abort();
    signal.addEventListener("abort", abort, { once: true });
    const timer = setTimeout(abort, remaining);
    try {
      const response = await this.request(url, controller.signal);
      const body = object(response.body, "assignment operation status");
      if (response.status !== 200 && response.status !== 202) {
        throw new AssignmentApiError("shape", "The assignment operation returned an unsupported status response.");
      }
      const status = body.status;
      if (status === "Succeeded") operation.terminal = { confirmed: true };
      else if (status === "Failed") operation.terminal = { confirmed: false, terminalState: "failed" };
      else if (status === "Canceled" || status === "Cancelled") operation.terminal = { confirmed: false, terminalState: "cancelled" };
      else if (!["Accepted", "Running", "InProgress", "NotStarted", "Updating"].includes(String(status))) {
        throw new AssignmentApiError("shape", "The assignment operation returned a missing or unrecognized status.");
      }
      if (operation.terminal) operation.url = "";
      return operation.terminal || { confirmed: false };
    } finally {
      clearTimeout(timer);
      signal.removeEventListener("abort", abort);
    }
  }
  assignmentServices(source: AssignmentSource): AssignmentServices {
    return {
      read: (target, signal) => this.recheck(target, source, signal),
      write: (target, id, signal, onDispatch, deadline) => this.patchReference(target, id, signal, source, onDispatch, deadline),
      trackingId: target => this.operationIds.get(target.id.toLowerCase()),
      trackingDeadline: target => {
        const id = this.operationIds.get(target.id.toLowerCase());
        const operation = id ? nativeAssignmentOperations.get(id) : undefined;
        if (!operation) throw new AssignmentTrackingUnavailableError();
        return operation.deadline;
      },
      observe: (target, receipt, signal, current, publish) => pollAssignmentReadback({
        operationId: receipt.operationId, modelId: receipt.modelId, targetIds: [target.id], targets: [target],
      }, [receipt], {
        current, publish,
        read: (value, modelId, abort, trackingId) => this.readAssignment(value, modelId, abort, trackingId),
      }, signal),
    };
  }

  private issue(inventory: AssignmentInventory, resource: string, error: unknown, signal: AbortSignal, scope: AssignmentIssueScope,
    view?: "group" | "individual"): void {
    cancelled(signal, this.options.isCurrent);
    inventory.issues.push({ resource, message: error instanceof AssignmentApiError ? error.message : "The resource could not be verified.",
      retryable: error instanceof AssignmentApiError && error.retryable, scope,
      ...(view ? { view } : {}),
      ...(error instanceof AssignmentApiError ? { kind: error.kind, ...(error.authScope ? { authScope: error.authScope } : {}) } : {}) });
  }
  async patchReference(target: AssignmentTarget, modelId: string, signal: AbortSignal, source: AssignmentSource,
    onDispatch?: () => void, deadline = Date.now() + ASSIGNMENT_OPERATION_TIMEOUT_MS): Promise<"accepted" | "completed"> {
    cancelled(signal);
    if (target.kind !== "device" || !armIdentity(target.id, "device") || !modelId.trim() || modelId.length > 512) {
      throw new AssignmentApiError("shape", "The selected assignment is not supported.");
    }
    const blocker = assignmentBlocker(target);
    if (blocker) throw new AssignmentApiError("unsupported", blocker);
    const controller = new AbortController();
    const abort = () => controller.abort();
    signal.addEventListener("abort", abort, { once: true });
    const timer = setTimeout(abort, Math.max(0, deadline - Date.now()));
    try {
      if (signal.aborted || Date.now() >= deadline) controller.abort();
      cancelled(controller.signal, this.options.isCurrent);
      const fresh = await awaitAssignment(this.recheck(target, source, controller.signal), controller.signal);
      cancelled(controller.signal, this.options.isCurrent);
      const freshBlocker = assignmentBlocker(fresh);
      if (freshBlocker) throw new AssignmentApiError("unsupported", freshBlocker);
      if (fresh.updatePermission !== "granted" || fresh.etag !== target.etag ||
          JSON.stringify(fresh.deviceIdentity) !== JSON.stringify(target.deviceIdentity) ||
          fresh.namespaceId.toLowerCase() !== target.namespaceId.toLowerCase() ||
          JSON.stringify(fresh.hubIds.map(id => id.toLowerCase()).sort()) !== JSON.stringify(target.hubIds.map(id => id.toLowerCase()).sort())) {
        throw new AssignmentApiError("request", "The device changed. Refresh and review the assignment again.", false, 412);
      }
      if (Date.now() >= deadline) controller.abort();
      cancelled(controller.signal, this.options.isCurrent);
      const result = await this.request(`${ARM}${target.id}?api-version=${PREVIEW}`, controller.signal, {
        method: "PATCH", headers: { "If-Match": specificEtag(fresh.etag) }, body: JSON.stringify({ properties: { deviceTypeRefs: [modelId] } }),
      }, onDispatch);
      if (result.status === 202) {
        if (!result.operationAddress) throw new AssignmentApiError("shape", "The accepted assignment did not expose operation tracking information. Completion is unknown; no assignment will be retried.");
        const url = operationUrl(result.operationAddress, target.id);
        const pollingDeadline = Date.now() + ASSIGNMENT_POLL_TIMEOUT_MS;
        for (const [id, operation] of nativeAssignmentOperations) if (operation.deadline <= Date.now()) nativeAssignmentOperations.delete(id);
        const id = crypto.randomUUID();
        nativeAssignmentOperations.set(id, { url, targetId: target.id, modelId, deadline: pollingDeadline });
        this.operationIds.set(target.id.toLowerCase(), id);
        return "accepted";
      }
      if (result.status !== 200 && result.status !== 204) throw new AssignmentApiError("shape", "The assignment returned an unsupported completion response.");
      return "completed";
    } finally {
      clearTimeout(timer);
      signal.removeEventListener("abort", abort);
    }
  }
}

export const ASSIGNMENT_POLL_INTERVAL_MS = 3000;
function assignmentPollDelay(completedPolls: number): number {
  return completedPolls < 3 ? ASSIGNMENT_POLL_INTERVAL_MS : completedPolls < 6 ? 5000 : 10000;
}
function waitForReadback(milliseconds: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    const abort = () => { clearTimeout(timer); reject(new DOMException("Assignment verification cancelled.", "AbortError")); };
    if (signal.aborted) { reject(new DOMException("Assignment verification cancelled.", "AbortError")); return; }
    const timer = setTimeout(() => { signal.removeEventListener("abort", abort); resolve(); }, milliseconds);
    signal.addEventListener("abort", abort, { once: true });
  });
}
/** Tracks only the original mutation's operation; never reads device state or replays assignments. */
export async function pollAssignmentReadback(attempt: AssignmentAttempt, receipts: AssignmentReceipt[],
  services: { read: (target: AssignmentTarget, modelId: string, signal: AbortSignal, trackingId?: string) => Promise<AssignmentReadback>;
    current: () => boolean; publish: (receipt: AssignmentReceipt) => void;
    now?: () => number; random?: () => number;
    wait?: (milliseconds: number, signal: AbortSignal) => Promise<void> }, signal: AbortSignal): Promise<void> {
  const now = services.now || (() => Date.now());
  const random = services.random || Math.random;
  const currentReceipts = new Map<string, AssignmentReceipt>();
  const targets = new Map((attempt.targets || []).map(target => [target.id.toLowerCase(), target]));
  for (const id of attempt.targetIds) {
    const receipt = receipts.find(value => value.operationId === attempt.operationId && sameResource(value.targetId, id) &&
      value.modelId === attempt.modelId && value.kind === "device");
    if (!receipt) continue;
    if (receipt.state === "failed" && receipt.writeDispatched !== false && receipt.readbackFailure?.kind === "shape" &&
        receipt.readbackFailure.retryable === false && receipt.message === ASSIGNMENT_TRACKING_UNAVAILABLE_MESSAGE) {
      cancelled(signal, services.current);
      services.publish({ ...receipt, state: "unknown" });
      continue;
    }
    if (!["accepted", "unknown"].includes(receipt.state)) continue;
    if (receipt.state === "unknown" && receipt.readbackFailure?.retryable === false) {
      if (receipt.writeDispatched !== false && receipt.readbackFailure.kind === "request" &&
          Number.isFinite(receipt.readbackDeadline) && receipt.readbackDeadline! <= now() &&
          receipt.message !== ASSIGNMENT_UNTRACKED_DISPATCH_MESSAGE &&
          receipt.message !== ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE) {
        cancelled(signal, services.current);
        services.publish({ ...receipt, message: ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE });
      }
      continue;
    }
    if (receipt.state === "unknown" && receipt.writeDispatched === true && !receipt.trackingId) {
      cancelled(signal, services.current);
      services.publish({ ...receipt, message: ASSIGNMENT_UNTRACKED_DISPATCH_MESSAGE,
        readbackFailure: { kind: "request", retryable: false } });
      continue;
    }
    if (!Number.isFinite(receipt.readbackDeadline) || receipt.readbackDeadline! <= now()) {
      cancelled(signal, services.current);
      const expired = Number.isFinite(receipt.readbackDeadline);
      services.publish({ ...receipt, state: "unknown",
        message: expired ? ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE : ASSIGNMENT_TRACKING_UNAVAILABLE_MESSAGE,
        readbackFailure: { kind: expired ? "request" : "shape", retryable: false } });
      continue;
    }
    const target = targets.get(id.toLowerCase());
    if (!target) {
      cancelled(signal, services.current);
      services.publish({ ...receipt, state: "failed", message: "The original device selection is unavailable. Verification cannot resume safely.",
        readbackFailure: { kind: "shape", retryable: false } });
      continue;
    }
    currentReceipts.set(id.toLowerCase(), receipt);
  }
  while (currentReceipts.size) {
    const jitter = [...currentReceipts.values()].some(receipt => receipt.readbackNotBefore === undefined)
      ? random() * 2000 - 1000 : 0;
    for (const [id, receipt] of currentReceipts) {
      if (receipt.readbackNotBefore !== undefined) continue;
      cancelled(signal, services.current);
      const scheduled: AssignmentReceipt = { ...receipt,
        readbackNotBefore: now() + assignmentPollDelay(receipt.readbackPollCount || 0) + jitter };
      currentReceipts.set(id, scheduled);
      services.publish(scheduled);
    }
    const delay = Math.max(...[...currentReceipts.values()].map(receipt => receipt.readbackNotBefore! - now()));
    const remaining = Math.min(...[...currentReceipts.values()].map(receipt => receipt.readbackDeadline! - now()));
    await (services.wait || waitForReadback)(Math.max(0, Math.min(delay, remaining)), signal);
    cancelled(signal, services.current);
    for (const [id, previous] of currentReceipts) {
      cancelled(signal, services.current);
      const controller = new AbortController();
      const abort = () => controller.abort();
      signal.addEventListener("abort", abort, { once: true });
      const remaining = previous.readbackDeadline! - now();
      const receipt: AssignmentReceipt = { ...previous, readbackNotBefore: undefined,
        readbackPollCount: (previous.readbackPollCount || 0) + (remaining > 0 ? 1 : 0) };
      const timer = setTimeout(abort, Math.max(0, remaining));
      try {
        if (remaining <= 0) throw new AssignmentApiError("request", ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
        services.publish(receipt);
        const result = await awaitAssignment(services.read(targets.get(id)!, attempt.modelId, controller.signal, receipt.trackingId), controller.signal);
        cancelled(signal, services.current);
        if (receipt.readbackDeadline! <= now()) {
          throw new AssignmentApiError("request", ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE);
        }
        if (result.confirmed) {
          services.publish({ ...receipt, state: "confirmed", readbackFailure: undefined,
            message: "The assignment operation completed. Current device state and matching data have not been verified." });
          currentReceipts.delete(id);
        } else if (result.terminalState) {
          services.publish({ ...receipt, state: result.terminalState, readbackFailure: undefined,
            message: result.terminalState === "cancelled" ? "The assignment operation was canceled." : "The assignment operation failed." });
          currentReceipts.delete(id);
        } else {
          const next: AssignmentReceipt = { ...receipt,
            readbackFailure: undefined, message: "Waiting for the assignment operation to complete." };
          currentReceipts.set(id, next);
          services.publish(next);
        }
      } catch (failure: unknown) {
        cancelled(signal, services.current);
        const error = controller.signal.aborted ? new AssignmentApiError("request",
          ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE)
          : failure instanceof AssignmentApiError ? failure
          : new AssignmentApiError("request", failure instanceof Error ? failure.message : "Assignment verification could not be completed.");
        const next: AssignmentReceipt = { ...receipt,
          state: error.retryable || failure instanceof AssignmentTrackingUnavailableError ||
            controller.signal.aborted || receipt.readbackDeadline! <= now() ? "unknown" : "failed",
          message: error.retryable ? `${error.message} Verification will retry; no assignment will be sent again.` : error.message,
          readbackFailure: { kind: error.kind, retryable: error.retryable } };
        services.publish(next);
        if (!error.retryable) currentReceipts.delete(id); else currentReceipts.set(id, next);
      } finally {
        clearTimeout(timer);
        signal.removeEventListener("abort", abort);
      }
    }
  }
}

/** Common discovery belongs to the source owner, not whichever tab is awaiting it. */
export class HubAssignmentReader {
  private readonly api: HubAssignmentApi;
  private readonly source: AssignmentSource;
  private readonly publish: (inventory: AssignmentInventory) => void;
  private readonly isCurrent: () => boolean;
  private readonly hubs: HubSummaryOptions;
  private inventory?: AssignmentInventory;
  private common?: { controller: AbortController; promise: Promise<AssignmentInventory> };
  private generation = 0;
  private disposed = false;
  constructor(api: HubAssignmentApi, source: AssignmentSource,
    publish: (inventory: AssignmentInventory) => void, isCurrent: () => boolean,
    hubs: HubSummaryOptions = {}) {
    this.api = api;
    this.source = source;
    this.publish = publish;
    this.isCurrent = isCurrent;
    this.hubs = hubs;
  }

  dispose(): void {
    this.disposed = true; this.generation++; this.common?.controller.abort();
  }
  updateTarget(target: AssignmentTarget): void {
    if (target.kind !== "device" || !this.current() || !this.inventory?.targets.some(item => item.id === target.id)) return;
    const inventory = { ...this.inventory, targets: this.inventory.targets.map(item => item.id === target.id ? target : item) };
    this.inventory = inventory;
    this.publish(inventory);
  }
  private current(): boolean { return !this.disposed && this.isCurrent(); }
  private discover(refresh: boolean): Promise<AssignmentInventory> {
    if (this.common && !refresh) return this.common.promise;
    this.common?.controller.abort();
    const controller = new AbortController();
    const promise = this.api.discover(this.source, controller.signal, undefined,
      { initial: refresh ? undefined : this.hubs.initial, publish: this.hubs.publish }).then(base => {
      cancelled(controller.signal, () => this.current());
      const value: AssignmentInventory = { ...base,
        targets: base.targets.filter(target => target.kind === "device"),
        issues: base.issues.filter(issue => issue.scope !== "enrollments" && issue.authScope !== "dps") };
      this.inventory = value;
      this.publish(value);
      return value;
    });
    this.common = { controller, promise };
    return promise;
  }
  async load(kind: AssignmentTarget["kind"], mode: "view" | "refresh" | "more", signal: AbortSignal): Promise<AssignmentInventory> {
    if (kind !== "device") throw new AssignmentApiError("unsupported", DPS_INFORMATIONAL_ONLY);
    const ticket = ++this.generation;
    const current = () => this.current() && ticket === this.generation;
    cancelled(signal, current);
    await awaitAssignment(this.discover(mode === "refresh"), signal);
    cancelled(signal, current);
    const value = this.inventory;
    if (!value) throw new AssignmentApiError("shape", "The source inventory has not been published.");
    const publication: ReadPublication = { isCurrent: current, publish: inventory => {
      this.inventory = inventory; this.publish(inventory);
    } };
    return mode === "more" ? this.api.loadMore(value, signal, publication) : value;
  }
}
