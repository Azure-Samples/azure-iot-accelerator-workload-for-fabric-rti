// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { canonicalJson } from "./ModelAuthoring.ts";
import { assignableReference, modelValidation, type HubModel } from "./HubModeling.ts";
import { sourceSelectionRevision, type SourceReadiness } from "./ModelSourceReadiness.ts";

export interface AssignmentAttempt { operationId: string; modelId: string; targetIds: string[]; targets?: AssignmentTarget[] }
export interface AssignmentIntent {
  owner: string; mode: "assign" | "skip"; attempt?: AssignmentAttempt; listLoading?: AssignmentListLoading;
}
export class AssignmentListLoading {
  readonly owner: string;
  private readonly requests = new Set<object>();
  private reserved = false;
  private active = true;
  private notify: (() => void) | undefined;
  constructor(owner: string, notify: () => void) { this.owner = owner; this.notify = notify; }
  activate(): void { this.active = true; }
  get busy(): boolean { return this.active && (this.reserved || this.requests.size > 0); }
  reserve(): void { if (this.active) { this.reserved = true; this.notify?.(); } }
  releaseReservation(): void { if (this.active && this.reserved) { this.reserved = false; this.notify?.(); } }
  begin(): () => void {
    if (!this.active) return () => undefined;
    const ticket = {};
    this.requests.add(ticket); this.reserved = false; this.notify?.();
    return () => { if (this.requests.delete(ticket)) this.notify?.(); };
  }
  dispose(): void { this.active = false; this.reserved = false; this.requests.clear(); }
}
export function assignmentIntentOwner(context: Record<string, any>, source: SourceReadiness | undefined): string {
  const model: HubModel | undefined = context.hubModel;
  if (!source?.ready || !source.key || !model || modelValidation(model).length) return "";
  return canonicalJson([source.key, source.revision, sourceSelectionRevision(context), context.hubModelRevision || 0,
    model.kind, model.id || "", model.document, context.hubAssignmentSource || null]);
}
export function currentAssignmentIntent(context: Record<string, any>, source: SourceReadiness | undefined): AssignmentIntent | undefined {
  const owner = assignmentIntentOwner(context, source), intent: AssignmentIntent | undefined = context.hubAssignmentIntent;
  return owner && intent?.owner === owner && (intent.mode === "skip" ||
    intent.mode === "assign" && context.hubModel?.kind === "wot")
    ? intent : undefined;
}
export function currentAssignmentListsLoading(context: Record<string, any>, source: SourceReadiness | undefined): boolean {
  const intent = currentAssignmentIntent(context, source), loading = intent?.listLoading;
  return intent?.mode === "assign" && loading instanceof AssignmentListLoading && loading.owner === intent.owner && loading.busy;
}
export interface AssignmentHubSummary {
  sourceSignature: string;
  hubs: AssignmentInventory["hubs"];
  issues: AssignmentIssue[];
  complete: boolean;
}
export function currentAssignmentHubSummary(context: Record<string, any>, source: SourceReadiness | undefined): AssignmentHubSummary | undefined {
  const owner = assignmentIntentOwner(context, source), cached = context.hubAssignmentHubSummary;
  const value: AssignmentHubSummary | undefined = cached?.value;
  return owner && context.hubModel?.kind === "wot" && cached?.owner === owner &&
    context.hubAssignmentSource?.key === source?.key && value?.sourceSignature === canonicalJson(context.hubAssignmentSource) &&
    Array.isArray(value.hubs) && Array.isArray(value.issues) && typeof value.complete === "boolean" ? value : undefined;
}
export function modelReferenceFilterAvailability(context: Record<string, any>, source: SourceReadiness | undefined):
  "available" | "dtdl" | "unlinked" | "unconfirmed" {
  if (context.hubModel?.kind === "dtdl") return "dtdl";
  const summary = currentAssignmentHubSummary(context, source);
  if (summary?.hubs.some(hub => hub.state === "linked")) return "available";
  return summary?.complete && !summary.issues.length && summary.hubs.length > 0 &&
    summary.hubs.every(hub => hub.state === "unlinked") ? "unlinked" : "unconfirmed";
}
export function allAssignmentHubsUnavailable(summary: AssignmentHubSummary | undefined): boolean {
  return summary?.complete === true && !summary.issues.length && summary.hubs.length > 0 &&
    summary.hubs.every(hub => hub.state === "unlinked" || hub.state === "pending");
}

export interface AssignmentSource {
  key: string;
  telemetry: { table: string; docstring: string };
  twin: { table: string; docstring: string };
}
export type ReferenceState =
  | { kind: "unassigned"; ids: [] }
  | { kind: "assigned"; ids: [string] }
  | { kind: "unknown"; reason: string };
export interface ReferenceEvidence {
  source: "adr-point-get" | "dps-api";
  observedAt: string;
  currentArmEtag?: string;
  apiVersion?: "2026-11-02-preview";
  rawCarrier?: "missing" | "present";
}
export interface AssignmentTarget {
  id: string;
  name: string;
  kind: "device" | "group" | "individual";
  hubIds: string[];
  namespaceId: string;
  reference: ReferenceState;
  referenceEvidence?: ReferenceEvidence;
  deviceIdentity?: { uuid: string; externalDeviceId: string };
  etag?: string;
  updatePermission?: "granted" | "not-granted" | "unknown";
  writeBlocker?: string;
}
export type AssignmentIssueScope = "source" | "devices" | "enrollments";
export interface AssignmentIssue {
  resource: string; message: string; retryable: boolean;
  kind?: "permission" | "authentication" | "shape" | "request" | "unsupported";
  scope?: AssignmentIssueScope;
  view?: "group" | "individual";
  authScope?: "arm" | "dps";
}
export interface AssignmentViewIssue {
  key: string; message: string; retryable: boolean; kind?: AssignmentIssue["kind"]; count: number;
  authScope?: AssignmentIssue["authScope"];
}
export function assignmentIssuesForView(issues: AssignmentIssue[], kind: AssignmentTarget["kind"]): AssignmentViewIssue[] {
  const groups = new Map<string, { issue: AssignmentIssue; resources: Set<string> }>();
  for (const issue of issues) {
    // Retained inventories without scope remain conservative until refreshed.
    const scope = issue.scope || "source";
    if (kind === "device" && issue.authScope === "dps") continue;
    if (issue.view && issue.view !== kind) continue;
    if (scope === "devices" && kind !== "device" || scope === "enrollments" && kind === "device") continue;
    const key = canonicalJson([scope, issue.view || "", issue.authScope || "", issue.kind || "", issue.message, issue.retryable]);
    const group = groups.get(key) || { issue, resources: new Set<string>() };
    group.resources.add(issue.resource);
    groups.set(key, group);
  }
  return [...groups].map(([key, group]) => ({ key, message: group.issue.message, retryable: group.issue.retryable,
    kind: group.issue.kind, count: group.resources.size, ...(group.issue.authScope ? { authScope: group.issue.authScope } : {}) }));
}
export interface AssignmentInventory {
  hasMore: boolean;
  hubs: { id: string; name: string; namespaceId?: string; state: "linked" | "unlinked" | "pending" | "unverified" }[];
  targets: AssignmentTarget[];
  issues: AssignmentIssue[];
}
export const DPS_INFORMATIONAL_ONLY = "DPS enrollments are updated outside this workload. Only ADR devices can be assigned here.";
export const UPDATE_PERMISSION_UNKNOWN = "Permission to update hasn't been confirmed.";
export const CURRENT_REFERENCE_FRESHNESS = "The current device reference couldn't be verified. Refresh devices before assigning a model. This device remains read-only.";
export function specificEtag(value: unknown): string {
  if (typeof value !== "string" || !value || /[\u0000-\u0020\u007f,]/.test(value) || value === "*" || value === '"*"') {
    throw new Error("A current specific conditional-update receipt is required.");
  }
  if (/^(W\/)?"[^"]+"$/.test(value)) return value;
  if (value.includes('"')) throw new Error("The conditional-update receipt is not valid.");
  return `"${value}"`;
}
export function permissionGrants(value: unknown, action: string): boolean {
  if (!Array.isArray(value)) throw new Error("The update-permission response is incomplete.");
  const matches = (pattern: string) => new RegExp("^" + pattern.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")
    .replace(/\\\*/g, ".*") + "$", "i").test(action);
  let allowed = false;
  for (const entry of value) {
    if (!entry || typeof entry !== "object" || !Array.isArray(entry.actions) || !Array.isArray(entry.notActions) ||
        [...entry.actions, ...entry.notActions].some(item => typeof item !== "string" || !item || /\s/.test(item))) {
      throw new Error("The update-permission response is incomplete.");
    }
    if (entry.actions.some(matches) && !entry.notActions.some(matches)) allowed = true;
  }
  return allowed;
}

export function sourceHubs(source: AssignmentSource): { sources: string[]; issues: string[] } {
  const sources = new Set<string>(), issues: string[] = [];
  for (const stream of [source.telemetry, source.twin]) {
    const lines = [...stream.docstring.matchAll(/^\s*@iot\.meta\s+source=(.*)$/gm)];
    if (!lines.length) issues.push(`No source Hub is recorded for ${stream.table}. You can continue modeling without assignment.`);
    for (const line of lines) for (const part of line[1].split(";")) {
      const value = part.trim();
      if (!value) continue;
      if (/^[a-z0-9-]+\.azure-devices\.net$/i.test(value) || armIdentity(value, "hub")) sources.add(value.toLowerCase());
      else issues.push(`The source Hub recorded for ${stream.table} is not a supported hostname or Hub resource ID.`);
    }
  }
  return { sources: [...sources], issues: [...new Set(issues)] };
}

const resourcePatterns = {
  hub: /^\/subscriptions\/[^/]+\/resourcegroups\/[^/]+\/providers\/microsoft\.devices\/iothubs\/[^/]+$/i,
  namespace: /^\/subscriptions\/[^/]+\/resourcegroups\/[^/]+\/providers\/microsoft\.deviceregistry\/namespaces\/[^/]+$/i,
  device: /^\/subscriptions\/[^/]+\/resourcegroups\/[^/]+\/providers\/microsoft\.deviceregistry\/namespaces\/[^/]+\/registrydevices\/[^/]+$/i,
};
export function armIdentity(value: unknown, kind: keyof typeof resourcePatterns): value is string {
  return typeof value === "string" && !/[?#\\%\u0000-\u0020]/.test(value) &&
    !value.split("/").some(part => part === "." || part === "..") && resourcePatterns[kind].test(value);
}
export const sameResource = (left: string, right: string): boolean => left.toLowerCase() === right.toLowerCase();

export function readReference(properties: Record<string, unknown>): ReferenceState {
  if (!Object.prototype.hasOwnProperty.call(properties, "deviceTypeRefs")) {
    return { kind: "unknown", reason: "No model reference was returned. Assignment status has not been verified." };
  }
  const value = properties.deviceTypeRefs;
  if (!Array.isArray(value) || value.length > 1 || value.some(id => typeof id !== "string" || !id.trim() || id.length > 512)) {
    return { kind: "unknown", reason: "The service returned an unsupported model-reference value." };
  }
  return value.length ? { kind: "assigned", ids: [value[0]] } : { kind: "unassigned", ids: [] };
}
export function assignmentReference(model: HubModel): string {
  const id = assignableReference(model).id;
  if (id.length > 512) throw new Error("This model id exceeds the assignment API's 512-character limit. The model can still be used for modeling.");
  return id;
}
export function assignmentOwner(source: AssignmentSource, revision: number, model: HubModel): string {
  return canonicalJson([source, revision, assignableReference(model).id, model.document]);
}
export function referenceLabel(reference: ReferenceState, modelId: string): string {
  return reference.kind === "unknown" ? "Unable to check"
    : reference.kind === "unassigned" ? "No model assigned"
      : reference.ids[0] === modelId ? "This model" : reference.ids[0];
}
export function assignmentBlocker(target: AssignmentTarget): string | undefined {
  if (target.kind !== "device") return DPS_INFORMATIONAL_ONLY;
  if (!armIdentity(target.id, "device")) return "Only verified ADR registry-device resources can be assigned here.";
  return target.writeBlocker || (target.reference.kind === "assigned" ? "Existing model assignments are read-only."
    : target.reference.kind === "unknown" ? target.reference.reason
    : !referenceIsCurrent(target) ? CURRENT_REFERENCE_FRESHNESS
    : !target.etag || target.etag === "*" || target.etag === '"*"' ? "A current conditional-update receipt is needed before assignment." : undefined);
}
export function referenceIsCurrent(target: AssignmentTarget): boolean {
  if (target.kind !== "device") return false;
  const evidence = target.referenceEvidence;
  if (evidence === undefined) return true;
  if (!evidence || evidence.source !== "adr-point-get" || evidence.apiVersion !== "2026-11-02-preview" ||
      evidence.rawCarrier !== "missing" && evidence.rawCarrier !== "present" || !target.deviceIdentity ||
      typeof target.deviceIdentity.uuid !== "string" ||
      !/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(target.deviceIdentity.uuid) ||
      target.deviceIdentity.uuid === "00000000-0000-0000-0000-000000000000" ||
      typeof target.deviceIdentity.externalDeviceId !== "string" || !target.deviceIdentity.externalDeviceId.trim()) return false;
  const etag = evidence.currentArmEtag;
  if (!etag || etag !== target.etag) return false;
  try { specificEtag(etag); return true; } catch { return false; }
}
export interface AssignmentReceipt {
  operationId: string;
  targetId: string;
  name: string;
  kind: AssignmentTarget["kind"];
  modelId: string;
  state: "blocked" | "checking" | "assigning" | "accepted" | "confirmed" | "failed" | "unknown" | "cancelled";
  message?: string;
  writeDispatched?: boolean;
  readbackFailure?: { kind: "permission" | "authentication" | "shape" | "request" | "unsupported"; retryable: boolean };
  readbackPollCount?: number;
  readbackNotBefore?: number;
  trackingId?: string;
  preparationDeadline?: number;
  readbackDeadline?: number;
}
export const ASSIGNMENT_OPERATION_TIMEOUT_MS = 30000;
export const ASSIGNMENT_POLL_TIMEOUT_MS = 30000;
export const ASSIGNMENT_WAVE_SIZE = 5;
export const ASSIGNMENT_WAVE_INTERVAL_MS = 10000;
export const ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE =
  "The assignment wasn't confirmed within 30 seconds. It may still complete in Azure, so check the device later. No assignment will be retried.";
export const ASSIGNMENT_UNTRACKED_DISPATCH_MESSAGE =
  "The assignment request may have reached Azure, but its result couldn't be confirmed. Check the device later. No assignment will be retried.";
export function confirmedAssignmentCount(intent: AssignmentIntent | undefined, history: AssignmentReceipt[], modelId: string): number {
  const attempt = intent?.mode === "assign" ? intent.attempt : undefined;
  if (!attempt?.operationId || attempt.modelId !== modelId || !Array.isArray(attempt.targetIds) || !attempt.targetIds.length ||
      attempt.targetIds.some(id => typeof id !== "string" || !armIdentity(id, "device")) ||
      new Set(attempt.targetIds.map(id => id.toLowerCase())).size !== attempt.targetIds.length) return 0;
  return attempt.targetIds.every(id => {
    const receipts = history.filter(receipt => receipt.operationId === attempt.operationId && sameResource(receipt.targetId, id));
    return receipts.length === 1 && receipts[0].modelId === modelId && receipts[0].kind === "device" && receipts[0].state === "confirmed";
  })
    ? attempt.targetIds.length : 0;
}
export function inventoryHasModelAssignments(inventory: AssignmentInventory | undefined): boolean {
  return !!inventory && !inventory.hasMore && !inventory.issues.length && inventory.hubs.length > 0 &&
    inventory.hubs.some(hub => hub.state === "linked") &&
    inventory.hubs.every(hub => hub.state === "linked" || hub.state === "unlinked") && inventory.targets.length > 0 &&
    inventory.targets.every(target => target.kind === "device" && armIdentity(target.id, "device") &&
      target.reference.kind === "assigned" && readReference({ deviceTypeRefs: target.reference.ids }).kind === "assigned" &&
      referenceIsCurrent(target) &&
      target.hubIds.length > 0 && target.hubIds.every(id => inventory.hubs.some(hub => hub.state === "linked" && sameResource(hub.id, id))));
}
export interface AssignmentServices {
  read: (target: AssignmentTarget, signal: AbortSignal) => Promise<AssignmentTarget>;
  /** Report dispatch immediately before invoking the mutation transport, after all internal prechecks. */
  write: (target: AssignmentTarget, modelId: string, signal: AbortSignal, onDispatch: () => void, deadline?: number) => Promise<"accepted" | "completed">;
  trackingId?: (target: AssignmentTarget) => string | undefined;
  trackingDeadline?: (target: AssignmentTarget) => number;
  observe?: (target: AssignmentTarget, receipt: AssignmentReceipt, signal: AbortSignal,
    current: () => boolean, publish: (receipt: AssignmentReceipt) => void) => Promise<void>;
}
export function awaitAssignment<T>(promise: Promise<T>, signal: AbortSignal): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(new DOMException("Assignment cancelled.", "AbortError"));
    if (signal.aborted) { abort(); return; }
    signal.addEventListener("abort", abort, { once: true });
    promise.then(value => { signal.removeEventListener("abort", abort); resolve(value); },
      error => { signal.removeEventListener("abort", abort); reject(error); });
  });
}
function waitForAssignmentWave(wave: Promise<void>, deadline: number, signal: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    let timer: ReturnType<typeof setTimeout>;
    const finish = (error?: unknown) => {
      clearTimeout(timer);
      signal.removeEventListener("abort", abort);
      if (error === undefined) resolve(); else reject(error);
    };
    const abort = () => finish();
    if (signal.aborted) { resolve(); return; }
    timer = setTimeout(() => finish(), Math.max(0, deadline - Date.now()));
    signal.addEventListener("abort", abort, { once: true });
    wave.then(() => finish(), finish);
  });
}

/** Owned, explicit operation; no automatic retries, replacement, rollback or adoption claim. */
export class ModelAssignmentOperation {
  private running = false;
  async run(targets: AssignmentTarget[], model: HubModel, confirmed: boolean, services: AssignmentServices,
    current: () => boolean, signal: AbortSignal, publish: (receipt: AssignmentReceipt) => void,
    operationId = crypto.randomUUID()): Promise<AssignmentReceipt[]> {
    if (this.running) throw new Error("An assignment is already running.");
    const modelId = assignmentReference(model);
    if (!confirmed) throw new Error("Confirm the selected first assignments before assigning.");
    if (!targets.length || new Set(targets.map(target => target.id.toLowerCase())).size !== targets.length) throw new Error("Select distinct devices for assignment.");
    targets = JSON.parse(JSON.stringify(targets));
    this.running = true;
    const receipts: AssignmentReceipt[] = targets.map(target => ({
      operationId, targetId: target.id, name: target.name, kind: target.kind, modelId, state: "checking", writeDispatched: false,
    }));
    const attempt = new AbortController(), active = new Set<AbortController>(), started = new Set<number>();
    const abort = () => { attempt.abort(); for (const controller of active) controller.abort(); };
    signal.addEventListener("abort", abort, { once: true });
    const isCurrent = () => {
      if (signal.aborted || !current()) abort();
      return !attempt.signal.aborted;
    };
    const assignTarget = async (index: number): Promise<void> => {
      const target = targets[index], receipt = receipts[index];
      started.add(index);
      let dispatched = false, awaitingWriteResponse = false;
      const controller = new AbortController();
      active.add(controller);
      const preparationDeadline = Date.now() + ASSIGNMENT_OPERATION_TIMEOUT_MS;
      receipt.preparationDeadline = preparationDeadline;
      const timer = setTimeout(() => controller.abort(), ASSIGNMENT_OPERATION_TIMEOUT_MS);
      let pollingTimer: ReturnType<typeof setTimeout> | undefined;
      const update = (state: AssignmentReceipt["state"], message?: string) => {
        Object.assign(receipt, { state, message, writeDispatched: dispatched }); publish({ ...receipt });
      };
      try {
        if (!isCurrent() || controller.signal.aborted || Date.now() >= preparationDeadline) throw new DOMException("Assignment cancelled.", "AbortError");
        const blocker = assignmentBlocker(target);
        if (blocker) { update("blocked", blocker); return; }
        update("checking");
        const fresh = await awaitAssignment(services.read(target, controller.signal), controller.signal);
        if (!isCurrent() || controller.signal.aborted || Date.now() >= preparationDeadline) throw new DOMException("Assignment cancelled.", "AbortError");
        if (fresh.reference.kind === "assigned") {
          update("blocked", fresh.reference.ids[0] === modelId ? "Already matches - no change made."
            : "This device now has a model reference and won't be changed.");
          return;
        }
        if (fresh.id !== target.id || canonicalJson([fresh.hubIds, fresh.namespaceId, fresh.reference, fresh.etag, fresh.deviceIdentity]) !==
            canonicalJson([target.hubIds, target.namespaceId, target.reference, target.etag, target.deviceIdentity])) {
          throw new Error("The device changed. Refresh its assignment and review the change again.");
        }
        const freshBlocker = assignmentBlocker(fresh);
        if (freshBlocker) throw new Error(freshBlocker);
        if (!isCurrent() || controller.signal.aborted || Date.now() >= preparationDeadline) throw new DOMException("Assignment cancelled.", "AbortError");
        update("assigning");
        awaitingWriteResponse = true;
        const outcome = await awaitAssignment(services.write(fresh, modelId, controller.signal, () => {
          if (!isCurrent() || controller.signal.aborted || Date.now() >= preparationDeadline) throw new DOMException("Assignment cancelled before dispatch.", "AbortError");
          dispatched = true;
          receipt.writeDispatched = true;
          publish({ ...receipt });
        }, preparationDeadline), controller.signal);
        awaitingWriteResponse = false;
        if (!isCurrent() || controller.signal.aborted || Date.now() >= preparationDeadline) throw new DOMException("Assignment stopped after dispatch.", "AbortError");
        clearTimeout(timer);
        if (outcome === "accepted") {
          receipt.trackingId = services.trackingId?.(fresh);
          receipt.readbackDeadline = services.trackingDeadline?.(fresh) ?? Date.now() + ASSIGNMENT_POLL_TIMEOUT_MS;
          pollingTimer = setTimeout(() => controller.abort(), Math.max(0, receipt.readbackDeadline - Date.now()));
          update("accepted", "The service accepted the assignment. Completion and matching data are not yet verified.");
          if (services.observe) await awaitAssignment(services.observe(fresh, { ...receipt }, controller.signal, isCurrent, next => {
            if (controller.signal.aborted || !isCurrent()) return;
            Object.assign(receipt, next);
            publish({ ...receipt });
          }), controller.signal);
        } else {
          update("confirmed", "The assignment operation completed. Current device state and matching data have not been verified.");
        }
      } catch (error: unknown) {
        if (receipt.state === "confirmed") return;
        const cancelled = !isCurrent();
        const status = typeof error === "object" && error !== null && "status" in error ? error.status : undefined;
        const unsupported = typeof error === "object" && error !== null && "kind" in error && error.kind === "unsupported";
        const notSent = typeof error === "object" && error !== null && "writeDispatched" in error && error.writeDispatched === false;
        const rejected = typeof status === "number" && ([400, 401, 403, 404, 409, 412, 422].includes(status) ||
          status === 429 && dispatched && awaitingWriteResponse);
        const uncertain = dispatched && !notSent && !rejected && !unsupported;
        const expired = receipt.readbackDeadline !== undefined && Date.now() >= receipt.readbackDeadline;
        const expiredUncertain = expired && uncertain;
        const untrackedUncertain = uncertain && !receipt.trackingId;
        const trackingShapeError = typeof error === "object" && error !== null && "kind" in error && error.kind === "shape";
        if (expiredUncertain || untrackedUncertain) receipt.readbackFailure = {
          kind: expiredUncertain || !trackingShapeError ? "request" : "shape", retryable: false,
        };
        update(expiredUncertain ? "unknown" : unsupported || status === 412 ? "blocked" : uncertain ? "unknown" : cancelled ? "cancelled" : "failed",
          Date.now() >= preparationDeadline && receipt.readbackDeadline === undefined && (!dispatched || notSent) ? "The assignment was not sent before the observation deadline. No device change was dispatched."
            : expiredUncertain ? ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE
            : untrackedUncertain && !trackingShapeError ? ASSIGNMENT_UNTRACKED_DISPATCH_MESSAGE
            : uncertain ? "The request may have reached the service. " +
            (trackingShapeError && error instanceof Error ? error.message
              : "Operation completion is unknown; no assignment will be retried.")
            : error instanceof Error ? error.message : "The assignment could not be checked.");
      } finally {
        clearTimeout(timer);
        if (pollingTimer !== undefined) clearTimeout(pollingTimer);
        active.delete(controller);
      }
    };
    const admitted: Promise<void>[] = [];
    try {
      for (let first = 0; first < targets.length && isCurrent(); first += ASSIGNMENT_WAVE_SIZE) {
        const nextWaveAt = Date.now() + ASSIGNMENT_WAVE_INTERVAL_MS;
        const wave = targets.slice(first, first + ASSIGNMENT_WAVE_SIZE).map((_, index) => assignTarget(first + index));
        admitted.push(...wave);
        if (first + ASSIGNMENT_WAVE_SIZE < targets.length)
          await waitForAssignmentWave(Promise.all(wave).then((): void => undefined), nextWaveAt, attempt.signal);
      }
      if (!current()) abort();
      await Promise.all(admitted);
      for (const [index, receipt] of receipts.entries()) if (!started.has(index)) {
        Object.assign(receipt, { state: "cancelled",
          message: "The assignment was not sent because the batch stopped before this device started." });
        publish({ ...receipt });
      }
    } finally {
      abort();
      await Promise.allSettled(admitted);
      signal.removeEventListener("abort", abort);
      this.running = false;
    }
    return receipts;
  }
}
