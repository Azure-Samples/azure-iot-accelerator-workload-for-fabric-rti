// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { getFabricCreationErrorMessage, parseFabricErrorPayload } from "./FabricCreationError.ts";
import { isJsonObject } from "./ModelAuthoring.ts";

const API_ORIGIN = "https://api.fabric.microsoft.com";
const REGIONAL_HOST = /^wabi-[a-z0-9]+(?:-[a-z0-9]+)*-primary-redirect\.analysis\.windows\.net$/;
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const UNCONFIRMED = "Data Agent creation could not be verified. Verify the existing operation or agent in Fabric before creating another.";

export class DataAgentCreationUnsuccessfulError extends Error {
  readonly status: "Failed" | "Cancelled";

  constructor(status: "Failed" | "Cancelled", message: string) {
    super(message);
    this.name = "DataAgentCreationUnsuccessfulError";
    this.status = status;
  }
}

function operationUrl(value: string, result: boolean, expectedOperation?: string): URL {
  const url = new URL(value, API_ORIGIN);
  const match = /^\/v1\/operations\/([^/]+)(\/result)?$/.exec(url.pathname);
  if (url.protocol !== "https:" || url.port || url.origin !== API_ORIGIN && !REGIONAL_HOST.test(url.hostname) ||
      url.username || url.password || url.search || url.hash ||
      !match || !UUID.test(match[1]) || !!match[2] !== result ||
      expectedOperation !== undefined && match[1].toLowerCase() !== expectedOperation.toLowerCase()) {
    throw new Error("Fabric returned an unsupported Data Agent operation URL. " + UNCONFIRMED);
  }
  return url;
}

function retryDelay(response: Response, previous = 2000): number {
  const header = response.headers.get("Retry-After");
  if (header === null) return previous;
  const milliseconds = Number(header) * 1000;
  if (!/^\d+$/.test(header) || !Number.isSafeInteger(milliseconds) || milliseconds > 2147483647) {
    throw new Error("Fabric returned an invalid Retry-After value. " + UNCONFIRMED);
  }
  return milliseconds;
}

function wait(milliseconds: number, signal: AbortSignal): Promise<void> {
  signal.throwIfAborted();
  return new Promise((resolve, reject) => {
    const abort = () => { clearTimeout(timer); reject(signal.reason); };
    const timer = setTimeout(() => { signal.removeEventListener("abort", abort); resolve(); }, milliseconds);
    signal.addEventListener("abort", abort, { once: true });
  });
}

function resourceId(payload: unknown, workspaceId: string): string {
  if (!isJsonObject(payload) || typeof payload.id !== "string" || !UUID.test(payload.id) ||
      payload.id === "00000000-0000-0000-0000-000000000000" ||
      payload.type !== undefined && payload.type !== "DataAgent" ||
      payload.workspaceId !== undefined && (typeof payload.workspaceId !== "string" ||
        payload.workspaceId.toLowerCase() !== workspaceId.toLowerCase())) {
    throw new Error("Fabric did not return a valid created Data Agent identity for this workspace. " + UNCONFIRMED);
  }
  return payload.id;
}

async function payload(response: Response, displayName: string, operationRead = false): Promise<unknown> {
  if (response.status !== 200 && response.status !== 201) {
    if (operationRead) throw new Error(`Unable to verify the Data Agent operation response (HTTP ${response.status}). ${UNCONFIRMED}`);
    throw new Error(getFabricCreationErrorMessage(response.status, parseFabricErrorPayload(await response.text()), "Data Agent", displayName));
  }
  try { return await response.json(); }
  catch { throw new Error("Fabric returned an invalid Data Agent creation response. " + UNCONFIRMED); }
}

/** Resolve the response to the one creation POST; operation-state IDs are never resource IDs. */
export async function resolveCreatedDataAgentId(response: Response, workspaceId: string, displayName: string,
  token: string, signal: AbortSignal): Promise<string> {
  signal.throwIfAborted();
  if (!workspaceId.trim()) throw new Error("A current workspace is required to verify the created Data Agent.");
  if (response.type === "opaqueredirect" || response.status === 0) {
    throw new Error("The browser did not expose a verifiable Fabric creation response or redirect target. " + UNCONFIRMED);
  }
  if (response.status !== 202) {
    const body = await payload(response, displayName);
    signal.throwIfAborted();
    return resourceId(body, workspaceId);
  }

  const operationHeader = response.headers.get("x-ms-operation-id");
  if (operationHeader !== null && !UUID.test(operationHeader)) {
    throw new Error("Fabric returned an invalid Data Agent operation ID. " + UNCONFIRMED);
  }
  const location = response.headers.get("Location") || (operationHeader ? `/v1/operations/${operationHeader}` : "");
  if (!location) throw new Error(UNCONFIRMED);
  const suppliedStateUrl = operationUrl(location, false, operationHeader || undefined);
  const operationId = suppliedStateUrl.pathname.split("/").pop()!;
  let stateUrl = operationHeader ? operationUrl(`/v1/operations/${operationHeader}`, false) : suppliedStateUrl;
  const regionalState = suppliedStateUrl.origin !== API_ORIGIN ? suppliedStateUrl : undefined;
  let delay = retryDelay(response);
  const get = async (initial: URL, result: boolean): Promise<{ response: Response; url: URL }> => {
    let url = initial;
    const visited = new Set<string>();
    for (let hop = 0; hop < 5; hop++) {
      if (visited.has(url.href)) break;
      visited.add(url.href);
      const read = await fetch(url.href, { method: "GET", signal, redirect: "manual", headers: { Authorization: `Bearer ${token}` } });
      signal.throwIfAborted();
      if (read.type !== "opaqueredirect" && ![301, 302, 303, 307, 308].includes(read.status)) return { response: read, url };
      const next = read.headers.get("Location");
      if (next) url = operationUrl(new URL(next, url).href, result, operationId);
      // Browsers hide redirect headers. Only a regional URL already supplied by this operation can be used.
      else if (read.type === "opaqueredirect" && url.origin === API_ORIGIN && regionalState) {
        url = operationUrl(regionalState.href + (result ? "/result" : ""), result, operationId);
      } else {
        throw new Error("The browser did not expose a verifiable Fabric operation redirect target. " + UNCONFIRMED);
      }
    }
    throw new Error("Fabric operation redirects did not resolve within the safe limit. " + UNCONFIRMED);
  };
  for (let attempt = 0; attempt < 30; attempt++) {
    await wait(delay, signal);
    const stateRead = await get(stateUrl, false);
    const stateResponse = stateRead.response;
    stateUrl = stateRead.url;
    signal.throwIfAborted();
    if (stateResponse.status === 202) {
      delay = retryDelay(stateResponse, delay);
      const next = stateResponse.headers.get("Location");
      if (next) stateUrl = operationUrl(next, false, operationId);
      continue;
    }
    const state = await payload(stateResponse, displayName, true);
    signal.throwIfAborted();
    if (!isJsonObject(state)) throw new Error("Fabric returned an invalid Data Agent operation state. " + UNCONFIRMED);
    if (state.status === "Succeeded") {
      const resultUrl = operationUrl(stateResponse.headers.get("Location") || `${stateUrl.href}/result`, true, operationId);
      const result = (await get(resultUrl, true)).response;
      signal.throwIfAborted();
      const body = await payload(result, displayName, true);
      signal.throwIfAborted();
      return resourceId(body, workspaceId);
    }
    if (state.status === "Failed" || state.status === "Cancelled") {
      throw new DataAgentCreationUnsuccessfulError(state.status, getFabricCreationErrorMessage(stateResponse.status,
        state.error ? state : { ...state, message: state.message || `The operation ${state.status.toLowerCase()}.` }, "Data Agent", displayName));
    }
    if (state.status !== "Running" && state.status !== "NotStarted") {
      throw new Error("Fabric returned an unsupported Data Agent operation state. " + UNCONFIRMED);
    }
    delay = retryDelay(stateResponse, delay);
    const next = stateResponse.headers.get("Location");
    if (next) stateUrl = operationUrl(next, false, operationId);
  }
  throw new Error("Data Agent creation is taking longer than expected. " + UNCONFIRMED);
}
