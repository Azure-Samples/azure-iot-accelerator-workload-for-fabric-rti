// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { DataAgentCreationUnsuccessfulError, resolveCreatedDataAgentId } from "./DataAgentCreation.ts";

const workspace = "11111111-1111-4111-8111-111111111111", agent = "22222222-2222-4222-8222-222222222222";
const operation = "33333333-3333-4333-8333-333333333333";
const stateUrl = `https://api.fabric.microsoft.com/v1/operations/${operation}`, resultUrl = stateUrl + "/result";
const created = { id: agent, type: "DataAgent", workspaceId: workspace };
const json = (body: unknown, status = 200, headers: HeadersInit = {}) => new Response(JSON.stringify(body), { status, headers });
const accepted = (headers: HeadersInit = {}) => new Response(null, { status: 202,
  headers: { Location: stateUrl, "Retry-After": "0", ...headers } });
const resolve = (response: Response, signal = new AbortController().signal) =>
  resolveCreatedDataAgentId(response, workspace, "Fixture agent", "mock-token-only", signal);

test("direct Data Agent 200 and 201 creation validates the resource identity, never an empty/path/operation-shaped fallback", async () => {
  for (const status of [200, 201]) assert.equal(await resolve(json(created, status)), agent);
  for (const body of [{}, { id: "" }, { id: "not-a-uuid" }, { id: `/workspaces/${workspace}/dataAgents/${agent}` },
    { resourceId: agent }, { ...created, id: "00000000-0000-0000-0000-000000000000" },
    { ...created, workspaceId: operation }, { ...created, type: "Notebook" }]) {
    await assert.rejects(() => resolve(json(body, 201)), /valid created Data Agent identity/);
  }
  await assert.rejects(() => resolve(new Response("{bad", { status: 201 })), /invalid Data Agent creation response/);
  await assert.rejects(() => resolveCreatedDataAgentId(json(created, 201), "", "Fixture", "mock-only", new AbortController().signal), /current workspace/);
});

test("status-only Succeeded is followed by the same trusted operation result, not its operation ID", async () => {
  const saved = globalThis.fetch, requests: string[] = [];
  try {
    globalThis.fetch = async (input, init) => {
      const url = String(input); requests.push(url);
      assert.equal(init?.method, "GET"); assert.equal(init?.redirect, "manual");
      assert.equal(new Headers(init?.headers).get("Authorization"), "Bearer mock-token-only");
      if (requests.length === 1) return json({ status: "Running", percentComplete: 25 }, 200, { "Retry-After": "0", Location: stateUrl });
      if (requests.length === 2) return json({ status: "Succeeded", id: operation, percentComplete: 100 }, 200, { Location: resultUrl });
      assert.equal(url, resultUrl); return json(created);
    };
    assert.equal(await resolve(accepted()), agent);
    assert.deepEqual(requests, [stateUrl, stateUrl, resultUrl]);
  } finally { globalThis.fetch = saved; }
});

test("documented operation header and result URL construction work without inventing a missing operation ID", async () => {
  const saved = globalThis.fetch, requests: string[] = [];
  try {
    globalThis.fetch = async input => {
      requests.push(String(input));
      return String(input).endsWith("/result") ? json(created) : json({ status: "Succeeded" });
    };
    const response = new Response(null, { status: 202, headers: { "x-ms-operation-id": operation, "Retry-After": "0" } });
    assert.equal(await resolve(response), agent);
    assert.deepEqual(requests, [stateUrl, resultUrl]);
    await assert.rejects(() => resolve(new Response(null, { status: 202 })), /could not be verified/);
    await assert.rejects(() => resolve(accepted({ "x-ms-operation-id": agent })), /unsupported Data Agent operation URL/);
  } finally { globalThis.fetch = saved; }
});

test("only trusted original-operation Failed or Cancelled results are definitive unsuccessful creation", async () => {
  const saved = globalThis.fetch, requests: string[] = [];
  try {
    for (const status of ["Failed", "Cancelled"] as const) {
      requests.length = 0;
      globalThis.fetch = async (input, init) => {
        requests.push(String(input));
        assert.equal(init?.method, "GET");
        assert.equal(String(input), stateUrl);
        return json({ status, ...(status === "Failed" ? { error: { message: "Synthetic provisioning failure." } } : {}) },
          200, { "Retry-After": "irrelevant-for-terminal-state" });
      };
      await assert.rejects(() => resolve(accepted()), error =>
        error instanceof DataAgentCreationUnsuccessfulError && error.status === status &&
        error.message.includes(status === "Failed" ? "Synthetic provisioning failure." : "cancelled") &&
        !error.message.includes("before creating another"));
      assert.deepEqual(requests, [stateUrl], "A terminal unsuccessful operation must not fetch a result or resend creation.");
    }
  } finally { globalThis.fetch = saved; }
});

test("invalid state/result transport errors and timeout remain unconfirmed rather than definitive failure", async () => {
  const saved = globalThis.fetch;
  try {
    for (const [body, expected] of [
      [{ status: "Unknown" }, /unsupported Data Agent operation state/],
      [[], /invalid Data Agent operation state/],
    ] as const) {
      globalThis.fetch = async () => json(body);
      await assert.rejects(() => resolve(accepted()), error => error instanceof Error &&
        !(error instanceof DataAgentCreationUnsuccessfulError) && expected.test(error.message));
    }
    for (const response of [json({}, 200), json({ resourceId: agent }, 200), new Response("{bad", { status: 200 }),
      json({ errorCode: "Forbidden", message: "Synthetic result read denied." }, 403)]) {
      globalThis.fetch = async input => String(input).endsWith("/result") ? response : json({ status: "Succeeded" });
      await assert.rejects(() => resolve(accepted()), error => error instanceof Error &&
        !(error instanceof DataAgentCreationUnsuccessfulError));
    }
    for (const response of [json({ status: "Failed" }, 503), json({ status: "Cancelled" }, 403)]) {
      globalThis.fetch = async () => response;
      await assert.rejects(() => resolve(accepted()), error => error instanceof Error &&
        !(error instanceof DataAgentCreationUnsuccessfulError));
    }
    globalThis.fetch = async () => { throw new TypeError("Synthetic lost operation response."); };
    await assert.rejects(() => resolve(accepted()), error => error instanceof TypeError &&
      !(error instanceof DataAgentCreationUnsuccessfulError));
    let polls = 0;
    globalThis.fetch = async () => { polls++; return json({ status: "Running" }); };
    await assert.rejects(() => resolve(accepted()), error => error instanceof Error &&
      !(error instanceof DataAgentCreationUnsuccessfulError) && /taking longer than expected/.test(error.message));
    assert.equal(polls, 30);
  } finally { globalThis.fetch = saved; }
});

test("authenticated follow-ups reject foreign, credentialed, mismatched and unrelated operation locations before dispatch", async () => {
  const saved = globalThis.fetch; let calls = 0;
  try {
    globalThis.fetch = async () => { calls++; throw new Error("Untrusted URL must not be fetched."); };
    for (const location of [`https://example.test/v1/operations/${operation}`, `https://name:password@api.fabric.microsoft.com/v1/operations/${operation}`,
      `http://api.fabric.microsoft.com/v1/operations/${operation}`, `https://api.fabric.microsoft.com/v1/workspaces/${workspace}`,
      stateUrl + "?token=not-real", stateUrl + "#fragment"]) {
      await assert.rejects(() => resolve(accepted({ Location: location })), /unsupported Data Agent operation URL/);
    }
    assert.equal(calls, 0);
    for (const location of [`https://example.test/v1/operations/${operation}/result`, `https://api.fabric.microsoft.com/v1/operations/${agent}/result`, stateUrl]) {
      globalThis.fetch = async () => { calls++; return json({ status: "Succeeded" }, 200, { Location: location }); };
      const before = calls;
      await assert.rejects(() => resolve(accepted()), /unsupported Data Agent operation URL/);
      assert.equal(calls, before + 1);
    }
    globalThis.fetch = async () => { calls++; return json({ status: "Running" }, 200, { "Retry-After": "invalid" }); };
    await assert.rejects(() => resolve(accepted()), /invalid Retry-After/);
  } finally { globalThis.fetch = saved; }
});

test("cancellation during the operation state read never fetches a result or returns a stale identity", async () => {
  const saved = globalThis.fetch, controller = new AbortController(); let calls = 0;
  try {
    globalThis.fetch = async () => { calls++; controller.abort(); return json({ status: "Succeeded" }, 200, { Location: resultUrl }); };
    await assert.rejects(() => resolve(accepted(), controller.signal), error => error instanceof DOMException && error.name === "AbortError");
    assert.equal(calls, 1);
    await assert.rejects(() => resolve(accepted(), controller.signal));
    assert.equal(calls, 1);
  } finally { globalThis.fetch = saved; }
});

function opaqueRedirect(): Response {
  const response = new Response(null);
  Object.defineProperties(response, { type: { value: "opaqueredirect" }, status: { value: 0 } });
  return response;
}

test("recorded regional Fabric202 status/result locations are supported without a canonical-only host assumption", async () => {
  const recordedOperation = "0e82388e-e811-459d-8f71-5f9048deb17a";
  const regional = `https://wabi-us-central-b-primary-redirect.analysis.windows.net/v1/operations/${recordedOperation}`;
  const savedFetch = globalThis.fetch, savedTimeout = globalThis.setTimeout;
  const requests: string[] = [], delays: number[] = [];
  globalThis.setTimeout = ((handler: (...args: unknown[]) => void, delay?: number, ...args: unknown[]) => {
    if (delay === 20000) { delays.push(delay); return savedTimeout(handler, 0, ...args); }
    return savedTimeout(handler, delay, ...args);
  }) as typeof setTimeout;
  try {
    globalThis.fetch = async (input, init) => {
      const url = String(input); requests.push(url);
      assert.equal(init?.redirect, "manual");
      return url.endsWith("/result") ? json(created) : json({ status: "Succeeded", percentComplete: 100 }, 200, { Location: regional + "/result" });
    };
    const response = new Response(null, { status: 202, headers: { Location: regional, "Retry-After": "20" } });
    assert.equal(await resolve(response), agent);
    assert.deepEqual(requests, [regional, regional + "/result"]);
    assert.deepEqual(delays, [20000]);
  } finally { globalThis.fetch = savedFetch; globalThis.setTimeout = savedTimeout; }
});

test("canonical operation-header polling uses only the supplied same-operation regional hint when a browser hides redirect headers", async () => {
  const saved = globalThis.fetch, requests: string[] = [];
  const regional = `https://wabi-us-central-b-primary-redirect.analysis.windows.net/v1/operations/${operation}`;
  try {
    globalThis.fetch = async input => {
      const url = String(input); requests.push(url);
      if (url === stateUrl) return opaqueRedirect();
      assert.ok(url === regional || url === regional + "/result");
      return url === regional ? json({ status: "Succeeded" }, 200, { Location: regional + "/result" }) : json(created);
    };
    assert.equal(await resolve(accepted({ Location: regional, "x-ms-operation-id": operation })), agent);
    assert.deepEqual(requests, [stateUrl, regional, regional + "/result"]);
  } finally { globalThis.fetch = saved; }
});

test("visible redirects are bounded to trusted service hosts and the exact operation route, never arbitrary Azure hosts", async () => {
  const saved = globalThis.fetch, regional = `https://wabi-us-central-b-primary-redirect.analysis.windows.net/v1/operations/${operation}`;
  try {
    const requests: string[] = [];
    globalThis.fetch = async (input, init) => {
      const url = String(input); requests.push(url); assert.equal(init?.redirect, "manual");
      if (url === stateUrl) return new Response(null, { status: 307, headers: { Location: regional } });
      return url === regional ? json({ status: "Succeeded" }, 200, { Location: regional + "/result" }) : json(created);
    };
    assert.equal(await resolve(accepted()), agent);
    assert.deepEqual(requests, [stateUrl, regional, regional + "/result"]);
    for (const target of [`https://evil.windows.net/v1/operations/${operation}`,
      `https://wabi-us-central-b-primary-redirect.analysis.windows.net.evil.test/v1/operations/${operation}`,
      `https://wabi-us-central-b-primary-redirect.analysis.windows.net:444/v1/operations/${operation}`,
      `https://wabi-us-central-b-primary-redirect.analysis.windows.net/v1/operations/${agent}`,
      `https://name:password@wabi-us-central-b-primary-redirect.analysis.windows.net/v1/operations/${operation}`,
      regional + "?query=value", regional + "#fragment"]) {
      let dispatched = 0;
      globalThis.fetch = async () => { dispatched++; return new Response(null, { status: 307, headers: { Location: target } }); };
      await assert.rejects(() => resolve(accepted()), /unsupported Data Agent operation URL/);
      assert.equal(dispatched, 1, "An unsafe redirect target received a follow-up request.");
    }
    let loops = 0;
    globalThis.fetch = async () => { loops++; return new Response(null, { status: 307, headers: { Location: stateUrl } }); };
    await assert.rejects(() => resolve(accepted()), /safe limit/);
    assert.equal(loops, 1);
  } finally { globalThis.fetch = saved; }
});

test("opaque redirects without a supplied trusted regional operation hint remain explicitly unverified and are not guessed or replayed", async () => {
  const saved = globalThis.fetch; let calls = 0;
  try {
    globalThis.fetch = async () => { calls++; return opaqueRedirect(); };
    await assert.rejects(() => resolve(accepted()), /did not expose a verifiable Fabric operation redirect target/);
    assert.equal(calls, 1);
    await assert.rejects(() => resolve(opaqueRedirect()), /did not expose a verifiable Fabric creation response/);
    assert.equal(calls, 1);
  } finally { globalThis.fetch = saved; }
});
