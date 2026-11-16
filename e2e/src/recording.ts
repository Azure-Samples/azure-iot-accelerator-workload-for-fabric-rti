import { chromium, type Page, type Request, type TestInfo, type Video } from "@playwright/test";
import { constants } from "node:fs";
import { copyFile, mkdir, unlink, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { openBrowser, signIn, grantDevelopmentAccess } from "./auth.js";
import { artifacts, profile, type RunConfig } from "./config.js";
import { requireDeploymentMode } from "./deployment.js";
import { enableVisualChecks } from "./visual-expect.js";
import { object, redactSensitiveText } from "./cloud.js";
import { closeOwnedItemTabs } from "./browser-tabs.js";
import { recordingDirectoryName, workflowRecording, workflowVideoName } from "./video-names.js";
import { readPlan } from "./plan.js";

export function assignmentRequestStage(address: string, method: string, group: string, operations: ReadonlySet<string>): string | undefined {
  const url = new URL(address), path = url.pathname.toLowerCase(), root = group.toLowerCase();
  if (url.origin !== "https://management.azure.com") return undefined;
  if (method === "GET" && operations.has(path)) return "original-operation-status";
  if (method === "POST" && path === "/providers/microsoft.resourcegraph/resources") return "resource-graph-read";
  if (!path.startsWith(`${root}/`)) return undefined;
  if (method === "PATCH" && /\/registrydevices\/[^/]+$/.test(path)) return "assignment-write";
  if (method !== "GET") return undefined;
  if (path.endsWith("/providers/microsoft.authorization/permissions")) return "permission-read";
  if (/\/registrydevices\/[^/]+\/capabilities(?:\/[^/]+)?$/.test(path)) return "capability-read";
  if (/\/registrydevices\/[^/]+$/.test(path)) return "device-point-read";
  if (/\/namespaces\/[^/]+$/.test(path)) return "namespace-read";
  if (/\/iothubs\/[^/]+$/.test(path)) return "source-hub-read";
  return undefined;
}

export function assignmentOperationState(value: unknown): string {
  const status = value && typeof value === "object" && !Array.isArray(value) && "status" in value ? value.status : undefined;
  return typeof status === "string" && ["Accepted", "Running", "InProgress", "NotStarted", "Updating", "Succeeded", "Failed", "Canceled", "Cancelled"].includes(status)
    ? status : "unrecognized";
}

interface AssignmentRequestTrace {
  stage: string;
  method: string;
  startedAt: string;
  elapsedMs?: number;
  responseStatus?: number;
  operationState?: string;
  operationHandleExposed?: boolean;
  operationHandleScoped?: boolean;
  requestFailed?: boolean;
  responseReadFailed?: boolean;
}

export async function saveWorkflowVideo(
  video: Pick<Video, "path">, directory: string, scenario: string,
  outcome: "passed" | "failed", pageIndex = 0,
): Promise<string> {
  const filename = workflowVideoName(scenario, outcome, pageIndex);
  // Closing a persistent context disconnects Video.saveAs; its finalized local file remains readable.
  const original = await video.path();
  await copyFile(original, join(directory, filename), constants.COPYFILE_EXCL);
  await unlink(original);
  return filename;
}

export function nativeResourceShape(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return { kind: value === null ? "null" : Array.isArray(value) ? "array" : typeof value };
  }
  const record = object(value, "native response diagnostic");
  return {
    kind: "object", keys: Object.keys(record),
    childKeys: Object.fromEntries(Object.entries(record).filter(([, child]) =>
      child !== null && typeof child === "object" && !Array.isArray(child))
      .map(([key, child]) => [key, Object.keys(object(child, "native diagnostic child"))])),
    hasNamespace: typeof record.fullyQualifiedNamespace === "string" && !!record.fullyQualifiedNamespace,
    hasEventHubName: typeof record.eventHubName === "string" && !!record.eventHubName,
    hasCustomEndpoint: Array.isArray(record.sources) && record.sources.some(source =>
      source !== null && typeof source === "object" && "type" in source && source.type === "CustomEndpoint"),
    hasCustomEndpointId: Array.isArray(record.sources) && record.sources.some(source =>
      source !== null && typeof source === "object" && "type" in source && source.type === "CustomEndpoint" &&
      "id" in source && typeof source.id === "string" && !!source.id),
  };
}

export async function openRecordedBrowser(
  config: Pick<RunConfig, "browserChannel" | "workspaceUrl" | "deployment">,
  directory: string, userDataDir = profile,
) {
  await mkdir(directory, { recursive: true, mode: 0o700 });
  const context = await chromium.launchPersistentContext(userDataDir, {
    args: ["--force-device-scale-factor=1", "--window-size=1456,1088"],
    ...(config.browserChannel === "msedge" ? { channel: "msedge" } : {}), headless: false,
    locale: "en-US", viewport: null,
    recordVideo: { dir: directory, size: { width: 1440, height: 1000 } },
  });
  const browser = context.browser();
  if (!browser) { await context.close(); throw new Error("[video] Recorded persistent browser is unavailable."); }
  try {
    context.setDefaultTimeout(30_000);
    context.setDefaultNavigationTimeout(60_000);
    await grantDevelopmentAccess(context, config);
    return { browser, context };
  } catch (error) {
    await browser.close();
    throw error;
  }
}

export async function authenticatedAppSessions(page: Page, workspaceUrl: string): Promise<Record<string, Record<string, string>>> {
  const sessions: Record<string, Record<string, string>> = {};
  for (const frame of page.frames()) {
    const url = new URL(frame.url());
    const appOrigin = url.origin === new URL(workspaceUrl).origin || url.protocol === "https:" &&
      (url.hostname.endsWith(".powerbi.com") || url.hostname.endsWith(".fabric.microsoft.com") ||
        url.hostname === "dataexplorer.azure.com" || url.hostname.endsWith(".dataexplorer.azure.com"));
    if (!appOrigin || frame.isDetached() ||
        frame !== page.mainFrame() && !await (await frame.frameElement()).isVisible()) continue;
    const values = await frame.locator("body").evaluate(() => Object.fromEntries(Object.entries(sessionStorage)), undefined, { timeout: 5000 });
    sessions[url.origin] = { ...sessions[url.origin], ...values };
  }
  return sessions;
}

export async function recordedWorkflow<T>(
  config: RunConfig, info: TestInfo, runId: string, scenario: string,
  action: (page: Page) => Promise<T>,
): Promise<T> {
  const recording = workflowRecording(scenario);
  const plan = await readPlan(config, runId);
  const assignmentGroup = scenario === "combined-modeling" ? plan.azure.resourceGroupId : undefined;
  const auth = await openBrowser(config);
  let sessionState: Record<string, Record<string, string>>;
  try {
    const page = auth.pages()[0] ?? await auth.newPage();
    await signIn(page, config);
    await requireDeploymentMode(page, config);
    await closeOwnedItemTabs(page, plan);
    sessionState = await authenticatedAppSessions(page, config.workspaceUrl);
  } finally {
    await auth.close();
  }
  const startedAt = new Date().toISOString();
  const directory = join(artifacts, "runs", runId, "videos", recordingDirectoryName(scenario, startedAt));
  let outcome: "passed" | "failed" = "failed";
  const { browser, context } = await openRecordedBrowser(config, directory);
  const assignmentRequests: AssignmentRequestTrace[] = [];
  const requestEntries = new WeakMap<Request, AssignmentRequestTrace>();
  const operationPaths = new Set<string>();
  const diagnosticReads: Promise<void>[] = [];
  context.on("request", request => {
    if (!assignmentGroup || assignmentRequests.length >= 200) return;
    const stage = assignmentRequestStage(request.url(), request.method(), assignmentGroup, operationPaths);
    if (!stage) return;
    const entry = { stage, method: request.method(), startedAt: new Date().toISOString() };
    assignmentRequests.push(entry);
    requestEntries.set(request, entry);
  });
  context.on("requestfailed", request => {
    const entry = requestEntries.get(request);
    if (entry) {
      entry.requestFailed = true;
      entry.elapsedMs = Date.now() - Date.parse(entry.startedAt);
    }
  });
  const nativeResponses: Record<string, unknown>[] = [];
  context.on("response", response => {
    const request = response.request(), diagnostic = requestEntries.get(request);
    if (diagnostic) {
      diagnostic.responseStatus = response.status();
      diagnostic.elapsedMs = Date.now() - Date.parse(diagnostic.startedAt);
      if (diagnostic.stage === "assignment-write" && response.status() === 202) {
        diagnosticReads.push((async () => {
          const handle = await response.headerValue("azure-asyncoperation") ?? await response.headerValue("location");
          diagnostic.operationHandleExposed = !!handle;
          if (!handle) return;
          const operation = new URL(handle, response.url());
          diagnostic.operationHandleScoped = operation.origin === "https://management.azure.com" &&
            operation.pathname.toLowerCase().startsWith(`/subscriptions/${config.subscriptionId}/`.toLowerCase());
          if (diagnostic.operationHandleScoped) {
            operationPaths.add(operation.pathname.toLowerCase());
          }
        })().catch(() => { diagnostic.responseReadFailed = true; }));
      }
      if (diagnostic.stage === "original-operation-status") {
        diagnosticReads.push(response.json().then(body => {
          diagnostic.operationState = assignmentOperationState(body);
        }).catch(() => { diagnostic.responseReadFailed = true; }));
      }
    }
    const url = new URL(response.url());
    if (nativeResponses.length >= 100 || url.origin !== "https://api.fabric.microsoft.com" ||
        !url.pathname.toLowerCase().startsWith(`/v1/workspaces/${config.workspaceId}/eventstreams/`.toLowerCase()) ||
        !/\/(?:topology|sources\/[0-9a-f-]{36}\/connection)$/i.test(url.pathname)) return;
    const entry: Record<string, unknown> = { receivedAt: new Date().toISOString(), path: url.pathname, status: response.status() };
    nativeResponses.push(entry);
    void response.json().then(body => { entry.shape = nativeResourceShape(body); }).catch(failure => {
      entry.bodyReadFailure = failure instanceof Error ? failure.name : "Unknown";
    });
  });
  // Preserve the authenticated tab's session storage in memory; never write it to an artifact.
  await context.addInitScript(sessions => {
    const values = sessions[location.origin];
    if (values && sessionStorage.length === 0) {
      for (const [key, value] of Object.entries(values)) sessionStorage.setItem(key, value);
    }
  }, sessionState);
  let authenticationBlocked = false;
  await context.route("**/*", async route => {
    const url = new URL(route.request().url());
    if (route.request().isNavigationRequest() &&
        ["login.microsoftonline.com", "login.live.com", "login.microsoft.com"].includes(url.hostname) &&
        !(url.searchParams.get("prompt") === "none" && route.request().frame().parentFrame())) {
      authenticationBlocked = true;
      console.error(`[video-auth] Blocked interactive authentication navigation to ${url.hostname}${url.pathname}.`);
      await route.abort("blockedbyclient");
    } else await route.continue();
  });
  const page = context.pages()[0] ?? await context.newPage();
  enableVisualChecks(page);
  const browserErrors: string[] = [];
  page.on("pageerror", error => {
    if (browserErrors.length < 100) browserErrors.push(redactSensitiveText(error.message).slice(0, 2000));
  });
  const pages: Page[] = [...context.pages()];
  context.on("page", popup => { pages.push(popup); enableVisualChecks(popup); });
  console.log(`[video] Workflow only; authentication is excluded. Files: ${directory}`);
  try {
    const result = await action(page);
    await closeOwnedItemTabs(page, await readPlan(config, runId));
    if (authenticationBlocked) throw new Error("[authentication] Recorded workflow required reauthentication. No credentials page was loaded; rerun the unrecorded auth checkpoint.");
    outcome = "passed";
    return result;
  } catch (error) {
    const capture = async (name: string, action: () => Promise<string>, contentType = "text/plain") => {
      try {
        await info.attach(`${scenario}-${name}`, { body: await action(), contentType });
      } catch (failure) {
        console.error(`[workflow-evidence] ${name}: ${redactSensitiveText(failure instanceof Error ? failure.message : String(failure)).split("\n")[0]}`);
      }

    };
    await capture("original-error", async () => redactSensitiveText(error instanceof Error ? error.stack ?? error.message : String(error)));
    await capture("native-resource-response-shapes", async () => JSON.stringify(nativeResponses, null, 2), "application/json");
    await capture("workflow-state", () => page.locator("body").ariaSnapshot({ timeout: 5000 }));
    const geometry: unknown[] = [];
    for (const frame of page.frames()) {
      if (frame.isDetached()) continue;
      await capture(`geometry-${geometry.length}`, async () => {
        const value = await frame.locator("body").evaluate(() => ({
          path: location.pathname, innerWidth, innerHeight, outerWidth, outerHeight, devicePixelRatio,
          visualViewport: visualViewport && { width: visualViewport.width, height: visualViewport.height, scale: visualViewport.scale },
          screen: { width: screen.width, height: screen.height, availWidth: screen.availWidth, availHeight: screen.availHeight },
        }), undefined, { timeout: 5000 });
        geometry.push(value);
        return JSON.stringify(value);
      }, "application/json");
    }
    await capture("viewport-geometry", async () => JSON.stringify({ configured: page.viewportSize(), frames: geometry }, null, 2), "application/json");
    await capture("browser-errors", async () => JSON.stringify(browserErrors, null, 2), "application/json");
    const workload = page.frames().find(frame => /\/IoTSolutionItem-editor\/[0-9a-f-]{36}/i.test(frame.url()));
    if (workload && !workload.isDetached()) {
      await capture("workload-state", () => workload.locator("body").ariaSnapshot({ timeout: 5000 }));
    }
    for (const [index, frame] of page.frames().entries()) {
      if (frame === page.mainFrame() || frame === workload || frame.isDetached()) continue;
      const url = new URL(frame.url());
      const nativeHost = url.hostname === "127.0.0.1" || url.hostname === "app.powerbi.com" ||
        url.hostname === "app.fabric.microsoft.com" || url.hostname.endsWith(".powerbi.com") ||
        url.hostname.endsWith(".fabric.microsoft.com") || url.hostname === "dataexplorer.azure.com";
      if (!nativeHost || !await (await frame.frameElement()).isVisible()) continue;
      await capture(`native-frame-${index}`, () => frame.locator("body").ariaSnapshot({ timeout: 5000 }));
    }
    console.error(`[workflow] Final frame paths: ${page.frames().map(frame => {
      const url = new URL(frame.url());
      return `${url.origin}${url.pathname}`;
    }).join(", ")}`);
    if (authenticationBlocked) throw new Error("[authentication] Recorded workflow required interactive authentication, which was blocked before rendering. Complete the unrecorded auth checkpoint.");
    throw error;
  } finally {
    await context.close();
    try {
      await Promise.all(diagnosticReads);
      if (assignmentGroup) {
        const path = join(directory, "assignment-native-requests.json");
        await writeFile(path, JSON.stringify({
          events: assignmentRequests, limitReached: assignmentRequests.length >= 200,
          requestsUnmodified: true, headersBodiesAndUrlsNotRecorded: true,
        }, null, 2));
        await info.attach("assignment-native-requests", { path, contentType: "application/json" });
      }
      const saved: string[] = [];
      for (const [index, recorded] of pages.entries()) {
        const video = recorded.video();
        if (video) {
          const filename = await saveWorkflowVideo(video, directory, scenario, outcome, index);
          saved.push(filename);
          await info.attach(`${recording.stem}-${outcome}-video-${index + 1}`, { path: join(directory, filename), contentType: "video/webm" });
        }
      }
      await writeFile(join(directory, "video-manifest.json"), JSON.stringify({
        runId, scenario, scenarioId: recording.id, title: recording.title, outcome,
        startedAt, endedAt: new Date().toISOString(), authenticationRecorded: false, videos: saved,
      }, null, 2));
    } finally {
      await browser.close();
    }
  }
}
