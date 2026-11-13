import { chromium, type BrowserContext, type Page } from "@playwright/test";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { artifacts, profile, type AuthConfig } from "./config.js";

export async function grantDevelopmentAccess(
  context: Pick<BrowserContext, "grantPermissions">,
  config: Pick<AuthConfig, "deployment" | "workspaceUrl">,
): Promise<void> {
  if (config.deployment === "development") {
    await context.grantPermissions(["local-network-access"], {
      origin: new URL(config.workspaceUrl).origin,
    });
  }
}

export async function openBrowser(config: AuthConfig): Promise<BrowserContext> {
  await mkdir(profile, { recursive: true, mode: 0o700 });
  const context = await chromium.launchPersistentContext(profile, {
    args: ["--force-device-scale-factor=1", "--window-size=1456,1088"],
    ...(config.browserChannel === "msedge" ? { channel: "msedge" } : {}),
    headless: false,
    locale: "en-US",
    viewport: null,
    acceptDownloads: false,
  });
  try {
    context.setDefaultTimeout(30_000);
    context.setDefaultNavigationTimeout(60_000);
    await grantDevelopmentAccess(context, config);
    return context;
  } catch (error) {
    await context.close();
    throw error;
  }
}

export async function signIn(page: Page, config: AuthConfig): Promise<void> {
  console.log("Complete Microsoft Fabric sign-in and MFA in the dedicated test browser. This authentication step creates no resources.");
  await page.goto(config.workspaceUrl, { waitUntil: "domcontentloaded" });
  await page.getByRole("button", { name: /^(New item|\+ New item)$/i })
    .waitFor({ state: "visible", timeout: config.loginTimeoutMs });
  const current = new URL(page.url());
  if (!["app.powerbi.com", "app.fabric.microsoft.com"].includes(current.hostname) ||
      !current.pathname.toLowerCase().startsWith(`/groups/${config.workspaceId}`)) {
    throw new Error("[authentication] Sign-in did not reach the requested Fabric workspace.");
  }
  // This is a UI-readiness receipt, not proof of ARM/Kusto permissions or tenant identity.
  await writeFile(join(artifacts, "auth-receipt.json"), JSON.stringify({
    tenantHint: config.tenantId,
    workspaceId: config.workspaceId,
    workspaceUiReadyAt: new Date().toISOString(),
    browser: config.browserChannel,
  }, null, 2), { mode: 0o600 });
  console.log("Fabric workspace UI is ready. Authentication remains in the local, ignored test-browser profile.");
}
