import { expect, type Locator, type Page } from "@playwright/test";
import type { AuthConfig, RunConfig } from "./config.js";

export function requireWorkloadOrigin(frameUrl: string, config: Pick<RunConfig, "deployment">): void {
  const url = new URL(frameUrl);
  if (config.deployment === "published" &&
      (url.protocol !== "https:" || ["localhost", "127.0.0.1", "[::1]"].includes(url.hostname))) {
    throw new Error("[deployment] A published run must load the HTTPS-hosted workload, not a local development frontend.");
  }
}

export async function developerSwitch(page: Page, config: AuthConfig): Promise<Locator> {
  const settings = new URL("/user/user-settings/developer-settings", config.workspaceUrl);
  settings.searchParams.set("ctid", config.tenantId);
  await page.goto(settings.href, { waitUntil: "domcontentloaded" });
  await expect(page.getByRole("tab", { name: "Developer settings", exact: true })).toBeVisible({ timeout: 60_000 });
  await expect(page.getByRole("main")).toContainText(
    "Turn on developer mode to enable your local hosted workload on Fabric.",
  );
  const control = page.getByRole("main").getByRole("switch");
  await expect(control,
    "[deployment] Expected one accessible Fabric developer switch. Refusing to guess or change another setting.",
  ).toHaveCount(1);
  return control;
}

export async function requireDeploymentMode(page: Page, config: RunConfig): Promise<void> {
  const control = await developerSwitch(page, config);
  await expect(control,
    `[deployment] E2E_DEPLOYMENT=${config.deployment} requires Fabric developer mode ` +
    `${config.deployment === "development" ? "on" : "off"} in the test profile. ` +
    "A published release run must never silently use the local development gateway.",
  ).toBeChecked({ checked: config.deployment === "development" });
  await page.goto(config.workspaceUrl, { waitUntil: "domcontentloaded" });
  await expect(page.getByRole("button", { name: /^(New item|\+ New item)$/i })).toBeVisible({ timeout: 60_000 });
}
