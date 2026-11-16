import { type Locator, type Page } from "@playwright/test";
import { expect } from "./visual-expect.js";
import type { RunConfig } from "./config.js";

export async function creationDialog(page: Page, config: RunConfig): Promise<Locator> {
  await page.getByRole("button", { name: /^(New item|\+ New item)$/i }).click();
  const catalog = page.getByRole("region", { name: "New item", exact: true });
  await catalog.getByPlaceholder("Filter by keyword").fill(config.itemLabel.replace(/\s+\(preview\)$/i, ""));
  const card = catalog.getByRole("button", { name: config.itemLabel, exact: true });
  await expect(card,
    `[deployment] "${config.itemLabel}" must be available in this workspace's New item catalog. ` +
    "Check workload installation, workspace binding and browser-specific developer mode.",
  ).toBeVisible();
  await card.click();
  const dialog = page.getByRole("dialog", { name: "New Azure IoT Solution Accelerator", exact: true });
  await expect(dialog).toBeVisible();
  return dialog;
}
