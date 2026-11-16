import { expect } from "@playwright/test";
import { openBrowser, signIn } from "./auth.js";
import { authConfig } from "./config.js";
import { developerSwitch } from "./deployment.js";

if (process.env.E2E_DEPLOYMENT !== "development") {
  throw new Error("[configuration] Browser developer setup requires E2E_DEPLOYMENT=development. Published runs must not enable local developer mode.");
}
const config = authConfig();
const context = await openBrowser(config);
try {
  const page = context.pages()[0] ?? await context.newPage();
  await signIn(page, config);
  const fabricDeveloperSwitch = await developerSwitch(page, config);
  await fabricDeveloperSwitch.setChecked(true);
  await expect(fabricDeveloperSwitch).toBeChecked();
  console.log("Fabric developer mode is enabled through the normal UI in the dedicated test profile.");
  console.log("The existing product frontend and development gateway must already be running and bound to this workspace.");
} finally {
  await context.close();
}
