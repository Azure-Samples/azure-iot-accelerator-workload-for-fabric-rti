import { test } from "@playwright/test";
import { openBrowser, signIn } from "../../src/auth.js";
import { runConfig } from "../../src/config.js";
import { requireDeploymentMode } from "../../src/deployment.js";
import { creationDialog } from "../../src/item.js";
import { scenarioTitle } from "../../src/scenarios.js";

test(scenarioTitle("deployment"), async ({}, info) => {
  const config = runConfig();
  const context = await openBrowser(config);
  try {
    const page = context.pages()[0] ?? await context.newPage();
    await test.step("Complete interactive sign-in before any resource costs", async () => {
      await signIn(page, config);
    });
    await test.step("Require the selected development or published target without fallback", async () => {
      await requireDeploymentMode(page, config);
    });
    const dialog = await test.step("Require the expected workload card and creation dialog", () =>
      creationDialog(page, config));
    await info.attach("creation-dialog", {
      body: await dialog.ariaSnapshot(),
      contentType: "text/plain",
    });
    // Preflight opens but never submits the create dialog.
  } finally {
    await context.close();
  }
});
