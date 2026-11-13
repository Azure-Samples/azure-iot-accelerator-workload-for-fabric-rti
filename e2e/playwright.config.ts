import { defineConfig } from "@playwright/test";
import { join } from "node:path";
import { artifacts } from "./src/config.js";

const reportDirectory = join(artifacts, "reports", `playwright-${Date.now()}-${process.pid}`);

export default defineConfig({
  globalSetup: "./tests/global-setup.ts",
  testDir: "./tests",
  testMatch: "**/*.spec.ts",
  outputDir: join(artifacts, "test-results"),
  timeout: 660_000,
  expect: { timeout: 30_000 },
  fullyParallel: false,
  workers: 1,
  retries: 0,
  reporter: [
    ["list", { printSteps: true }],
    ["junit", { outputFile: join(reportDirectory, "results.junit.xml") }],
    ["json", { outputFile: join(reportDirectory, "results.json") }],
    ["html", { outputFolder: join(reportDirectory, "html"), open: "never" }],
    ["./src/scenario-reporter.ts"],
  ],
  use: {
    trace: "off",
    screenshot: "off",
    video: "off",
  },
});
