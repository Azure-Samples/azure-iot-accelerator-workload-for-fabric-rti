import { type FrameLocator, type Locator, type Page } from "@playwright/test";
import { expect } from "../../src/visual-expect.js";
import { requireGuid } from "../../src/config.js";
import { text } from "../../src/cloud.js";
import type { RunPlan } from "../../src/plan.js";
import { requireWorkloadOrigin } from "../../src/deployment.js";

type Entrypoint = "Ingest device data from your IoT Hub" | "Model your device data" |
  "Generate Real-Time Dashboards" | "Set up Data Agent" | "Set up Activator rules";

export const documentationLinks = {
  adr: "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal",
  dps: "https://aka.ms/iot-hub-agent-ready-data-doc",
  activator: "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators",
  dataAgent: "https://learn.microsoft.com/en-us/fabric/data-science/concept-data-agent",
} as const;

export const modelingAdrPrerequisite = "Optionally, an Azure Device Registry (ADR) namespace linked to your IoT Hub. Linking an ADR namespace to your IoT Hub allows devices to reference your model.";
export const ingestionAdrPrerequisite = "Optionally, an Azure Device Registry (ADR) namespace linked to your IoT Hub. Linking an ADR namespace to your IoT Hub lets you assign a model to your devices later, when you model your data.";

export async function expectExternalDocumentationLink(link: Locator, href: string): Promise<void> {
  await expect(link).toHaveCount(1);
  await expect(link).toBeVisible();
  await expect(link).toBeEnabled();
  await expect(link).toHaveAttribute("href", href);
  await expect(link).toHaveAttribute("target", "_blank");
  await expect(link).toHaveAttribute("rel", "noopener noreferrer");
  await expect(link).not.toHaveAttribute("aria-disabled", "true");
  await expect(link).not.toHaveAttribute("aria-describedby", /.+/);
}

export async function solutionHome(page: Page, plan: RunPlan, entrypoint: Entrypoint): Promise<FrameLocator> {
  const itemId = requireGuid(plan.observed.itemId, "owned item ID");
  const url = new URL(text(plan.observed.itemUrl, "owned item URL"));
  const expectedPath = `/groups/${plan.target.workspaceId}/Org.IoTSolutionAccelerator.IoTSolution/${itemId}`;
  if (url.origin !== new URL(plan.target.workspaceUrl).origin || url.username || url.password ||
      url.pathname.toLowerCase() !== expectedPath.toLowerCase() || url.searchParams.get("ctid") !== plan.target.tenantId) {
    throw new Error("[ownership] Recorded Fabric URL does not identify this run's exact workspace/item/tenant.");
  }
  await page.goto(url.href, { waitUntil: "domcontentloaded" });
  await expect.poll(() => page.frames().some(frame => frame.url().includes(`/IoTSolutionItem-editor/${itemId}`)),
    { timeout: 90_000 }).toBe(true);
  const frame = page.frames().find(candidate => candidate.url().includes(`/IoTSolutionItem-editor/${itemId}`));
  if (!frame) throw new Error("[deployment] The owned workload frame disappeared before it was ready.");
  requireWorkloadOrigin(frame.url(), plan.target);
  const workload = page.getByTestId("iframe-page-Org.IoTSolutionAccelerator").contentFrame();
  const ready = workload.getByRole("button", { name: entrypoint, exact: true });
  const loadFailure = workload.getByRole("heading", { name: "This item could not be loaded", exact: true });
  await expect(ready.or(loadFailure).describe("Workload item readiness response")).toBeVisible({ timeout: 90_000 });
  await expect(loadFailure.describe("No workload item loading error"),
    `[deployment] The workload item context failed to load. ${plan.target.deployment === "development"
      ? "Check the existing development gateway." : "Check workload installation and item permissions."}`,
  ).toHaveCount(0, { timeout: 1000 });
  await ready.click();
  return workload;
}
