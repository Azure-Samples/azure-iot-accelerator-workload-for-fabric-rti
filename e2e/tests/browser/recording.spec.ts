import { expect, test } from "@playwright/test";
import { readdir, stat } from "node:fs/promises";
import { join } from "node:path";
import { stripVTControlCharacters } from "node:util";
import { artifacts } from "../../src/config.js";
import { openRecordedBrowser, saveWorkflowVideo } from "../../src/recording.js";
import { scenarioTitle } from "../../src/scenarios.js";
import { enableVisualChecks, expect as visualExpect } from "../../src/visual-expect.js";

for (const outcome of ["pass", "failure"] as const) {
  test(scenarioTitle(outcome === "pass" ? "videoPass" : "videoFailure"), async ({}, info) => {
    const directory = join(artifacts, "video-check", `${outcome}-${Date.now()}`);
    const { browser, context } = await openRecordedBrowser({
      browserChannel: "msedge", workspaceUrl: "https://fabric-browser-check.example", deployment: "development",
    }, directory, join(artifacts, "video-check", `profile-${outcome}`));
    const page = context.pages()[0] ?? await context.newPage();
    const window = await context.newCDPSession(page);
    const { windowId } = await window.send("Browser.getWindowForTarget");
    await window.send("Browser.setWindowBounds", { windowId, bounds: { windowState: "normal" } });
    await window.send("Browser.setWindowBounds", { windowId, bounds: { width: 910, height: 709 } });
    await window.detach();
    await context.route("**/*", route => route.abort("blockedbyclient"));
    enableVisualChecks(page, 1000);
    const video = page.video();
    try {
      await context.route("https://fabric-browser-check.example/**", route => route.fulfill({ contentType: "text/html",
        body: '<h1>Local video check</h1><iframe id="wizard" src="http://127.0.0.1:60006/local-video-check" style="position:fixed;left:0;top:80px;width:100%;height:calc(100% - 80px);border:0"></iframe>' }));
      await context.route("http://127.0.0.1:60006/local-video-check", route => route.fulfill({ contentType: "text/html",
        body: '<button id="continue" style="position:fixed;right:24px;bottom:24px" onclick="this.textContent=\'Continued\'">Continue</button>' }));
      await page.goto("https://fabric-browser-check.example/");
      const frame = page.frames()[1];
      if (!frame) throw new Error("Missing local recorded wizard frame.");
      const scale = { top: await page.evaluate(() => devicePixelRatio), frame: await frame.evaluate(() => devicePixelRatio) };
      await info.attach("recorded-window-geometry", { body: JSON.stringify(await page.evaluate(() => ({
        innerWidth, innerHeight, outerWidth, outerHeight,
      }))), contentType: "application/json" });
      expect(page.viewportSize()).toBeNull();
      expect(await page.evaluate(() => innerWidth <= outerWidth && innerHeight <= outerHeight)).toBe(true);
      await info.attach("recorded-frame-scale", { body: JSON.stringify(scale), contentType: "application/json" });
      expect(scale.frame).toBeCloseTo(scale.top, 2);
      const next = page.locator("#wizard").contentFrame().locator("#continue").describe("Bottom-right wizard navigation");
      await visualExpect(page.getByRole("heading", { name: "Local video check" }).describe("Recorded workflow heading")).toBeVisible();
      await visualExpect(next).toBeEnabled();
      await next.click();
      await visualExpect(next).toHaveText("Continued");
      await context.route("https://fabric-browser-check.example/native-dashboard", route => route.fulfill({
        contentType: "text/html",
        body: '<h1>Generated native dashboard</h1><span data-testid="dashboard-temperature">21.5 C</span>',
      }));
      await page.goto("https://fabric-browser-check.example/native-dashboard");
      await visualExpect(page.getByRole("heading", { name: "Generated native dashboard" })
        .describe("Native dashboard after creation")).toBeVisible();
      const value = page.getByTestId("dashboard-temperature").describe("Native dashboard field value");
      await visualExpect(value).toHaveText("21.5 C");
      expect(page.video()).toBe(video);
      expect(context.pages()).toEqual([page]);
      if (outcome === "failure") {
        let failure = "";
        try {
          await visualExpect(value.describe("Intentional local failed native verification"))
            .toHaveText("Expected unavailable value", { timeout: 100 });
        } catch (error) {
          if (!(error instanceof Error)) throw error;
          failure = stripVTControlCharacters(error.message);
        }
        expect(failure).toContain("Expected unavailable value");
        expect(failure).toContain("21.5 C");
      }
    } finally {
      await context.close();
      await browser.close();
    }
    if (!video) throw new Error("The workflow context did not create a video.");
    const filename = await saveWorkflowVideo(video, directory, outcome === "pass" ? "video-pass" : "video-failure",
      outcome === "pass" ? "passed" : "failed");
    const path = join(directory, filename);
    expect((await stat(path)).size).toBeGreaterThan(1000);
    expect((await readdir(directory)).filter(file => file.endsWith(".webm"))).toEqual([filename]);
    await info.attach(`${outcome}-video`, { path, contentType: "video/webm" });
  });
}
