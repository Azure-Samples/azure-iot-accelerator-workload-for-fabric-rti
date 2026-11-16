import { expect as baseExpect, type ExpectMatcherState, type Locator, type Page } from "@playwright/test";
import { setTimeout as pause } from "node:timers/promises";

type Assertions = ReturnType<typeof baseExpect<Locator>>;
const recordedPages = new WeakMap<Page, number>();

export function enableVisualChecks(page: Page, holdMs = 450): void {
  recordedPages.set(page, holdMs);
}

async function overlay(locator: Locator, caption: string, state: "checking" | "passed" | "failed"): Promise<void> {
  await locator.evaluateAll((elements, { caption, state }) => {
    const id = "fabric-e2e-verification-overlay";
    document.getElementById(id)?.remove();
    const host = document.createElement("div");
    host.id = id;
    host.setAttribute("aria-hidden", "true");
    host.setAttribute("data-check-state", state);
    host.style.cssText = "position:fixed;inset:0;pointer-events:none;z-index:2147483647;";
    // Closed shadow content remains visible in video but cannot satisfy product text assertions.
    const shadow = host.attachShadow({ mode: "closed" });
    const color = state === "checking" ? "#ffb900" : state === "passed" ? "#107c10" : "#d13438";
    const banner = document.createElement("div");
    banner.style.cssText = `position:fixed;left:8px;bottom:8px;max-width:calc(100vw - 48px);padding:10px 14px;border:3px solid ${color};border-radius:6px;background:#111;color:#fff;font:14px/1.4 Segoe UI,sans-serif;overflow-wrap:anywhere;box-shadow:0 2px 8px #0008;`;
    banner.textContent = `${state.toUpperCase()}: ${caption}${elements.length ? "" : " (no matching elements)"}`;
    shadow.append(banner);
    for (const element of elements) {
      const rect = element.getBoundingClientRect();
      if (!rect.width || !rect.height) continue;
      const box = document.createElement("div");
      box.style.cssText = `position:fixed;left:${rect.left - 3}px;top:${rect.top - 3}px;width:${rect.width + 6}px;height:${rect.height + 6}px;box-sizing:border-box;border:3px solid ${color};border-radius:3px;`;
      shadow.append(box);
    }
    document.documentElement.append(host);
  }, { caption, state });
}

function plain(locator: Locator, negated: boolean, timeout: number): Assertions {
  const assertion = baseExpect.configure({ timeout })(locator);
  return negated ? assertion.not : assertion;
}

function visualMatcher<Args extends unknown[]>(
  name: string, check: (assertion: Assertions, ...args: Args) => Promise<void>,
) {
  return async function (this: ExpectMatcherState, locator: Locator, ...args: Args) {
    const holdMs = recordedPages.get(locator.page());
    const expected = args.filter(value => ["string", "number"].includes(typeof value) || Array.isArray(value) || value instanceof RegExp)
      .slice(0, 2).map(value => String(value).slice(0, 100)).join(" ");
    const caption = `${this.isNot ? "not " : ""}${name} ${expected} - ${locator.description() ?? locator.toString()}`.slice(0, 350);
    let failure: Error | undefined;
    try {
      if (holdMs !== undefined) await overlay(locator, caption, "checking");
      try {
        await check(plain(locator, this.isNot, this.timeout), ...args);
      } catch (error) {
        if (!(error instanceof Error) || !("matcherResult" in error)) throw error;
        failure = error;
      }
      if (holdMs !== undefined) {
        await overlay(locator, caption, failure ? "failed" : "passed");
        await pause(failure ? Math.max(holdMs, 900) : holdMs);
      }
    } finally {
      if (holdMs !== undefined && !locator.page().isClosed()) {
        await locator.evaluateAll(() => document.getElementById("fabric-e2e-verification-overlay")?.remove());
      }
    }
    return {
      name, pass: this.isNot ? !!failure : !failure,
      message: () => failure?.message ?? `${caption}: assertion unexpectedly passed.`,
    };
  };
}

export const expect = baseExpect.extend({
  toBeVisible: visualMatcher("is visible", (assertion, ...args: Parameters<Assertions["toBeVisible"]>) => assertion.toBeVisible(...args)),
  toBeEnabled: visualMatcher("is enabled", (assertion, ...args: Parameters<Assertions["toBeEnabled"]>) => assertion.toBeEnabled(...args)),
  toBeDisabled: visualMatcher("is disabled", (assertion, ...args: Parameters<Assertions["toBeDisabled"]>) => assertion.toBeDisabled(...args)),
  toBeChecked: visualMatcher("is checked", (assertion, ...args: Parameters<Assertions["toBeChecked"]>) => assertion.toBeChecked(...args)),
  toHaveText: visualMatcher("text equals", (assertion, ...args: Parameters<Assertions["toHaveText"]>) => assertion.toHaveText(...args)),
  toContainText: visualMatcher("text contains", (assertion, ...args: Parameters<Assertions["toContainText"]>) => assertion.toContainText(...args)),
  toHaveCount: visualMatcher("matching element count", (assertion, ...args: Parameters<Assertions["toHaveCount"]>) => assertion.toHaveCount(...args)),
  toHaveValue: visualMatcher("value equals", (assertion, ...args: Parameters<Assertions["toHaveValue"]>) => assertion.toHaveValue(...args)),
  toHaveAttribute: visualMatcher("attribute", (assertion, name: string, value: string | RegExp, options?: { timeout?: number; ignoreCase?: boolean }) =>
    assertion.toHaveAttribute(name, value, options)),
});
