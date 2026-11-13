// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Isolated browser component tests. Never uses the signed-in Fabric browser or production endpoints.
const path = require("path");
const fs = require("fs");
const os = require("os");
const http = require("http");
const { spawn } = require("child_process");
const webpack = require("webpack");
const HtmlPlugin = require("html-webpack-plugin");
const BROWSER_BOOT_TIMEOUT_MS = 15000;
const CASE_TIMEOUT_MS = 30000;
const RUN_OVERHEAD_MS = 10000;
const SELECTED_CASE_BUDGET_MS = 5000;
if (process.argv.some(argument => argument.startsWith("--exact-name") && !argument.startsWith("--exact-name="))) {
  throw new Error("Use --exact-name=<full registered name> for each exact test.");
}
const exactNames = process.argv.filter(argument => argument.startsWith("--exact-name="))
  .map(argument => argument.slice("--exact-name=".length));
const expectedNames = process.env.MODELING_TEST_EXPECTED_NAMES ? JSON.parse(process.env.MODELING_TEST_EXPECTED_NAMES) : undefined;
if (expectedNames !== undefined && (!Array.isArray(expectedNames) || !expectedNames.length ||
    expectedNames.some(name => typeof name !== "string" || !name) || new Set(expectedNames).size !== expectedNames.length)) {
  throw new Error("Expected test names must be a nonempty distinct list.");
}
if (exactNames.some(name => !name) || new Set(exactNames).size !== exactNames.length) {
  throw new Error("Exact test names must be nonempty and distinct.");
}
if (exactNames.length && process.argv.some(argument => argument.startsWith("--filter="))) {
  throw new Error("Use either exact test names or the incumbent prefix filter, not both.");
}
async function capturePreviews(send, sessionId, url, directory, modelFile) {
  fs.mkdirSync(directory, { recursive: true });
  const evaluate = async expression => {
    const value = await send("Runtime.evaluate", { expression, returnByValue: true }, sessionId);
    if (value.exceptionDetails) throw new Error(JSON.stringify(value.exceptionDetails));
    return value.result.value;
  };
  const waitFor = async expression => {
    const deadline = Date.now() + 15000;
    while (Date.now() < deadline) {
      const failure = await evaluate("document.documentElement.dataset.previewError || null");
      if (failure) throw new Error(`Mock preview failed: ${failure}`);
      if (await evaluate(expression)) return;
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    throw new Error(`Mock preview did not become ready: ${expression}`);
  };
  const metrics = [];
  const recovery = process.argv.includes("--capture-recovery");
  const parity = process.argv.includes("--capture-parity");
  const alignment = process.argv.includes("--capture-alignment");
  const sourceLabels = process.argv.includes("--capture-source-labels");
  const sourceAnalysis = process.argv.includes("--capture-source-analysis");
  const sourceReselection = process.argv.includes("--capture-source-reselection");
  const sourceValidation = process.argv.includes("--capture-source-validation");
  const modelFormat = process.argv.includes("--capture-model-format");
  const modelPromptCopied = process.argv.includes("--capture-model-prompt-copied");
  const modelPrompt = modelFormat || modelPromptCopied || process.argv.includes("--capture-model-prompt");
  const inlineModel = process.argv.includes("--capture-inline-model");
  const acceptedModel = process.argv.includes("--capture-accepted-model");
  const fieldPresentation = process.argv.includes("--capture-field-presentation");
  const complexField = process.argv.includes("--capture-complex-field");
  const calibrationLayout = process.argv.includes("--capture-calibration-layout");
  const reviewBatch = process.argv.includes("--capture-review-batch");
  const reviewCoverage = process.argv.includes("--capture-review-coverage");
  const reviewDiagnostics = process.argv.includes("--capture-review-diagnostics");
  const reviewCreation = process.argv.includes("--capture-review-creation");
  const reviewReady = process.argv.includes("--capture-review-ready");
  const reviewHeadings = process.argv.includes("--capture-review-headings");
  const reviewAssignmentSpacing = process.argv.includes("--capture-review-assignment-spacing");
  const reviewCoverageFollowup = process.argv.includes("--capture-review-coverage-followup");
  const reviewEmptyPrompt = process.argv.includes("--capture-review-empty-prompt");
  const emptyPrompt = reviewEmptyPrompt || process.argv.includes("--capture-empty-prompt");
  const uploadSpacing = process.argv.includes("--capture-upload-spacing");
  const assignmentIntent = process.argv.includes("--capture-assignment-intent");
  const modelCoverageResult = process.argv.includes("--capture-model-coverage-result");
  const modelCoverageCompact = process.argv.includes("--capture-model-coverage-compact");
  const modelCoverage = modelCoverageCompact || modelCoverageResult || process.argv.includes("--capture-model-coverage");
  const assignmentPolling = process.argv.includes("--capture-assignment-polling");
  const assignmentWarning = process.argv.includes("--capture-assignment-warning");
  const deviceToolbar = process.argv.includes("--capture-device-toolbar");
  const deviceTable = deviceToolbar || process.argv.includes("--capture-device-table");
  const hubWarning = process.argv.includes("--capture-hub-warning");
  const hubSummary = hubWarning || process.argv.includes("--capture-hub-summary");
  const readyLearnLinks = [
    ["Learn about Activator", "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators"],
    ["Learn about Fabric data agents", "https://learn.microsoft.com/en-us/fabric/data-science/concept-data-agent"],
  ];
  const adrLearnLink = ["Learn how to link an IoT hub to ADR", "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal"];
  const dpsLearnLink = ["Learn how to update DPS enrollments", "https://aka.ms/iot-hub-agent-ready-data-doc"];
  for (const width of reviewBatch || reviewCoverage || reviewDiagnostics || reviewCreation || reviewReady || reviewHeadings || reviewAssignmentSpacing || reviewCoverageFollowup || emptyPrompt || modelFormat || assignmentWarning ? [1440, 520] : modelCoverageResult ? [1440] : calibrationLayout || uploadSpacing || hubSummary || deviceTable || assignmentPolling || modelCoverage ? [1440, 520] : alignment || sourceLabels || sourceAnalysis || sourceReselection || sourceValidation || modelPrompt || inlineModel || acceptedModel || fieldPresentation || complexField || assignmentIntent ? [1440, 700] : recovery || parity ? [1440, 732] : [1440, 800, 732]) {
    await send("Emulation.setDeviceMetricsOverride", { width, height: 900, deviceScaleFactor: 1, mobile: false }, sessionId);
    for (const preview of assignmentWarning ? ["assignment-warning-light", "assignment-warning-dark"]
      : reviewAssignmentSpacing ? ["assignment-spacing-active", "assignment-spacing-confirmed",
      "assignment-spacing-partial", "assignment-spacing-unknown", ...(emptyPrompt ? ["model-prompt-empty"] : []),
      ...(reviewCoverageFollowup ? ["model-coverage-followup-reference-results", "model-coverage-followup-data-results", "model-coverage-followup-skipped-results", "model-coverage-followup-dtdl-results"] : [])]
      : reviewCoverageFollowup ? ["model-coverage-followup-reference-results", "model-coverage-followup-data-results", "model-coverage-followup-skipped-results", "model-coverage-followup-dtdl-results"]
      : emptyPrompt ? ["model-prompt-empty"]
      : reviewHeadings ? ["model-headings-wot", "model-headings-dtdl"]
      : reviewReady ? ["model-ready"] : reviewCreation ? ["model-creation-form", "model-creation-complete"]
      : reviewDiagnostics ? ["model-diagnostic-json", "model-diagnostic-missing-id",
      "model-diagnostic-warnings", "model-diagnostic-dtdl-role"]
      : reviewCoverage ? ["model-coverage-available", "model-coverage-conditions",
      "model-coverage-positive-results", "model-coverage-mixed-results", "model-coverage-dense-results"]
      : reviewBatch ? ["calibration-layout", "hub-summary-mixed", "hub-summary-unlinked",
      "assignment-confirmation", "assignment-poll-progress", "assignment-poll-success", "device-table-already", "assignment-intent-dtdl"]
      : modelFormat ? ["model-prompt-collapsed", "model-prompt-expanded"]
      : modelCoverageCompact ? ["model-coverage-mixed-results", "model-coverage-dense-results"]
      : modelCoverageResult ? ["model-coverage-results"] : modelCoverage ? ["model-coverage-available", "model-coverage-results"]
      : assignmentPolling ? ["assignment-poll-progress", "assignment-poll-success"]
      : deviceToolbar ? ["device-table-mixed"]
      : deviceTable ? ["device-table-active", "device-table-mixed", "device-table-confirmed", "device-table-already"]
      : hubWarning ? ["hub-summary-mixed", "hub-summary-unlinked"]
      : hubSummary ? ["hub-summary-initial", "hub-summary-loading", "hub-summary-mixed", "hub-summary-unlinked"]
      : assignmentIntent ? ["assignment-intent-initial", "assignment-intent-active", "assignment-intent-dtdl"]
      : uploadSpacing ? ["upload-spacing-default", "upload-spacing-accepted"] : calibrationLayout ? ["calibration-layout"] : complexField ? ["field-complex"] : fieldPresentation ? ["field-presentation-overview", "field-presentation-fields"]
      : acceptedModel ? ["accepted-model"] : inlineModel ? ["inline-model-wot", "inline-model-ambiguous", "inline-model-dtdl", "inline-model-failed"]
      : modelPrompt ? (modelPromptCopied ? ["model-prompt-copied"] : ["model-prompt-collapsed", "model-prompt-expanded", "model-prompt-partial",
      "model-prompt-empty", "model-prompt-error", "model-prompt-copy-failed"])
      : sourceValidation ? ["source-validation-initial", "source-validation-pending", "source-validation-success",
      "source-validation-failed", "source-validation-selection-reset", "source-validation-model-intro"]
      : sourceReselection ? ["source-reselection-eventhouses", "source-reselection-databases",
      "source-reselection-tables", "source-reselection-recovered"]
      : sourceAnalysis ? ["source-analysis-prerequisites", "source-analysis-selected", "source-analysis-loading",
      "source-analysis-analyzed", "source-analysis-menu", "source-analysis-model", "source-analysis-wot",
      "source-analysis-dtdl", "source-analysis-fields"]
      : sourceLabels ? ["source-labels", "source-labels-telemetry", "source-labels-properties", "source-labels-model"]
      : alignment ? ["alignment-fields", "alignment-review", "alignment-ready"]
      : parity ? ["loaded", "parity-fields", "parity-scope", "parity-review"]
      : recovery ? ["error-dtdl", "error-wot", "error-size", "error-read", "loaded"] : ["raw", "model", "loaded", "schema", "dtdl"]) {
      if (preview === "loaded" && !modelFile) continue;
      await send("Page.navigate", { url: `${url}/?preview=${preview}` }, sessionId);
      await waitFor(`document.documentElement.dataset.previewReady === "true" && document.documentElement.dataset.previewKind === ${JSON.stringify(preview)}`);
      let modelFormatKeyboard = null;
      if (modelFormat && preview === "model-prompt-collapsed") {
        const before = await evaluate("window.__modelFormatPreviewCounts()");
        const key = async (value, code) => {
          await send("Input.dispatchKeyEvent", { type: "keyDown", key: value, code: value,
            windowsVirtualKeyCode: code, nativeVirtualKeyCode: code }, sessionId);
          await send("Input.dispatchKeyEvent", { type: "keyUp", key: value, code: value,
            windowsVirtualKeyCode: code, nativeVirtualKeyCode: code }, sessionId);
        };
        await evaluate('document.querySelector(".iot-solution-model-format-choices input[value=wot]").focus()');
        await key(width > 620 ? "ArrowRight" : "ArrowDown", width > 620 ? 39 : 40);
        await waitFor('document.querySelector(".iot-solution-model-format-choices input[value=dtdl]").checked && !document.getElementById("model-preparation-heading")');
        await key(width > 620 ? "ArrowLeft" : "ArrowUp", width > 620 ? 37 : 38);
        await waitFor('document.querySelector(".iot-solution-model-format-choices input[value=wot]").checked && !!document.getElementById("model-preparation-heading")');
        await key("Tab", 9);
        await waitFor('document.activeElement?.textContent === "Get Copilot prompt"');
        const ax = await send("Accessibility.getFullAXTree", {}, sessionId);
        for (const [label, description] of [
          ["WoT Thing Model (recommended)", "Lets you assign the model to devices."],
          ["Digital Twin Definition Language (DTDL)", "Use an existing DTDL model or IoT Central template."],
        ]) {
          if (!ax.nodes.some(node => node.role?.value === "radio" && node.name?.value === label && node.description?.value === description)) {
            throw new Error(`The format radio lost its exact accessible name/description: ${label}`);
          }
        }
        if (JSON.stringify(before) !== JSON.stringify(await evaluate("window.__modelFormatPreviewCounts()"))) {
          throw new Error("Native format keyboard interaction triggered data, assignment or authorization calls.");
        }
        modelFormatKeyboard = { trustedArrowKeys: true, trustedTabToHelper: true,
          exactRadioNamesAndDescriptions: true, extraReadsOrAuth: 0 };
        await send("Page.navigate", { url: `${url}/?preview=${preview}` }, sessionId);
        await waitFor(`document.documentElement.dataset.previewReady === "true" && document.documentElement.dataset.previewKind === ${JSON.stringify(preview)}`);
      }
      let coverageKeyboard = null;
      if (modelCoverage && preview === "model-coverage-available" && width === 1440) {
        const before = await evaluate("window.__coveragePreviewCounts()");
        const key = async (value, code) => {
          await send("Input.dispatchKeyEvent", { type: "keyDown", key: value, code: value, windowsVirtualKeyCode: code,
            nativeVirtualKeyCode: code, ...(value === "Enter" ? { text: "\r", unmodifiedText: "\r" } : {}) }, sessionId);
          await send("Input.dispatchKeyEvent", { type: "keyUp", key: value, code: value, windowsVirtualKeyCode: code,
            nativeVirtualKeyCode: code }, sessionId);
        };
        await evaluate('document.querySelector(".iot-solution-coverage-choices input").focus()');
        await key("ArrowDown", 40);
        await waitFor('document.querySelector(".iot-solution-coverage-choices input[value=and]").checked');
        await key("ArrowUp", 38);
        await waitFor('document.querySelector(".iot-solution-coverage-choices input[value=data]").checked');
        const point = await evaluate(`(() => {
          const row = document.querySelectorAll(".iot-solution-coverage-choice")[2];
          row.scrollIntoView({block:"center"});
          const box = row.getBoundingClientRect();
          return {x:box.right-20,y:box.top+box.height/2};
        })()`);
        await send("Input.dispatchMouseEvent", { type: "mousePressed", button: "left", clickCount: 1, ...point }, sessionId);
        await send("Input.dispatchMouseEvent", { type: "mouseReleased", button: "left", clickCount: 1, ...point }, sessionId);
        await waitFor('document.querySelector(".iot-solution-coverage-choices input[value=or]").checked');
        await evaluate('document.querySelector(\'[aria-labelledby="additional-conditions-heading"] input[value=none]\').focus()');
        await key("ArrowDown", 40);
        await waitFor('document.querySelector(\'[aria-labelledby="additional-conditions-heading"] input[value=specified]\').checked');
        await key("Tab", 9);
        await waitFor('document.activeElement === document.getElementById("hub-telemetry-scope")');
        const ax = await send("Accessibility.getFullAXTree", {}, sessionId);
        for (const label of ["Telemetry", "Reported properties"]) {
          if (!ax.nodes.some(node => node.role?.value === "textbox" && node.name?.value.includes(label))) {
            throw new Error(`The raw predicate editor lost its accessible label: ${label}`);
          }
        }
        const after = await evaluate("window.__coveragePreviewCounts()");
        if (JSON.stringify(before) !== JSON.stringify(after)) throw new Error("Keyboard/row/editor interaction triggered reads or authorization.");
        coverageKeyboard = { trustedArrowKeys: true, trustedWholeRowClick: true, trustedTabToEditor: true,
          accessibleEditorLabels: true, implicitReads: 0 };
        await evaluate('delete document.documentElement.dataset.previewReady');
        await send("Page.navigate", { url: `${url}/?preview=${preview}` }, sessionId);
        await waitFor('document.documentElement.dataset.previewReady === "true"');
        await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: 0, y: 0 }, sessionId);
      }
      if (sourceValidation && width === 700) {
        await evaluate(`document.querySelector(".iot-solution-loading, .iot-solution-success, [role='alert']")?.scrollIntoView({block:"nearest"})`);
      }
      if (modelPrompt && width === 700) {
        await evaluate(`document.querySelector("[aria-labelledby='model-preparation-heading']")?.scrollIntoView({block:"start"})`);
      }
      if (modelPromptCopied) {
        await evaluate('document.querySelector("#model-preparation-content .iot-solution-model-actions button[aria-live]").scrollIntoView({block:"center"})');
      }
      if (["model-coverage-results", "model-coverage-mixed-results", "model-coverage-dense-results"].includes(preview)) {
        await evaluate('document.getElementById("coverage-results-heading").scrollIntoView({block:"start"})');
      }
      if (inlineModel) {
        await evaluate(`(document.querySelector("[aria-label='Accepted model']") || document.querySelector("[role='alert']"))?.scrollIntoView({block:"start"})`);
      }
      if (acceptedModel || fieldPresentation || complexField || uploadSpacing) {
        await evaluate(`document.querySelector(${JSON.stringify(uploadSpacing ? ".iot-solution-model-upload-area"
          : complexField ? '[data-field-id="/properties/config"]'
          : preview === "field-presentation-fields" ? "[data-field-id]" : ".iot-solution-model-upload-area--loaded")}).scrollIntoView({block:"start"})`);
      }
      if (calibrationLayout || reviewBatch && preview === "calibration-layout") {
        await evaluate(`document.querySelector("[data-field-id]").scrollIntoView({block:"start"})`);
      }
      if (reviewBatch && preview.startsWith("hub-summary-")) {
        await evaluate('document.querySelector(".iot-solution-assignment > h2").scrollIntoView({block:"start"})');
      }
      if (reviewCoverage) await evaluate(preview.endsWith("-results")
        ? 'document.getElementById("coverage-results-heading").scrollIntoView({block:"start"})'
        : 'document.querySelector(".iot-solution-model-coverage > h2").scrollIntoView({block:"start"})');
      if (reviewReady) {
        await evaluate('document.querySelector(".iot-solution-model-ready h2").scrollIntoView({block:"start"})');
      }
      const expectedLearnLinks = reviewReady ? [...readyLearnLinks]
        : hubSummary && ["hub-summary-mixed", "hub-summary-unlinked"].includes(preview) ? [adrLearnLink] : [];
      if (await evaluate('!!document.querySelector(\'[aria-label="DPS enrollment guidance"]\')')) expectedLearnLinks.push(dpsLearnLink);
      const learnLinkAccessibility = [];
      const learnLinkInteractions = [];
      if (expectedLearnLinks.length) {
        const ax = await send("Accessibility.getFullAXTree", {}, sessionId);
        for (const [name, href] of expectedLearnLinks) {
          const nodes = ax.nodes.filter(node => node.role?.value === "link" && node.name?.value === name);
          const node = nodes[0], property = name => node?.properties?.find(value => value.name === name)?.value?.value;
          if (nodes.length !== 1 || property("url") !== href || property("disabled") === true ||
              property("focusable") !== true || node.description?.value) {
            throw new Error(`Native documentation link lost its exact accessible name, URL or enabled state: ${name}`);
          }
          learnLinkAccessibility.push({ name: node.name.value, href: property("url"),
            focusable: property("focusable"), disabled: property("disabled") === true, description: node.description?.value || null });
          const beforeNavigation = await evaluate("window.__documentationNavigationState()");
          const point = await evaluate(`(() => {
            const link = [...document.querySelectorAll("a")].find(node => node.textContent.trim() === ${JSON.stringify(name)});
            link.scrollIntoView({block:"center",behavior:"instant"});
            const events = [];
            const observe = event => {
              if (!link.contains(event.target)) return;
              events.push({trusted:event.isTrusted,defaultPreventedByProduct:event.defaultPrevented});
              if (!event.defaultPrevented) event.preventDefault();
            };
            document.addEventListener("click", observe);
            window.__documentationNativeProbe = { events, cleanup: () => document.removeEventListener("click", observe) };
            link.focus();
            const rect = link.getClientRects()[0];
            return {x:rect.left+rect.width/2,y:rect.top+rect.height/2};
          })()`);
          try {
            await send("Input.dispatchMouseEvent", { type: "mousePressed", button: "left", clickCount: 1, ...point }, sessionId);
            await send("Input.dispatchMouseEvent", { type: "mouseReleased", button: "left", clickCount: 1, ...point }, sessionId);
            const afterMouse = await evaluate("window.__documentationNavigationState()");
            if (JSON.stringify(afterMouse.browserTabs.slice(beforeNavigation.browserTabs.length)) !==
                JSON.stringify([{ url: href }]) || afterMouse.sideEffects !== beforeNavigation.sideEffects) {
              throw new Error(`Native documentation mouse click must call Fabric exactly once: ${name}`);
            }
            await send("Input.dispatchKeyEvent", { type: "keyDown", key: "Enter", code: "Enter",
              windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13, text: "\r", unmodifiedText: "\r" }, sessionId);
            await send("Input.dispatchKeyEvent", { type: "keyUp", key: "Enter", code: "Enter",
              windowsVirtualKeyCode: 13, nativeVirtualKeyCode: 13 }, sessionId);
            const afterNavigation = await evaluate("window.__documentationNavigationState()");
            const events = await evaluate("window.__documentationNativeProbe.events");
            const actualCalls = afterNavigation.browserTabs.slice(beforeNavigation.browserTabs.length);
            if (events.length !== 2 || events.some(event => !event.trusted || !event.defaultPreventedByProduct) ||
                JSON.stringify(actualCalls) !== JSON.stringify([{ url: href }, { url: href }]) ||
                afterNavigation.sideEffects !== beforeNavigation.sideEffects) {
              throw new Error(`Native mouse/Enter documentation activation did not use only Fabric navigation: ${JSON.stringify({name,href,events,actualCalls})}`);
            }
            learnLinkInteractions.push({ name, href, nativeMouseAndEnter: true, events,
              actualOpenBrowserTabCalls: actualCalls, unrelatedSideEffects: 0 });
          } finally {
            await evaluate("window.__documentationNativeProbe.cleanup(); delete window.__documentationNativeProbe; document.activeElement?.blur()");
          }
        }
      }
      if (reviewHeadings) {
        await evaluate('document.getElementById("model-format-heading").scrollIntoView({block:"start"})');
        const ax = await send("Accessibility.getFullAXTree", {}, sessionId);
        for (const name of ["1. Choose a model format", "2. Upload your model", "3. Select fields"]) {
          if (!ax.nodes.some(node => node.role?.value === "heading" && node.name?.value === name &&
              node.properties?.some(property => property.name === "level" && property.value?.value === 3)))
            throw new Error(`Model section lost its peer heading hierarchy: ${name}`);
        }
        if (ax.nodes.some(node => node.role?.value === "heading" && node.name?.value === "What's supported"))
          throw new Error("What's supported still appears as a peer heading.");
      }
      if (reviewAssignmentSpacing && preview.startsWith("assignment-spacing-"))
        await evaluate('document.querySelector(\'[aria-label="Device assignment progress"]\').scrollIntoView({block:"end",behavior:"instant"})');
      if (emptyPrompt && preview === "model-prompt-empty") {
        await evaluate('document.getElementById("model-preparation-content").scrollIntoView({block:"center",behavior:"instant"})');
        await waitFor(`(() => {
          const region = document.getElementById("model-preparation-content");
          const panel = document.querySelector(".wizard-control__content-panel").getBoundingClientRect();
          const footer = document.querySelector(".wizard-control__footer").getBoundingClientRect();
          return region.getBoundingClientRect().top >= panel.top &&
            region.querySelector(".iot-solution-model-prompt-disclaimer").getBoundingClientRect().bottom <= footer.top;
        })()`);
      }
      if (reviewCoverageFollowup && preview.startsWith("model-coverage-followup-"))
        await evaluate('(document.querySelector(\'[aria-labelledby="model-reference-filter-heading"] .iot-solution-assignment-notice\') || document.querySelector(".iot-solution-coverage-assignment-hint") || document.querySelector("[data-coverage-stream]")).scrollIntoView({block:"start",behavior:"instant"})');
      if (reviewAssignmentSpacing || reviewCoverageFollowup || emptyPrompt)
        await send("Runtime.evaluate", { expression: "new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve(true))))",
          awaitPromise: true, returnByValue: true }, sessionId);
      if (assignmentWarning) {
        await evaluate('document.querySelector(\'[aria-label="Device assignment progress"]\').scrollIntoView({block:"end",behavior:"instant"})');
        await send("Runtime.evaluate", { expression: "new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve(true))))",
          awaitPromise: true, returnByValue: true }, sessionId);
        await waitFor(`(() => {
          const progress = document.querySelector('[aria-label="Device assignment progress"]'), panel = document.querySelector(".wizard-control__content-panel");
          const footer = document.querySelector(".wizard-control__footer").getBoundingClientRect();
          return progress.querySelector("h3").getBoundingClientRect().top >= panel.getBoundingClientRect().top &&
            progress.getBoundingClientRect().bottom <= footer.top;
        })()`);
      }
      if (reviewDiagnostics) {
        if (preview === "model-diagnostic-dtdl-role") {
          await evaluate(`(() => {
            const button = [...document.querySelectorAll('button[aria-label="About DTDL roles"]')]
              .find(node => node.getBoundingClientRect().width > 0);
            button.scrollIntoView({block:"center"}); button.focus();
          })()`);
          for (const modifiers of [8, 0]) {
            await send("Input.dispatchKeyEvent", { type: "keyDown", key: "Tab", code: "Tab", modifiers, windowsVirtualKeyCode: 9, nativeVirtualKeyCode: 9 }, sessionId);
            await send("Input.dispatchKeyEvent", { type: "keyUp", key: "Tab", code: "Tab", modifiers, windowsVirtualKeyCode: 9 }, sessionId);
          }
          if (reviewCreation) await evaluate(preview === "model-creation-complete"
            ? 'document.querySelector(\'[aria-label="Setup progress"]\').scrollIntoView({block:"start"})'
            : 'document.querySelector(".iot-solution-step > h2").scrollIntoView({block:"start"})');
          await waitFor(`document.activeElement?.getAttribute("aria-label") === "About DTDL roles" &&
            [...document.querySelectorAll('[role=tooltip]')].some(node => node.textContent === "Roles come from your DTDL model" &&
              node.getBoundingClientRect().width > 10 && node.getBoundingClientRect().height > 10 &&
              getComputedStyle(node).visibility === "visible" && Number(getComputedStyle(node).opacity) > 0.95)`);
        } else await evaluate(`document.querySelector(${JSON.stringify(preview === "model-diagnostic-warnings"
          ? "#model-upload-warnings" : "#model-upload-error")}).scrollIntoView({block:"start"})`);
      }
      if (reviewBatch && preview === "assignment-confirmation") {
        await waitFor('document.querySelector("[role=dialog]")?.contains(document.activeElement)');
        for (let index = 0; index < 3; index++) {
          await send("Input.dispatchKeyEvent", { type: "keyDown", key: "Tab", code: "Tab", windowsVirtualKeyCode: 9, nativeVirtualKeyCode: 9 }, sessionId);
          await send("Input.dispatchKeyEvent", { type: "keyUp", key: "Tab", code: "Tab", windowsVirtualKeyCode: 9 }, sessionId);
          if (!await evaluate('document.querySelector("[role=dialog]").contains(document.activeElement) && window.__confirmationPreviewWrites() === 0')) {
            throw new Error("Confirmation keyboard focus escaped or review dispatched an assignment.");
          }
        }
      }
      let deviceKeyboard = null;
      if (preview === "hub-summary-devices" || preview === "device-table-mixed") {
        await evaluate('document.getElementById("assignment-devices-heading").scrollIntoView({block:"start"}); document.querySelector(\'input[aria-label="Select Fixture device"]\').focus()');
        const toggle = async () => {
          await send("Input.dispatchKeyEvent", { type: "keyDown", key: " ", code: "Space", windowsVirtualKeyCode: 32,
            nativeVirtualKeyCode: 32, text: " ", unmodifiedText: " " }, sessionId);
          await send("Input.dispatchKeyEvent", { type: "keyUp", key: " ", code: "Space", windowsVirtualKeyCode: 32 }, sessionId);
        };
        await toggle();
        await waitFor('document.querySelector(\'input[aria-label="Select Fixture device"]\').checked === true');
        await toggle();
        await waitFor('document.querySelector(\'input[aria-label="Select Fixture device"]\').checked === false');
        deviceKeyboard = { trustedSpaceSelectedAndDeselected: true };
        await evaluate('document.getElementById("assignment-devices-heading").focus()');
      }
      if (deviceTable) await evaluate(preview === "device-table-mixed"
        ? 'document.getElementById("assignment-devices-heading").scrollIntoView({block:"start"})'
        : 'document.querySelector(\'[aria-label="Selected device model"]\').scrollIntoView({block:"start"})');
      if (assignmentPolling || reviewBatch && (preview.startsWith("assignment-poll-") || preview === "device-table-already")) await evaluate('(document.querySelector(\'[aria-label="Device assignment progress"]\') || document.querySelector(\'[aria-label="Selected device model"]\')).scrollIntoView({block:"start"})');
      if (preview === "loaded") {
        const { root } = await send("DOM.getDocument", {}, sessionId);
        const { nodeId } = await send("DOM.querySelector", { nodeId: root.nodeId, selector: "#thing-model-upload" }, sessionId);
        await send("DOM.setFileInputFiles", { nodeId, files: [modelFile] }, sessionId);
        await waitFor('!!document.querySelector(".iot-solution-model-upload-area--loaded") && !document.querySelector(\'[data-testid="wizard-next-btn"]\').disabled');
      }
      let uploadActivation = null;
      if (acceptedModel || uploadSpacing) {
        await evaluate(`(() => {
          const input = document.getElementById("thing-model-upload");
          const button = document.querySelector(".iot-solution-model-upload-area[role='button']");
          const probe = { chooser: 0, buttons: [] };
          const inputClick = event => { event.preventDefault(); probe.chooser++; };
          const buttonClick = event => probe.buttons.push({trusted:event.isTrusted,type:event.type,detail:event.detail});
          const tileKey = event => {
            if (!event.repeat && (event.key === "Enter" || event.key === " ")) probe.buttons.push({trusted:event.isTrusted,type:event.type,key:event.key});
          };
          input.addEventListener("click", inputClick); button.addEventListener("click", buttonClick); button.addEventListener("keydown", tileKey);
          probe.cleanup = () => {
            input.removeEventListener("click", inputClick); button.removeEventListener("click", buttonClick); button.removeEventListener("keydown", tileKey);
          };
          window.__uploadButtonProbe = probe;
        })()`);
        try {
          const clickAt = async (x, y) => {
            await send("Input.dispatchMouseEvent", { type: "mousePressed", x, y, button: "left", clickCount: 1 }, sessionId);
            await send("Input.dispatchMouseEvent", { type: "mouseReleased", x, y, button: "left", clickCount: 1 }, sessionId);
          };
          const geometry = await evaluate(`(() => {
            const rect = document.querySelector(".iot-solution-model-upload-area").getBoundingClientRect();
            return {background:{x:rect.left+6,y:rect.top+6}};
          })()`);
          await clickAt(geometry.background.x, geometry.background.y);
          if (await evaluate("window.__uploadButtonProbe.chooser") !== 1) throw new Error("Tile click must open chooser exactly once.");
          await evaluate('document.querySelector(".iot-solution-model-upload-area[role=button]").focus()');
          for (const [key, code, windowsVirtualKeyCode, expected] of [["Enter", "Enter", 13, 2], [" ", "Space", 32, 3]]) {
            await send("Input.dispatchKeyEvent", { type: "keyDown", key, code, windowsVirtualKeyCode,
              nativeVirtualKeyCode: windowsVirtualKeyCode, text: key === "Enter" ? "\r" : " ",
              unmodifiedText: key === "Enter" ? "\r" : " " }, sessionId);
            await send("Input.dispatchKeyEvent", { type: "keyUp", key, code, windowsVirtualKeyCode }, sessionId);
            if (await evaluate("window.__uploadButtonProbe.chooser") !== expected) {
              const state = await evaluate("({chooser:window.__uploadButtonProbe.chooser,buttons:window.__uploadButtonProbe.buttons,active:document.activeElement?.outerHTML})");
              throw new Error(`${code} must open chooser exactly once: ${JSON.stringify(state)}`);
            }
          }
          uploadActivation = await evaluate("({chooser:window.__uploadButtonProbe.chooser,buttons:window.__uploadButtonProbe.buttons})");
          if (uploadActivation.buttons.length !== 3 || uploadActivation.buttons.some(event => !event.trusted)) {
            throw new Error("Expected trusted mouse, Enter and Space upload-tile activations.");
          }
        } finally {
          await evaluate("window.__uploadButtonProbe.cleanup(); delete window.__uploadButtonProbe");
        }
      }
      if (uploadSpacing) {
        await send("Input.dispatchMouseEvent", { type: "mouseMoved", x: 10, y: 10 }, sessionId);
        await waitFor(`(() => {
          const tile = document.querySelector(".iot-solution-model-upload-area");
          const style = getComputedStyle(tile), normalized = document.createElement("span").style;
          normalized.backgroundColor = style.getPropertyValue(tile.classList.contains("iot-solution-drop-zone--loaded")
            ? "--colorPaletteGreenBackground1" : "--colorNeutralBackground2").trim();
          return !tile.matches(":hover") && style.backgroundColor === normalized.backgroundColor;
        })()`);
      }
      const values = await evaluate(`(() => {
        const area = document.getElementById("test-surface");
        const main = area.querySelector(".wizard-control__content-panel");
        const footer = area.querySelector(".wizard-control__footer").getBoundingClientRect();
        const guide = area.querySelector("[aria-labelledby='model-preparation-heading']");
        const upload = area.querySelector(".iot-solution-model-upload-area");
        const formatPicker = area.querySelector(".iot-solution-model-format-choices");
        const supportedModels = area.querySelector('ul[aria-label^="Supported "]');
        const next = area.querySelector('[data-testid="wizard-next-btn"]');
        const tile = area.querySelector(".iot-solution-model-upload-area--loaded");
        const fieldGroup = area.querySelector('[aria-labelledby="model-fields-group-heading"]');
        const field = area.querySelector(${JSON.stringify(complexField ? '[data-field-id="/properties/config"]' : "[data-field-id]")});
        const title = field?.querySelector(".iot-solution-model-field-heading > span");
        const description = field?.querySelector(".iot-solution-model-field-description");
        const metadata = field?.querySelector(".iot-solution-model-field-metadata");
        const role = field?.querySelector(".iot-solution-model-field-role");
        const roleControl = role?.querySelector(".fui-Dropdown, .fui-Text:not(.iot-solution-model-field-mobile-label)");
        return { heading: area.querySelector("h2").textContent, surfaceOverflow: area.scrollWidth > area.clientWidth,
          contentOverflow: main.scrollWidth > main.clientWidth, footerVisible: footer.bottom <= innerHeight + 1,
          guideHeading: area.querySelector("#model-preparation-heading")?.textContent || null,
          guideBeforeUpload: guide && upload ? !!(guide.compareDocumentPosition(upload) & Node.DOCUMENT_POSITION_FOLLOWING) : null,
          formatPickerBeforeGuide: guide && formatPicker ? !!(formatPicker.compareDocumentPosition(guide) & Node.DOCUMENT_POSITION_FOLLOWING) : null,
          guideBeforeSupportedModels: guide && supportedModels ? !!(guide.compareDocumentPosition(supportedModels) & Node.DOCUMENT_POSITION_FOLLOWING) : null,
          preparationExpanded: area.querySelector('button[aria-controls="model-preparation-content"]')?.getAttribute("aria-expanded") || null,
          failedCandidate: area.querySelector("#model-upload-error-file")?.textContent || null,
          fileError: area.querySelector("#model-upload-error-message")?.textContent || null,
          creationPresentation: ${JSON.stringify(reviewCreation)} ? {
            summary: [...area.querySelectorAll(".iot-solution-validation-row")].map(node => node.textContent),
            inputOrder: [...area.querySelectorAll(".iot-solution-form input[required]")].map(node => node.id),
            fields: [...area.querySelectorAll(".iot-solution-form > .iot-solution-field")].map(node => ({
              label: node.querySelector("label").textContent, value: node.querySelector("input").value,
              required: node.querySelector("input").required, help: node.querySelector(".iot-solution-field-hint").textContent,
              overflow: node.scrollWidth > node.clientWidth + 1,
            })),
            progress: [...area.querySelectorAll('[aria-label="Setup progress"] [data-status]')].map(node => ({
              status: node.dataset.status, label: node.querySelector(".iot-solution-creation-step-label").textContent.trim(),
              overflow: node.scrollWidth > node.clientWidth + 1,
            })),
            success: area.querySelector("#modeled-data-success")?.textContent || null,
            footer: next.textContent, proof: JSON.parse(document.documentElement.dataset.creationPresentationProof || "{}")
          } : null,
          modelHeadingTypography: ${JSON.stringify(reviewHeadings)}
            ? JSON.parse(document.documentElement.dataset.modelHeadingTypography || "{}") : null,
          promptEmptyPresentation: ${JSON.stringify(emptyPrompt)} && ${JSON.stringify(preview)} === "model-prompt-empty" ? (() => {
            const region = area.querySelector("#model-preparation-content");
            const message = region.querySelector(':scope > [role="status"]');
            const actions = region.querySelector(".iot-solution-model-actions"), disclaimer = region.querySelector(".iot-solution-model-prompt-disclaimer");
            const type = node => { const style = getComputedStyle(node); return {family:style.fontFamily,size:style.fontSize,weight:style.fontWeight,lineHeight:style.lineHeight}; };
            return { message: message.textContent, messageType: type(message),
              messageToActions: actions.getBoundingClientRect().top - message.getBoundingClientRect().bottom,
              actionsToDisclaimer: disclaimer.getBoundingClientRect().top - actions.getBoundingClientRect().bottom,
              buttons: [...actions.querySelectorAll("button")].map(node => node.textContent),
              disclaimer: disclaimer.textContent, disclaimerType: type(disclaimer),
              promptPresent: !!region.querySelector("#model-preparation-prompt"),
              hidden: region.hidden,
              overflow: [region, message, actions, disclaimer].some(node => node.scrollWidth > node.clientWidth + 1) };
          })() : null,
          readyPresentation: ${JSON.stringify(reviewReady)} ? (() => {
            const ready = area.querySelector(".iot-solution-model-ready");
            const text = node => node?.textContent?.replace(/\\s+/g, " ").trim();
            const rows = [...ready.querySelectorAll(".iot-solution-validation-row")];
            const nodes = [...ready.querySelectorAll(".iot-solution-architecture-node")];
            const footer = area.querySelector('[data-testid="wizard-complete-btn"]');
            return {
              heading: text(ready.querySelector("h2")),
              intro: text(ready.querySelector(".iot-solution-success-heading p")),
              headings: [...ready.querySelectorAll("h3")].map(text),
              database: text(ready.querySelector("#ready-resources-heading + p")),
              resources: rows.map(row => ({
                label: text(row.querySelector(".iot-solution-validation-label")),
                name: text(row.querySelector(".iot-solution-validation-value > strong")),
                help: text(row.querySelector(".iot-solution-validation-value > div")),
                labelLeft: row.querySelector(".iot-solution-validation-label").getBoundingClientRect().left,
                valueLeft: row.querySelector(".iot-solution-validation-value").getBoundingClientRect().left,
                font: getComputedStyle(row.querySelector(".iot-solution-validation-value")).fontSize,
                overflow: [row, ...row.children].some(node => node.scrollWidth > node.clientWidth + 1),
              })),
              nodes: nodes.map(node => ({
                title: text(node.querySelector(".iot-solution-flow-node-title")),
                detail: text(node.querySelector(".iot-solution-flow-node-detail")),
                top: node.getBoundingClientRect().top, left: node.getBoundingClientRect().left,
                width: node.getBoundingClientRect().width, height: node.getBoundingClientRect().height,
                overflow: node.scrollWidth > node.clientWidth + 1,
              })),
              diagramDirection: getComputedStyle(ready.querySelector(".iot-solution-architecture")).flexDirection,
              decorativeNativeArrows: ready.querySelectorAll('.iot-solution-flow-arrow[aria-hidden="true"] > svg').length,
              links: [...ready.querySelectorAll("a.fui-Link")].map(node => ({
                text: text(node), nativeAnchor: node.tagName === "A" && node.classList.contains("fui-Link"),
                disabled: node.hasAttribute("disabled") || node.getAttribute("aria-disabled") === "true",
                href: node.getAttribute("href"), target: node.target, rel: node.rel,
                ariaDescribedby: node.getAttribute("aria-describedby"),
                description: document.getElementById(node.getAttribute("aria-describedby"))?.textContent || null,
              })),
              unavailableGuidance: !!ready.querySelector("#ready-activator-guidance-unavailable, #ready-data-agent-guidance-unavailable"),
              finish: text(footer), finishEnabled: !!footer && !footer.disabled,
              font: getComputedStyle(ready.querySelector("p")).fontFamily,
              proof: JSON.parse(document.documentElement.dataset.readyPresentationProof || "{}"),
            };
          })() : null,
          modelDiagnostics: ${JSON.stringify(reviewDiagnostics)} ? {
            errorTitle: area.querySelector("#model-upload-error-file")?.textContent || null,
            error: area.querySelector("#model-upload-error-message")?.firstChild?.textContent || null,
            warnings: [...area.querySelectorAll("#model-upload-warnings li")].map(node => node.textContent.trim()),
            warningsIntro: area.querySelector("#model-upload-warnings .fui-Text")?.textContent || null,
            roleTooltip: [...document.querySelectorAll('[role="tooltip"]')].filter(node =>
              node.getBoundingClientRect().width > 10 && node.getBoundingClientRect().height > 10 &&
              getComputedStyle(node).visibility === "visible" && Number(getComputedStyle(node).opacity) > 0.95)
              .map(node => node.textContent).find(text => text === "Roles come from your DTDL model") || null,
            roleInfoGeometry: document.documentElement.dataset.roleInfoGeometry ? JSON.parse(document.documentElement.dataset.roleInfoGeometry) : null,
            roleInfoFocused: document.activeElement?.getAttribute("aria-label") === "About DTDL roles",
            roleDropdowns: area.querySelectorAll('[role="combobox"][id^="field-role-"]').length,
            warningOverflow: (() => { const warning = area.querySelector("#model-upload-warnings");
              return warning ? warning.scrollWidth > warning.clientWidth + 1 : false; })()
          } : null,
          fieldRoles: [...area.querySelectorAll('[role="combobox"][id^="field-role-"]')].map(control =>
            ({label:control.getAttribute("aria-label"),role:control.textContent,invalid:control.getAttribute("aria-invalid")})),
          normalizedName: area.querySelector("#twin-changes-name")?.value || null,
          lkvName: area.querySelector("#twin-state-name")?.value || null,
          continueEnabled: !!next && !next.disabled,
          modelLoaded: area.querySelector(".iot-solution-model-upload-area--loaded")?.textContent || null,
          calibrationLayout: document.documentElement.dataset.calibrationLayout ? JSON.parse(document.documentElement.dataset.calibrationLayout) : null,
          assignmentDialog: (() => {
            const dialog = document.querySelector("[role=dialog]");
            return dialog ? { title: dialog.querySelector(".fui-DialogTitle").textContent,
              modelId: dialog.querySelector(".fui-Text").textContent,
              rows: [...dialog.querySelectorAll("li")].map(node => node.textContent.trim()),
              actions: [...dialog.querySelectorAll("button")].map(node => node.textContent),
              focusInside: dialog.contains(document.activeElement),
              overflow: [dialog, ...dialog.querySelectorAll(".fui-DialogTitle,.fui-Text,li")].some(node => node.scrollWidth > node.clientWidth + 1),
              font: getComputedStyle(dialog.querySelector(".fui-Text")).fontFamily } : null;
          })(),
          uploadSpacing: document.documentElement.dataset.uploadSpacing ? JSON.parse(document.documentElement.dataset.uploadSpacing) : null,
          loadedTile: tile ? { role: tile.getAttribute("role"), filename: tile.querySelector("#model-upload-file-name")?.textContent,
            tabIndex: tile.getAttribute("tabindex"), replacementButton: tile.querySelector("button")?.textContent,
            replacementHelp: tile.querySelector("#model-upload-replace-help")?.textContent,
            background: getComputedStyle(tile).backgroundColor, border: getComputedStyle(tile).borderColor,
            borderStyle: getComputedStyle(tile).borderStyle, align: getComputedStyle(tile).textAlign,
            nestedActions: tile.querySelectorAll("button,input,[role='button']").length,
            filenameOverflow: tile.querySelector("#model-upload-file-name").scrollWidth > tile.clientWidth,
            paragraphsOutsideFieldGroup: fieldGroup?.previousElementSibling?.tagName === "UL" &&
              fieldGroup?.previousElementSibling?.previousElementSibling?.textContent.startsWith("Choose the fields to include as typed columns.") } : null,
          fieldPresentation: ${JSON.stringify(fieldPresentation || complexField)} && field && fieldGroup ? {
            groupHeading: fieldGroup.querySelector("#model-fields-group-heading")?.textContent,
            externalLabel: !fieldGroup.contains(area.querySelector("#model-fields-heading")),
            modelIdInsideTile: tile?.contains(area.querySelector("#model-upload-reference")),
            separateIdBox: !!area.querySelector('[aria-label="Accepted model"]'),
            searchPresent: !!area.querySelector("#find-model-field"),
            jsonPreview: [...area.querySelectorAll("summary")].some(node => node.textContent === "Uploaded Model (JSON)"),
            routineMessages: area.textContent.includes("Roles were suggested automatically from the analyzed data.") ||
              area.textContent.includes("Your explicit role choices are kept."),
            modelAsLabel: field.textContent.includes("Model As"),
            titleFont: getComputedStyle(title).fontSize, titleWeight: getComputedStyle(title).fontWeight,
            qualifierFont: getComputedStyle(field.querySelector(".iot-solution-field-qualifier")).fontSize,
            metadataFont: getComputedStyle(metadata.firstElementChild).fontSize,
            groupFont: getComputedStyle(fieldGroup.querySelector("#model-fields-group-heading")).fontSize,
            statusFont: getComputedStyle(fieldGroup.querySelector('[role="status"]')).fontSize,
            rowPadding: getComputedStyle(field).padding, rowGap: getComputedStyle(field.parentElement).gap,
            panelPadding: getComputedStyle(fieldGroup).padding, tilePadding: getComputedStyle(tile).padding,
            tileGap: getComputedStyle(tile.querySelector(".iot-solution-model-upload-identity")).gap,
            sectionMargin: getComputedStyle(area.querySelector(".iot-solution-model-fields-inline")).marginTop,
            introFont: getComputedStyle(area.querySelector(".iot-solution-step-description")).fontSize,
            introLineHeight: getComputedStyle(area.querySelector(".iot-solution-step-description")).lineHeight,
            bulletsFont: getComputedStyle(area.querySelector(".iot-solution-outcome-list")).fontSize,
            metadataCount: metadata.children.length,
            complexShapeVisible: /(^|, )(properties|items|required|oneOf|fields|mapKey|mapValue|elementSchema):/.test(metadata.textContent),
            declaredObjectTypeVisible: metadata.textContent.split(" \u00b7 ")[0] === "object",
            oneEntryPerLine: [...metadata.children].every(node => getComputedStyle(node).display === "block"),
            metadataAligned: metadata.parentElement === field.querySelector(".iot-solution-model-field-affordances") &&
              Math.abs(title.getBoundingClientRect().left - description.getBoundingClientRect().left) < 1 &&
              (innerWidth <= 620 ? metadata.getBoundingClientRect().top >= description.getBoundingClientRect().bottom
                : metadata.getBoundingClientRect().left >= field.querySelector(".iot-solution-model-field-content").getBoundingClientRect().right),
            descriptionBelow: description.getBoundingClientRect().top >= title.getBoundingClientRect().bottom,
            roleAligned: innerWidth <= 620 ? roleControl.getBoundingClientRect().top >= metadata.getBoundingClientRect().bottom
              : Math.abs(roleControl.getBoundingClientRect().top + roleControl.getBoundingClientRect().bottom -
                  metadata.getBoundingClientRect().top - metadata.getBoundingClientRect().bottom) < 2
          } : null,
          assignmentChoice: area.querySelector(".iot-solution-assignment")?.textContent || null,
          refreshControls: [...area.querySelectorAll("button[title]")].map(button => ({ label: button.getAttribute("aria-label"), title: button.title })),
          sourceSelectors: [...area.querySelectorAll('[id^="hub-model-"]')].filter(node => node.getAttribute("role") === "combobox")
            .map(node => ({ id: node.id, disabled: node.disabled, value: node.textContent })),
          analyzingDeviceData: area.textContent.includes("Analyzing device data"),
          sourceValidationAction: [...area.querySelectorAll("button")].find(node => /^(Validate source tables|Validate again)$/.test(node.textContent))?.textContent || null,
          sourceValidationSuccess: area.querySelector(".iot-solution-success-text")?.textContent || null,
          sourceValidationLoading: area.querySelector("#source-tables-validation-loading")?.textContent || null,
          promptPresent: !!area.querySelector("#model-preparation-prompt"),
          promptHeading: area.querySelector("#model-preparation-heading")?.textContent || null,
          promptActions: [...area.querySelectorAll("button")].map(node => node.textContent).filter(value => ["Get Copilot prompt","Copy prompt","Copying...","Copied","Hide prompt"].includes(value)),
          inlineFieldHeading: area.querySelector("#model-fields-heading")?.textContent || null,
          suppliedModelId: area.querySelector("#model-upload-reference")?.textContent || null,
          modelCoverage: area.querySelector(".iot-solution-model-coverage") ? {
            followupProof: document.documentElement.dataset.coverageFollowupProof ? JSON.parse(document.documentElement.dataset.coverageFollowupProof) : null,
            skippedNotice: (() => { const notice = area.querySelector('[aria-labelledby="model-reference-filter-heading"] .iot-solution-assignment-notice');
              if (!notice?.textContent.startsWith("Model ID filtering isn't available because model assignment was skipped.")) return null;
              const style = getComputedStyle(notice.querySelector(".fui-MessageBarBody"));
              return {text:notice.textContent,role:notice.getAttribute("role"),padding:getComputedStyle(notice).padding,
                family:style.fontFamily,size:style.fontSize,weight:style.fontWeight,lineHeight:style.lineHeight,style:style.fontStyle,color:style.color,
                actions:notice.querySelectorAll("button,a,[href]").length}; })(),
            dtdlNotice: (() => { const notice = area.querySelector('[aria-labelledby="model-reference-filter-heading"] .iot-solution-assignment-notice');
              if (!notice?.textContent.startsWith("Model ID filtering isn't available for DTDL models.")) return null;
              const style = getComputedStyle(notice.querySelector(".fui-MessageBarBody"));
              return {text:notice.textContent,role:notice.getAttribute("role"),padding:getComputedStyle(notice).padding,
                family:style.fontFamily,size:style.fontSize,weight:style.fontWeight,lineHeight:style.lineHeight,style:style.fontStyle,color:style.color,
                actions:notice.querySelectorAll("button,a,[href]").length}; })(),
            cardGeometry: document.documentElement.dataset.coverageCardGeometry ? JSON.parse(document.documentElement.dataset.coverageCardGeometry) : null,
            sections: [...area.querySelectorAll(".iot-solution-model-coverage h3")].map(node => node.textContent),
            options: [...area.querySelectorAll('[aria-labelledby="model-reference-filter-heading"] input[type=radio]')].map(node => ({
              checked: node.checked, described: (node.getAttribute("aria-describedby") || "").split(" ").every(id => !!document.getElementById(id)),
              value: node.value
            })),
            additionalOptions: [...area.querySelectorAll('[aria-labelledby="additional-conditions-heading"] input[type=radio]')].map(node => ({
              checked: node.checked, value: node.value
            })),
            checks: [...area.querySelectorAll(".iot-solution-model-coverage button")].filter(node =>
              node.textContent === "Check coverage").map(node => ({ text: node.textContent, disabled: node.disabled })),
            results: [...area.querySelectorAll("[data-coverage-stream]")].map(node => ({
              stream: node.dataset.coverageStream, text: node.textContent, missingFields: node.querySelectorAll("[data-missing-field]").length,
              section: node.tagName === "SECTION", details: node.querySelectorAll("dl > div").length,
              tableName: node.querySelector(".iot-solution-coverage-source")?.textContent.slice(1,-1),
              sampleCount: node.querySelector("[data-sample-count]")?.getAttribute("aria-label")?.startsWith("Matching events checked:")
                ? Number(node.querySelector("[data-sample-count]").textContent.trim()) : null,
              title: node.querySelector("h4").textContent,
              status: node.querySelector(".iot-solution-coverage-result-status").textContent,
              missingGuidance: node.querySelector("[data-missing-fields-guidance]")?.textContent || null,
              extraLimitSentence: node.textContent.includes("Results are limited to the newest 10,000 ingested raw rows, not all historical events."),
              reassuranceAfterList: (() => {
                const table = node.querySelector(".iot-solution-coverage-missing-fields");
                const guidance = node.querySelector("[data-missing-fields-guidance]");
                return guidance ? guidance.previousElementSibling === table : null;
              })(),
              reassuranceGap: (() => {
                const table = node.querySelector(".iot-solution-coverage-missing-fields");
                const guidance = node.querySelector("[data-missing-fields-guidance]");
                return guidance ? guidance.getBoundingClientRect().top - table.getBoundingClientRect().bottom : null;
              })(),
              overflow: node.scrollWidth > node.clientWidth + 1,
              missing: [...node.querySelectorAll("[data-missing-field]")].map(item => {
                const title = item.querySelector(".iot-solution-coverage-missing-title");
                const technical = item.querySelector(".iot-solution-coverage-missing-name");
                return { id:item.dataset.missingField, title:title.textContent, technical:technical?.textContent,
                  titleSize:getComputedStyle(title).fontSize, technicalSize:technical && getComputedStyle(technical).fontSize,
                  technicalStyle:technical && getComputedStyle(technical).fontStyle,
                  rowHeight:item.getBoundingClientRect().height, cells:item.children.length,
                  clipped:[...item.children].some(cell => cell.scrollWidth>cell.clientWidth+1) };
              }),
              missingColumns: node.querySelector(".iot-solution-coverage-missing-fields")
                ? node.querySelectorAll(".iot-solution-coverage-missing-fields thead th[scope=col]").length : 0,
              semanticMissingTable: node.querySelector(".iot-solution-coverage-missing-fields")?.tagName === "TABLE"
            })),
            editorColumns: area.querySelector(".iot-solution-coverage-editors")
              ? getComputedStyle(area.querySelector(".iot-solution-coverage-editors")).gridTemplateColumns.split(" ").length : 0,
            choiceWidths: [...area.querySelectorAll(".iot-solution-coverage-choice")].map(node => node.getBoundingClientRect().width),
            resultColumns: area.querySelector(".iot-solution-coverage-streams")
              ? getComputedStyle(area.querySelector(".iot-solution-coverage-streams")).gridTemplateColumns.split(" ").length : 0,
            removedSecondaryControls: !area.querySelector(".iot-solution-coverage-example, .iot-solution-coverage-help, input[type=checkbox]"),
            sampleCopy: area.querySelector("#coverage-sample-note").textContent,
            sampleBodyMatchesIntro: ["fontFamily", "fontSize", "lineHeight", "fontStyle"].every(key =>
              getComputedStyle(area.querySelector("#coverage-sample-note"))[key] ===
                getComputedStyle(area.querySelector(".iot-solution-step-description"))[key]),
            sampleFontSize: getComputedStyle(area.querySelector("#coverage-sample-note")).fontSize,
            assignmentHintCount: [...area.querySelectorAll('[role=note]')].filter(node =>
              node.textContent === "New model assignments can take some time to show up in device events").length,
            placeholders: [...area.querySelectorAll(".iot-solution-coverage-editors textarea")].map(node => ({
              text: node.placeholder, normal: getComputedStyle(node, "::placeholder").fontStyle === "normal",
              fontSize: getComputedStyle(node, "::placeholder").fontSize,
              textNormal: getComputedStyle(node).fontStyle === "normal"
            })),
            standaloneIdentity: !!area.querySelector('[aria-label="Current device model"]'),
            referenceIdsInDescriptions: area.querySelector("#reference-and-description")?.textContent.includes(window.__coveragePreviewCounts?.().modelId) &&
              area.querySelector("#reference-or-description")?.textContent === "Include events with this model's ID, plus any events that meet the additional conditions",
            streamDividers: [...area.querySelectorAll(".iot-solution-model-coverage .iot-solution-stream-filter")].map(node => getComputedStyle(node).borderTopWidth),
            dismiss: [...area.querySelectorAll("button")].some(node => node.textContent === "Dismiss"),
            hubList: !!area.querySelector('[aria-label="Source IoT Hubs"]'),
            unlinkedNotice: area.textContent.includes("None of the IoT Hubs in your dataset are linked")
          } : null,
          assignmentChoice: {
            buttons: [...area.querySelectorAll("button")].filter(button =>
              ["Assign model to devices", "Skip assignment"].includes(button.textContent)).map(button => ({
                text: button.textContent, disabled: button.disabled
              })),
            loading: !!area.querySelector('[role="progressbar"]')
          },
          hubSummary: area.querySelector('[aria-label="Source IoT Hubs"]') ? {
            kind: area.querySelector('[aria-label="Source IoT Hubs"]').dataset.summaryKind,
            headings: [...area.querySelectorAll('[aria-label="Source IoT Hubs"] h3')].map(node => node.textContent),
            names: [...area.querySelectorAll('[aria-label="Source IoT Hubs"] li')].map(node => node.textContent),
            documentationLink: (() => {
              const box = area.querySelector('[aria-label="Source IoT Hubs"]'), links = box.querySelectorAll("a.fui-Link"), node = links[0];
              return node ? { count: links.length, text: node.textContent.trim(),
                nativeAnchor: node.tagName === "A" && node.classList.contains("fui-Link"),
                disabled: node.hasAttribute("disabled") || node.getAttribute("aria-disabled") === "true",
                href: node.getAttribute("href"), target: node.target, rel: node.rel,
                ariaDescribedby: node.getAttribute("aria-describedby"),
                description: document.getElementById(node.getAttribute("aria-describedby"))?.textContent || null } : null;
            })(),
            unavailableGuidance: !!area.querySelector("#adr-link-guidance-unavailable"),
            continueInstruction: area.textContent.includes("Select Continue to skip this step."),
            geometry: (() => {
              const box = area.querySelector('[aria-label="Source IoT Hubs"]');
              const body = box.querySelector(".fui-MessageBarBody");
              const icon = box.querySelector(".fui-MessageBar__icon");
              const title = body.querySelector("h3");
              return {
                padding: getComputedStyle(box).padding,
                sectionGap: getComputedStyle(body).gap,
                withinSectionGaps: [...body.querySelectorAll("section")].map(section => getComputedStyle(section).gap),
                iconToHeadingTop: icon.getBoundingClientRect().top - title.getBoundingClientRect().top,
                iconTextGap: body.getBoundingClientRect().left - icon.getBoundingClientRect().right,
                contentTopInset: body.getBoundingClientRect().top - box.getBoundingClientRect().top,
                contentBottomInset: box.getBoundingClientRect().bottom - body.getBoundingClientRect().bottom
              };
            })()
          } : null,
          deviceList: area.querySelector('[aria-labelledby="assignment-devices-heading"]') ? {
            searchPresent: !!area.querySelector('input[aria-label="Find a Target"]'),
            refreshText: area.querySelector('button[aria-label="Refresh devices"]')?.textContent,
            refreshIcon: !!area.querySelector('button[aria-label="Refresh devices"] svg'),
            panelBackground: getComputedStyle(area.querySelector('[aria-labelledby="assignment-devices-heading"]')).backgroundColor,
            columns: [...area.querySelectorAll(".iot-solution-assignment-table th")].map(node => ({
              text: node.textContent, scope: node.scope,
              weight: getComputedStyle(node.querySelector(".iot-solution-assignment-column-label")).fontWeight,
              size: getComputedStyle(node.querySelector(".iot-solution-assignment-column-label")).fontSize,
              labelX: node.querySelector(".iot-solution-assignment-column-label").getBoundingClientRect().left
            })),
            toolbar: (() => {
              const toolbar = area.querySelector(".iot-solution-assignment-toolbar");
              const actions = toolbar.querySelector(".iot-solution-model-actions");
              const buttons = [...actions.querySelectorAll("button")];
              const panel = toolbar.parentElement, style = getComputedStyle(panel);
              return {
                labels: buttons.map(button => button.textContent || button.getAttribute("aria-label")),
                refreshRight: buttons.at(-1).getBoundingClientRect().right,
                panelContentRight: panel.getBoundingClientRect().right - parseFloat(style.paddingRight) - parseFloat(style.borderRightWidth),
                actionsTop: actions.getBoundingClientRect().top,
                toolbarTop: toolbar.getBoundingClientRect().top,
                statusBottom: toolbar.firstElementChild.getBoundingClientRect().bottom,
                statusCenter: (toolbar.firstElementChild.getBoundingClientRect().top + toolbar.firstElementChild.getBoundingClientRect().bottom) / 2,
                controlsCenter: (actions.getBoundingClientRect().top + actions.getBoundingClientRect().bottom) / 2,
                oldDropdown: !!toolbar.querySelector("select"),
                overflow: actions.scrollWidth > actions.clientWidth + 1
              };
            })(),
            counts: area.querySelector("#assignment-device-counts")?.textContent,
            rows: [...area.querySelectorAll('tr.iot-solution-assignment-device')].map(row => ({
              name: row.querySelector(".iot-solution-assignment-device-identity").children[1].firstElementChild.textContent,
              assignment: row.children[1].textContent,
              cells: row.children.length,
              padding: getComputedStyle(row.children[0]).padding, background: getComputedStyle(row.children[0]).backgroundColor,
              primaryFont: getComputedStyle(row.querySelector(".iot-solution-assignment-device-identity").children[1].firstElementChild).fontSize,
              secondaryFont: getComputedStyle(row.querySelector(".iot-solution-assignment-device-identity").children[1].lastElementChild).fontSize,
              overflow: row.scrollWidth > row.clientWidth + 1,
              identityLeft: row.querySelector(".iot-solution-assignment-device-identity").children[1].getBoundingClientRect().left,
              assignmentLeft: row.children[1].firstElementChild.getBoundingClientRect().left,
              checkbox: !!row.querySelector('input[type="checkbox"]')
            }))
          } : null,
          deviceSuccess: [...area.querySelectorAll('[aria-label="Device model assignments"] .fui-MessageBarBody')]
            .map(node => node.textContent).find(text => text.includes("assigned successfully") || text === "All devices are already assigned to this model.") || null,
          existingAssignmentsState: (() => {
            const notice = area.querySelector("[data-all-devices-assigned]");
            return notice ? { text: notice.textContent, checks: notice.querySelectorAll("svg").length } : null;
          })(),
          completedAssignments: (() => {
            const section = area.querySelector("section[aria-label='Device assignment progress'], section[aria-label='Device model assignments']");
            const rows = [...(section?.querySelectorAll(".iot-solution-creation-step") || [])];
            const list = section?.querySelector(".iot-solution-creation-steps");
            const dps = section?.querySelector('[aria-label="DPS enrollment guidance"]');
            return rows.length && rows.every(row => row.dataset.assignmentState === "confirmed") && dps ? {
              heading: section.querySelector("h3").textContent,
              names: rows.map(row => row.textContent),
              checks: rows.filter(row => row.querySelector('svg[aria-label="Assignment confirmed"]')).length,
              noBanner: !section.querySelector(".fui-MessageBar"),
              dpsAfterList: dps.previousElementSibling === list?.parentElement
            } : null;
          })(),
          dpsGuidance: (() => {
            const box = area.querySelector('[aria-label="DPS enrollment guidance"]');
            if (!box) return null;
            return {
              title: box.querySelector("h3").textContent,
              plainSection: box.tagName === "SECTION" && !box.closest(".fui-MessageBar") && !box.querySelector("svg") &&
                getComputedStyle(box).borderTopWidth === "0px" && getComputedStyle(box).padding === "0px",
              documentationLink: (() => {
                const links = box.querySelectorAll("a.fui-Link"), node = links[0];
                return node ? { count: links.length, text: node.textContent.trim(),
                  nativeAnchor: node.tagName === "A" && node.classList.contains("fui-Link"),
                  disabled: node.hasAttribute("disabled") || node.getAttribute("aria-disabled") === "true",
                  href: node.getAttribute("href"), target: node.target, rel: node.rel,
                  ariaDescribedby: node.getAttribute("aria-describedby"),
                  description: document.getElementById(node.getAttribute("aria-describedby"))?.textContent || null } : null;
              })(),
              unavailableGuidance: !!box.querySelector("#dps-guidance-unavailable")
            };
          })(),
          assignmentProgress: (() => {
            const progress = area.querySelector('[aria-label="Device assignment progress"]');
            return progress ? {
              heading: progress.querySelector("h3").textContent,
              warningProof: document.documentElement.dataset.assignmentWarningProof ? JSON.parse(document.documentElement.dataset.assignmentWarningProof) : null,
              count: progress.querySelector('[role="status"]').textContent,
              presentationProof: document.documentElement.dataset.assignmentSpacingProof
                ? JSON.parse(document.documentElement.dataset.assignmentSpacingProof) : null,
              spacing: (() => {
                const heading = progress.querySelector("h3"), summary = progress.querySelector("[data-assignment-summary]");
                const list = progress.querySelector(".iot-solution-creation-steps");
                const rows = [...list.querySelectorAll(".iot-solution-creation-step")];
                const icon = node => node.querySelector(".fui-Spinner") || node.querySelector("svg");
                return {
                  sectionGap: getComputedStyle(progress).rowGap,
                  listPaddingTop: getComputedStyle(list).paddingTop, listPaddingBottom: getComputedStyle(list).paddingBottom,
                  titleToAggregate: summary ? summary.getBoundingClientRect().top - heading.getBoundingClientRect().bottom : null,
                  aggregateToFirstDevice: summary ? rows[0].getBoundingClientRect().top - summary.getBoundingClientRect().bottom : null,
                  titleToFirstDevice: rows[0].getBoundingClientRect().top - heading.getBoundingClientRect().bottom,
                  rowGaps: rows.slice(1).map((row, index) => row.getBoundingClientRect().top - rows[index].getBoundingClientRect().bottom),
                  childIndent: list.getBoundingClientRect().left - heading.getBoundingClientRect().left,
                  rows: rows.map(row => {
                    const indicator = icon(row), content = row.querySelector(".iot-solution-creation-step-content");
                    return { state: row.dataset.assignmentState, name: row.querySelector(".iot-solution-creation-step-label").textContent,
                      indicatorWidth: indicator.getBoundingClientRect().width, indicatorHeight: indicator.getBoundingClientRect().height,
                      textLeft: content.getBoundingClientRect().left, iconToTextGap: content.getBoundingClientRect().left - indicator.getBoundingClientRect().right,
                      indicatorTopOffset: indicator.getBoundingClientRect().top - content.getBoundingClientRect().top,
                      rowHeight: row.getBoundingClientRect().height,
                      overflow: [row, content].some(node => node.scrollWidth > node.clientWidth + 1) };
                  }),
                };
              })(),
              aggregate: (() => {
                const summary = progress.querySelector("[data-assignment-summary]");
                const list = progress.querySelector(".iot-solution-creation-steps");
                return summary ? { state: summary.dataset.assignmentSummary,
                  message: summary.querySelector('[role="status"]').textContent.trim(),
                  spinner: !!summary.querySelector('[role="progressbar"]'),
                  checkmark: !!summary.querySelector("svg"),
                  childIndent: list.getBoundingClientRect().left - summary.getBoundingClientRect().left,
                  liveRegions: progress.querySelectorAll('[role="status"],[aria-live="polite"]').length } : null;
              })(),
              rows: [...progress.querySelectorAll(".iot-solution-creation-step")].map(row => ({
                name: row.querySelector(".iot-solution-creation-step-label").textContent,
                state: row.dataset.assignmentState,
                confirmedIcon: !!row.querySelector('svg[aria-label="Assignment confirmed"]'),
                iconColor: row.querySelector('svg') ? getComputedStyle(row.querySelector('svg')).color : null,
                overflow: row.scrollWidth > row.clientWidth + 1
              })),
              spinner: !!progress.querySelector('[role="progressbar"]'),
              table: !!area.querySelector("table"),
              modal: !!document.querySelector('[role="dialog"]')
            } : null;
          })(),
          modelSummary: (() => {
            const summary = area.querySelector('[aria-label="Selected device model"]');
            const warning = area.querySelector('[aria-label="Source IoT Hubs"]');
            return summary ? {
              labels: [...summary.querySelectorAll("dt")].map(node => node.textContent),
              id: summary.querySelectorAll("dd")[1]?.textContent,
              beforeWarning: !warning || !!(summary.compareDocumentPosition(warning) & Node.DOCUMENT_POSITION_FOLLOWING)
            } : null;
          })(),
          assignmentChoiceButtons: [...area.querySelectorAll("button")].filter(node =>
            ["Assign model to devices","Skip assignment"].includes(node.textContent)).map(node => node.textContent),
          assignmentTargetsVisible: !!area.querySelector('[aria-label="Device model assignments"]'),
          assignmentTargetCheckboxes: [...area.querySelectorAll('[aria-label="Device model assignments"] input[type="checkbox"]')]
            .map(node => ({label:node.getAttribute("aria-label"),checked:node.checked})),
          dtdlAssignmentInfo: area.textContent.includes("Model assignment isn't available for DTDL models."),
          assignmentStep: [...area.querySelectorAll(".wizard-control__step")].filter(node =>
            node.querySelector(".wizard-control__step-title")?.textContent === "Assign model")
            .map(node => ({number:node.querySelector(".wizard-control__step-number")?.textContent,
              description:node.querySelector(".wizard-control__step-description")?.textContent,
              unavailable:node.classList.contains("wizard-control__step--unavailable")})),
          inlineFieldCount: area.querySelectorAll("[data-field-id]").length,
          standaloneFieldsStep: [...area.querySelectorAll(".wizard-control__step-title")].some(node => node.textContent === "Fields"),
          sourceAnalysisStatus: [...area.querySelectorAll('[role="status"]')].map(node => node.textContent),
          selectedFormats: [...area.querySelectorAll('input[type="radio"]:checked')].map(node => node.value),
          modelFileInputPresent: !!area.querySelector("#thing-model-upload"),
          formatGuide: area.querySelector('ul[aria-label^="Supported "]')?.textContent || null,
          sourceColumns: area.querySelector(".iot-solution-model-source-grid") ? getComputedStyle(area.querySelector(".iot-solution-model-source-grid")).gridTemplateColumns : null };
      })()`);
      if (uploadActivation) values.uploadActivation = uploadActivation;
      if (deviceKeyboard) values.deviceKeyboard = deviceKeyboard;
      if (coverageKeyboard) values.coverageKeyboard = coverageKeyboard;
      if (modelFormatKeyboard) values.modelFormatKeyboard = modelFormatKeyboard;
      if (expectedLearnLinks.length) {
        values.learnLinkAccessibility = learnLinkAccessibility;
        values.learnLinkInteractions = learnLinkInteractions;
      }
      let name = null;
      if (!uploadSpacing || width === 1440) {
        const { data } = await send("Page.captureScreenshot", { format: "png" }, sessionId);
        name = `${width}-${preview}.png`;
        fs.writeFileSync(path.join(directory, name), Buffer.from(data, "base64"));
      }
      metrics.push({ width, preview, screenshot: name, ...values });
      if (assignmentWarning && (!values.assignmentProgress?.warningProof?.unknownStillNonretryable ||
          !values.assignmentProgress.warningProof.noSuccessAggregate || !values.assignmentProgress.warningProof.noRefreshOrReplay ||
          values.assignmentProgress.warningProof.states.join("|") !== "confirmed|failed|unknown" ||
          values.assignmentProgress.warningProof.indicators.some(indicator => indicator.iconWidth < 23 || indicator.iconHeight < 23 || !indicator.noOverflow))) {
        throw new Error(`Unknown warning appearance changed its state, bounds, alignment or outcomes: ${JSON.stringify(values.assignmentProgress)}`);
      }
      if (reviewCoverageFollowup && preview.startsWith("model-coverage-followup-") &&
          (!values.modelCoverage || !values.modelCoverage.followupProof ||
            values.modelCoverage.results.some(result => result.overflow || result.extraLimitSentence ||
              result.missingGuidance && (!result.reassuranceAfterList || Math.abs(result.reassuranceGap - 8) > 1)) ||
            values.modelCoverage.assignmentHintCount !== (preview === "model-coverage-followup-reference-results" ? 1 : 0) ||
            !!values.modelCoverage.skippedNotice !== (preview === "model-coverage-followup-skipped-results") ||
            !!values.modelCoverage.dtdlNotice !== (preview === "model-coverage-followup-dtdl-results") ||
            values.modelCoverage.followupProof.counts.join("|") !== "0|100" ||
            Math.abs(values.modelCoverage.followupProof.cardGeometry.heights[0] - values.modelCoverage.followupProof.cardGeometry.heights[1]) > 1)) {
        throw new Error(`Coverage followup list/notice/equal-height proof failed: ${JSON.stringify(values.modelCoverage)}`);
      }
      if (reviewEmptyPrompt && preview === "model-prompt-empty" &&
          (!values.promptEmptyPresentation || values.promptEmptyPresentation.messageToActions !== 12 ||
            values.promptEmptyPresentation.actionsToDisclaimer !== 12 || values.promptEmptyPresentation.overflow ||
            values.promptEmptyPresentation.messageType.size !== "14px" || values.promptEmptyPresentation.messageType.weight !== "400" ||
            values.promptEmptyPresentation.disclaimerType.size !== "12px" || values.promptEmptyPresentation.disclaimerType.weight !== "400" ||
            values.promptEmptyPresentation.buttons.join("|") !== "Hide prompt" || values.promptEmptyPresentation.promptPresent ||
            values.promptEmptyPresentation.hidden)) {
        throw new Error(`Exceptional no-sample spacing/copy/type/action proof failed: ${JSON.stringify(values.promptEmptyPresentation)}`);
      }
      if (reviewAssignmentSpacing && preview.startsWith("assignment-spacing-") &&
          (!values.assignmentProgress || values.assignmentProgress.heading !== "Assigned devices" ||
          !values.assignmentProgress.presentationProof?.geometry.noOverflow ||
          values.assignmentProgress.spacing.listPaddingTop !== "0px" || values.assignmentProgress.spacing.listPaddingBottom !== "0px" ||
          values.assignmentProgress.spacing.childIndent !== 16 || values.assignmentProgress.spacing.rowGaps.some(gap => Math.abs(gap - 8) > 1) ||
          values.assignmentProgress.rows.some(row => row.overflow) ||
          values.assignmentProgress.presentationProof.syntheticWrites !== 3 ||
          (["assignment-spacing-active", "assignment-spacing-confirmed"].includes(preview)
            ? Math.abs(values.assignmentProgress.spacing.titleToAggregate - 12) > 1 ||
              Math.abs(values.assignmentProgress.spacing.aggregateToFirstDevice - 12) > 1
            : Math.abs(values.assignmentProgress.spacing.titleToFirstDevice - 12) > 1))) {
        throw new Error(`Assignment grouping, wrapping or unchanged outcome proof failed: ${JSON.stringify(values.assignmentProgress)}`);
      }
      if (reviewHeadings && (!values.modelHeadingTypography.fieldsVisible ||
          values.modelHeadingTypography.headings.length !== 3 ||
          !values.modelHeadingTypography.noOverflow ||
          values.modelHeadingTypography.headings.some(heading => heading.tag !== "H3" || heading.size !== "16px" || heading.weight !== "600") ||
          values.modelHeadingTypography.support.tag !== "SPAN" || values.modelHeadingTypography.support.size !== "14px" ||
          !values.modelHeadingTypography.gaps)) {
        throw new Error(`Model heading/field-spacing proof failed: ${JSON.stringify(values.modelHeadingTypography)}`);
      }
      if (reviewReady && (!values.readyPresentation || values.readyPresentation.heading !== "Your modeled data is ready" ||
          values.readyPresentation.headings.join("|") !== "What you can do next|What was created|How your modeled data is built" ||
          values.readyPresentation.resources.map(row => row.label).join("|") !==
            "Modeled data table|Normalized properties table|Latest property values (materialized view)" ||
          values.readyPresentation.resources.some(row => row.overflow || row.font !== "14px" ||
            Math.abs(row.valueLeft - values.readyPresentation.resources[0].valueLeft) > 1) ||
          values.readyPresentation.nodes.map(node => node.title).join("|") !==
            "Raw IoT tables|Device model|Latest property values|Modeled data" ||
          values.readyPresentation.nodes.some(node => node.overflow) || values.readyPresentation.decorativeNativeArrows !== 3 ||
          values.readyPresentation.diagramDirection !== (width > 700 ? "row" : "column") ||
          values.readyPresentation.links.length !== readyLearnLinks.length ||
          values.readyPresentation.links.some((link, index) => !link.nativeAnchor || link.disabled ||
            link.text !== readyLearnLinks[index][0] || link.href !== readyLearnLinks[index][1] ||
            link.target !== "_blank" || link.rel !== "noopener noreferrer" || link.ariaDescribedby || link.description) ||
          values.readyPresentation.unavailableGuidance ||
          values.learnLinkAccessibility.length !== readyLearnLinks.length ||
          values.readyPresentation.finish !== "Finish and Close" || !values.readyPresentation.finishEnabled ||
          !values.readyPresentation.proof.exactCommandsUnchangedAndTransportMatched ||
          values.readyPresentation.proof.commandCount !== 9 || values.readyPresentation.proof.readyReadsOrAuthOrWrites !== 0)) {
        throw new Error(`Ready content/wrapping/guard/transport proof failed: ${JSON.stringify(values.readyPresentation)}`);
      }
      if (reviewCreation && (values.creationPresentation.inputOrder.join("|") !== "modeled-output-name|twin-changes-name|twin-state-name" ||
          values.creationPresentation.footer !== "Continue" || values.creationPresentation.fields.some(field => !field.required || field.overflow) ||
          values.creationPresentation.progress.some(row => row.overflow) ||
          preview === "model-creation-complete" && (values.creationPresentation.progress.some(row => row.status !== "passed") ||
            values.creationPresentation.success !== "Modeled data created. Select Continue to see what was created and what to do next." ||
            !values.creationPresentation.proof.exactCommandsUnchangedAndTransportMatched))) {
        throw new Error(`Creation presentation/dependency proof failed: ${JSON.stringify(values.creationPresentation)}`);
      }
      if (reviewDiagnostics && (values.modelDiagnostics.warningOverflow ||
          preview === "model-diagnostic-dtdl-role" && (!values.modelDiagnostics.roleTooltip ||
            !values.modelDiagnostics.roleInfoFocused || values.modelDiagnostics.roleDropdowns || values.fileError) ||
          ["model-diagnostic-json", "model-diagnostic-missing-id"].includes(preview) &&
            (!values.modelDiagnostics.errorTitle || !values.modelDiagnostics.error || values.continueEnabled))) {
        throw new Error(`Diagnostic wrapping/state/accessibility failed: ${JSON.stringify(values.modelDiagnostics)}`);
      }
      if (reviewCoverage && (values.heading !== "Filter events and check model coverage" || !values.modelCoverage ||
          values.modelCoverage.sections.join("|") !== "1. Filter by model ID|2. Add conditions (optional)|3. Check model coverage" ||
          !values.modelCoverage.referenceIdsInDescriptions ||
          values.modelCoverage.results.some(result => result.overflow ||
            !!result.missingGuidance !== (result.sampleCount > 0 && result.missingFields > 0) ||
            result.extraLimitSentence || result.missingGuidance && (!result.reassuranceAfterList || Math.abs(result.reassuranceGap - 8) > 1) ||
            result.missing.some(field => field.clipped)))) {
        throw new Error(`Coverage copy/state/wrapping capture failed: ${JSON.stringify(values.modelCoverage)}`);
      }
      if (reviewBatch && preview === "assignment-confirmation" && (!values.assignmentDialog ||
          values.assignmentDialog.rows.length !== 2 || values.assignmentDialog.actions.join("|") !== "Cancel|Assign" ||
          values.assignmentDialog.overflow || !values.assignmentDialog.focusInside)) {
        throw new Error(`Confirmation review capture lost content, wrapping or native focus: ${JSON.stringify(values.assignmentDialog)}`);
      }
      if (values.surfaceOverflow || values.contentOverflow || !values.footerVisible || values.guideBeforeUpload === false ||
          values.formatPickerBeforeGuide === false || values.guideBeforeSupportedModels === false) {
        throw new Error(`Mock ${width}/${preview} has overflow, an inaccessible footer or incorrect helper order: ${JSON.stringify(values)}`);
      }
      if (acceptedModel && (!values.loadedTile || values.loadedTile.role !== "button" || values.loadedTile.tabIndex !== "0" ||
          values.loadedTile.nestedActions !== 0 || values.loadedTile.replacementButton ||
          values.loadedTile.replacementHelp !== "Click or drop a new file to replace" ||
          values.loadedTile.filenameOverflow || !values.loadedTile.paragraphsOutsideFieldGroup ||
          values.loadedTile.filename !== "thermoProX200.json" || values.loadedTile.borderStyle !== "solid" ||
          values.loadedTile.align !== "center")) {
        throw new Error(`Accepted model success tile failed: ${JSON.stringify(values)}`);
      }
      if ((fieldPresentation || complexField) && (!values.fieldPresentation || !values.fieldPresentation.externalLabel ||
          !values.fieldPresentation.modelIdInsideTile || values.fieldPresentation.separateIdBox ||
          values.fieldPresentation.searchPresent || values.fieldPresentation.jsonPreview || values.fieldPresentation.routineMessages ||
          values.fieldPresentation.modelAsLabel || values.fieldPresentation.titleWeight !== "400" ||
          !values.fieldPresentation.oneEntryPerLine || !values.fieldPresentation.metadataAligned ||
          !values.fieldPresentation.descriptionBelow || !values.fieldPresentation.roleAligned ||
          values.fieldPresentation.groupHeading !== "Fields")) {
        throw new Error(`Field presentation failed: ${JSON.stringify(values)}`);
      }
      if (complexField && (values.fieldPresentation.complexShapeVisible || !values.fieldPresentation.declaredObjectTypeVisible)) {
        throw new Error(`Complex field shape not hidden at render time: ${JSON.stringify(values)}`);
      }
      if (assignmentIntent && (values.heading !== "Assign the model to devices" ||
          values.assignmentStep.length !== 1 || values.assignmentStep[0].number !== "3" ||
          values.assignmentStep[0].unavailable ||
          values.assignmentStep[0].description !== "Optionally assign the model to your IoT devices" ||
          (preview === "assignment-intent-active"
            ? !values.assignmentTargetsVisible || values.assignmentChoiceButtons.length !== 0 || !values.continueEnabled ||
              values.assignmentTargetCheckboxes.some(item => item.checked)
            : values.assignmentTargetsVisible ||
              preview === "assignment-intent-initial" && (values.continueEnabled || values.assignmentChoiceButtons.join("|") !== "Assign model to devices|Skip assignment") ||
              preview === "assignment-intent-dtdl" && (!values.continueEnabled || !values.dtdlAssignmentInfo || values.assignmentChoiceButtons.length !== 0)))) {
        throw new Error(`Assignment intent preview mismatch: ${JSON.stringify(values)}`);
      }
      if (hubSummary) {
        const initial = preview === "hub-summary-initial", loading = preview === "hub-summary-loading";
        const summary = values.hubSummary, geometry = summary?.geometry, link = summary?.documentationLink;
        if (values.heading !== "Assign the model to devices" || values.continueEnabled !== !(initial || loading) ||
            values.assignmentChoice.buttons.length !== (initial ? 2 : 0) ||
            values.assignmentChoice.buttons.some(button => button.disabled) ||
            (initial || loading ? !!summary || values.assignmentChoice.loading !== loading
              : !summary || summary.kind !== "warning" ||
                !link || link.count !== 1 || !link.nativeAnchor || link.disabled ||
                link.text !== adrLearnLink[0] || link.href !== adrLearnLink[1] ||
                link.target !== "_blank" || link.rel !== "noopener noreferrer" ||
                link.ariaDescribedby || link.description || summary.unavailableGuidance ||
                values.learnLinkAccessibility.length !== 1 || geometry.padding !== "12px 16px" ||
                geometry.sectionGap !== "16px" || geometry.withinSectionGaps.some(gap => gap !== "8px") ||
                Math.abs(geometry.iconToHeadingTop) > 1 || Math.abs(geometry.iconTextGap - 8) > 1 ||
                Math.abs(geometry.contentTopInset - 13) > 1 || Math.abs(geometry.contentBottomInset - 13) > 1 ||
                summary.headings.length !== 1 || summary.headings[0] !== (preview === "hub-summary-unlinked"
                  ? "Your IoT hubs aren't linked to Azure Device Registry" : "Some IoT hubs aren't linked to Azure Device Registry") ||
                summary.names.length !== 1 || summary.names[0] !== (preview === "hub-summary-unlinked" ? "Fixture Hub" : "Other source Hub") ||
                summary.continueInstruction !== (preview === "hub-summary-unlinked"))) {
          throw new Error(`Explicit assignment choice/summary layout mismatch: ${JSON.stringify(values)}`);
        }
      }
      if (preview === "hub-summary-devices" && (!values.deviceList || values.deviceList.searchPresent ||
          values.deviceList.refreshText !== "" || !values.deviceList.refreshIcon || values.deviceList.rows.length !== 4 ||
          values.deviceList.rows.filter(row => row.checkbox).length !== 1 ||
          values.deviceList.rows.some(row => row.overflow || row.padding !== "6px 12px" || row.primaryFont !== "14px" ||
            row.secondaryFont !== "12px" || row.background === values.deviceList.panelBackground ||
            Math.abs(row.identityLeft - values.deviceList.rows[0].identityLeft) > 1))) {
        throw new Error(`Native device list preview mismatch: ${JSON.stringify(values)}`);
      }
      if (values.dpsGuidance) {
        const link = values.dpsGuidance.documentationLink;
        if (!link || link.count !== 1 || !link.nativeAnchor || link.disabled ||
            link.text !== dpsLearnLink[0] || link.href !== dpsLearnLink[1] ||
            link.target !== "_blank" || link.rel !== "noopener noreferrer" ||
            link.ariaDescribedby || link.description || values.dpsGuidance.unavailableGuidance ||
            !values.learnLinkAccessibility?.some(node => node.name === dpsLearnLink[0] && node.href === dpsLearnLink[1])) {
          throw new Error(`DPS documentation lost its enabled native external link semantics: ${JSON.stringify(values.dpsGuidance)}`);
        }
      }
      if (deviceTable) {
        const success = preview === "device-table-confirmed" || preview === "device-table-already";
        const completionMatches = preview === "device-table-confirmed"
          ? values.completedAssignments?.heading === "Assigned devices" && values.completedAssignments.names.length === 3 &&
            values.completedAssignments.checks === 3 && values.completedAssignments.noBanner && values.completedAssignments.dpsAfterList
          : values.existingAssignmentsState?.text === "All devices already have a model assigned" && values.existingAssignmentsState.checks === 1;
        if (!values.modelSummary || values.modelSummary.labels.join("|") !== "Model|Model ID" ||
            !values.modelSummary.beforeWarning || !values.continueEnabled ||
            (success ? !completionMatches || !!values.deviceList ||
                values.dpsGuidance?.title !== "Assign the model to new devices automatically" ||
                !values.dpsGuidance.plainSection ||
                !values.dpsGuidance.documentationLink?.nativeAnchor || values.dpsGuidance.documentationLink.disabled :
              !values.deviceList || values.deviceList.columns.map(column => column.text).join("|") !== "Device|Current model" ||
              !!values.dpsGuidance ||
              values.deviceList.columns.some(column => column.scope !== "col") ||
              values.deviceList.rows.some(row => row.cells !== 2 || row.overflow || row.padding !== "6px 12px" ||
                row.primaryFont !== "14px" || row.secondaryFont !== "12px")) ||
            (preview === "device-table-mixed" && (values.deviceList.rows.length !== 5 ||
              values.deviceList.counts !== "0 selected" ||
              !values.deviceKeyboard?.trustedSpaceSelectedAndDeselected))) {
          throw new Error(`Two-column device assignment preview mismatch: ${JSON.stringify(values)}`);
        }
        const coverageChecked = ["model-coverage-results", "model-coverage-mixed-results", "model-coverage-dense-results"].includes(preview);
        const coverageMixed = preview === "model-coverage-mixed-results", coverageDense = preview === "model-coverage-dense-results";
        const expectedSamples = coverageMixed ? [0,4] : coverageDense ? [4,4] : [0,0];
        const expectedMissing = coverageMixed ? [0,15] : coverageDense ? [14,15] : [0,0];
        if (modelCoverage && (values.heading !== "Filter events and check model coverage" ||
            !values.modelCoverage || values.modelCoverage.sections.join("|") !== "1. Filter by model ID|2. Add conditions (optional)|3. Check model coverage" ||
            values.modelCoverage.options.length !== 3 ||
            values.modelCoverage.options.some(option => option.checked !== (coverageChecked && option.value === "and") || !option.described) ||
            values.modelCoverage.additionalOptions.length !== 2 ||
            values.modelCoverage.additionalOptions.some(option => option.checked !== (coverageChecked && option.value === "specified")) ||
            values.modelCoverage.checks.length !== 1 ||
            values.modelCoverage.checks[0].disabled !== !coverageChecked ||
            values.modelCoverage.checks[0].text !== "Check coverage" ||
            values.modelCoverage.standaloneIdentity || !values.modelCoverage.referenceIdsInDescriptions || values.modelCoverage.results.length !== (coverageChecked ? 2 : 0) ||
            values.modelCoverage.resultColumns !== (coverageChecked ? width > 1100 ? 2 : 1 : 0) || !values.modelCoverage.removedSecondaryControls ||
            !values.modelCoverage.sampleBodyMatchesIntro || values.modelCoverage.sampleFontSize !== "14px" ||
            values.modelCoverage.sampleCopy !== "Check up to 100 recent matching events from each table to see which of your selected fields appear" ||
            values.modelCoverage.assignmentHintCount !== (coverageChecked && !coverageDense ? 1 : 0) ||
            values.modelCoverage.results.some(result => !result.section || result.details !== 3 || !result.tableName) ||
            Math.max(...values.modelCoverage.choiceWidths) - Math.min(...values.modelCoverage.choiceWidths) > 1 ||
            values.modelCoverage.editorColumns !== (!coverageChecked ? 0 : width > 1100 ? 2 : 1) ||
            (coverageChecked && (!values.modelCoverage.results.every((result,index) =>
              result.sampleCount === expectedSamples[index] && result.missingFields === expectedMissing[index] &&
              (expectedMissing[index] ? result.semanticMissingTable && result.missingColumns === 1 : !result.semanticMissingTable)) ||
                values.modelCoverage.placeholders.length !== 2 ||
                values.modelCoverage.placeholders.some((placeholder, index) => !placeholder.normal || placeholder.fontSize !== "12px" || !placeholder.textNormal ||
                  !placeholder.text.startsWith("Condition for " + values.modelCoverage.results[index].tableName + "\nExample: ") ||
                  placeholder.text.includes("parse_json")) ||
                values.modelCoverage.results.some(result => result.missing.some(item => item.clipped || item.titleSize !== "14px" ||
                  item.technicalSize !== "14px" || item.technicalStyle !== "normal")))) ||
            values.modelCoverage.streamDividers.some(border => border !== "0px") ||
            values.modelCoverage.dismiss || values.modelCoverage.hubList ||
            values.modelCoverage.unlinkedNotice || values.continueEnabled !== coverageChecked)) {
          throw new Error(`Native Model Coverage preview mismatch: ${JSON.stringify(values)}`);
        }
      }
        if (assignmentPolling && (preview === "assignment-poll-progress"
          ? !values.assignmentProgress || values.assignmentProgress.heading !== "Assigned devices" ||
            values.assignmentProgress.count.trim() !== "Assigning the model to 3 devices..." ||
            values.assignmentProgress.rows.length !== 3 || values.assignmentProgress.rows.filter(row => row.confirmedIcon).length !== 1 ||
            !values.assignmentProgress.spinner || values.assignmentProgress.table || values.assignmentProgress.modal ||
            values.assignmentProgress.rows.some(row => row.overflow) || !!values.dpsGuidance || !!values.deviceSuccess
          : !values.assignmentProgress || !!values.deviceList || !!values.deviceSuccess ||
            values.assignmentProgress.heading !== "Assigned devices" || values.assignmentProgress.spinner ||
            values.completedAssignments?.checks !== 3 || !values.completedAssignments?.dpsAfterList ||
            values.dpsGuidance?.title !== "Assign the model to new devices automatically")) {
          throw new Error(`Assignment polling progress/success mismatch: ${JSON.stringify(values)}`);
        }
        if (deviceToolbar && (values.deviceList.columns.some(column => column.weight !== "600" || column.size !== "14px") ||
            values.deviceList.rows.some(row => Math.abs(row.identityLeft - values.deviceList.columns[0].labelX) > 0.5 ||
              Math.abs(row.assignmentLeft - values.deviceList.columns[1].labelX) > 0.5) ||
            values.deviceList.toolbar.labels.join("|") !== "Select all|Show unassigned devices|Refresh devices" ||
            values.deviceList.toolbar.oldDropdown || values.deviceList.toolbar.overflow ||
            Math.abs(values.deviceList.toolbar.refreshRight - values.deviceList.toolbar.panelContentRight) > 0.5 ||
            (values.deviceList.toolbar.actionsTop < values.deviceList.toolbar.statusBottom + 7 &&
              Math.abs(values.deviceList.toolbar.statusCenter - values.deviceList.toolbar.controlsCenter) > 0.5))) {
          throw new Error(`Device table header/toolbar geometry mismatch: ${JSON.stringify(values.deviceList)}`);
        }
      if (preview.startsWith("error-") && (!values.failedCandidate || values.continueEnabled || values.modelLoaded)) {
        throw new Error(`Rejected-file preview is not an identified failure: ${JSON.stringify(values)}`);
      }
      if (preview === "schema") {
        await evaluate('document.querySelector("#model-preparation-prompt").scrollIntoView({block:"start"})');
        const { data: resultData } = await send("Page.captureScreenshot", { format: "png" }, sessionId);
        fs.writeFileSync(path.join(directory, `${width}-schema-result.png`), Buffer.from(resultData, "base64"));
        await evaluate('[...document.querySelectorAll("button")].find(button => button.textContent === "Hide prompt").click()');
        await waitFor('document.activeElement?.textContent === "Get Copilot prompt"');
        const { data: uploadData } = await send("Page.captureScreenshot", { format: "png" }, sessionId);
        fs.writeFileSync(path.join(directory, `${width}-upload-anchor.png`), Buffer.from(uploadData, "base64"));
      }
      if (alignment) {
        const layout = await evaluate(`(() => {
          const area = document.getElementById("test-surface");
          return { width: innerWidth, preview: ${JSON.stringify(preview)},
            fieldHeaders: [...area.querySelectorAll(".iot-solution-model-field-heading")].map(node => node.textContent),
            qualifiers: [...area.querySelectorAll(".iot-solution-field-qualifier")].map(node => ({
              text: node.textContent, breaks: node.querySelectorAll("wbr").length,
              overflow: node.scrollWidth > node.clientWidth + 1, width: node.getBoundingClientRect().width })),
            checkboxFlex: [...area.querySelectorAll(".iot-solution-model-field-include")].map(node => getComputedStyle(node).flexShrink),
            outputNamesHelp: area.querySelector("#output-names-help")?.textContent || null,
            createdOutputs: [...area.querySelectorAll(".iot-solution-validation-value strong")].map(node => node.textContent),
            contentOverflow: area.querySelector(".wizard-control__content-panel").scrollWidth > area.querySelector(".wizard-control__content-panel").clientWidth };
        })()`);
        fs.writeFileSync(path.join(directory, `${width}-${preview}-geometry.json`), JSON.stringify(layout, null, 2));
        if (layout.contentOverflow || layout.qualifiers.some(value => value.overflow) || layout.checkboxFlex.some(value => value !== "0"))
          throw new Error(`Alignment content/qualifier/checkbox geometry failed: ${JSON.stringify(layout)}`);
        const focus = preview === "alignment-fields" ? '.iot-solution-field-qualifier' : preview === "alignment-review" ? '.iot-solution-form' : null;
        if (focus) {
          await evaluate(`document.querySelector(${JSON.stringify(focus)}).scrollIntoView({block:"center"})`);
          const { data: focused } = await send("Page.captureScreenshot", { format: "png" }, sessionId);
          fs.writeFileSync(path.join(directory, `${width}-${preview}-focus.png`), Buffer.from(focused, "base64"));
        }
      }
      if (reviewBatch && ["calibration-layout", "hub-summary-mixed", "hub-summary-unlinked"].includes(preview)) {
        const anchors = preview === "calibration-layout" ? [["select-fields-guidance", "#model-fields-heading"]]
          : [["hub-link-warning", '[aria-label="Source IoT Hubs"]'],
            ...(preview === "hub-summary-mixed" ? [["select-devices", "#assignment-devices-heading"]] : [])];
        const additional = [];
        for (const [label, selector] of anchors) {
          await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:"start"})`);
          const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
          const file = `${width}-${preview}-${label}.png`;
          fs.writeFileSync(path.join(directory, file), Buffer.from(image.data, "base64"));
          additional.push(file);
        }
        metrics[metrics.length - 1].supplementaryCaptures = additional;
      }
      if (reviewCoverage) {
        const anchors = preview.endsWith("-results")
          ? [["properties", '[data-coverage-stream="twin"]'],
            ...values.modelCoverage.results.filter(result => result.missingFields > 0).map(result =>
              [`${result.stream}-fields-end`, `[data-coverage-stream="${result.stream}"] [data-missing-field]:last-child`])]
          : [["conditions", "#additional-conditions-heading"], ["check-coverage", "#coverage-results-heading"]];
        const additional = [];
        for (const [label, selector] of anchors) {
          await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:${JSON.stringify(label.endsWith("-end") ? "end" : "start")}})`);
          const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
          const file = `${width}-${preview}-${label}.png`;
          fs.writeFileSync(path.join(directory, file), Buffer.from(image.data, "base64"));
          additional.push(file);
        }
        metrics[metrics.length - 1].supplementaryCaptures = additional;
      }
      if (reviewDiagnostics && ["model-diagnostic-warnings", "model-diagnostic-dtdl-role"].includes(preview)) {
        await evaluate('document.getElementById("model-fields-heading").scrollIntoView({block:"start"})');
        const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
        const file = `${width}-${preview}-fields.png`;
        fs.writeFileSync(path.join(directory, file), Buffer.from(image.data, "base64"));
        metrics[metrics.length - 1].supplementaryCaptures = [file];
      }
      if (reviewCreation) {
        const anchors = preview === "model-creation-form"
          ? [["resources", "#creation-resources-heading"], ["names", 'label[for="modeled-output-name"]']]
          : [["success", "#modeled-data-success"]];
        const additional = [];
        for (const [label, selector] of anchors) {
          await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:"start"})`);
          const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
          const file = `${width}-${preview}-${label}.png`;
          fs.writeFileSync(path.join(directory, file), Buffer.from(image.data, "base64"));
          additional.push(file);
        }
        metrics[metrics.length - 1].supplementaryCaptures = additional;
      }
      if (reviewReady) {
        const additional = [];
        for (const [label, selector] of [["resources", "#ready-resources-heading"], ["diagram", "#ready-architecture-heading"]]) {
          await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:"start"})`);
          const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
          const file = `${width}-${preview}-${label}.png`;
          fs.writeFileSync(path.join(directory, file), Buffer.from(image.data, "base64"));
          additional.push(file);
        }
        metrics[metrics.length - 1].supplementaryCaptures = additional;
      }
      if (reviewHeadings) {
        const additional = [];
        const anchors = [["support-upload", "#model-supported-heading"], ["fields", "#model-fields-heading"],
          ...(preview === "model-headings-wot" ? [["missing-roles", '[data-selection-needed="true"]']] : [])];
        for (const [label, selector] of anchors) {
          await evaluate(`document.querySelector(${JSON.stringify(selector)}).scrollIntoView({block:"start"})`);
          const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
          const file = `${width}-${preview}-${label}.png`;
          fs.writeFileSync(path.join(directory, file), Buffer.from(image.data, "base64"));
          additional.push(file);
        }
        metrics[metrics.length - 1].supplementaryCaptures = additional;
      }
      if (reviewCoverageFollowup && preview.startsWith("model-coverage-followup-") && width === 520) {
        await evaluate('document.querySelector(\'[data-coverage-stream="twin"]\').scrollIntoView({block:"center",behavior:"instant"})');
        await send("Runtime.evaluate", { expression: "new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(() => resolve(true))))",
          awaitPromise: true, returnByValue: true }, sessionId);
        const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
        const file = `${width}-${preview}-properties.png`;
        fs.writeFileSync(path.join(directory, file), Buffer.from(image.data, "base64"));
        metrics[metrics.length - 1].supplementaryCaptures = [file];
      }
    }
  }
  fs.writeFileSync(path.join(directory, "metrics.json"), JSON.stringify(metrics, null, 2));
  console.log(JSON.stringify({ mockCaptures: directory, count: metrics.length, images: metrics.filter(value => value.screenshot).length }));
}
async function runHeadless(browser, profile, url) {
  if (typeof WebSocket !== "function") throw new Error("Headless checks require Node 22's built-in WebSocket. Manual browser checks support the existing runtime.");
  let stderr = "", socket;
  const runtimeErrors = [];
  let blockedExternalRequests = 0;
  const child = spawn(browser, ["--headless=new", "--disable-gpu", "--no-first-run", "--no-default-browser-check",
    "--disable-background-networking", "--disable-extensions", "--remote-debugging-port=0",
    `--user-data-dir=${profile}`, "about:blank"], { windowsHide: true });
  console.log(JSON.stringify({ ownedBrowserPid: child.pid, profile }));
  const pending = new Map();
  let sequence = 0;
  try {
    const address = await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(new Error(`Owned browser startup timed out: ${stderr.slice(-2000)}`)), BROWSER_BOOT_TIMEOUT_MS);
      child.stderr.on("data", data => {
        stderr += data;
        const match = /DevTools listening on (ws:\/\/[^\s]+)/.exec(stderr);
        if (match) { clearTimeout(timeout); resolve(match[1]); }
      });
      child.once("error", error => { clearTimeout(timeout); reject(error); });
      child.once("exit", code => { clearTimeout(timeout); reject(new Error(`Owned browser exited ${code}: ${stderr.slice(-2000)}`)); });
    });
    socket = new WebSocket(address);
    await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => reject(new Error("Owned browser debugger connection timed out.")), BROWSER_BOOT_TIMEOUT_MS);
      socket.addEventListener("open", () => { clearTimeout(timeout); resolve(); }, { once: true });
      socket.addEventListener("error", error => { clearTimeout(timeout); reject(error); }, { once: true });
    });
    socket.addEventListener("message", event => {
      const value = JSON.parse(event.data);
      if (value.method === "Fetch.requestPaused") {
        const allowed = new URL(value.params.request.url).origin === new URL(url).origin;
        if (!allowed) blockedExternalRequests++;
        void send(allowed ? "Fetch.continueRequest" : "Fetch.failRequest", {
          requestId: value.params.requestId, ...(!allowed ? { errorReason: "BlockedByClient" } : {}),
        }, value.sessionId).catch(() => runtimeErrors.push("Owned test network isolation could not complete."));
        return;
      }
      if (value.method === "Runtime.exceptionThrown") {
        const detail = value.params?.exceptionDetails;
        runtimeErrors.push((detail?.exception?.description || detail?.text || "Browser exception").slice(0, 2000));
      }
      const call = pending.get(value.id);
      if (!call) return;
      pending.delete(value.id); clearTimeout(call.timeout);
      value.error ? call.reject(new Error(JSON.stringify(value.error))) : call.resolve(value.result);
    });
    const send = (method, params = {}, sessionId) => new Promise((resolve, reject) => {
      const id = ++sequence;
      const timeout = setTimeout(() => { pending.delete(id); reject(new Error(`Owned browser command timed out: ${method}`)); }, 30000);
      pending.set(id, { resolve, reject, timeout });
      socket.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
    });
    const { targetId } = await send("Target.createTarget", { url: "about:blank" });
    const { sessionId } = await send("Target.attachToTarget", { targetId, flatten: true });
    await send("Runtime.enable", {}, sessionId);
    const viewportWidth = process.argv.find(argument => argument.startsWith("--viewport-width="))?.slice("--viewport-width=".length);
    if (viewportWidth !== undefined) {
      if (!/^\d+$/.test(viewportWidth) || Number(viewportWidth) < 320 || Number(viewportWidth) > 4000) throw new Error("Viewport width must be 320-4000.");
      await send("Emulation.setDeviceMetricsOverride", { width: Number(viewportWidth), height: 900, deviceScaleFactor: 1, mobile: false }, sessionId);
    }
    await send("Fetch.enable", { patterns: [{ urlPattern: "http://*" }, { urlPattern: "https://*" }] }, sessionId);
    console.log(JSON.stringify({ networkIsolation: "HTTP(S) restricted to owned loopback origin", allowedOrigin: url }));
    const filter = process.argv.find(argument => argument.startsWith("--filter="))?.slice("--filter=".length);
    const exactQuery = exactNames.map(name => `&exact-name=${encodeURIComponent(name)}`).join("");
    await send("Page.navigate", { url: `${url}/?run=1${filter ? `&filter=${encodeURIComponent(filter)}` : ""}${exactQuery}&selectionAck=1${process.env.MODELING_TEST_DIAGNOSTICS ? "&captureFailures=1" : ""}` }, sessionId);
    let deadline = Date.now() + BROWSER_BOOT_TIMEOUT_MS;
    let selectedCount = 0, runBudgetMs = 0, currentCase, caseDeadline = deadline, timeoutReason;
    let failuresCaptured = 0;
    while (Date.now() < deadline) {
      const value = await send("Runtime.evaluate", { expression: `(() => {
        const report = document.getElementById("test-results");
        if (report?.dataset.selectionPending === "true") return JSON.stringify({
          selectionPending: true, names: JSON.parse(report.dataset.selectedNames)
        });
        if (report?.dataset.failurePending === "true") return JSON.stringify({
          pendingFailure: true, name: report.dataset.currentTest, results: report.textContent,
          dom: document.getElementById("test-surface")?.outerHTML || null
        });
        if (report?.dataset.complete === "true") return report.textContent;
        return report ? JSON.stringify({ running: true, currentTest: report.dataset.currentTest || null }) : null;
      })()`, returnByValue: true }, sessionId);
      if (value.exceptionDetails) throw new Error(JSON.stringify(value.exceptionDetails));
      if (value.result.value) {
        const results = JSON.parse(value.result.value);
        if (results.selectionPending) {
          const requested = exactNames.length ? exactNames : expectedNames;
          if (!Array.isArray(results.names) || !results.names.length ||
              new Set(results.names).size !== results.names.length ||
              requested && (results.names.length !== requested.length || results.names.some(name => !requested.includes(name))) ||
              filter && results.names.some(name => !name.startsWith(filter))) {
            throw new Error("The real case registry did not select exactly the requested names.");
          }
          selectedCount = results.names.length;
          runBudgetMs = RUN_OVERHEAD_MS + selectedCount * SELECTED_CASE_BUDGET_MS;
          deadline = Date.now() + runBudgetMs;
          caseDeadline = Date.now() + CASE_TIMEOUT_MS;
          console.log(JSON.stringify({ selectedCount, selectedNames: results.names,
            timeoutBudgets: { bootMs: BROWSER_BOOT_TIMEOUT_MS, caseMs: CASE_TIMEOUT_MS,
              runMs: runBudgetMs, runOverheadMs: RUN_OVERHEAD_MS, perSelectedCaseMs: SELECTED_CASE_BUDGET_MS } }));
          await send("Runtime.evaluate", { expression: 'document.getElementById("test-results").dataset.selectionPending = "false"' }, sessionId);
          continue;
        }
        if (results.running) {
          if (selectedCount) {
            if (results.currentTest !== currentCase) {
              currentCase = results.currentTest;
              caseDeadline = Date.now() + CASE_TIMEOUT_MS;
            }
            if (Date.now() >= caseDeadline) {
              timeoutReason = `Case ${JSON.stringify(currentCase || "starting the first case")} exceeded ${CASE_TIMEOUT_MS} ms`;
              break;
            }
          }
          await new Promise(resolve => setTimeout(resolve, 100));
          continue;
        }
        if (results.pendingFailure) {
          failuresCaptured++;
          if (process.env.MODELING_TEST_DIAGNOSTICS) {
            const directory = path.resolve(process.env.MODELING_TEST_DIAGNOSTICS);
            const prefix = process.env.MODELING_TEST_DIAGNOSTICS_PREFIX || "assignment-failure";
            if (!/^[a-z0-9-]+$/i.test(prefix)) throw new Error("Invalid diagnostic filename prefix.");
            fs.writeFileSync(path.join(directory, `${prefix}-${failuresCaptured}.json`), JSON.stringify({ ...results, runtimeErrors }, null, 2));
            const image = await send("Page.captureScreenshot", { format: "png" }, sessionId);
            fs.writeFileSync(path.join(directory, `${prefix}-${failuresCaptured}.png`), Buffer.from(image.data, "base64"));
          }
          await send("Runtime.evaluate", { expression: 'document.getElementById("test-results").dataset.failurePending = "false"' }, sessionId);
          caseDeadline = Date.now() + CASE_TIMEOUT_MS;
          continue;
        }
        const capture = process.argv.find(argument => argument.startsWith("--capture-dir="))?.slice("--capture-dir=".length);
        const model = process.argv.find(argument => argument.startsWith("--model-file="))?.slice("--model-file=".length);
        if (process.env.MODELING_TEST_DIAGNOSTICS) {
          const prefix = process.env.MODELING_TEST_DIAGNOSTICS_PREFIX || "assignment";
          if (!/^[a-z0-9-]+$/i.test(prefix)) throw new Error("Invalid diagnostic filename prefix.");
          fs.writeFileSync(path.join(path.resolve(process.env.MODELING_TEST_DIAGNOSTICS), `${prefix}-interaction-results.json`),
            JSON.stringify(results, null, 2));
        }
        if (capture && results.every(result => result.passed)) {
          await capturePreviews(send, sessionId, url, path.resolve(capture), model ? path.resolve(model) : undefined);
        }
        console.log(JSON.stringify({ blockedExternalRequests, runtimeErrors: runtimeErrors.length }));
        return results;
      }
      await new Promise(resolve => setTimeout(resolve, 100));
    }
    const state = await send("Runtime.evaluate", { expression: `JSON.stringify({
      currentTest: document.getElementById("test-results")?.dataset.currentTest || null,
      currentTestStartedAt: document.getElementById("test-results")?.dataset.currentTestStartedAt || null,
      completed: document.getElementById("test-results")?.textContent || null,
      title: document.title, surface: document.getElementById("test-surface")?.outerHTML?.slice(0, 20000) || null
    })`, returnByValue: true }, sessionId);
    if (process.env.MODELING_TEST_DIAGNOSTICS) {
      const directory = path.resolve(process.env.MODELING_TEST_DIAGNOSTICS);
      fs.writeFileSync(path.join(directory, "assignment-browser-timeout.json"),
        JSON.stringify({ ...JSON.parse(state.result.value), runtimeErrors }, null, 2));
      const screenshot = await send("Page.captureScreenshot", { format: "png" }, sessionId);
      fs.writeFileSync(path.join(directory, "assignment-browser-timeout.png"), Buffer.from(screenshot.data, "base64"));
    }
    timeoutReason ||= selectedCount ? `Run exceeded ${runBudgetMs} ms for ${selectedCount} selected cases`
      : `Case registry bootstrap exceeded ${BROWSER_BOOT_TIMEOUT_MS} ms`;
    throw new Error(`Headless component checks timed out: ${timeoutReason}. ${state.result.value}; runtimeErrors=${JSON.stringify(runtimeErrors)}`);
  } finally {
    for (const call of pending.values()) { clearTimeout(call.timeout); call.reject(new Error("Owned browser closed.")); }
    if (socket?.readyState === WebSocket.OPEN) socket.send(JSON.stringify({ id: ++sequence, method: "Browser.close" }));
    await new Promise(resolve => {
      if (child.exitCode !== null) { resolve(); return; }
      const timeout = setTimeout(() => { child.kill(); resolve(); }, 5000);
      child.once("exit", () => { clearTimeout(timeout); resolve(); });
    });
    socket?.close();
  }
}
const workload = path.resolve(__dirname, "..");
const output = fs.mkdtempSync(path.join(os.tmpdir(), "iot-modeling-regressions-"));
const compiler = webpack({
  mode: "development", context: workload, entry: "./tests/modeling-regressions.test.tsx",
  output: { path: output, filename: "checks.js" },
  resolve: { extensions: [".tsx", ".ts", ".js", ".json"], extensionAlias: { ".js": [".js", ".ts", ".tsx"] } },
  plugins: [new webpack.DefinePlugin({ "process.env.WORKLOAD_NAME": JSON.stringify("Local.MockOnly") }),
    new webpack.ProvidePlugin({ process: "process/browser.js" }),
    new HtmlPlugin({ templateContent: '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Isolated modeling component regressions - no live APIs</title></head><body><p>Mock-only component regression checks. No Azure/Fabric writes.</p></body></html>' })],
  module: { rules: [
    { test: /\.tsx?$/, exclude: /node_modules/, use: { loader: "ts-loader", options: { transpileOnly: true } } },
    { test: /\.scss$/, use: ["style-loader", "css-loader", "sass-loader"] },
    { test: /\.(png|svg|jpe?g)$/, type: "asset/resource" },
  ] },
});
compiler.run((error, stats) => {
  if (error || stats.hasErrors()) { console.error(error || stats.toString("errors-only")); process.exitCode = 1; compiler.close(() => {}); return; }
  compiler.close(() => {});
  const server = http.createServer((request, response) => {
    const relative = new URL(request.url, "http://localhost").pathname === "/" ? "index.html" : new URL(request.url, "http://localhost").pathname.slice(1);
    const file = path.resolve(output, relative);
    if (!file.startsWith(output + path.sep)) { response.writeHead(403).end(); return; }
    fs.readFile(file, (failure, data) => {
      if (failure) { response.writeHead(404).end(); return; }
      response.writeHead(200, { "Content-Type": file.endsWith(".js") ? "text/javascript" : file.endsWith(".html") ? "text/html" : "application/octet-stream" }); response.end(data);
    });
  });
  server.listen(0, "127.0.0.1", async () => {
    const url = `http://127.0.0.1:${server.address().port}`;
    console.log(JSON.stringify({ url, output }));
    if (!process.argv.includes("--headless")) return;
    const browser = [
      path.join(process.env["ProgramFiles(x86)"] || "", "Microsoft", "Edge", "Application", "msedge.exe"),
      path.join(process.env.ProgramFiles || "", "Microsoft", "Edge", "Application", "msedge.exe"),
    ].find(file => fs.existsSync(file));
    if (!browser) {
      console.error("Headless checks require an installed Microsoft Edge. Open the printed URL manually instead.");
      process.exitCode = 1; server.close(); fs.rmSync(output, { recursive: true }); return;
    }
    const profile = fs.mkdtempSync(path.join(os.tmpdir(), "iot-modeling-browser-"));
    try {
      await new Promise((resolve, reject) => http.get(url, response => {
        response.resume(); response.on("end", () => response.statusCode === 200 ? resolve() : reject(new Error("Mock server is not ready.")));
      }).on("error", reject));
      const results = await runHeadless(browser, profile, url);
      if (process.env.MODELING_TEST_DIAGNOSTICS) {
        const prefix = process.env.MODELING_TEST_DIAGNOSTICS_PREFIX || "assignment";
        if (!/^[a-z0-9-]+$/i.test(prefix)) throw new Error("Invalid diagnostic filename prefix.");
        fs.writeFileSync(path.join(path.resolve(process.env.MODELING_TEST_DIAGNOSTICS), `${prefix}-results.json`), JSON.stringify(results, null, 2));
      }
      for (const result of results) console.log(`${result.passed ? "PASS" : "FAIL"}: ${result.name}${result.error ? `\n${result.error}` : ""}`);
      if (!results.length || !results.every(result => result.passed)) process.exitCode = 1;
    } catch (error) {
      console.error(error); process.exitCode = 1;
    } finally {
      server.close();
      server.closeAllConnections();
      fs.rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 });
      fs.rmSync(output, { recursive: true, force: true });
    }
  });
});
