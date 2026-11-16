// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
import React from "react";
import { createRoot, type Root } from "react-dom/client";
import { createMemoryHistory } from "history";
import { act } from "react-dom/test-utils";
import { FluentProvider, MessageBar, MessageBarBody, webLightTheme, webDarkTheme } from "@fluentui/react-components";
import type { WorkloadClientAPI, ThemeConfiguration } from "@ms-fabric/workload-client";
import { PayloadType } from "@ms-fabric/workload-client";
import { IoTSolutionItemModelWizard } from "../app/items/IoTSolutionItem/IoTSolutionItemModelWizard";
import { HubModelFieldsStep, HubModelScopeStep, HubModelReviewStep, hubInput } from "../app/items/IoTSolutionItem/steps/HubModelWorkflowSteps";
import { HubModelAssignmentStep } from "../app/items/IoTSolutionItem/steps/HubModelAssignmentStep";
import { HubAssignmentApi, AssignmentApiError } from "../app/items/IoTSolutionItem/HubAssignmentApi";
import { ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE, allAssignmentHubsUnavailable, assignmentBlocker, assignmentIntentOwner, currentAssignmentHubSummary, currentAssignmentIntent, confirmedAssignmentCount, assignmentOwner, type AssignmentInventory, type AssignmentTarget } from "../app/items/IoTSolutionItem/HubAssignment";
import { HubModelSourceStep } from "../app/items/IoTSolutionItem/steps/HubModelSourceStep";
import { ModelPreparationStep } from "../app/items/IoTSolutionItem/steps/ModelPreparationStep";
import { currentSourceAnalysis, modelSourceKey, SOURCE_ANALYSIS_LIMIT, type SourceReadiness } from "../app/items/IoTSolutionItem/ModelSourceReadiness";
import { canonicalJson, decodeDoSamples, inferPayloadSchema, MODEL_SEMANTIC_INFO_PREFIX, parseIotMeta } from "../app/items/IoTSolutionItem/ModelAuthoring";
import { FIELD_SEMANTIC_INFO_PREFIX, readColumnSemantics } from "../app/items/IoTSolutionItem/FieldSemantics";
import { semanticDocumentation } from "../app/items/IoTSolutionItem/ModeledDataPlan";
import { metadataThingModel } from "./fieldSemanticFixtures";
import { IoTSolutionItemRibbon } from "../app/items/IoTSolutionItem/IoTSolutionItemRibbon";
import { FabricHostThemeProvider } from "../app/FabricHostThemeProvider";
import { effectiveScope, fieldValidation, fromDtdlModel, fromThingModel, inspectFieldEvidence, inspectScopedCoverage, modelingFingerprint, prepareRoleEvidence, setFieldRole } from "../app/items/IoTSolutionItem/HubModeling";
import { buildModeledDataPlan } from "../app/items/IoTSolutionItem/ModeledDataPlan";
import { parseDtdlModel, UnsupportedDtdlModelFileError } from "../app/items/IoTSolutionItem/DtdlModelParser";
import { resolveThingModel } from "../app/items/IoTSolutionItem/ThingModelParser";
import { DataAgentSetupStep } from "../app/items/IoTSolutionItem/steps/DataAgentSetupStep";
import { buildDataAgentContext, dataAgentExamplesMarkdown } from "../app/items/IoTSolutionItem/DataAgentContext";
import { dataAgentSourceFixture } from "./dataAgentFixtures";
import { ActivatorCreateResourcesStep } from "../app/items/IoTSolutionItem/steps/ActivatorCreateResourcesStep";
import { IoTSolutionItemActivatorWizard } from "../app/items/IoTSolutionItem/IoTSolutionItemActivatorWizard";
import { IoTSolutionItemDataAgentWizard } from "../app/items/IoTSolutionItem/IoTSolutionItemDataAgentWizard";
import { WizardControl, type WizardStepProps } from "../app/components/Wizard";
import type { ItemWithDefinition } from "../app/controller/ItemCRUDController";
import { convertGetItemResultToWorkloadItem } from "../app/controller/ItemCRUDController";
import type { IoTSolutionItemDefinition } from "../app/items/IoTSolutionItem/IoTSolutionItemDefinition";
import type { ViewContext } from "../app/components/ItemEditor";
import type { ModelField } from "../app/items/IoTSolutionItem/HubModeling";
import { createCreationMock, creationRequest } from "./creationMock";
import { mountedWotVariationKeys, wotVariations, wotVariationMockPayloads, wotVariationRequest } from "./wotVariationFixtures";
import { mixedAssignmentFixture, mixedAssignmentSource, mixedDtdlDocument, mixedDtdlRequest } from "./mixedSourceDtdlFixtures";
import { localFixtureRequest, localFixtureValues, localNegativeFixtures, localPositiveFixtures, localReferenceWorkModel, localRepeatedReferenceModel, localSingleModel } from "./localWotFixtures";
import { App } from "../app/App";
import { encodeJsonToBase64 } from "../app/utils/base64";
import { ModeledDataCreator, type CreationResult } from "../app/items/IoTSolutionItem/ModeledDataCreation";
import { creationFingerprint } from "../app/items/IoTSolutionItem/ModeledDataCreation";
import { HubModelConfiguredStep } from "../app/items/IoTSolutionItem/steps/HubModelConfiguredStep";
import { ModelOverviewStep } from "../app/items/IoTSolutionItem/steps/ModelOverviewStep";
import { ActivatorOverviewStep } from "../app/items/IoTSolutionItem/steps/ActivatorOverviewStep";
import { DashboardOverviewStep } from "../app/items/IoTSolutionItem/steps/DashboardOverviewStep";
import { DataAgentOverviewStep } from "../app/items/IoTSolutionItem/steps/DataAgentOverviewStep";
import { IngestionOverviewStep } from "../app/items/IoTSolutionItem/steps/IngestionOverviewStep";
import { ActivatorSummaryStep } from "../app/items/IoTSolutionItem/steps/ActivatorSummaryStep";
import { DashboardSummaryStep } from "../app/items/IoTSolutionItem/steps/DashboardSummaryStep";
import { DashboardDataStep } from "../app/items/IoTSolutionItem/steps/DashboardDataStep";
import { DataAgentSummaryStep } from "../app/items/IoTSolutionItem/steps/DataAgentSummaryStep";
import { IngestionSummaryStep } from "../app/items/IoTSolutionItem/steps/IngestionSummaryStep";
import { beginStreamCheck, finishStreamCheck, coverageConditionExample, currentStreamChecks, modelCoverageReady, streamCoverageReady, type StreamChecks } from "../app/items/IoTSolutionItem/StreamCoverage";
import { IoTSolutionItemDashboardWizard } from "../app/items/IoTSolutionItem/IoTSolutionItemDashboardWizard";
import { IoTSolutionItemDefaultView } from "../app/items/IoTSolutionItem/IoTSolutionItemDefaultView";
import { IoTHubStep } from "../app/items/IoTSolutionItem/steps/IoTHubStep";
import { dashboardClock, dashboardRow, replayDashboardQuery } from "./dashboardFixtures";
const area = document.createElement("div");
area.id = "test-surface";
document.body.appendChild(area);
let root: Root | undefined;
let currentContext: Record<string, any>;
let observedWizardProps: WizardStepProps | undefined;
let hostChange: ((theme: ThemeConfiguration) => void) | undefined;
let leaveGuard: (() => Promise<{
    canLeave: boolean;
}>) | undefined;
let tokenRequests = 0;
let consentRequests = 0;
const frontendScopeRequests: string[][] = [];
let assignmentAuthFailure = false;
let assignmentResponse: typeof fetch | undefined;
const assignmentQueries: string[] = [];
let caseObservations: Record<string, unknown> | undefined;
let deferredToken = false;
let tokenReleases: (() => void)[] = [];
let setStepVisible: (visible: boolean) => void = () => undefined;
const errors: string[] = [];
window.addEventListener("error", event => errors.push(event.error?.stack || event.message));
window.addEventListener("unhandledrejection", event => {
    if (new URLSearchParams(location.search).has("preview"))
        document.documentElement.dataset.previewError = event.reason instanceof Error ? event.reason.stack || event.reason.message : String(event.reason);
});
const assert = (condition: unknown, message: string) => { if (!condition)
    throw new Error(message); };
const pause = () => new Promise(resolve => setTimeout(resolve, 15));
async function settle() { await act(async () => { await pause(); }); }
async function until(condition: () => unknown, message: string | (() => string)) {
    for (let index = 0; index < 80; index++) {
        await settle();
        if (condition())
            return;
    }
    throw new Error(typeof message === "function" ? message() : message);
}
async function click(element: Element | null | undefined) {
    assert(element, "Expected control is missing.");
    await act(async () => { (element as HTMLElement).click(); });
    await settle();
}
function assignmentClock() {
    const originalSet = window.setTimeout, originalClear = window.clearTimeout;
    const pending = new Map<number, () => void>();
    let sequence = -1000;
    const delays: number[] = [];
    window.setTimeout = ((callback: TimerHandler, delay?: number, ...args: unknown[]) => {
        if (typeof callback === "function" && delay !== undefined && delay >= 1900 && delay <= 11000) {
            const id = sequence--;
            pending.set(id, () => callback(...args)); delays.push(delay); return id;
        }
        return originalSet(callback, delay, ...args);
    }) as typeof window.setTimeout;
    window.clearTimeout = id => { if (typeof id === "number" && pending.has(id)) pending.delete(id); else originalClear(id); };
    return {
        delays, pending,
        async tick() {
            assert(pending.size === 1, `Expected one nonoverlapping readback timer, found ${pending.size}.`);
            const [id, callback] = [...pending][0]; pending.delete(id);
            await act(async () => callback()); await settle();
        },
        async tickAll() {
            assert(pending.size > 0, "Expected active independent operation timers.");
            const callbacks = [...pending.values()]; pending.clear();
            await act(async () => callbacks.forEach(callback => callback())); await settle();
        },
        restore() { window.setTimeout = originalSet; window.clearTimeout = originalClear; pending.clear(); },
    };
}
function assignmentWaveClock() {
    const originalSet = window.setTimeout, originalClear = window.clearTimeout, originalNow = Date.now;
    const pending = new Map<number, { at: number; callback: () => void }>();
    let time = originalNow(), sequence = -10000;
    const start = time;
    Date.now = () => time;
    window.setTimeout = ((callback: TimerHandler, delay?: number, ...args: unknown[]) => {
        if (typeof callback === "function" && delay !== undefined && delay >= 1000 && delay <= 60000) {
            const id = sequence--;
            pending.set(id, { at: time + delay, callback: () => callback(...args) }); return id;
        }
        return originalSet(callback, delay, ...args);
    }) as typeof window.setTimeout;
    window.clearTimeout = id => { if (typeof id === "number" && pending.has(id)) pending.delete(id); else originalClear(id); };
    return {
        pending, elapsed: () => time - start,
        async advance(milliseconds: number) {
            const end = time + milliseconds;
            for (let turn = 0; turn < 1000; turn++) {
                const next = [...pending].filter(([, timer]) => timer.at <= end).sort((left, right) => left[1].at - right[1].at)[0];
                if (!next) { time = end; await settle(); return; }
                time = next[1].at; pending.delete(next[0]);
                await act(async () => next[1].callback()); await settle();
            }
            throw new Error("The assignment wave fixture did not settle its bounded timer queue.");
        },
        restore() { window.setTimeout = originalSet; window.clearTimeout = originalClear; Date.now = originalNow; pending.clear(); },
    };
}
const button = (name: string) => [...area.querySelectorAll("button")].find(element => element.textContent === name || element.getAttribute("aria-label") === name);
const coverageSampleCount = (row: Element | null | undefined) => {
    const cell = row?.querySelector("[data-sample-count]");
    return cell?.getAttribute("aria-label")?.startsWith("Matching events checked:") ? Number(cell.textContent?.trim()) : undefined;
};
const coverageFieldLabel = (row: Element | null | undefined) => row?.querySelector("[data-observed-fields]")?.getAttribute("aria-label");
async function chooseAssignmentView(view: "all" | "unassigned") {
    const action = button(view === "all" ? "Show all devices" : "Show unassigned devices");
    if (action) await click(action);
    assert(button(view === "all" ? "Show unassigned devices" : "Show all devices"), "View toggle did not expose its next action.");
}
function assertDpsGuidance(visible: boolean) {
    const guidance = area.querySelector('[aria-label="DPS enrollment guidance"]');
    assert(!button("Copy Model Reference") && !summary("Update DPS Enrollments Manually"),
        "Retired technical/copy DPS disclosure remains.");
    if (!visible) { assert(!guidance, "DPS guidance appeared before current verified success."); return; }
    const link = guidance?.querySelector<HTMLAnchorElement>("a.fui-Link");
    assert(guidance?.tagName === "SECTION" && !guidance.closest(".fui-MessageBar") && !guidance.querySelector("svg") &&
        guidance.getAttribute("aria-labelledby") === guidance.querySelector("h3")?.id &&
        getComputedStyle(guidance).borderTopWidth === "0px" && getComputedStyle(guidance).padding === "0px" &&
        getComputedStyle(guidance).backgroundColor === "rgba(0, 0, 0, 0)" &&
        getComputedStyle(guidance.querySelector(".fui-Text")!).fontSize === "14px" &&
        guidance.querySelector("h3")?.textContent === "Assign the model to new devices automatically" &&
        guidance.querySelector(".fui-Text")?.textContent ===
        "Update the enrollment groups and individual enrollments in your Device Provisioning Service (DPS) instance. New devices then get this model when they're provisioned, and their data flows into your modeled data." &&
        link?.textContent === "Learn how to update DPS enrollments" &&
        guidance.querySelector(".fui-Text")?.nextElementSibling?.contains(link) &&
        guidance.textContent?.split("Learn how to update DPS enrollments").length === 2 &&
        guidance.querySelectorAll("a.fui-Link").length === 1 &&
        link?.getAttribute("href") === "https://aka.ms/iot-hub-agent-ready-data-doc" &&
        link.target === "_blank" && link.rel === "noopener noreferrer" &&
        !link.hasAttribute("disabled") && link.getAttribute("aria-disabled") !== "true" &&
        !link.hasAttribute("aria-describedby") && !guidance.querySelector("#dps-guidance-unavailable") &&
        !guidance.textContent?.includes("Enrollment guidance coming soon."),
        "Success-only DPS information lost exact copy or its enabled native external-documentation semantics.");
    assert(getComputedStyle(guidance.parentElement!).borderTopWidth === "0px",
        "Assignment success sections retained the removed horizontal separator.");
}
function completedAssignmentList(): HTMLElement | undefined {
    const section = area.querySelector<HTMLElement>("section[aria-label='Device assignment progress'], section[aria-label='Device model assignments']");
    const rows = [...(section?.querySelectorAll<HTMLElement>(".iot-solution-creation-step") || [])];
    return section && rows.length > 0 && section.getAttribute("aria-busy") === "false" && !section.querySelector("table") &&
        rows.every(row => row.dataset.assignmentState === "confirmed") &&
        !!section.querySelector('[aria-label="DPS enrollment guidance"]') ? section : undefined;
}
function assertCompletedAssignments(names: string[]) {
    const section = completedAssignmentList();
    const list = section?.querySelector(".iot-solution-creation-steps");
    const rows = [...(list?.querySelectorAll(".iot-solution-creation-step") || [])];
    const dps = section?.querySelector('[aria-label="DPS enrollment guidance"]');
    assert(section?.querySelector("h3")?.textContent === "Assigned devices" &&
        !section.querySelector("[data-all-devices-assigned]") &&
        !section.querySelector(".fui-MessageBar") && !section.textContent?.includes("assigned successfully") &&
        !section.textContent?.includes("All devices are already assigned to this model.") &&
        JSON.stringify(rows.map(row => row.textContent)) === JSON.stringify(names) &&
        rows.every(row => row.querySelector('svg[aria-label="Assignment confirmed"]') && !row.querySelector('[role="progressbar"]')) &&
        dps?.previousElementSibling === list?.parentElement && list?.parentElement?.classList.contains("iot-solution-assignment-device-results") &&
        !button("Assign to selected devices") && !button("Select all"),
        "Completed assignment did not retain the exact checked-name list and stable heading before plain DPS guidance.");
    assertAssignmentSummary("confirmed", names.length);
    assertDpsGuidance(true);
}
function confirmedDeviceRows() {
    return area.querySelectorAll('[aria-label="Device assignment progress"] [data-assignment-state="confirmed"]').length;
}
function assertAssignmentSummary(state: "active" | "confirmed" | "incomplete", total: number) {
    const progress = area.querySelector<HTMLElement>('[aria-label="Device assignment progress"]')!;
    const summary = progress.querySelector<HTMLElement>("[data-assignment-summary]");
    const list = progress.querySelector<HTMLElement>(".iot-solution-creation-steps")!;
    const status = progress.querySelector<HTMLElement>('[role="status"]')!;
    const noun = total === 1 ? "device" : "devices";
    assert(progress.querySelectorAll('[role="status"]').length === 1 && !list.hasAttribute("aria-live") &&
        list.querySelectorAll("[data-assignment-state]").length === total,
        "Assignment announcements duplicate live regions or lost frozen target identity.");
    assignmentProgressGeometry();
    if (state === "incomplete") {
        assert(!summary && progress.getAttribute("aria-busy") === "false" &&
            !progress.textContent?.includes(`Model assigned to ${total} ${noun}`),
            "A settled incomplete batch retained an aggregate spinner or fabricated aggregate success.");
        return;
    }
    assert(summary?.dataset.assignmentSummary === state && status === summary.querySelector('[role="status"]') &&
        status.textContent?.trim() === (state === "active" ? `Assigning the model to ${total} ${noun}...` : `Model assigned to ${total} ${noun}`) &&
        status.getAttribute("aria-hidden") !== "true" && getComputedStyle(status).clipPath !== "inset(50%)" &&
        !!summary.querySelector('[role="progressbar"]') === (state === "active") &&
        list.getBoundingClientRect().left - summary.getBoundingClientRect().left > 0 &&
        Math.abs(list.getBoundingClientRect().left - summary.getBoundingClientRect().left -
            parseFloat(getComputedStyle(list.parentElement!).paddingInlineStart)) < 1,
        "Aggregate text/native indicator, child indentation or accessible status is incorrect.");
    if (state === "confirmed") {
        const label = summary.querySelector<HTMLElement>(".success")!, rowLabel = list.querySelector<HTMLElement>(".success")!;
        assert(confirmedDeviceRows() === total && summary.querySelector("svg")?.getAttribute("viewBox") ===
            list.querySelector("svg")?.getAttribute("viewBox") &&
            getComputedStyle(label).color === getComputedStyle(rowLabel).color &&
            getComputedStyle(label).fontFamily === getComputedStyle(rowLabel).fontFamily &&
            getComputedStyle(label).fontSize === getComputedStyle(rowLabel).fontSize,
            "Aggregate success differs from the native individual success style or lacks complete confirmation.");
    }
}
function assignmentProgressGeometry() {
    const progress = area.querySelector<HTMLElement>('[aria-label="Device assignment progress"]')!;
    const heading = progress.querySelector<HTMLElement>("h3")!, summary = progress.querySelector<HTMLElement>("[data-assignment-summary]");
    const list = progress.querySelector<HTMLElement>(".iot-solution-creation-steps")!;
    const rows = [...list.querySelectorAll<HTMLElement>(".iot-solution-creation-step")];
    const token = (name: string) => parseFloat(getComputedStyle(progress).getPropertyValue(name));
    const gaps = {
        titleToAggregate: summary ? summary.getBoundingClientRect().top - heading.getBoundingClientRect().bottom : null,
        aggregateToFirstDevice: summary ? rows[0].getBoundingClientRect().top - summary.getBoundingClientRect().bottom : null,
        titleToFirstDevice: rows[0].getBoundingClientRect().top - heading.getBoundingClientRect().bottom,
        rowGaps: rows.slice(1).map((row, index) => row.getBoundingClientRect().top - rows[index].getBoundingClientRect().bottom),
        childIndent: list.getBoundingClientRect().left - heading.getBoundingClientRect().left,
    };
    const indicators = rows.map(row => {
        const indicator = row.querySelector<HTMLElement>(".fui-Spinner") || row.querySelector<SVGElement>("svg")!;
        const content = row.querySelector<HTMLElement>(".iot-solution-creation-step-content")!;
        const gap = content.getBoundingClientRect().left - indicator.getBoundingClientRect().right;
        const topOffset = indicator.getBoundingClientRect().top - content.getBoundingClientRect().top;
        assert(Math.abs(gap - token("--spacingHorizontalS")) < 1 && Math.abs(topOffset) < 1 &&
            [row, content, content.firstElementChild!].every(node => node.scrollWidth <= node.clientWidth + 1),
            "Assignment indicator/text alignment or wrapped device labels broke in a current result row.");
        return { state: row.dataset.assignmentState, indicatorWidth: indicator.getBoundingClientRect().width,
            indicatorHeight: indicator.getBoundingClientRect().height, gap, topOffset,
            rowHeight: row.getBoundingClientRect().height, noOverflow: true };
    });
    assert(getComputedStyle(list).paddingTop === "0px" && getComputedStyle(list).paddingBottom === "0px" &&
        (summary ? Math.abs(gaps.titleToAggregate! - token("--spacingVerticalM")) < 1 &&
            Math.abs(gaps.aggregateToFirstDevice! - token("--spacingVerticalM")) < 1
            : Math.abs(gaps.titleToFirstDevice - token("--spacingVerticalM")) < 1) &&
        gaps.rowGaps.every(gap => Math.abs(gap - token("--spacingVerticalS")) < 1) &&
        Math.abs(gaps.childIndent - token("--spacingHorizontalL")) < 1 &&
        progress.scrollWidth <= progress.clientWidth + 1,
        `Assignment results have stacked padding, inconsistent group/row gaps or lost16px nesting: ${JSON.stringify(gaps)}`);
    return { viewport: innerWidth, ...gaps, indicators, noOverflow: true };
}
function assertAssignmentSemanticIndicators() {
    const progress = area.querySelector<HTMLElement>('[aria-label="Device assignment progress"]')!;
    const rows = [...progress.querySelectorAll<HTMLElement>("[data-assignment-state]")];
    const color = (node: Element, value: string) => {
        const probe = document.createElement("span"); probe.style.color = `var(${value})`; node.appendChild(probe);
        const resolved = getComputedStyle(probe).color; probe.remove(); return resolved;
    };
    const luminance = (value: string) => value.match(/[\d.]+/g)!.slice(0, 3).map(Number)
        .map(channel => channel / 255).map(channel => channel <= 0.04045 ? channel / 12.92 : ((channel + 0.055) / 1.055) ** 2.4)
        .reduce((sum, channel, index) => sum + channel * [0.2126, 0.7152, 0.0722][index], 0);
    const background = color(progress, "--colorNeutralBackground1");
    const contrast = (foreground: string) => {
        const a = luminance(foreground), b = luminance(background);
        return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
    };
    const indicators = rows.flatMap(row => {
        const state = row.dataset.assignmentState, label = row.querySelector<HTMLElement>(".iot-solution-creation-step-label")!;
        const icon = row.querySelector<SVGElement>("svg");
        if (!icon || !["unknown", "failed", "blocked", "cancelled", "confirmed"].includes(state!)) return [];
        const token = state === "unknown" ? "--colorPaletteYellowForeground1"
            : state === "confirmed" ? "--colorPaletteGreenForeground1" : "--colorPaletteRedForeground1";
        const expected = color(row, token), labelColor = getComputedStyle(label).color, iconColor = getComputedStyle(icon).fill;
        const box = icon.getBoundingClientRect();
        assert(labelColor === expected && iconColor === expected &&
            !!icon.hasAttribute("data-assignment-warning") === (state === "unknown") &&
            !row.querySelector('[role="progressbar"]') &&
            Math.abs(box.width - 24) <= 1 && Math.abs(box.height - 24) <= 1 &&
            icon.getAttribute("aria-label") === (state === "confirmed" ? "Assignment confirmed" : "Assignment not confirmed"),
            `Assignment ${state} lost its truthful native semantic color/icon or24px box: ${JSON.stringify({ labelColor, iconColor, expected, box })}`);
        if (state === "unknown") assert(contrast(labelColor) >= 4.5 && contrast(iconColor) >= 3 &&
            !label.classList.contains("success") && !label.classList.contains("error"),
            "The warning lacks readable light/dark contrast or styles an unknown outcome as confirmed/failed.");
        return [{ state, token, color: labelColor, iconColor, contrast: contrast(labelColor),
            iconWidth: box.width, iconHeight: box.height,
            fontSize: getComputedStyle(label).fontSize, fontWeight: getComputedStyle(label).fontWeight,
            lineHeight: getComputedStyle(label).lineHeight, noOverflow: row.scrollWidth <= row.clientWidth + 1 }];
    });
    return { viewport: innerWidth, background, indicators, geometry: assignmentProgressGeometry() };
}
function existingAssignmentsNotice(): HTMLElement | undefined {
    const notice = area.querySelector<HTMLElement>("[data-all-devices-assigned]");
    return notice?.closest("section")?.querySelector('[aria-label="DPS enrollment guidance"]') ? notice : undefined;
}
const acceptedModelTile = () => area.querySelector<HTMLElement>(".iot-solution-model-upload-area--loaded");
const modelUploadAction = (name: string) => area.querySelector<HTMLElement>(`.iot-solution-model-upload-area[aria-label="${name}"]`);
const sourceValidationAction = () => button("Validate source tables") || button("Validate again");
const tableChoicesReady = () => !!area.querySelector<HTMLButtonElement>("#hub-model-telemetry") &&
    !area.querySelector<HTMLButtonElement>("#hub-model-telemetry")!.disabled;
const next = () => area.querySelector<HTMLButtonElement>('[data-testid="wizard-next-btn"]');
const step = (name: string) => [...area.querySelectorAll<HTMLElement>(".wizard-control__step")]
    .find(element => element.querySelector(".wizard-control__step-title")?.textContent === name);
function assertDecorativeDiagramArrows(expectedArrows: number, expectedPlus = 0) {
    const diagram = area.querySelector<HTMLElement>(".iot-solution-overview-flow, .iot-solution-architecture");
    const markers = [...(diagram?.querySelectorAll<HTMLElement>(".iot-solution-flow-arrow") || [])];
    const plus = markers.filter(marker => marker.textContent?.trim() === "+");
    const arrows = markers.filter(marker => marker.textContent?.trim() !== "+");
    assert(diagram && arrows.length === expectedArrows && plus.length === expectedPlus &&
        diagram.querySelectorAll(".iot-solution-flow-node, .iot-solution-architecture-node").length === markers.length + 1,
        "Diagram stages, separators or plus operators changed.");
    for (const marker of arrows) {
        const svg = marker.querySelector("svg");
        assert(marker.getAttribute("aria-hidden") === "true" && marker.textContent?.trim() === "" &&
            marker.querySelectorAll("svg").length === 1 && svg?.getAttribute("viewBox") === "0 0 20 20" &&
            !!svg.querySelector("path[d]") && svg.getAttribute("aria-hidden") === "true" &&
            !svg.hasAttribute("style") &&
            !svg.querySelector("title") && getComputedStyle(svg).fontSize === getComputedStyle(marker).fontSize &&
            Math.abs(svg.getBoundingClientRect().width - parseFloat(getComputedStyle(marker).fontSize)) < 1 &&
            Math.abs(svg.getBoundingClientRect().height - parseFloat(getComputedStyle(marker).fontSize)) < 1,
            "The diagram arrow is not a decorative Fluent SVG at the inherited wrapper size.");
    }
    assert(plus.every(marker => marker.getAttribute("aria-hidden") === "true" && !marker.querySelector("svg")) &&
        getComputedStyle(diagram!).flexDirection === (innerWidth <= 700 ? "column" : "row") &&
        arrows.every(marker => innerWidth > 700 || getComputedStyle(marker).transform !== "none"),
        "Diagram SVG replacement changed its responsive direction, wrapper rotation or decorative plus marker.");
}
const summary = (name: string) => [...area.querySelectorAll<HTMLElement>("summary")].find(element => element.textContent === name);
function assertFieldSelectionCopy() {
    const heading = area.querySelector("#model-fields-heading")!, container = heading?.parentElement;
    const paragraphs = [...(container?.children || [])].filter(node => node.matches(".iot-solution-step-description"));
    const list = paragraphs[0]?.nextElementSibling;
    assert(heading?.textContent === (container?.classList.contains("iot-solution-model-fields-inline") ? "3. Select fields" : "Select fields") &&
        paragraphs.length === 1 &&
        paragraphs[0].textContent === "Choose the fields to include as typed columns. Fields found in your recent device events have a suggested role. Select a role for any others. Fields you leave out stay in your raw tables." &&
        list?.tagName === "UL" && list.classList.contains("iot-solution-outcome-list") &&
        JSON.stringify([...list.children].map(node => node.textContent)) === JSON.stringify([
            "Telemetry: Readings that change over time, like temperature",
            "Twin property: Device settings or state, like firmware version",
        ]) && JSON.stringify([...list.querySelectorAll("strong")].map(node => node.textContent)) ===
            JSON.stringify(["Telemetry:", "Twin property:"]) &&
        [...list.children].every(node => node.firstElementChild?.tagName === "STRONG" && node.childNodes.length === 2) &&
        list.nextElementSibling?.getAttribute("aria-labelledby") === "model-fields-group-heading",
        "Shared field selection lost its exact heading, single paragraph, ordered role bullets or label-only emphasis.");
    const region = container!.querySelector('[aria-labelledby="model-fields-group-heading"]')!;
    assert(region.querySelector("#model-fields-group-heading")?.textContent === "Fields" &&
        region.querySelector('[role="status"]')?.textContent?.trim() === "Select a role for every field to continue" &&
        !region.querySelector(".iot-solution-selection-recovery") && !button("Review Selections") &&
        !/\d+ (selected|role selections needed|selections needed)/.test(region.textContent || "") &&
        JSON.stringify([...region.querySelectorAll(".iot-solution-model-field-columns > span")].map(node => node.textContent)) ===
            JSON.stringify(["Field", "Type", "Role"]),
        "Fields lost the exact instruction/column titles or retained the removed counts/review action.");
}
function assertModelPreparationHeadingTypography() {
    const ids = ["model-format-heading", "model-file-upload-heading", "model-fields-heading"];
    const titles = ["1. Choose a model format", "2. Upload your model", "3. Select fields"];
    const headings = ids.map(id => area.querySelector<HTMLElement>(`#${id}`)).filter((node): node is HTMLElement => !!node);
    const style = (node: Element) => {
        const computed = getComputedStyle(node);
        return { family: computed.fontFamily, size: computed.fontSize, weight: computed.fontWeight,
            lineHeight: computed.lineHeight, letterSpacing: computed.letterSpacing };
    };
    const styles = headings.map(style);
    assert(headings.length >= 2 && headings.every((node, index) => node.tagName === "H3" &&
        node.classList.contains("iot-solution-outcome-heading") && node.textContent === titles[index] &&
        node.scrollWidth <= node.clientWidth + 1) &&
        styles.every(value => JSON.stringify(value) === JSON.stringify(styles[0])) &&
        styles[0].size === "16px" && styles[0].weight === "600",
        `Model preparation section headings lost their consistent native type/hierarchy: ${JSON.stringify(styles)}`);
    const support = area.querySelector<HTMLElement>("#model-supported-heading")!, supportStyle = style(support);
    assert(support.tagName === "SPAN" && support.textContent === "What's supported" &&
        support.classList.contains("iot-solution-field-label") && !support.matches("h1,h2,h3,h4,h5,h6,[role=heading]") &&
        !support.hasAttribute("aria-level") && supportStyle.size === "14px" && supportStyle.weight === "600" &&
        supportStyle.family === styles[0].family && parseFloat(supportStyle.size) < parseFloat(styles[0].size),
        `What's supported is not the native subordinate label: ${JSON.stringify(supportStyle)}`);
    const guidance = area.querySelector<HTMLElement>(".iot-solution-model-fields-inline");
    let gaps: { headingToDescription: number; descriptionToBullets: number; betweenBullets: number; bulletsToCard: number } | undefined;
    if (guidance) {
        const heading = guidance.querySelector<HTMLElement>("#model-fields-heading")!;
        const description = guidance.querySelector<HTMLElement>(":scope > .iot-solution-step-description")!;
        const list = guidance.querySelector<HTMLElement>(":scope > .iot-solution-outcome-list")!;
        const card = guidance.querySelector<HTMLElement>(':scope > [aria-labelledby="model-fields-group-heading"]')!;
        const bullets = [...list.children];
        gaps = { headingToDescription: description.getBoundingClientRect().top - heading.getBoundingClientRect().bottom,
            descriptionToBullets: list.getBoundingClientRect().top - description.getBoundingClientRect().bottom,
            betweenBullets: bullets[1].getBoundingClientRect().top - bullets[0].getBoundingClientRect().bottom,
            bulletsToCard: card.getBoundingClientRect().top - list.getBoundingClientRect().bottom };
        const token = (name: string) => parseFloat(getComputedStyle(guidance).getPropertyValue(name));
        assert(Math.abs(gaps.headingToDescription - token("--spacingVerticalS")) < 1 &&
            Math.abs(gaps.descriptionToBullets - token("--spacingVerticalS")) < 1 &&
            Math.abs(gaps.betweenBullets - token("--spacingVerticalXS")) < 1 &&
            Math.abs(gaps.bulletsToCard - token("--spacingVerticalL")) < 1 &&
            gaps.betweenBullets < gaps.descriptionToBullets && gaps.bulletsToCard > gaps.descriptionToBullets &&
            [description, list, ...list.children].every(node => node.scrollWidth <= node.clientWidth + 1),
            `Field guidance no longer has grouped native heading/description/list/card rhythm: ${JSON.stringify(gaps)}`);
        assertFieldSelectionCopy();
        for (const row of guidance.querySelectorAll<HTMLElement>('[data-selection-needed="true"]')) {
            const control = row.querySelector<HTMLElement>('[role="combobox"]')!;
            const cue = row.querySelector<HTMLElement>(".iot-solution-selection-cue")!;
            const field = (currentContext.hubFields as ModelField[]).find(value => value.id === row.dataset.fieldId)!;
            assert(cue.firstChild?.textContent?.trim() === "Select a role" &&
                cue.querySelector(".iot-solution-model-visually-hidden") &&
                cue.textContent?.trim() === `Select a role for ${field.title}. Choose Telemetry or Twin property.` &&
                control.getAttribute("aria-invalid") === "true" && control.getAttribute("aria-describedby") === cue.id &&
                field.included && !field.role && next()?.disabled,
                "The concise visible role cue lost its field-specific accessibility or unresolved-field continuation gate.");
        }
    }
    return { viewport: innerWidth, headings: headings.map((node, index) => ({ text: node.textContent, tag: node.tagName, ...styles[index] })),
        support: { text: support.textContent, tag: support.tagName, ...supportStyle }, gaps,
        fieldsVisible: !!guidance, visibleMissingRoleCues: [...(guidance?.querySelectorAll<HTMLElement>(".iot-solution-selection-cue") || [])]
            .map(node => node.firstChild?.textContent?.trim()), noOverflow: true };
}
function emptyPromptGeometry() {
    const region = area.querySelector<HTMLElement>("#model-preparation-content")!;
    const message = region.querySelector<HTMLElement>(":scope > .iot-solution-model-prompt-empty")!;
    const actions = region.querySelector<HTMLElement>(".iot-solution-model-actions")!;
    const disclaimer = region.querySelector<HTMLElement>(".iot-solution-model-prompt-disclaimer")!;
    const messageToActions = actions.getBoundingClientRect().top - message.getBoundingClientRect().bottom;
    const actionsToDisclaimer = disclaimer.getBoundingClientRect().top - actions.getBoundingClientRect().bottom;
    const space = parseFloat(getComputedStyle(region).getPropertyValue("--spacingVerticalM"));
    const type = (node: Element) => {
        const style = getComputedStyle(node);
        return { family: style.fontFamily, size: style.fontSize, weight: style.fontWeight, lineHeight: style.lineHeight };
    };
    const messageType = type(message), disclaimerType = type(disclaimer);
    assert(!region.hidden && message.getAttribute("role") === "status" &&
        message.textContent === "No sample records were found during source validation. Choose Validate again in Select data source after data arrives." &&
        disclaimer.textContent === "Copilot can make mistakes. Check the model before you upload it." &&
        Math.abs(messageToActions - space) < 1 && Math.abs(actionsToDisclaimer - space) < 1 &&
        messageType.size === "14px" && messageType.weight === "400" && messageType.lineHeight === "20px" &&
        disclaimerType.size === "12px" && disclaimerType.weight === "400" && disclaimerType.lineHeight === "16px" &&
        actions.querySelectorAll("button").length === 1 && actions.querySelector("button")?.textContent === "Hide prompt" &&
        !region.querySelector("#model-preparation-prompt") && !button("Copy prompt") &&
        [region, message, actions, disclaimer].every(node => node.scrollWidth <= node.clientWidth + 1),
        "The exceptional no-sample helper lost exact copy/type/actions or balanced native message/button/disclaimer spacing.");
    return { viewport: innerWidth, messageToActions, actionsToDisclaimer, messageType, disclaimerType, noOverflow: true };
}
function assertCoverageMissingGuidance() {
    const panels = [...area.querySelectorAll("[data-coverage-stream]")];
    for (const panel of panels) {
        const count = coverageSampleCount(panel), missing = panel.querySelectorAll("[data-missing-field]").length;
        const guidance = panel.querySelector("[data-missing-fields-guidance]");
        assert(!!guidance === (count !== undefined && count > 0 && missing > 0) &&
            (!guidance || guidance.textContent === "You can still continue. These fields stay empty until devices send them." &&
                getComputedStyle(guidance).fontSize === "12px"),
            "Missing-fields guidance is absent, duplicated or shown for zero, all-found or unverified evidence.");
        if (guidance) {
            const table = panel.querySelector(".iot-solution-coverage-missing-fields")!;
            assert(guidance.previousElementSibling === table &&
               !!(table.compareDocumentPosition(guidance) & Node.DOCUMENT_POSITION_FOLLOWING) &&
               Math.abs(guidance.getBoundingClientRect().top - table.getBoundingClientRect().bottom -
                   parseFloat(getComputedStyle(guidance).getPropertyValue("--spacingVerticalS"))) < 1,
               "Missing-fields reassurance is not after the entire actual field list with the native small gap.");
        }
        assert(!panel.textContent?.includes("Results are limited to the newest 10,000 ingested raw rows, not all historical events."),
            "The unsolicited limit sentence remains in a coverage card.");
    }
    if (!panels.length) assert(!area.querySelector("[data-missing-fields-guidance]"),
        "Failed, stale or unchecked coverage retained missing-fields guidance.");
}
function coverageCardGeometry(checkGrowth = false) {
    const coverage = area.querySelector<HTMLElement>(".iot-solution-model-coverage")!;
    const cards = [...coverage.querySelectorAll<HTMLElement>("[data-coverage-stream]")];
    assert(cards.length === 2, "Equal-height coverage needs both actual current stream cards.");
    const measure = () => {
        const naturalHeights = cards.map(card => {
            const probe = document.createElement("div"), copy = card.cloneNode(true) as HTMLElement;
            Object.assign(probe.style, { position: "absolute", visibility: "hidden", pointerEvents: "none",
                width: `${card.getBoundingClientRect().width}px` });
            Object.assign(copy.style, { width: "100%", height: "auto", boxSizing: "border-box" });
            probe.appendChild(copy); coverage.appendChild(probe);
            const height = copy.getBoundingClientRect().height; probe.remove(); return height;
        });
        const heights = cards.map(card => card.getBoundingClientRect().height), tallest = Math.max(...naturalHeights);
        assert(Math.abs(heights[0] - heights[1]) <= 1 && heights.every(height => Math.abs(height - tallest) <= 1) &&
            cards.every(card => card.scrollWidth <= card.clientWidth + 1 && card.scrollHeight <= card.clientHeight + 1),
            `Coverage cards do not naturally match the tallest complete content without clipping: ${JSON.stringify({ heights, naturalHeights })}`);
        return { viewport: innerWidth, heights, naturalHeights, equalWithinPx: 1, noOverflow: true };
    };
    const initial = measure();
    if (checkGrowth) {
        const guidance = cards[1].querySelector<HTMLElement>("[data-missing-fields-guidance]")!, text = guidance.textContent!;
        try {
            guidance.textContent = Array.from({ length: 8 }, () => text).join(" ");
            const grown = measure();
            assert(grown.heights[0] > initial.heights[0], "Equal-height cards did not grow with the tallest helper content.");
        } finally { guidance.textContent = text; }
        measure();
    }
    return initial;
}
function assertCreationDisplayVerbs(plan: ReturnType<typeof buildModeledDataPlan>) {
    const expected = new Map<string, { before: string; success: string }>();
    for (const command of plan.commands) {
        const metadata = ["tableDoc", "viewDoc", "columnDocs"].includes(command.kind);
        const key = `${metadata ? "description" : command.kind}:${command.target}`;
        expected.set(key, {
            before: metadata ? `Set docstring on ${command.target}` : command.label,
            success: metadata ? `Added description to ${command.target}` : command.kind === "policy"
                ? `Added update policy to ${command.target}` : `Created ${command.target}`,
        });
    }
    const rows = [...area.querySelectorAll<HTMLElement>('[aria-label="Setup progress"] [data-status]')];
    assert(rows.length === expected.size, "Creation presentation changed the logical operation groups.");
    for (const row of rows) {
        const label = row.querySelector(".iot-solution-creation-step-label")!.textContent!.trim();
        const value = [...expected.values()].find(item => item.before === label || item.success === label);
        assert(value && label === (row.dataset.status === "passed" ? value.success : value.before),
            `Unconfirmed ${row.dataset.status} effects claimed a successful display verb: ${label}`);
    }
    assert(currentContext.hubCreation.steps.every((step: { label: string }) => !/^(Created |Added update policy to |Added description to )/.test(step.label)),
        "Display verbs changed underlying execution-record labels/keys.");
}
function assertReadyPresentation(input: ReturnType<typeof creationRequest>["input"], result: CreationResult) {
    const ready = area.querySelector<HTMLElement>(".iot-solution-model-ready");
    assert(ready, "A confirmed creation did not show the current Ready presentation.");
    const text = (node: Element | null | undefined) => node?.textContent?.replace(/\s+/g, " ").trim();
    const telemetry = input.fields.filter(field => field.included && field.role === "telemetry").length;
    const twin = input.fields.filter(field => field.included && field.role === "twin").length;
    const hasProperties = !!result.plannedOutputs?.some(output => output.name === result.names.twinState && output.kind === "view");
    const intro = result.awaitingData ? telemetry === 0
        ? "Your modeled data is ready to add each device's latest property values to new telemetry events when they arrive. It stays empty for device types that send no telemetry."
        : hasProperties && twin > 0
            ? "Your modeled data is ready to turn new telemetry into typed columns, along with each device's latest property values, when it arrives."
            : "Your modeled data is ready to turn new telemetry into typed columns when it arrives."
        : twin === 0 ? "New telemetry now flows into typed columns."
        : telemetry > 0 ? "New telemetry now flows into typed columns, along with each device's latest property values."
        : "Your modeled data is ready to add each device's latest property values to telemetry events. It stays empty for device types that send no telemetry.";
    assert(text(ready.querySelector("h2")) === "Your modeled data is ready" &&
        text(ready.querySelector(".iot-solution-success-heading p")) ===
            `${intro} Query it by the field names and types from your device model.` &&
        !!ready.querySelector('.fui-MessageBar[role="status"]') === result.awaitingData &&
        !!ready.textContent?.includes("This setup will use matching telemetry when it arrives.") === result.awaitingData,
        "Ready copy confused confirmed configuration, awaiting data or property-only processing.");
    const headings = [...ready.querySelectorAll("h3")].map(text);
    assert(JSON.stringify(headings) === JSON.stringify(["What you can do next", "What was created", "How your modeled data is built"]),
        "Ready sections did not follow the supplied content order.");
    const bullets = [...ready.querySelectorAll(".iot-solution-outcome-list > li")].map(item =>
        [...item.childNodes].filter(node => node.nodeType === Node.TEXT_NODE)
            .map(node => node.textContent).join(" ").replace(/\s+/g, " ").trim());
    assert(JSON.stringify(bullets) === JSON.stringify([
        hasProperties && twin > 0
            ? "Set alerts with Activator that account for device state, like different temperature limits for each operating mode"
            : "Set alerts with Activator using the typed fields from your device model",
        "Build a Fabric data agent that answers questions about your devices in plain language",
        "Keep using your raw tables for fields that aren't in this model",
    ]), "Ready next-step copy changed its actual data meaning or kept superseded paragraphs.");
    const links = [...ready.querySelectorAll<HTMLAnchorElement>("a.fui-Link")];
    const learnHrefs = [
        "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators",
        "https://learn.microsoft.com/en-us/fabric/data-science/concept-data-agent",
    ];
    assert(links.length === 2 && JSON.stringify(links.map(text)) ===
        JSON.stringify(["Learn about Activator", "Learn about Fabric data agents"]) &&
        links.every((link, index) => link.getAttribute("href") === learnHrefs[index] && link.target === "_blank" &&
            link.rel === "noopener noreferrer" && !link.hasAttribute("disabled") && link.getAttribute("aria-disabled") !== "true" &&
            !link.hasAttribute("aria-describedby")) &&
        !ready.querySelector("#ready-activator-guidance-unavailable, #ready-data-agent-guidance-unavailable") &&
        !ready.querySelector("input,[role=combobox]") &&
        !ready.textContent?.includes("Generated modeled entities") && !ready.textContent?.includes("What you can do now") &&
        !ready.textContent?.includes("foundation for context-aware alerts"),
        "Ready documentation links lost their exact enabled external destinations or retained unavailable semantics.");
    const resourceRows = [...ready.querySelectorAll<HTMLElement>(".iot-solution-validation-row")];
    const expectedRows = [
        ["Modeled data table", result.names.modeled,
            `Your typed data: ${telemetry} telemetry ${telemetry === 1 ? "column" : "columns"} and ${twin} property ${twin === 1 ? "column" : "columns"}, plus device ID and event time, for each event`],
        ...(hasProperties ? [
            ["Normalized properties table", result.names.twinChanges, "Stores each property update as its own row: device, property, value, and time"],
            ["Latest property values (materialized view)", result.names.twinState, "Keeps each device's latest reported value for every property from the last 30 days"],
        ] : []),
    ];
    const resources = resourceRows.map(row => [
        text(row.querySelector(".iot-solution-validation-label")),
        text(row.querySelector(".iot-solution-validation-value > strong")),
        text(row.querySelector(".iot-solution-validation-value > div")),
    ]);
    assert(JSON.stringify(resources) === JSON.stringify(expectedRows) &&
        text(ready.querySelector("#ready-resources-heading + p")) ===
            `These resources were created in the ${result.databaseName} KQL database` &&
        !ready.textContent?.includes("Destination KQL database"),
        "Ready omitted actual verified names/database/included-role counts or invented unplanned resources.");
    const positions = resourceRows.map(row => {
        const label = row.querySelector<HTMLElement>(".iot-solution-validation-label")!;
        const value = row.querySelector<HTMLElement>(".iot-solution-validation-value")!;
        assert(getComputedStyle(row).display === "grid" && label.getBoundingClientRect().right <= value.getBoundingClientRect().left &&
            [row, label, value].every(node => node.scrollWidth <= node.clientWidth + 1),
            "Ready resources lost native two-column layout or clipped long names/helpers.");
        return { label: label.getBoundingClientRect().left, value: value.getBoundingClientRect().left };
    });
    assert(positions.every(position => Math.abs(position.label - positions[0].label) <= 1 &&
        Math.abs(position.value - positions[0].value) <= 1), "Ready resource label/name columns did not share real alignment.");
    const nodes = [...ready.querySelectorAll<HTMLElement>(".iot-solution-architecture-node")];
    const expectedNodes = [
        ["Raw IoT tables", "Telemetry and property JSON"], ["Device model", "Field names, types, and units"],
        ...(hasProperties && twin > 0 ? [["Latest property values", "Each device's latest state"]] : []),
        ["Modeled data", twin === 0 ? "Typed telemetry" : telemetry > 0 ? "Typed telemetry with device state" : "Telemetry events with device state"],
    ];
    assert(JSON.stringify(nodes.map(node => [text(node.querySelector(".iot-solution-flow-node-title")),
        text(node.querySelector(".iot-solution-flow-node-detail"))])) === JSON.stringify(expectedNodes) &&
        ready.querySelectorAll('.iot-solution-flow-arrow[aria-hidden="true"] > svg').length === nodes.length - 1 &&
        [...ready.querySelectorAll(".iot-solution-flow-arrow")].every(arrow => !arrow.textContent?.trim()) &&
        nodes.every(node => node.scrollWidth <= node.clientWidth + 1) &&
        text(ready.querySelector(".iot-solution-finish-hint")) === "Select Finish and Close when you are done.",
        "Ready diagram changed conceptual ordering, fabricated state enrichment or lost native arrows/Finish guidance.");
    const vertical = innerWidth <= 700;
    assert(getComputedStyle(ready.querySelector(".iot-solution-architecture")!).flexDirection === (vertical ? "column" : "row") &&
        nodes.every((node, index) => index === 0 || (vertical
            ? node.getBoundingClientRect().top > nodes[index - 1].getBoundingClientRect().bottom
            : node.getBoundingClientRect().left > nodes[index - 1].getBoundingClientRect().right)),
        "Ready diagram did not retain its actual responsive native flow.");
    return { viewport: innerWidth, telemetry, twin, awaitingData: result.awaitingData, hasPropertyOutputs: hasProperties,
        headings, intro: `${intro} Query it by the field names and types from your device model.`,
        resources, positions, nodes: expectedNodes, verticalDiagram: vertical, disabledLearnLinks: 2, noOverflow: true };
}
const documentElementRole = (id: string) => [...area.querySelectorAll<HTMLElement>("[data-field-id]")]
    .find(element => element.dataset.fieldId === id)?.querySelector<HTMLButtonElement>('[role="combobox"]');
const controlDisabled = (control: Element | null | undefined) => !!control &&
    (control.matches(":disabled") || control.getAttribute("aria-disabled") === "true");
const activeOptions = (id: string) => {
    const control = area.querySelector(`#${id}`);
    const listbox = document.getElementById(control?.getAttribute("aria-controls") || "");
    assert(listbox, `The ${id} dropdown has no active listbox.`);
    return [...listbox!.querySelectorAll('[role="option"]')];
};
async function chooseOption(id: string, label: string) {
    const control = area.querySelector(`#${id}`);
    if (control?.getAttribute("aria-expanded") !== "true")
        await click(control);
    await click(activeOptions(id).find(option => option.textContent === label));
}
async function enter(id: string, value: string) {
    if (id.startsWith("field-role-")) {
        assert(value === "telemetry" || value === "twin", "Role dropdowns contain only actual role choices.");
        await chooseOption(id, value === "telemetry" ? "Telemetry" : "Twin Property");
        return;
    }
    const element = area.querySelector<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>(`#${id}`);
    assert(element, `Missing ${id}.`);
    await act(async () => {
        const prototype = element instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype
            : element instanceof HTMLSelectElement ? HTMLSelectElement.prototype : HTMLInputElement.prototype;
        Object.getOwnPropertyDescriptor(prototype, "value")!.set!.call(element, value);
        element.dispatchEvent(new Event(element instanceof HTMLSelectElement ? "change" : "input", { bubbles: true }));
    });
    await settle();
}
async function file(id: string, content: string, name = "model.json") {
    const transfer = new DataTransfer();
    transfer.items.add(new File([content], name, { type: "application/json" }));
    const element = area.querySelector<HTMLInputElement>(`#${id}`)!;
    await act(async () => { element.files = transfer.files; element.dispatchEvent(new Event("change", { bubbles: true })); });
    await settle();
}
async function unreadableFile(id: string, name: string) {
    const element = area.querySelector<HTMLInputElement>(`#${id}`)!;
    const candidate = new File(["{}"], name, { type: "application/json" });
    Object.defineProperty(candidate, "text", { value: async () => { throw new Error("Synthetic file read failure"); } });
    const original = Object.getOwnPropertyDescriptor(element, "files");
    Object.defineProperty(element, "files", { configurable: true, value: [candidate] });
    try {
        await act(async () => element.dispatchEvent(new Event("change", { bubbles: true })));
        await settle();
    }
    finally {
        if (original)
            Object.defineProperty(element, "files", original);
        else
            Reflect.deleteProperty(element, "files");
    }
}
const browserTabCalls: { url: string }[] = [];
const client = {
    auth: { acquireFrontendAccessToken: async (params: { scopes: string[] }) => {
            tokenRequests++;
            frontendScopeRequests.push([...params.scopes]);
            if (assignmentAuthFailure) throw { error: 1 };
            if (deferredToken)
                await new Promise<void>(resolve => tokenReleases.push(resolve));
            return { token: "mock-only" };
        }, acquireAccessToken: async () => {
            consentRequests++;
            assignmentAuthFailure = false; return { token: "mock-only" };
        } },
    theme: { get: async () => hostTheme(false), onChange: (callback: (value: ThemeConfiguration) => void) => { hostChange = callback; } },
    navigation: { openBrowserTab: async (params: { url: string }) => { browserTabCalls.push({ ...params }); },
        onBeforeNavigateAway: (callback: () => Promise<{
            canLeave: boolean;
        }>) => { leaveGuard = callback; } },
} as WorkloadClientAPI;
function documentationSideEffects() {
    return canonicalJson({ queries: calls.length, resources: resourceCalls.length, tokens: tokenRequests,
        consent: consentRequests, assignmentQueries, writes: creationBackend?.writes.length || 0,
        wizardContext: currentContext || {}, activeStep: area.querySelector('[aria-current="step"]')?.textContent,
        nextDisabled: next()?.disabled, finishDisabled: button("Finish and Close")?.disabled });
}
Reflect.set(window, "__documentationNavigationState", () => ({
    browserTabs: browserTabCalls.map(params => ({ ...params })), sideEffects: documentationSideEffects(),
}));
async function documentationClick(link: HTMLAnchorElement | null | undefined, href: string) {
    assert(link?.getAttribute("href") === href, `Missing exact documentation destination: ${href}`);
    const anchor = link!;
    const attributes = () => canonicalJson([anchor.tagName, anchor.textContent, anchor.getAttribute("href"),
        anchor.target, anchor.rel, anchor.getAttribute("style")]);
    const before = { calls: browserTabCalls.length, sideEffects: documentationSideEffects(), attributes: attributes() };
    let defaultPreventedByProduct = false;
    const observe = (event: MouseEvent) => {
        if (!(event.target instanceof Node) || !anchor.contains(event.target)) return;
        defaultPreventedByProduct = event.defaultPrevented;
        // Record product cancellation first; only suppress a broken link to prevent external test navigation.
        if (!event.defaultPrevented) event.preventDefault();
    };
    document.addEventListener("click", observe);
    try { await click(anchor); }
    finally { document.removeEventListener("click", observe); }
    const hostCalls = browserTabCalls.slice(before.calls);
    return { href, hostCalls, defaultPreventedByProduct,
        exactlyOneOpenBrowserTabWithExactUrl: canonicalJson(hostCalls) === canonicalJson([{ url: href }]),
        noWizardAuthDiscoveryAssignmentEnrollmentOrResourceEffects: documentationSideEffects() === before.sideEffects,
        nativeLinkAttributesAndCopyUnchanged: attributes() === before.attributes };
}
function assertDocumentationClicks(observations: Awaited<ReturnType<typeof documentationClick>>[]) {
    caseObservations = { documentationClicks: observations, externalNavigationNeverPerformedByTests: true };
    assert(observations.every(value => value.defaultPreventedByProduct && value.exactlyOneOpenBrowserTabWithExactUrl &&
        value.noWizardAuthDiscoveryAssignmentEnrollmentOrResourceEffects && value.nativeLinkAttributesAndCopyUnchanged),
        `Documentation must call navigation.openBrowserTab exactly once and cancel native navigation without side effects: ${JSON.stringify(observations)}`);
}
function hostTheme(dark: boolean): ThemeConfiguration {
    const theme = dark ? webDarkTheme : webLightTheme;
    return { name: dark ? "dark" : "light", colorScheme: dark ? "dark" : "light",
        accentColors: { 10: "#061724", 20: "#082338", 30: "#0A2E4A", 40: "#0C3B5E", 50: "#0E4775", 60: "#0F548C",
            70: "#115EA3", 80: "#0F6CBD", 90: "#2886DE", 100: "#479EF5", 110: "#62ABF5", 120: "#77B7F7", 130: "#96C6FA",
            140: "#B4D6FA", 150: "#CFE4FA", 160: "#EBF3FC" }, tokens: { ...theme } };
}
const dtdl = { "@id": "dtmi:example:Device;1", "@type": "Interface", displayName: "Device",
    contents: [{ "@type": "Telemetry", name: "temperature", schema: "double", unit: "degreeCelsius" },
        { "@type": "Property", name: "state", schema: "string" }] };
const tm = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel", id: "urn:example:device", title: "Device",
    properties: { temperature: { type: "number", description: "Ambient temperature.", unit: "degC" }, state: { type: "string" } } };
const coveragePreviewTelemetryTable = "TelemetryRawIoTHub_combined_hubs_for_coverage_with_a_long_table_name";
const coveragePreviewPropertiesTable = "PropertiesRawIoTHub_combined_hubs_for_coverage_with_a_long_table_name";
const headingPresentationDocument = { ...tm, properties: { ...tm.properties,
    environmentalComfortProfile: { type: "object", title: "Environmental comfort profile",
        properties: { setpoint: { type: "number" } } },
    predictiveMaintenanceAlert: { type: "boolean", title: "Predictive maintenance alert" },
} };
function context(wot = false): Record<string, any> {
    const model = wot ? fromThingModel(resolveThingModel(tm)) : fromDtdlModel(parseDtdlModel(dtdl), dtdl);
    return { modelRawDataSelected: true, modelEhId: "eh", modelEhName: "Eventhouse", modelDbId: "db", modelDbName: "Database",
        modelQueryServiceUri: "https://mock.kusto.fabric.microsoft.com", rawTelemetryTable: "Telemetry", rawPropertiesTable: "Twins",
        preparationFileName: "seed.json", hubModel: model, hubSelectedFormat: model.kind,
        hubFields: model.fields.map(field => ({ ...field, role: field.path[0] === "state" ? "twin" : "telemetry",
            roleSource: wot ? "manual" : "model" })), hubScope: { mode: "data", additional: "specified", operator: "and", telemetry: "", twin: "" } };
}
function coverageMissingFixture(): Record<string, any> {
    const telemetry = [["temperature", "Temperature"], ["humidity", "Humidity"], ["calibration", "Calibration"],
        ["pressure", "Pressure"], ["co2Level", "CO2 Level"], ["airQualityIndex", "Air Quality Index"],
        ["batteryVoltage", "Battery Voltage"], ["signalStrength", "Signal Strength"], ["uptime", "Uptime"],
        ["status", "Status"], ["operatingMode", "Operating Mode"], ["alertActive", "Alert Active"],
        ["timestamp", "Timestamp"], ["gpsLocation", "GPS Location"]];
    const properties = [["customerName", "Customer Name"], ["firmwareVersion", "Firmware Version"], ["location", "Location"],
        ["deviceModel", "Device Model"], ["serialNumber", "Serial Number"], ["isOnline", "Is Online"],
        ["lastCalibrated", "Last Calibrated"], ["installYear", "Install Year"], ["maintenanceSchedule", "Maintenance Schedule"],
        ["lastKnownTemperature", "Last Known Temperature"], ["lastKnownStatus", "Last Known Status"],
        ["lastReportedAt", "Last Reported At"], ["totalMessagesSent", "Total Messages Sent"],
        ["thermostat", "Thermostat"], ["powerModule", "Power Module"]];
    const document = { ...tm, id: "urn:example:ThermoProX200", title: "ThermoPro X200", properties:
        Object.fromEntries([...telemetry, ...properties].map(([name, title]) => [name, { type: "string", title }])) };
    const model = fromThingModel(resolveThingModel(document));
    const telemetryNames = new Set(telemetry.map(([name]) => name));
    return { ...context(true), hubModel: model, hubFields: model.fields.map(field =>
        setFieldRole(field, telemetryNames.has(field.path[0]) ? "telemetry" : "twin")) };
}
function result(columns: [
    string,
    string
][], rows: unknown[][]) {
    return new Response(JSON.stringify({ Tables: [{ Columns: columns.map(([ColumnName, ColumnType]) => ({ ColumnName, ColumnType })), Rows: rows }] }));
}
function entryMetadata(id: string) {
    return { item: { id, workspaceId: "workspace", type: "Local.MockOnly.IoTSolution",
            displayName: "Entry control", description: "", folderId: "", tags: [] } };
}
function emptyAppDefinition() {
    return { definition: { format: "", parts: [{ path: ".platform",
                    payload: encodeJsonToBase64({ metadata: { type: "Local.MockOnly.IoTSolution", displayName: "Entry control" } }),
                    payloadType: PayloadType.InlineBase64 }] } };
}
let deferredScoped = false;
let releases: (() => void)[] = [];
let sampleState: "normal" | "missing-identity" | "missing-hub" | "cross-hub" = "normal";
let creationBackend: ReturnType<typeof createCreationMock> | undefined;
let deferredCreate = false;
let creationReleases: (() => void)[] = [];
let emptyProcessing = false;
let emptyProcessingStream: "telemetry" | "twin" | undefined;
let coverageReadFailure: "telemetry" | "twin" | "both" | undefined;
let coverageReadFailureStatus = 403;
let coverageSampleRows = 1;
let managementStatusFailure = false;
let resourceResponse: (() => Promise<Response>) | undefined;
let tableMetadataResponse: (() => Promise<Response>) | undefined;
let eventhouseMetadataResponse: (() => Promise<Response>) | undefined;
let databaseItemIds = ["db"];
let emptySamples: "none" | "both" | "telemetry" | "twin" = "none";
let sourceSampleRows = 1;
let sourceSampleStatus = 0;
let deferSourceSamples = false;
const sourceSampleReleases: (() => void)[] = [];
const sourceSampleSignals: AbortSignal[] = [];
let parityPayloads: {
    telemetry: object;
    twin: object;
} | undefined;
const calls: string[] = [];
const resourceCalls: string[] = [];
const mockFetch: typeof fetch = async (url, options) => {
    const address = String(url);
    if (address.startsWith("https://management.azure.com/")) {
        if (address.includes("ResourceGraph")) assignmentQueries.push(JSON.parse(String(options?.body)).query);
        if (assignmentResponse) return assignmentResponse(url, options);
        assert(address === "https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01" &&
            options?.method === "POST" && String(options.body).includes("fixture-hub.azure-devices.net"),
            "A nonfixture ARM request was rejected.");
        return new Response(JSON.stringify({ data: [], resultTruncated: false }));
    }
    if (address.startsWith("http://ipc.localhost/"))
        return originalFetch(url, options);
    assert(address.startsWith("https://api.fabric.microsoft.com/") || address.startsWith("https://mock.kusto.fabric.microsoft.com/"),
        "A nonfixture Azure/Fabric/DPS request was rejected.");
    if (address.includes("/workspaces/"))
        resourceCalls.push(address);
    if (address.includes("/items"))
        return resourceResponse ? resourceResponse() : new Response(JSON.stringify({ value: [
                { id: "eh", type: "Eventhouse", displayName: "Eventhouse" }, { id: "db", type: "KQLDatabase", displayName: "Database" }
            ] }));
    if (address.includes("/eventhouses/eh"))
        return eventhouseMetadataResponse ? eventhouseMetadataResponse() : new Response(JSON.stringify({
            properties: { queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databasesItemIds: databaseItemIds }
        }));
    const csl = JSON.parse(String(options?.body)).csl as string;
    const coverageClock = csl.includes("__coverage_ingestion_time");
    const rawColumns: [string, string][] = [["data", "string"], ["headers", "dynamic"], ["user_headers", "dynamic"],
        ...(coverageClock ? [["__coverage_ingestion_time", "datetime"] as [string, string]] : [])];
    calls.push(csl);
    const sourceAnalysisRead = csl.includes("| top 1000 by ingestion_time() desc") && !csl.includes("__observation");
    if (sourceAnalysisRead && options?.signal) sourceSampleSignals.push(options.signal);
    if (sourceAnalysisRead && sourceSampleStatus) return new Response("Synthetic source-analysis request failure", { status: sourceSampleStatus });
    if (csl === ".show tables details | project TableName, DocString")
        return tableMetadataResponse ? tableMetadataResponse()
            : result([["TableName", "string"], ["DocString", "string"]], [
                ["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw\n@iot.meta source=fixture-hub.azure-devices.net"],
                ["Twins", "@iot.meta tableRole=iothub-properties-raw\n@iot.meta source=fixture-hub.azure-devices.net"], ["NewTelemetry", null],
            ]);
    if (creationBackend && !csl.includes("| top ")) {
        if (managementStatusFailure && csl.startsWith(".alter table ")) {
            return new Response(JSON.stringify({ Tables: [
                    { TableName: "PrimaryResult", Columns: [{ ColumnName: "TableName" }], Rows: [["Modeled_Output"]] },
                    { TableName: "QueryStatus", Columns: [{ ColumnName: "Severity" }, { ColumnName: "StatusCode" }], Rows: [[2, -1]] },
                ] }));
        }
        if (deferredCreate && csl.startsWith(".set ")) {
            await new Promise<void>(resolve => creationReleases.push(resolve));
        }
        const metadata = await creationBackend.service.request(csl, address.endsWith("/query") ? "query" : "mgmt");
        return new Response(JSON.stringify({ Tables: [metadata] }));
    }
    if (csl === ".show tables | project TableName")
        return result([["TableName", "string"]], [["Telemetry"], ["Twins"], ["NewTelemetry"]]);
    assert(csl.includes("| top "), "Only mocked, bounded reads are allowed in this test.");
    const twin = csl.includes("['Twins']") || csl.includes("['PropertiesRawIoTHub_from_canary']") ||
        csl.includes(`['${coveragePreviewPropertiesTable}']`);
    if ((coverageReadFailure === "both" || coverageReadFailure === (twin ? "twin" : "telemetry")) && csl.includes("__observation"))
        return new Response(JSON.stringify({ error: { code: coverageReadFailureStatus === 400 ? "SyntaxError" : "Forbidden",
            message: coverageReadFailureStatus === 400 ? "Synthetic current KQL syntax error." : "Synthetic current coverage denial." } }),
            { status: coverageReadFailureStatus });
    if (emptySamples === "both" || emptySamples === (twin ? "twin" : "telemetry")) {
        return result(rawColumns, []);
    }
    if ((emptyProcessing || emptyProcessingStream === (twin ? "twin" : "telemetry")) && csl.includes("__observation"))
        return result(rawColumns, []);
    const body = { schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
        event_time: "2026-10-01T12:00:00Z", event_kind: twin ? "lifecycle" : "measurement",
        payload: { value: parityPayloads ? twin ? parityPayloads.twin : parityPayloads.telemetry
                : twin ? { properties: { reported: { state: "idle" } } } : { temperature: 22 } } };
    const response = () => result(rawColumns,
        Array.from({ length: sourceAnalysisRead ? sourceSampleRows : coverageSampleRows }, () => [JSON.stringify(body), sampleState === "missing-identity" ? {} : {
                IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "", IoTSubject: ""
            },
            sampleState === "missing-hub" ? {} : { hubName: sampleState === "cross-hub" ? twin ? "twin-hub" : "telemetry-hub" : "mock-hub" },
            ...(coverageClock ? ["2026-10-02T12:00:00Z"] : [])]));
    if (sourceAnalysisRead && deferSourceSamples) return new Promise(resolve => sourceSampleReleases.push(() => resolve(response())));
    if (deferredScoped && csl.includes("filter-A") && csl.includes("__observation")) {
        // Deliberately ignore AbortSignal: the component's revision fence must still reject a late response.
        return new Promise(resolve => releases.push(() => resolve(response())));
    }
    return response();
};
async function mount(element: React.ReactNode, host = false) {
    await act(async () => { root?.unmount(); });
    root = createRoot(area);
    await act(async () => root!.render(host ? element : <FluentProvider theme={webLightTheme} style={{ height: "100%" }}>{element}</FluentProvider>));
    await settle();
}
const sourceFixture: SourceReadiness = { ready: true, reason: "Source tables selected.", key: "mock-source", revision: 1 };
let setTestSourceState: (state: SourceReadiness) => void = () => undefined;
type TestStep = React.ComponentType<WizardStepProps & {
    workloadClient: WorkloadClientAPI;
    workspaceId: string;
    sourceState?: SourceReadiness;
    getSourceState?: () => SourceReadiness;
}>;
function Harness({ component: Component, initial }: {
    component: TestStep;
    initial: Record<string, any>;
}) {
    const [state, setState] = React.useState(() => {
        if (Component !== HubModelReviewStep || !initial.hubCoverage || initial.hubStreamChecks !== undefined)
            return initial;
        const input = hubInput(initial, sourceFixture), coverage = initial.hubCoverage;
        let checks: StreamChecks = currentStreamChecks(input, 0);
        for (const stream of ["telemetry", "twin"] as const) {
            const started = beginStreamCheck(input, 0, stream, checks);
            checks = finishStreamCheck(input, started.checks, started.ticket, { result: {
                sample: coverage[stream], probe: coverage.sourceProbe?.[stream] || coverage[stream], checkedAt: coverage.checkedAt,
            } });
        }
        return { ...initial, hubStreamChecks: checks };
    });
    currentContext = state;
    const [sourceState, updateSourceState] = React.useState(sourceFixture);
    setTestSourceState = updateSourceState;
    const sourceRef = React.useRef(sourceState);
    sourceRef.current = sourceState;
    const [visible, setVisible] = React.useState(true);
    setStepVisible = setVisible;
    const updateContext = React.useCallback((key: string, value: unknown) => setState(previous => ({ ...previous, [key]: value })), []);
    return visible ? <Component workloadClient={client} workspaceId="workspace" wizardContext={state} updateContext={updateContext} sourceState={sourceState} getSourceState={() => sourceRef.current} step={{ id: "test", title: "Test", component: Component }} stepIndex={1} totalSteps={5} isFirstStep={false} isLastStep={false} onNext={async () => undefined} onPrevious={() => undefined} resetStepsFrom={() => undefined} onGoToStep={(id: string) => { currentContext.requestedStep = id; }}/>
        : <div data-testid="disposed-step">Review view disposed; journey context retained.</div>;
}
async function wizard(initial = context(), onBack: () => void = () => undefined, dark = false) {
    const content = <IoTSolutionItemModelWizard workloadClient={client} item={{ id: "mock", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>} initialContext={initial} onBack={onBack}/>;
    await mount(dark ? <FluentProvider theme={webDarkTheme} style={{ height: "100%" }}>{content}</FluentProvider> : content, dark);
    if (initial.modelRawDataSelected) {
        await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Current source inventory did not settle.");
        if (next()!.disabled) await click(sourceValidationAction());
        await until(() => !next()!.disabled, "Explicit current-source analysis did not enable continuation.");
        await click(next());
        await until(() => {
            const choice = area.querySelector<HTMLInputElement>('input[value="wot"]');
            return choice && !choice.disabled;
        }, "Source inventories did not enable model actions.");
    }
}
function AnalysisJourney({ initial }: { initial: Record<string, any> }) {
    const [open, setOpen] = React.useState(true);
    setStepVisible = setOpen;
    return open ? <IoTSolutionItemModelWizard workloadClient={client}
        item={{ id: "mock", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
        initialContext={initial} onBack={() => setOpen(false)}/>
        : <div>Analysis journey closed.</div>;
}
async function openAnalysisSource(initial = context(true)) {
    await mount(<AnalysisJourney initial={initial}/>);
    if (!initial.modelRawDataSelected) await click(next());
    await until(() => area.querySelector("h2")?.textContent === "Select your data source", "The source stage did not open.");
}
const analysisQueries = () => calls.filter(query => query.includes("| top 1000 by ingestion_time() desc") && !query.includes("__observation"));
function withAnalyzedSamples(initial: Record<string, any>): Record<string, any> {
    const samples = creationRequest().coverage.sourceProbe!;
    return { ...initial, hubSourceAnalysis: { key: modelSourceKey("workspace", initial), selectionRevision: 0,
        limit: SOURCE_ANALYSIS_LIMIT, telemetry: samples.telemetry, twin: samples.twin,
        roleEvidence: prepareRoleEvidence(samples.telemetry, samples.twin), analyzedAt: "2026-10-25T00:00:00Z" } };
}
async function observeWizardContext(run: () => Promise<void>) {
    const original = React.createElement;
    currentContext = {};
    observedWizardProps = undefined;
    const observer = new Proxy(original, {
        apply(target, thisArg, args) {
            const props = args[1];
            if (props && Object.prototype.hasOwnProperty.call(props, "wizardContext") && props.wizardContext) {
                currentContext = props.wizardContext;
                if (typeof props.onNext === "function") observedWizardProps = props;
            }
            return Reflect.apply(target, thisArg, args);
        },
    });
    assert(Reflect.set(React, "createElement", observer), "The test-only guarded context observer could not be installed.");
    try { await run(); }
    finally { Reflect.set(React, "createElement", original); }
}
async function reachScope() {
    if (area.querySelector("h2")?.textContent === "Add a device model") {
        assert(!next()!.disabled, "Combined model/inline field validation must pass before advancing.");
        await click(next());
    }
    if (area.querySelector("h2")?.textContent === "Assign the model to devices") {
        if (next()!.disabled && button("Skip assignment")) await click(button("Skip assignment"));
        else {
            await until(() => !next()!.disabled, "Current assignment lists did not settle.");
            await click(next());
        }
    }
    assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage", "The combined model flow did not reach Data Filters.");
}
async function reachReview() {
    await reachScope();
    await verifyCoverage();
    await until(() => next() && !next()!.disabled, "Both checks never enabled review.");
    await click(next());
    assert(area.querySelector("h2")?.textContent === "Create modeled data", "Review was not reached.");
}
async function verifyCoverage() {
    if (area.querySelector('input[type="radio"][value="specified"]') &&
        !area.querySelector('input[type="radio"][value="specified"]:checked, input[type="radio"][value="none"]:checked')) {
        await click(area.querySelector('input[type="radio"][value="specified"]'));
    }
    await click(button("Check coverage"));
    await until(() => currentContext.hubStreamChecks?.telemetry.phase === "checked" &&
        currentContext.hubStreamChecks?.twin.phase === "checked", "Both independent coverage checks did not finish.");
}
function assignmentContext() {
    const initial = { ...context(true), hubAssignmentSource: { key: sourceFixture.key,
        telemetry: { table: "Telemetry", docstring: "@iot.meta source=fixture-hub.azure-devices.net" },
        twin: { table: "Twins", docstring: "@iot.meta source=fixture-hub.azure-devices.net" } } };
    return { ...initial, hubAssignmentHubSummary: { owner: assignmentIntentOwner(initial, sourceFixture),
        value: { sourceSignature: canonicalJson(initial.hubAssignmentSource), hubs: assignmentInventory().hubs, issues: [], complete: true } } };
}
function withCurrentHubSummary(initial: Record<string, any>, hubs = assignmentInventory().hubs,
    issues: AssignmentInventory["issues"] = [], complete = true, sourceState = sourceFixture) {
    const source = initial.hubAssignmentSource || { ...assignmentContext().hubAssignmentSource, key: sourceState.key };
    const context = { ...initial, hubAssignmentSource: source };
    return { ...context, hubAssignmentHubSummary: { owner: assignmentIntentOwner(context, sourceState),
        value: { sourceSignature: canonicalJson(source), hubs, issues, complete } } };
}
const CoverageView = (props: React.ComponentProps<typeof HubModelScopeStep>) =>
    <HubModelScopeStep {...props} sourceRevision={props.sourceState?.revision || 0} />;
async function beginAssignmentTargets() {
    await until(() => !!button("Assign model to devices") && !button("Assign model to devices")!.disabled,
        "Positive current Hub eligibility did not enable the assignment action.");
    await click(button("Assign model to devices"));
}
async function openAssignmentTargets(initial = assignmentContext()) {
    await mount(<Harness component={HubModelAssignmentStep} initial={initial}/>);
    if (initial.hubModel?.kind === "wot") await beginAssignmentTargets();
}
function alignmentCreationContext(request: ReturnType<typeof creationRequest>) {
    const initial = { ...context(true), hubModel: request.input.model, hubFields: request.input.fields, hubScope: request.input.scope,
        hubCoverage: request.coverage, hubAllowWaiting: request.allowWaiting, hubOutputNames: request.names,
        hubSelectedFormat: request.input.model.kind };
    let checks = currentStreamChecks(request.input, 0);
    for (const stream of ["telemetry", "twin"] as const) {
        const started = beginStreamCheck(request.input, 0, stream, checks);
        checks = finishStreamCheck(request.input, started.checks, started.ticket, { result: {
            sample: request.coverage[stream], probe: request.coverage.sourceProbe![stream], checkedAt: request.coverage.checkedAt,
        } });
    }
    return { ...initial, hubStreamChecks: checks };
}
function assignmentInventory(): AssignmentInventory {
    const root = "/subscriptions/00000000-0000-0000-0000-000000000001/resourceGroups/fixture";
    const hub = `${root}/providers/Microsoft.Devices/IotHubs/fixture-hub`;
    const namespace = `${root}/providers/Microsoft.DeviceRegistry/namespaces/fixture-namespace`;
    return { hasMore: false, issues: [], hubs: [{ id: hub, name: "Fixture Hub", namespaceId: namespace, state: "linked" }],
        targets: [{ id: `${namespace}/registryDevices/fixture-device`, name: "Fixture device", kind: "device",
            hubIds: [hub], namespaceId: namespace, etag: '"synthetic-etag"',
            reference: { kind: "unknown", reason: "No reference returned." }, writeBlocker: "Safe updates are not verified." }] };
}
function adrOnlyResponse(requests: string[], assignment?: { modelId?: string; patches: { body: unknown; ifMatch: string | null }[] }): typeof fetch {
    const inventory = assignmentInventory(), hub = inventory.hubs[0], device = inventory.targets[0];
    return async (url, init) => {
        const address = String(url), path = new URL(address).pathname; requests.push(address);
        const reply = (body: unknown) => new Response(JSON.stringify(body), { status: 200, headers: { "Content-Type": "application/json" } });
        assert(address.startsWith("https://management.azure.com/"), "Only fixture ARM requests are allowed.");
        if (address.includes("ResourceGraph")) {
            const query = JSON.parse(String(init?.body)).query as string;
            assert(!query.includes("registrydevices"), "Device reference authority must not invoke a registry-device ARG query.");
            assert(query.includes("microsoft.devices/iothubs") && query.includes("fixture-hub.azure-devices.net"),
                "Unexpected DPS or unscoped discovery query.");
            return reply({ data: [{ id: hub.id, name: hub.name, hostName: "fixture-hub.azure-devices.net" }] });
        }
        if (path === hub.id) return reply({ id: hub.id, name: hub.name, properties: {
            hostName: "fixture-hub.azure-devices.net", deviceRegistry: { namespaceResourceId: hub.namespaceId, linkingProperties: { state: "Succeeded" } },
        } });
        if (path === hub.namespaceId) return reply({ id: hub.namespaceId, properties: { messaging: { endpoints: {
            "fixture-hub": { endpointType: "Microsoft.Devices/IotHubs", resourceId: hub.id, address: "fixture-hub.service.azure-devices.net" },
        } } } });
        if (path === `${hub.namespaceId}/registryDevices`) return reply({ value: [{ id: device.id }] });
        if (path === `${device.id}/capabilities`) return reply({ value: [{
            id: `${device.id}/capabilities/fixture-hub`, properties: { capabilityType: "Microsoft.IoTHub", messagingEndpointName: "fixture-hub" },
        }] });
        if (path === `${device.id}/providers/Microsoft.Authorization/permissions`) return reply({ value: [{ actions: ["*"], notActions: [] }] });
        if (path === device.id && init?.method === "PATCH") {
            assert(assignment && !assignment.modelId, "The isolated fixture must be authoritatively unset before a first assignment.");
            const body = JSON.parse(String(init.body));
            const ifMatch = new Headers(init.headers).get("If-Match");
            assert(ifMatch === '"fixture-current"' && JSON.stringify(body) === JSON.stringify({ properties: { deviceTypeRefs: [tm.id] } }),
                "The real adapter changed its exact conditional refs-only request.");
            assignment!.patches.push({ body, ifMatch }); assignment!.modelId = tm.id;
            return new Response("", { status: 202, headers: { "Azure-AsyncOperation":
                `https://management.azure.com/subscriptions/${device.id.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/australiaeast/asyncOperationStatuses/fixture-assignment?api-version=2026-11-02-preview` } });
        }
        if (path.includes("/asyncOperationStatuses/fixture-assignment")) return reply({ status: "Succeeded" });
        if (path === device.id && (!init?.method || init.method === "GET")) return reply({ id: device.id, name: device.name,
            etag: assignment?.modelId ? '"fixture-stored"' : '"fixture-current"',
            ...(assignment ? { type: "Microsoft.DeviceRegistry/namespaces/registryDevices" } : {}),
            properties: { externalDeviceId: "fixture-device", ...(assignment ? {
                uuid: "11111111-1111-4111-8111-111111111111", provisioningState: "Succeeded",
                ...(assignment.modelId ? { deviceTypeRefs: [assignment.modelId] } : {}),
            } : {}) } });
        throw new Error("Unexpected non-ADR fixture request.");
    };
}
type HubSummaryMode = "linked" | "mixed" | "omitted" | "null" | "pending" | "failed" | "authentication" | "permission" |
    "network" | "malformed" | "partial" | "truncated" | "missing";
function sourceHubFixture(initialMode: HubSummaryMode) {
    const inventory = assignmentInventory(), hub = inventory.hubs[0];
    const second = { id: hub.id + "-other", host: "other-hub.azure-devices.net", name: "Other source Hub" };
    const calls: { url: string; method: string; query?: string }[] = [], patches: { body: unknown; ifMatch: string | null }[] = [];
    const base = adrOnlyResponse([], { patches });
    const fixture = { mode: initialMode, calls, patches, hub, second,
        fetcher: (async (input, init) => {
            const url = String(input), address = new URL(url);
            const query = url.includes("ResourceGraph") ? JSON.parse(String(init?.body)).query as string : undefined;
            calls.push({ url, method: init?.method || "GET", query });
            if (query?.includes("microsoft.devices/iothubs")) {
                if (fixture.mode === "truncated") return new Response(JSON.stringify({ data: [], resultTruncated: true }));
                return new Response(JSON.stringify({ data: [
                    { id: hub.id, name: hub.name, hostName: "fixture-hub.azure-devices.net" },
                    ...(["mixed", "partial"].includes(fixture.mode) ? [{ id: second.id, name: second.name, hostName: second.host }] : []),
                ], resultTruncated: false }));
            }
            if (address.pathname === hub.id || address.pathname === second.id) {
                const other = address.pathname === second.id;
                if (fixture.mode === "network") throw new Error("Owned synthetic transport failure.");
                if (fixture.mode === "authentication" || fixture.mode === "permission" || fixture.mode === "partial" && other) {
                    return new Response("Owned unavailable Hub metadata", { status: fixture.mode === "authentication" ? 401 : 403 });
                }
                const registry = fixture.mode === "malformed" ? { namespaceResourceId: "invalid" }
                    : other || fixture.mode === "null" ? null
                    : { namespaceResourceId: hub.namespaceId, linkingProperties: {
                        state: fixture.mode === "pending" ? "Pending" : fixture.mode === "failed" ? "Failed" : "Succeeded",
                    } };
                return new Response(JSON.stringify({ id: other ? second.id : hub.id, name: other ? second.name : hub.name,
                    properties: { hostName: other ? second.host : "fixture-hub.azure-devices.net",
                        ...(fixture.mode === "omitted" ? {} : { deviceRegistry: registry }) } }));
            }
            return base(input, init);
        }) as typeof fetch,
    };
    tableMetadataResponse = () => {
        const provenance = fixture.mode === "missing" ? "" : `\n@iot.meta source=fixture-hub.azure-devices.net${
            ["mixed", "partial"].includes(fixture.mode) ? "; other-hub.azure-devices.net" : ""}`;
        return result([["TableName", "string"], ["DocString", "string"]], [
            ["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw" + provenance],
            ["Twins", "@iot.meta tableRole=iothub-properties-raw" + provenance], ["NewTelemetry", null],
        ]);
    };
    assignmentResponse = fixture.fetcher;
    return fixture;
}
const deviceDiscoveryCalls = (calls: { url: string; query?: string }[]) => calls.filter(call =>
    /registryDevices|capabilities|Authorization\/permissions/.test(call.url) || call.query?.includes("registrydevices"));

function deviceListInventory() {
    const value = assignmentInventory();
    value.targets[0] = { ...value.targets[0], writeBlocker: undefined, reference: { kind: "unassigned", ids: [] } };
    value.targets.push({ ...value.targets[0], id: value.targets[0].id + "-assigned",
        name: "Existing-reference-device-with-a-long-name-that-wraps-without-truncation",
        reference: { kind: "assigned", ids: [`https://models.example.test/${"ExactReference-".repeat(12)}V1`] } });
    value.targets.push({ ...value.targets[0], id: value.targets[0].id + "-unknown", name: "Unknown reference",
        reference: { kind: "unknown", reason: "Fixture unknown" } });
    value.hubs.push({ id: value.hubs[0].id + "-other", name: "Other source Hub", state: "unlinked" });
    value.targets.push({ ...value.targets[0], id: value.targets[0].id + "-unlinked", name: "Read-only source device",
        hubIds: [value.hubs[1].id], writeBlocker: "The source IoT Hub is not linked to an ADR namespace." });
    value.targets.push({ ...value.targets[0], id: value.targets[0].id + "-current", name: "Already assigned device",
        reference: { kind: "assigned", ids: [tm.id] } });
    value.targets.push({ ...value.targets[0], id: value.targets[0].id.toUpperCase() });
    return value;
}

async function assignmentConfirmationCheck(captureOnly = false) {
    let result: Record<string, unknown> = {};
    await observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover, hubs = HubAssignmentApi.prototype.discoverHubs,
            services = HubAssignmentApi.prototype.assignmentServices;
        const inventory = assignmentInventory(), writes: string[] = [];
        inventory.hubs.push({ ...inventory.hubs[0], id: inventory.hubs[0].id + "-secondary",
            name: "Secondary source Hub with a longer name" });
        inventory.targets = [0, 1].map((index): AssignmentTarget => ({ ...inventory.targets[0], id: inventory.targets[0].id + "-" + index,
            name: index ? "Calibration device with a longer name for confirmation review" : "Calibration device 01",
            hubIds: index ? inventory.hubs.map(hub => hub.id) : [inventory.hubs[0].id],
            writeBlocker: undefined, reference: { kind: "unassigned", ids: [] } }));
        const model = fromThingModel(resolveThingModel({ ...tm,
            title: "Calibration model with a longer friendly name for confirmation review",
            id: `urn:fixture:${"calibration-reference-".repeat(7)}v1` }));
        try {
            HubAssignmentApi.prototype.discoverHubs = async source => ({
                sourceSignature: canonicalJson(source), hubs: inventory.hubs, issues: [], complete: true,
            });
            HubAssignmentApi.prototype.discover = async () => inventory;
            HubAssignmentApi.prototype.assignmentServices = () => ({
                read: async target => target,
                write: async (target, id, _signal, dispatch) => {
                    assert(id === model.id, "Confirmation changed the exact assignment reference.");
                    dispatch(); writes.push(target.id); return "completed";
                },
            });
            await wizard({ ...context(true), hubModel: model, hubFields: model.fields.map(field =>
                setFieldRole(field, field.path[0] === "state" ? "twin" : "telemetry")), preparationFileName: "confirmation-model.json" });
            await click(next()); await beginAssignmentTargets();
            await until(() => !!button("Select all") && !button("Select all")!.disabled, "Confirmation candidates did not settle.");
            await click(button("Select all")); await click(button("Assign to selected devices"));
            const dialog = () => document.querySelector<HTMLElement>('[role="dialog"]')!;
            const assertReview = (selected: AssignmentTarget[]) => {
                const surface = dialog(), title = surface.querySelector(".fui-DialogTitle")!;
                assert(title.textContent === `Assign ${model.title} to ${selected.length} ${selected.length === 1 ? "device" : "devices"}?` &&
                    surface.querySelector(".fui-Text")?.textContent === `Model ID: ${model.id}` &&
                    JSON.stringify([...surface.querySelectorAll("li")].map(node => node.textContent?.trim())) === JSON.stringify(
                        selected.map(target => `${target.name} \u00b7 ${target.hubIds.map(id => inventory.hubs.find(hub => hub.id === id)!.name).join(", ")}`)) &&
                    JSON.stringify([...surface.querySelectorAll("button")].map(node => node.textContent)) === JSON.stringify(["Cancel", "Assign"]) &&
                    !surface.textContent?.includes("(Devices)") && !surface.textContent?.includes(inventory.targets[0].namespaceId),
                    "Confirmation lost its friendly-name/count title, exact ID, device/Hub provenance or native actions.");
                assert(surface.scrollWidth <= surface.clientWidth + 1 &&
                    [...surface.querySelectorAll<HTMLElement>(".fui-DialogTitle,.fui-Text,li")].every(node => node.scrollWidth <= node.clientWidth + 1),
                    "Long confirmation model/device/Hub names or the exact model ID overflow.");
            };
            assertReview(inventory.targets);
            assert(writes.length === 0, "Opening the two-device review dispatched an assignment.");
            if (captureOnly) { Reflect.set(window, "__confirmationPreviewWrites", () => writes.length); return; }
            await click([...dialog().querySelectorAll("button")].find(node => node.textContent === "Cancel"));
            assert(writes.length === 0 && document.activeElement === button("Assign to selected devices"),
                "Cancel dispatched a mutation or lost its native return-focus target.");
            await click(area.querySelector('input[aria-label="Select Calibration device 01"]'));
            await click(button("Assign to selected devices"));
            assertReview([inventory.targets[1]]);
            await click([...dialog().querySelectorAll("button")].find(node => node.textContent === "Assign"));
            await until(() => !!completedAssignmentList(), "The explicit one-device confirmation did not complete.");
            assertAssignmentSummary("confirmed", 1);
            assert(writes.length === 1 && writes[0] === inventory.targets[1].id &&
                currentContext.hubAssignmentHistory.filter((receipt: AssignmentReceipt) => receipt.state === "confirmed").length === 1,
                "Presentation changes changed the confirmed target, dispatched on Cancel or replayed a write.");
            result = { friendlyModelName: model.title, modelReference: model.id, pluralAndSingularTitles: true,
                exactDeviceAndMultipleHubRows: true, cancelWrites: 0, cancelFocusRestored: true, explicitSyntheticWrites: 1, actualCloudWrites: 0 };
        } finally {
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.discoverHubs = hubs;
            HubAssignmentApi.prototype.assignmentServices = services;
        }
    });
    return result;
}

async function creationPresentationCheck(stage: "form" | "complete" | "test" | "ready") {
    let result: Record<string, unknown> = {};
    await observeWizardContext(async () => {
        await wizard(context(true)); await reachReview();
        if (stage === "ready") {
            await enter("modeled-output-name", "ModeledData_Devices_With_A_Long_Verified_Resource_Name_For_The_Ready_Results_Table");
            await enter("twin-changes-name", "NormalizedProperties_Devices_With_A_Long_Verified_Resource_Name_For_The_Ready_Results_Table");
            await enter("twin-state-name", "LatestProperties_Devices_With_A_Long_Verified_Resource_Name_For_The_Ready_Results_Table");
        }
        const source = (observedWizardProps as React.ComponentProps<typeof HubModelReviewStep>).sourceState;
        const request = { input: hubInput(currentContext, source), coverage: currentContext.hubCoverage,
            names: currentContext.hubOutputNames, allowWaiting: true };
        const plan = buildModeledDataPlan(request.input, request.coverage, true, request.names);
        const immutableCommands = canonicalJson(plan.commands);
        assert(JSON.stringify([...area.querySelectorAll(".iot-solution-form input[required]")].map(node => node.id)) ===
            JSON.stringify(["modeled-output-name", "twin-changes-name", "twin-state-name"]) &&
            area.querySelector("#creation-resources-heading")?.textContent === "Resources to create" &&
            area.querySelector("#output-names-help")?.textContent === "You can rename these resources before you create them" &&
            next()!.textContent === "Continue" && next()!.disabled,
            "The rendered resource order/copy changed names, required state or shared navigation.");
        const reads = calls.length, tokens = tokenRequests, contextFingerprint = canonicalJson({
            fields: currentContext.hubFields, scope: currentContext.hubScope, names: currentContext.hubOutputNames });
        await settle();
        assert(calls.length === reads && tokenRequests === tokens && contextFingerprint === canonicalJson({
            fields: currentContext.hubFields, scope: currentContext.hubScope, names: currentContext.hubOutputNames }),
            "Presentation dispatched work or mutated source/fields/names.");
        if (stage === "form") return;
        creationBackend = createCreationMock(request);
        await click(button("Create modeled data"));
        await until(() => currentContext.hubCreation?.phase === "completed" && !next()!.disabled,
            "Guarded local creation did not reach confirmed completion.");
        assertCreationDisplayVerbs(plan);
        assert(JSON.stringify(creationBackend.writes) === JSON.stringify(plan.commands.map(command => command.csl)) &&
            canonicalJson(buildModeledDataPlan(request.input, request.coverage, true, request.names).commands) === immutableCommands &&
            plan.commands.filter(command => command.kind === "table" || command.kind === "view").map(command => command.target).join("|") ===
                [request.names.twinChanges, request.names.twinState, request.names.modeled].join("|") &&
            area.querySelector("#modeled-data-success")?.textContent ===
                "Modeled data created. Select Continue to see what was created and what to do next." &&
            next()!.textContent === "Continue",
            "Input/display reordering changed actual command sequence, SQL, dependencies or navigation.");
        result = { renderedInputOrder: ["modeled", "twinChanges", "twinState"],
            actualResourceCreationOrder: [request.names.twinChanges, request.names.twinState, request.names.modeled],
            exactCommandsUnchangedAndTransportMatched: true, commandCount: creationBackend.writes.length,
            confirmedOnlyPastTense: true, footerContinueUnchanged: true, actualCloudWrites: 0 };
        if (stage === "ready") {
            const before = { calls: calls.length, resourceCalls: resourceCalls.length, tokens: tokenRequests,
                writes: creationBackend.writes.length, context: canonicalJson(currentContext) };
            await click(next());
            const presentation = assertReadyPresentation(request.input, currentContext.hubCreation);
            const documentationClicks = [];
            for (const link of area.querySelectorAll<HTMLAnchorElement>(".iot-solution-model-ready a.fui-Link"))
                documentationClicks.push(await documentationClick(link, link.getAttribute("href")!));
            assertDocumentationClicks(documentationClicks);
            assert(!button("Finish and Close")!.disabled && calls.length === before.calls &&
                resourceCalls.length === before.resourceCalls && tokenRequests === before.tokens &&
                creationBackend.writes.length === before.writes && canonicalJson(currentContext) === before.context,
                "Ready documentation interaction performed work, changed the result or altered the verified Finish gate.");
            result = { ...result, ready: presentation, documentationClicks, readyReadsOrAuthOrWrites: 0, verifiedFinishUnchanged: true };
        }
    });
    return result;
}

async function assignmentSpacingPresentation(outcome: "active" | "confirmed" | "partial" | "unknown") {
    const discover = HubAssignmentApi.prototype.discover, services = HubAssignmentApi.prototype.assignmentServices,
        readback = HubAssignmentApi.prototype.readAssignment, clock = assignmentClock();
    const value = assignmentInventory(), writes: string[] = [], statusReads: string[] = [];
    value.targets = [0, 1, 2].map(index => ({ ...value.targets[0], id: value.targets[0].id + `-${index}`,
        name: index < 2 ? `ThermoPro device ${index + 1}` :
            "Assembly line environment sensor with a deliberately long device label that wraps across multiple lines on narrow displays",
        writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
        deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: `device-${index}` } }));
    try {
        sourceHubFixture("linked");
        HubAssignmentApi.prototype.discover = async () => value;
        HubAssignmentApi.prototype.assignmentServices = () => ({
            read: async target => target, write: async (target, _id, _signal, dispatch) => {
                dispatch(); writes.push(target.id);
                if (target.id.endsWith("-2") && (outcome === "partial" || outcome === "unknown"))
                    throw Object.assign(new Error("Synthetic device assignment could not be confirmed."), { status: outcome === "partial" ? 403 : 503 });
                return outcome === "partial" || outcome === "unknown" || target.id.endsWith("-0") ? "completed" : "accepted";
            },
        });
        HubAssignmentApi.prototype.readAssignment = async target => {
            statusReads.push(target.id);
            return { target: { ...target, reference: { kind: "assigned", ids: [tm.id] } }, confirmed: true };
        };
        let geometry: ReturnType<typeof assignmentProgressGeometry> | undefined;
        await observeWizardContext(async () => {
            await wizard(context(true)); await click(next()); await beginAssignmentTargets();
            await until(() => !!button("Select all") && !button("Select all")!.disabled, "Spacing fixture candidates did not settle.");
            await click(button("Select all")); await click(button("Assign to selected devices"));
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(button => button.textContent === "Assign"));
            if (outcome === "active" || outcome === "confirmed") {
                await until(() => writes.length === 3 && clock.pending.size === 1, "Spacing fixture did not reach active observation.");
                if (outcome === "confirmed") {
                    await clock.tick();
                    await until(() => !!completedAssignmentList(), "Spacing fixture did not reach confirmed completion.");
                }
            } else await until(() => writes.length === 3 && !currentContext.hubAssignmentBusy &&
                area.querySelector('[aria-label="Device assignment progress"]')?.getAttribute("aria-busy") === "false",
                "Spacing fixture did not settle its partial/unknown outcome.");
            assertAssignmentSummary(outcome === "active" ? "active" : outcome === "confirmed" ? "confirmed" : "incomplete", 3);
            geometry = assignmentProgressGeometry();
            assert(writes.length === 3 && (outcome === "confirmed" ? statusReads.length === 2 : statusReads.length === 0) &&
                (outcome !== "partial" || area.querySelector('[data-assignment-state="failed"]')) &&
                (outcome !== "unknown" || area.querySelector('[data-assignment-state="unknown"]') && !button("Refresh results")) &&
                (outcome === "active" ? next()!.disabled : !next()!.disabled),
                "Assignment presentation changed current outcome, guarded continuation, request count or retry behavior.");
        });
        return { outcome, geometry, syntheticWrites: writes.length, statusReads: statusReads.length, actualCloudWrites: 0 };
    } finally {
        clock.restore(); HubAssignmentApi.prototype.discover = discover;
        HubAssignmentApi.prototype.assignmentServices = services; HubAssignmentApi.prototype.readAssignment = readback;
    }
}

async function assignmentWarningPresentation(dark = false) {
    const discover = HubAssignmentApi.prototype.discover, hubs = HubAssignmentApi.prototype.discoverHubs;
    const recheck = HubAssignmentApi.prototype.recheck, clock = assignmentWaveClock(), originalRandom = Math.random;
    const inventory = assignmentInventory(), writes: string[] = [], polls: string[] = [];
    Math.random = () => 0.5;
    inventory.targets = [0, 1, 2].map(index => ({ ...inventory.targets[0], id: inventory.targets[0].id + `-warning-${index}`,
        name: ["Confirmed device", "Failed device",
            "Timeout device with a longer name that wraps without clipping the warning icon or original explanation"][index],
        updatePermission: "granted", writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
        deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: `warning-${index}` } }));
    try {
        HubAssignmentApi.prototype.discoverHubs = async source => ({
            sourceSignature: canonicalJson(source), hubs: inventory.hubs, issues: [], complete: true,
        });
        HubAssignmentApi.prototype.discover = async () => inventory;
        HubAssignmentApi.prototype.recheck = async target => target;
        assignmentResponse = async (input, init) => {
            const address = new URL(String(input));
            if (init?.method === "PATCH") {
                const index = inventory.targets.findIndex(target => target.id === address.pathname);
                assert(index >= 0 && new Headers(init.headers).get("If-Match") === inventory.targets[index].etag &&
                    JSON.stringify(JSON.parse(String(init.body))) === JSON.stringify({ properties: { deviceTypeRefs: [tm.id] } }),
                    "The warning fixture changed the conditional refs-only assignment.");
                writes.push(address.pathname);
                return new Response(null, { status: 202, headers: { Location:
                    `https://management.azure.com/subscriptions/${inventory.targets[index].id.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/warning-${index}?api-version=2026-11-02-preview` } });
            }
            assert(address.pathname.includes("/operationStatuses/warning-"), "Warning presentation read mutable device state.");
            polls.push(address.pathname);
            return new Response(JSON.stringify({ status: address.pathname.endsWith("-0") ? "Succeeded"
                : address.pathname.endsWith("-1") ? "Failed" : "Running" }));
        };
        let proof: Record<string, unknown> = {};
        await observeWizardContext(async () => {
            await wizard(context(true), () => undefined, dark); await click(next()); await beginAssignmentTargets();
            await until(() => !!button("Select all") && !button("Select all")!.disabled, "Warning fixture devices did not settle.");
            await click(button("Select all")); await click(button("Assign to selected devices"));
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(button => button.textContent === "Assign"));
            await until(() => writes.length === 3 && currentContext.hubAssignmentBusy, "The warning fixture did not start owned observations.");
            await clock.advance(29999);
            assert(next()!.disabled && currentContext.hubAssignmentBusy, "An amber pending outcome released Continue before all work settled.");
            await clock.advance(1);
            await until(() => !currentContext.hubAssignmentBusy && !next()!.disabled && clock.pending.size === 0,
                "The actual original polling cutoff did not settle the warning fixture.");
            const timeout = area.querySelector<HTMLElement>('[data-assignment-state="unknown"]')!;
            assert(timeout?.textContent?.includes(ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE) &&
                area.querySelectorAll('[data-assignment-state="confirmed"]').length === 1 &&
                area.querySelectorAll('[data-assignment-state="failed"]').length === 1 &&
                !completedAssignmentList() && !area.querySelector("[data-assignment-summary]") &&
                confirmedAssignmentCount(currentContext.hubAssignmentIntent, currentContext.hubAssignmentHistory, tm.id) === 0 &&
                area.textContent?.includes("Not all assignments are confirmed. Review the device results before continuing.") &&
                !button("Refresh results") && writes.length === 3,
                "Warning styling changed receipt/outcome gating, unconfirmed guidance, recovery or mutation counts.");
            const styles = assertAssignmentSemanticIndicators(), history = canonicalJson(currentContext.hubAssignmentHistory), readCount = polls.length;
            await settle();
            assert(canonicalJson(currentContext.hubAssignmentHistory) === history && polls.length === readCount && writes.length === 3,
                "Rendering the warning modified receipts or restarted native operation reads/writes.");
            proof = { theme: dark ? "dark" : "light", ...styles, exactTimeoutMessage: ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE,
                states: ["confirmed", "failed", "unknown"], unchangedNativeState: true,
                observationSeconds: 30, mutations: writes.length, ownOperationReads: polls.length, pendingTimers: 0,
                unknownStillNonretryable: timeout.dataset.assignmentState === "unknown" &&
                    currentContext.hubAssignmentHistory.some((receipt: AssignmentReceipt) =>
                        receipt.state === "unknown" && receipt.readbackFailure?.retryable === false),
                noSuccessAggregate: true, noRefreshOrReplay: true, wholeAttemptStillNotConfirmed: true, actualCloudWrites: 0 };
        });
        return proof;
    } finally {
        clock.restore(); Math.random = originalRandom;
        HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.discoverHubs = hubs;
        HubAssignmentApi.prototype.recheck = recheck; assignmentResponse = undefined;
    }
}

async function coverageResultFollowupPresentation(mode: "reference" | "data" | "skipped" | "dtdl") {
    const discover = HubAssignmentApi.prototype.discover;
    try {
        const fixture = sourceHubFixture("linked"), requireModelId = mode === "reference", skipped = mode === "skipped", dtdlMode = mode === "dtdl";
        HubAssignmentApi.prototype.discover = async () => assignmentInventory();
        const model = fromThingModel(resolveThingModel({ ...tm, properties: { ...tm.properties,
            predictiveMaintenanceAlert: { type: "boolean", title: "Predictive maintenance alert" },
        } }));
        const seed = dtdlMode ? { ...context(), hubScope: undefined } :
            { ...context(true), hubModel: model, hubFields: model.fields.map(field =>
                setFieldRole(field, field.path[0] === "temperature" ? "telemetry" : "twin")), hubScope: undefined };
        parityPayloads = { telemetry: { temperature: 22 }, twin: { properties: { reported: { state: "running" } } } };
        emptyProcessingStream = "telemetry"; coverageSampleRows = 100;
        let proof: Record<string, unknown> = {};
        await observeWizardContext(async () => {
            await wizard(seed); await click(next());
            if (skipped) await click(button("Skip assignment"));
            else if (dtdlMode) await click(next());
            else {
                await beginAssignmentTargets();
                await until(() => !next()!.disabled, "Coverage followup assignment inventory did not settle.");
                await click(next());
                await click(area.querySelector(`input[value="${requireModelId ? "and" : "data"}"]`));
            }
            await click(area.querySelector('input[value="specified"]'));
            if (!requireModelId) await enter("hub-telemetry-scope", "false");
            await verifyCoverage();
            assertCoverageMissingGuidance();
            const cards = [...area.querySelectorAll<HTMLElement>("[data-coverage-stream]")];
            const notices = area.querySelectorAll(".iot-solution-coverage-assignment-hint");
            assert(coverageSampleCount(cards[0]) === 0 && coverageSampleCount(cards[1]) === 100 &&
                cards[1].querySelectorAll("[data-missing-field]").length === (dtdlMode ? 0 : 1) &&
                notices.length === (requireModelId ? 1 : 0) && !currentContext.hubAssignmentHistory?.length &&
                !next()!.disabled,
                `A current0/100 result lost reassurance/order or changed assignment-notice filtering eligibility: ${JSON.stringify({
                    counts: cards.map(coverageSampleCount), hints: notices.length, mode: currentContext.hubScope.mode, nextDisabled: next()!.disabled,
                })}`);
            if (skipped) assertSkippedFilteringNotice();
            if (dtdlMode) assertDtdlFilteringNotice();
            proof = { mode: currentContext.hubScope.mode, counts: [0, 100], selectedFields: [1, dtdlMode ? 1 : 2],
                missing: [0, dtdlMode ? 0 : 1], priorAssignmentReceiptCount: 0, assignmentHintCount: notices.length,
                explicitSkipNotice: skipped, dtdlNotice: dtdlMode, skipHubReads: skipped || dtdlMode ? fixture.calls.length : undefined,
                cardGeometry: coverageCardGeometry(), actualCloudWrites: 0 };
        });
        return proof;
    } finally { HubAssignmentApi.prototype.discover = discover; }
}
function assertSkippedFilteringNotice() {
    const notice = area.querySelector<HTMLElement>('[aria-labelledby="model-reference-filter-heading"] .iot-solution-assignment-notice')!;
    const body = notice?.querySelector(".fui-MessageBarBody"), type = body && getComputedStyle(body);
    assert(notice?.textContent ===
        "Model ID filtering isn't available because model assignment was skipped. To enable model ID filtering, go back and assign the model to your devices." &&
        body && !body.querySelector(".fui-Text") && !notice.querySelector("button,a,[href]") &&
        notice.scrollWidth <= notice.clientWidth + 1,
        "An actual skipped-assignment state lost exact copy, original native body rendering or introduced navigation.");
    return { family: type!.fontFamily, size: type!.fontSize, weight: type!.fontWeight,
        lineHeight: type!.lineHeight, style: type!.fontStyle, color: type!.color, padding: getComputedStyle(notice).padding };
}
function assertDtdlFilteringNotice() {
    const notice = area.querySelector<HTMLElement>('[aria-labelledby="model-reference-filter-heading"] .iot-solution-assignment-notice')!;
    const body = notice?.querySelector(".fui-MessageBarBody"), type = body && getComputedStyle(body);
    assert(notice?.textContent ===
        "Model ID filtering isn't available for DTDL models. To enable model ID filtering, go back and upload a WoT Thing Model." &&
        body && !body.querySelector(".fui-Text") && !notice.querySelector("button,a,[href]") &&
        !area.querySelector(".iot-solution-coverage-skipped-assignment") && notice.scrollWidth <= notice.clientWidth + 1,
        "The DTDL-only wording changed its native inherited rendering, distinct cause or actions.");
    return { family: type!.fontFamily, size: type!.fontSize, weight: type!.fontWeight,
        lineHeight: type!.lineHeight, style: type!.fontStyle, color: type!.color, padding: getComputedStyle(notice).padding };
}

const diagnosticWarningDocument = { ...tm, id: "urn:fixture:diagnostic-warning",
    properties: { ...tm.properties,
        calibration: { type: "number", title: "Calibration value with a long friendly field name that needs an explicit role" },
        null_sensor_configuration_with_a_long_name_for_wrapping: { type: null },
        conditional_sensor_value: { oneOf: [{ type: "number" }, { type: "string" }] } },
    links: [{ rel: "tm:submodel", href: "https://models.example.test/unused-child",
        instanceName: "Auxiliary device module with a longer name" }] };
async function assertDtdlRoleInformation(leaveOpen = false) {
    await until(() => [...area.querySelectorAll<HTMLButtonElement>('button[aria-label="About DTDL roles"]')]
        .some(node => node.getBoundingClientRect().width > 0),
        "The accepted DTDL fields did not finish loading their native role information.");
    const before = JSON.stringify(currentContext.hubFields), reads = calls.length + resourceCalls.length, tokens = tokenRequests;
    const trigger = [...area.querySelectorAll<HTMLButtonElement>('button[aria-label="About DTDL roles"]')]
        .find(node => node.getBoundingClientRect().width > 0 && node.getBoundingClientRect().height > 0)!;
    assert(trigger && !trigger.disabled && !area.querySelector('[role="combobox"][id^="field-role-"]') &&
        !area.querySelector('[data-selection-needed="true"]') && !area.querySelector("#model-upload-error"),
        "Normal DTDL roles are editable, treated as a validation error or lack native column information.");
    await act(async () => trigger.focus());
    await until(() => (trigger.getAttribute("aria-describedby") || "").split(" ").some(id =>
        document.getElementById(id)?.textContent === "Roles come from your DTDL model"),
        "Keyboard focus did not expose the native DTDL role description.");
    assert(document.activeElement === trigger && JSON.stringify(currentContext.hubFields) === before &&
        calls.length + resourceCalls.length === reads && tokenRequests === tokens,
        "Role information changed declared roles or triggered source/auth work.");
    const offsets = [...area.querySelectorAll<HTMLElement>("[data-field-id]")].map(row => {
        const headers = innerWidth <= 620 ? row : area.querySelector(".iot-solution-model-field-columns")!;
        return [
            ["field", row.querySelector(".iot-solution-model-field-heading")!],
            ["type", row.querySelector(".iot-solution-model-field-metadata > .fui-Text")!],
            ["role", row.querySelector(".iot-solution-model-field-role > .fui-Text:not(.iot-solution-model-field-mobile-label)")!],
        ].map(([column, value]) => headers.querySelector(`[data-field-column="${column}"]`)!.getBoundingClientRect().left -
            (value as Element).getBoundingClientRect().left);
    });
    assert(offsets.every(values => values.every(value => Math.abs(value) < 1)),
        `Native role information shifted the actual Field/Type/Role columns: ${JSON.stringify(offsets)}`);
    document.documentElement.dataset.roleInfoGeometry = JSON.stringify({ viewport: innerWidth, offsets });
    if (!leaveOpen) await act(async () => trigger.dispatchEvent(new KeyboardEvent("keydown", {
        key: "Escape", code: "Escape", bubbles: true, cancelable: true })));
}

const calibrationLayoutDocument = { ...tm, properties: {
    calibration: { title: "Calibration", description: "Current calibration.", readOnly: true, observable: false,
        default: 0, minimum: 0, maximum: 10, multipleOf: 0.5, unit: "mV", type: "number" },
    settings: { title: "Calibration settings", description: "Current calibration settings with a longer explanation that wraps naturally without being pushed down by a multiline role hint.",
        readOnly: false, default: null, const: { offset: 0, label: "", enabled: false }, type: "object",
        properties: { offset: { type: "number", minimum: 0 } }, required: ["offset"] },
    signalStrength: { title: "Signal strength", description: "Received signal strength.", type: "integer", unit: "dBm" },
} };
async function calibrationRowLayoutCheck() {
    let result: Record<string, unknown> = {};
    await observeWizardContext(async () => {
        const model = fromThingModel(resolveThingModel(calibrationLayoutDocument));
        const seed = withAnalyzedSamples({ ...context(true), hubModel: model, hubFields: model.fields, preparationFileName: "calibration.json" });
        const startReads = analysisQueries().length;
        await wizard(seed);
        assertFieldSelectionCopy();
        const rows = () => [...area.querySelectorAll<HTMLElement>("[data-field-id]")];
        const snapshots = (state: string) => rows().map(row => {
            const title = row.querySelector<HTMLElement>(".iot-solution-model-field-heading")!;
            const description = row.querySelector<HTMLElement>(".iot-solution-model-field-description")!;
            const metadata = row.querySelector<HTMLElement>(".iot-solution-model-field-metadata")!;
            const role = row.querySelector<HTMLElement>(".iot-solution-model-field-role")!;
            const affordances = row.querySelector(".iot-solution-model-field-affordances")!;
            const control = role.querySelector(".fui-Dropdown"), hint = role.querySelector(".iot-solution-selection-cue");
            const roleControl = control || role.querySelector(".fui-Text:not(.iot-solution-model-field-mobile-label)")!;
            const metadataStyle = getComputedStyle(metadata.firstElementChild!);
            const descriptionStyle = getComputedStyle(description);
            const content = row.querySelector(".iot-solution-model-field-content")!;
            const checkbox = row.querySelector(".iot-solution-model-field-include")!;
            const primaryStyle = getComputedStyle(title.querySelector(".fui-Text")!);
            const qualifierStyle = getComputedStyle(title.querySelector(".iot-solution-field-qualifier")!);
            const rect = (node: Element) => node.getBoundingClientRect();
            const header = (column: string) => (innerWidth <= 620 ? row : row.parentElement!.querySelector(".iot-solution-model-field-columns")!)
                .querySelector<HTMLElement>(`[data-field-column="${column}"]`)!;
            const headerLeftOffsets = [rect(header("field")).left - rect(title).left,
                rect(header("type")).left - rect(metadata.firstElementChild!).left,
                rect(header("role")).left - rect(roleControl).left];
            return { state, field: row.dataset.fieldId, viewport: innerWidth,
                headerLeftOffsets, headersLeftAligned: ["field", "type", "role"].every(column =>
                    getComputedStyle(header(column)).textAlign === "left" && rect(header(column)).width > 0),
                gridRows: getComputedStyle(row).gridTemplateRows, roleHeight: rect(affordances).height,
                descriptionGap: rect(description).top - rect(title).bottom,
                controlHintGap: control && hint ? rect(hint).top - rect(control).bottom : null,
                hintAligned: !control || !hint || Math.abs(rect(hint).left - rect(control).left) < 1,
                hintMarginTop: hint ? getComputedStyle(hint).marginTop : null,
                roleGap: getComputedStyle(affordances).rowGap,
                roleAlignment: getComputedStyle(affordances).alignSelf,
                roleCenterOffset: (rect(affordances).top + rect(affordances).bottom - rect(row).top - rect(row).bottom) / 2,
                roleBelowContent: rect(roleControl).top >= rect(content).bottom,
                contentAligned: Math.abs(rect(description).left - rect(title).left) < 1,
                metadataCenterOffset: (rect(metadata).top + rect(metadata).bottom - rect(row).top - rect(row).bottom) / 2,
                metadataControlCenterOffset: (rect(metadata).top + rect(metadata).bottom -
                    rect(roleControl).top - rect(roleControl).bottom) / 2,
                metadataAlignment: getComputedStyle(metadata).alignSelf,
                metadataBetweenColumns: rect(metadata).left >= rect(content).right && rect(roleControl).left >= rect(metadata).right,
                metadataBetweenRows: rect(metadata).top >= rect(content).bottom && rect(roleControl).top >= rect(metadata).bottom,
                metadataDomOrder: !!(content.compareDocumentPosition(metadata) & Node.DOCUMENT_POSITION_FOLLOWING) &&
                    !!(metadata.compareDocumentPosition(role) & Node.DOCUMENT_POSITION_FOLLOWING) &&
                    metadata.parentElement === affordances,
                checkboxTitleAligned: Math.abs(rect(checkbox).top - rect(content).top) < 1,
                overlap: rect(roleControl).left < rect(title).right && rect(roleControl).right > rect(title).left &&
                    rect(roleControl).top < rect(description).bottom && rect(roleControl).bottom > rect(title).top,
                overflow: row.scrollWidth > row.clientWidth + 1,
                descriptionFont: getComputedStyle(description).fontStyle,
                descriptionSize: getComputedStyle(description).fontSize,
                descriptionLineHeight: getComputedStyle(description).lineHeight,
                descriptionColor: descriptionStyle.color, metadataSize: metadataStyle.fontSize,
                metadataLineHeight: metadataStyle.lineHeight, metadataColor: metadataStyle.color,
                descriptionMatchesMetadata: (["fontFamily", "fontSize", "fontWeight", "lineHeight", "color"] as const)
                    .every(property => descriptionStyle[property] === metadataStyle[property]),
                titleSize: primaryStyle.fontSize, titleWeight: primaryStyle.fontWeight,
                qualifierSize: qualifierStyle.fontSize, qualifierFont: qualifierStyle.fontStyle,
                primaryMatchesQualifier: (["fontFamily", "fontSize", "fontWeight", "lineHeight", "color"] as const)
                    .every(property => primaryStyle[property] === qualifierStyle[property]),
                titleFont: getComputedStyle(title).fontStyle, metadataFont: getComputedStyle(metadata).fontStyle };
        });
        assert(next()!.disabled, "Missing Calibration roles passed the combined Continue gate.");
        const before = snapshots("missing-role");
        const longHint = rows()[1].querySelector(".iot-solution-selection-cue")!;
        longHint.firstChild!.textContent = "Select a role for. Choose the appropriate role for this calibration field before continuing.";
        await settle();
        const wrapped = snapshots("multiline-hint");
        for (const field of currentContext.hubFields as ModelField[]) {
            const control = documentElementRole(field.id)!;
            const hint = document.getElementById(control.getAttribute("aria-describedby")!);
            assert(control.getAttribute("aria-invalid") === "true" && !!hint &&
                hint.textContent?.includes("Select a role for"), "The unresolved hint lost its accessible linkage.");
            await enter(control.id, "telemetry");
        }
        await until(() => !next()!.disabled, "Explicit Calibration role choices did not satisfy Continue.");
        const after = snapshots("selected-role");
        const measurements = [...before, ...wrapped, ...after];
        assert(measurements.every(value => Math.abs(value.descriptionGap) < 1 && value.contentAligned &&
            !value.overlap && !value.overflow), `Role/hint stretched field text spacing: ${JSON.stringify(measurements)}`);
        assert(measurements.every(value => value.headersLeftAligned && value.headerLeftOffsets.every(offset => Math.abs(offset) < 1)),
            `Field/Type/Role headings are not left aligned with the actual row text/values/controls: ${JSON.stringify(measurements)}`);
        assert(measurements.every(value => value.hintAligned &&
            (value.controlHintGap === null || Math.abs(value.controlHintGap - 4) < 1 && value.hintMarginTop === "0px")),
            `Role validation hint is detached from its control: ${JSON.stringify(measurements)}`);
        const rolePositioned = (value: typeof measurements[number]) => value.viewport <= 620
            ? value.roleAlignment === "start" && value.roleBelowContent &&
                value.metadataAlignment === "start" && value.metadataBetweenRows && value.metadataDomOrder && value.checkboxTitleAligned
            : value.roleAlignment === "center" && Math.abs(value.roleCenterOffset) < 1 &&
                value.metadataAlignment === "center" && Math.abs(value.metadataControlCenterOffset) < 1 &&
                value.metadataBetweenColumns && value.metadataDomOrder && value.checkboxTitleAligned;
        assert(measurements.every(rolePositioned),
            `Role/hint unit is not centered on desktop or naturally below content on narrow layouts: ${JSON.stringify(measurements)}`);
        assert(measurements.every(value => value.descriptionFont === "normal" && value.descriptionSize === "12px" &&
            value.descriptionLineHeight === "16px" && value.descriptionMatchesMetadata && value.titleFont === "normal" &&
            value.metadataFont === "normal" && value.titleSize === "14px" && value.titleWeight === "400" &&
            value.qualifierSize === "14px" && value.qualifierFont === "normal" && value.primaryMatchesQualifier),
            `Description/title/metadata typography changed: ${JSON.stringify(measurements)}`);
        const compact = rows()[0].querySelector(".iot-solution-model-field-metadata")!;
        assert(compact.children.length === 1 && compact.textContent === "number \u00b7 mV" &&
            rows()[1].querySelector(".iot-solution-model-field-metadata")!.textContent === "object" &&
            rows()[2].querySelector(".iot-solution-model-field-metadata")!.textContent === "integer \u00b7 dBm" &&
            getComputedStyle(compact.firstElementChild!).whiteSpace !== "nowrap",
            "Compact metadata order, comma flow or explicit scalar declarations changed.");
        const settings = currentContext.hubFields[1] as ModelField;
        const docs = semanticDocumentation(currentContext.hubModel, currentContext.hubFields);
        assert(readColumnSemantics(docs.columns[settings.column])!.version === 3 && docs.columns[settings.column].length <= 1000 &&
            !rows()[1].querySelector(".iot-solution-model-field-metadata")!.textContent?.includes("properties:"),
            "Compact display leaked shape or published invalid/unbounded metadata.");
        await click(rows()[0].querySelector('input[type="checkbox"]'));
        assert(controlDisabled(documentElementRole(model.fields[0].id)), "Excluded field role picker stopped being disabled.");
        await click(rows()[0].querySelector('input[type="checkbox"]'));
        const dtdlDocument = { ...dtdl, contents: [
            { "@type": "Telemetry", name: "calibration", displayName: "Calibration", description: "Current calibration.",
                schema: "double", unit: "degreeCelsius" },
            { "@type": "Property", name: "settings", displayName: "Calibration settings",
                description: calibrationLayoutDocument.properties.settings.description,
                schema: { "@type": "Object", fields: [{ name: "offset", schema: "double" }] } },
        ] };
        const dtdlModel = fromDtdlModel(parseDtdlModel(dtdlDocument), dtdlDocument);
        await wizard(withAnalyzedSamples({ ...context(), hubModel: dtdlModel, hubFields: dtdlModel.fields, preparationFileName: "calibration-dtdl.json" }));
        const declared = snapshots("dtdl-readonly");
        assertFieldSelectionCopy();
        assert(!area.textContent?.includes("DTDL roles are defined by the model."), "DTD L retained a redundant role notice.");
        await assertDtdlRoleInformation();
        assert(declared.every(value => Math.abs(value.descriptionGap) < 1 && value.contentAligned &&
            !value.overlap && !value.overflow && value.descriptionFont === "normal" && value.titleFont === "normal" &&
            value.metadataFont === "normal" && value.descriptionMatchesMetadata && value.primaryMatchesQualifier &&
            value.titleSize === "14px" && value.qualifierSize === "14px" && value.titleWeight === "400" &&
            value.qualifierFont === "normal" && value.controlHintGap === null) && !next()!.disabled &&
            !area.querySelector('[role="combobox"][id^="field-role-"]') &&
            rows()[0].querySelector(".iot-solution-model-field-metadata")!.textContent === "double \u00b7 degreeCelsius" &&
            rows()[1].querySelector(".iot-solution-model-field-metadata")!.textContent === "Object",
            `DTD L compact metadata/style/read-only roles drifted: ${JSON.stringify(declared)}`);
        assert(declared.every(rolePositioned), "Readonly DTDL roles no longer share centered/narrow-flow placement.");
        assert(declared.every(value => value.headersLeftAligned && value.headerLeftOffsets.every(offset => Math.abs(offset) < 1)),
            `Readonly DTDL headers drifted from their actual columns: ${JSON.stringify(declared)}`);
        const dtdlDocs = semanticDocumentation(dtdlModel, dtdlModel.fields);
        assert(readColumnSemantics(dtdlDocs.columns[dtdlModel.fields[1].column])!.version === 3 &&
            dtdlDocs.columns[dtdlModel.fields[1].column].includes('"fields":['),
            "DTD L UI simplification lost declared nested facts.");
        result = { measurements: [...measurements, ...declared], extraRawReads: analysisQueries().length - startReads,
            missingGateAndAccessibleHint: true, selectedContinue: true, compactOrderedMetadata: true, fullMetadataRetained: true,
            exactTypeUnitText: ["number \u00b7 mV", "object", "integer \u00b7 dBm", "double \u00b7 degreeCelsius", "Object"] };
        await wizard(seed);
        const finalHint = rows()[1].querySelector(".iot-solution-selection-cue")!;
        finalHint.firstChild!.textContent = "Select a role for. Choose the appropriate role for this calibration field before continuing.";
    });
    return result;
}
async function uploadSpacingCheck(leaveAccepted = true) {
    let result: Record<string, unknown> = {};
    await observeWizardContext(async () => {
        const seed = withAnalyzedSamples(context(true));
        delete seed.hubModel; seed.hubFields = []; seed.preparationFileName = ""; seed.hubSelectedFormat = "wot";
        const Prepared = (props: WizardStepProps) => <ModelPreparationStep {...props} workloadClient={client} workspaceId="workspace"
            sourceState={sourceFixture} getSourceState={() => sourceFixture}/>;
        const mountUpload = () => mount(<div className="iot-solution-hub-model-wizard" style={{ height: "100%" }}>
            <WizardControl title="Upload spacing" initialContext={seed} steps={[
                { id: "model", title: "Device Model", component: Prepared, validate: () => false },
            ]} onComplete={() => { throw new Error("No writes in upload spacing checks."); }}/>
        </div>);
        const reads = calls.length;
        await mountUpload();
        const measure = () => {
            const tile = area.querySelector<HTMLElement>(".iot-solution-model-upload-area")!;
            const identity = tile.querySelector<HTMLElement>(".iot-solution-model-upload-identity")!;
            const style = getComputedStyle(tile), children = [...identity.children];
            return { padding: style.padding, marginTop: style.marginTop, outerGap: style.rowGap,
                identityGap: getComputedStyle(identity).rowGap,
                actualIdentityGaps: children.slice(1).map((child, index) => child.getBoundingClientRect().top - children[index].getBoundingClientRect().bottom),
                nestedButtons: tile.querySelectorAll("button").length, role: tile.getAttribute("role"), tabIndex: tile.tabIndex,
                overflow: tile.scrollWidth > tile.clientWidth + 1, background: style.backgroundColor, action: tile.getAttribute("aria-label") };
        };
        const initial = measure();
        assert(initial.action === "Choose File", "Default upload action changed.");
        const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
        let chooser = 0;
        const observeChooser = (event: Event) => { event.preventDefault(); chooser++; };
        input.addEventListener("click", observeChooser);
        let accepted: ReturnType<typeof measure>, retry: ReturnType<typeof measure>;
        const activate = async (label: string) => {
            const tile = modelUploadAction(label)!, before = chooser;
            assert(tile?.getAttribute("role") === "button" && tile.tabIndex === 0 && !tile.querySelector("button"),
                "Upload states must expose one keyboard-accessible tile, not a nested button.");
            await click(tile);
            for (const key of ["Enter", " "]) {
                const event = new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true });
                await act(async () => tile.dispatchEvent(event));
                assert(event.defaultPrevented, "Upload key activation retained its default scrolling behavior.");
            }
            assert(chooser === before + 3, "Mouse/Enter/Space must each open exactly one guarded chooser.");
        };
        try {
            await activate("Choose File");
            await file("thing-model-upload", JSON.stringify(tm), "thermoProX200.json");
            accepted = measure();
            assert(["padding", "marginTop", "outerGap", "identityGap"].every(key =>
                initial[key as keyof typeof initial] === accepted[key as keyof typeof accepted]) &&
                accepted.actualIdentityGaps.every(gap => Math.abs(gap - Number.parseFloat(initial.identityGap)) < 1) &&
                initial.actualIdentityGaps.every(gap => Math.abs(gap - Number.parseFloat(initial.identityGap)) < 1),
                `Loaded upload spacing differs from default: ${JSON.stringify({ initial, accepted })}`);
            const tile = acceptedModelTile()!;
            assert(accepted.action === "Replace Model" && tile.classList.contains("iot-solution-drop-zone--loaded"),
                "Shared spacing changed success styling or upload intent.");
            await activate("Replace Model");
            const transfer = new DataTransfer();
            transfer.items.add(new File([JSON.stringify(tm)], "dropped.json", { type: "application/json" }));
            await act(async () => tile.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: transfer })));
            await until(() => area.querySelector("#model-upload-file-name")?.textContent === "dropped.json", "Shared layout lost validated drop.");
            assert(chooser === 6, "Dropping a file opened the picker.");
            await file("thing-model-upload", "{bad", "retry.json");
            retry = measure();
            assert(retry.padding === initial.padding && retry.identityGap === initial.identityGap &&
                area.querySelector('[role="alert"]') && !acceptedModelTile(), "Retry retained accepted identity or mismatched spacing.");
            await activate("Replace Model");
        } finally { input.removeEventListener("click", observeChooser); }
        await file("thing-model-upload", JSON.stringify({ ...tm, id: `https://models.example.test/${"canonical-identifier-".repeat(15)}v1` }),
            `${"accepted-model-file-".repeat(12)}.json`);
        const long = measure();
        assert(!long.overflow && calls.length === reads, "Shared spacing broke wrapping or issued a source read.");
        result = { initial, accepted, retry, long, sourceReads: 0, chooserRequests: chooser,
            allStatesClickable: true, noNestedButtons: true, dropWithoutChooser: true };
        if (leaveAccepted) await file("thing-model-upload", JSON.stringify(tm), "thermoProX200.json");
        else await mountUpload();
    });
    return result;
}
const tests: [
    string,
    () => Promise<void>
][] = [
    ["Documentation links working Activator summary uses the existing Fabric browser-tab contract", async () => {
        await mount(<Harness component={ActivatorSummaryStep} initial={{ actActivatorId: "fixture-activator" }}/>);
        assertDocumentationClicks([await documentationClick(area.querySelector("a"),
            "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators")]);
    }],
    ["Documentation links both production overview roots pass their Fabric client without changing wizard state", async () => {
        const key = "iot-wizard-v5-ingestion-docs-overview", saved = sessionStorage.getItem(key);
        const item = { id: "docs-overview", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>;
        const observations = [];
        try {
            sessionStorage.removeItem(key);
            for (const view of [
                <IoTSolutionItemDefaultView workloadClient={client} item={item}
                    onWizardComplete={() => { throw new Error("Documentation must not save ingestion."); }}/>,
                <IoTSolutionItemModelWizard workloadClient={client} item={item}/>,
            ]) {
                await mount(view);
                observations.push(await documentationClick(area.querySelector(".iot-solution-prerequisites a"),
                    "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal"));
            }
            assertDocumentationClicks(observations);
        } finally {
            await mount(<div>Documentation overview fixture closed.</div>);
            if (saved === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, saved);
        }
    }],
    ["Documentation links Eventhouse help uses the production ingestion step client without extra discovery", async () => {
        const key = "iot-wizard-v5-ingestion-docs-eventhouse", saved = sessionStorage.getItem(key);
        try {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "eventhouse", completedSteps: [],
                wizardContext: { iotHubValidated: true } }));
            await mount(<IoTSolutionItemDefaultView workloadClient={client}
                item={{ id: "docs-eventhouse", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onWizardComplete={() => { throw new Error("Documentation must not save ingestion."); }}/>);
            await until(() => !area.textContent?.includes("Loading Eventhouses"), "Eventhouse fixture did not settle.");
            assertDocumentationClicks([await documentationClick(area.querySelector(".iot-solution-info a"),
                "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/create-eventhouse")]);
        } finally {
            await mount(<div>Eventhouse documentation fixture closed.</div>);
            if (saved === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, saved);
        }
    }],
    ["Documentation links routing permission help navigates without retrying permission or routing operations", async () => {
        const key = "iot-wizard-v5-ingestion-docs-routing", saved = sessionStorage.getItem(key), savedFetch = window.fetch;
        const requests: string[] = [];
        try {
            window.fetch = async (input, init) => {
                const url = String(input); requests.push(`${init?.method || "GET"} ${url}`);
                assert(url === "https://api.fabric.microsoft.com/v1/workspaces/workspace/roleAssignments",
                    "Routing documentation fixture reached an unrelated endpoint.");
                return init?.method === "POST"
                    ? new Response(JSON.stringify({ message: "Synthetic permission failure" }), { status: 403 })
                    : new Response(JSON.stringify({ value: [] }), { status: 200 });
            };
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "routing", completedSteps: [],
                wizardContext: { hubName: "Fixture Hub", iotHubPrincipalId: "fixture-principal" } }));
            await mount(<IoTSolutionItemDefaultView workloadClient={client}
                item={{ id: "docs-routing", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onWizardComplete={() => { throw new Error("Documentation must not save ingestion."); }}/>);
            await click(button("Check Access"));
            await until(() => !!button("Add as Contributor"), "Synthetic missing role did not expose its existing action.");
            await click(button("Add as Contributor"));
            await until(() => !!area.querySelector("a"), "Synthetic permission error did not expose documentation.");
            const count = requests.length;
            assertDocumentationClicks([await documentationClick(area.querySelector("a"),
                "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/event-streams/connect-using-managed-identity#step-2-assign-fabric-workspace-permissions")]);
            assert(requests.length === count && count === 2, "Documentation retried permission or routing operations.");
        } finally {
            window.fetch = savedFetch;
            await mount(<div>Routing documentation fixture closed.</div>);
            if (saved === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, saved);
        }
    }],
    ["Documentation links assignment ADR help receives the production model wizard client without mutations", async () => {
        const fixture = sourceHubFixture("omitted");
        await wizard(context(true)); await click(next()); await beginAssignmentTargets();
        await until(() => !!area.querySelector('[aria-label="Source IoT Hubs"] a'), "ADR warning did not settle.");
        const reads = fixture.calls.length;
        assertDocumentationClicks([await documentationClick(area.querySelector('[aria-label="Source IoT Hubs"] a'),
            "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal")]);
        assert(fixture.calls.length === reads && fixture.patches.length === 0, "ADR help performed discovery or assignment.");
    }],
    ["Documentation links success-only DPS guidance uses Fabric navigation for existing and newly confirmed devices", async () => {
        const discover = HubAssignmentApi.prototype.discover;
        const observations = [];
        try {
            const fixture = sourceHubFixture("linked"), inventory = assignmentInventory();
            inventory.targets = inventory.targets.map(target => ({ ...target, writeBlocker: undefined,
                reference: { kind: "assigned", ids: ["urn:already-assigned-model"] } }));
            HubAssignmentApi.prototype.discover = async () => inventory;
            await wizard(context(true)); await click(next()); await beginAssignmentTargets();
            await until(() => !!area.querySelector('[aria-label="DPS enrollment guidance"] a'), "Existing assignments did not show success-only DPS guidance.");
            assertDpsGuidance(true);
            const reads = fixture.calls.length;
            observations.push(await documentationClick(area.querySelector('[aria-label="DPS enrollment guidance"] a'),
                "https://aka.ms/iot-hub-agent-ready-data-doc"));
            assert(fixture.calls.length === reads && fixture.patches.length === 0, "DPS help performed discovery or assignment.");
        } finally { HubAssignmentApi.prototype.discover = discover; }
        await assignmentSpacingPresentation("confirmed");
        assertDpsGuidance(true);
        observations.push(await documentationClick(area.querySelector('[aria-label="DPS enrollment guidance"] a'),
            "https://aka.ms/iot-hub-agent-ready-data-doc"));
        assertDocumentationClicks(observations);
    }],
    ["Documentation links verified Ready receives the production model wizard client for both Learn actions", async () => {
        const proof = await creationPresentationCheck("ready");
        caseObservations = proof;
    }],
    ["Data Agent actual asynchronous creation binds the operation result before enabling Ready and host navigation", () => observeWizardContext(async () => {
        const key = "iot-wizard-v4-dataagent-created-id", savedStorage = sessionStorage.getItem(key), savedFetch = window.fetch;
        const savedSetTimeout = window.setTimeout;
        const workspaceId = "11111111-1111-4111-8111-111111111111", agentId = "22222222-2222-4222-8222-222222222222";
        const operationId = "0e82388e-e811-459d-8f71-5f9048deb17a";
        const stateUrl = `https://api.fabric.microsoft.com/v1/operations/${operationId}`;
        const regionalStateUrl = `https://wabi-us-central-b-primary-redirect.analysis.windows.net/v1/operations/${operationId}`;
        const resultUrl = regionalStateUrl + "/result";
        const source = dataAgentSourceFixture("wot");
        const requests: { method: string; url: string; redirect?: RequestRedirect }[] = [], delays: number[] = [];
        const navigations: { target: string; path: string }[] = [];
        let resultRelease: (() => void) | undefined, stateReads = 0, posts = 0;
        const rootClient: WorkloadClientAPI = { ...client, navigation: { ...client.navigation,
            navigate: async (target, params) => { navigations.push({ target, path: params.path }); },
        } };
        window.setTimeout = ((callback: TimerHandler, delay?: number, ...args: unknown[]) => {
            if (delay === 20000 || delay === 3000) { delays.push(delay); return savedSetTimeout(callback, 0, ...args); }
            return savedSetTimeout(callback, delay, ...args);
        }) as typeof window.setTimeout;
        window.fetch = async (input, init) => {
            const url = String(input); requests.push({ method: init?.method || "GET", url, redirect: init?.redirect });
            if (url === `https://api.fabric.microsoft.com/v1/workspaces/${workspaceId}/dataAgents` && init?.method === "POST") {
                posts++;
                return new Response(null, { status: 202, headers: { Location: regionalStateUrl, "x-ms-operation-id": operationId, "Retry-After": "20" } });
            }
            assert((init?.method || "GET") === "GET" &&
                new Headers(init?.headers).get("Authorization") === "Bearer mock-only",
                "LRO follow-up did not use the bounded authenticated GET contract.");
            if (url === stateUrl) {
                const redirect = new Response(null);
                Object.defineProperties(redirect, { type: { value: "opaqueredirect" }, status: { value: 0 } });
                return redirect;
            }
            if (url === regionalStateUrl) {
                stateReads++;
                return stateReads === 1 ? new Response(JSON.stringify({ status: "Running", percentComplete: 25 }), {
                    status: 200, headers: { "Retry-After": "3", Location: regionalStateUrl },
                }) : new Response(JSON.stringify({ status: "Succeeded", createdTimeUtc: "2026-11-02T00:00:00Z",
                    lastUpdatedTimeUtc: "2026-11-02T00:00:01Z", percentComplete: 100, error: null }), {
                    status: 200, headers: { Location: resultUrl },
                });
            }
            assert(url === resultUrl, "LRO requested an unrelated source or lookup endpoint.");
            return new Promise(resolve => { resultRelease = () => resolve(new Response(JSON.stringify({
                id: agentId, workspaceId, type: "DataAgent", displayName: "Fixture async agent",
            }), { status: 200 })); });
        };
        try {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "configure-agent", completedSteps: ["overview", "data-source"], wizardContext: {
                agentEhId: "eh", agentDbId: "db", agentDbName: source.database, agentTableName: source.table,
                agentQueryServiceUri: "https://mock.kusto.fabric.microsoft.com", agentName: "Fixture async agent",
                agentTableDocstring: source.docstring, agentColumns: source.columns,
            } }));
            await mount(<IoTSolutionItemDataAgentWizard workloadClient={rootClient}
                item={{ id: "created-id", workspaceId } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onWizardComplete={() => { throw new Error("No live item saves in identity flow."); }}/>);
            assert(requests.length === 0, "Configure dispatched a resource read before explicit creation.");
            await click(button("Create Data Agent"));
            await until(() => !!resultRelease || !!currentContext.agentCreated, "Async creation reached neither result nor terminal state.");
            assert(!!resultRelease && !currentContext.agentCreated && !currentContext.agentId && next()!.disabled,
                "Succeeded operation state was treated as the created agent without reading its resource result.");
            assert(posts === 1 && stateReads === 2 && JSON.stringify(delays) === JSON.stringify([20000, 3000]),
                "Creation repeated its POST or ignored initial/updated Retry-After.");
            assert(requests.slice(1).every(request => request.redirect === "manual"), "Authenticated LRO reads allowed automatic redirects.");
            resultRelease!();
            await until(() => currentContext.agentCreated === true && !next()!.disabled, "Validated resource result did not complete creation.");
            assert(currentContext.agentId === agentId && currentContext.agentId !== operationId && posts === 1,
                "Created context used the operation ID, a missing ID or a second agent POST.");
            await click(next());
            await until(() => !!button("Open Data Agent in Fabric"), "Actual Ready action did not render.");
            assert(!button("Open Data Agent in Fabric")!.disabled, "Valid async created resource still leaves the Ready action disabled.");
            await click(button("Open Data Agent in Fabric"));
            assert(JSON.stringify(navigations) === JSON.stringify([{ target: "host", path: `/groups/${workspaceId}/aiskills/${agentId}` }]) &&
                posts === 1, "Ready navigation did not use the actual returned resource/current workspace.");
            caseObservations = { actualRoot: true, requests, delays, posts, stateReads, resultReads: 1,
                recordedRegionalLocation: regionalStateUrl, canonicalHiddenRedirectUsesOnlySuppliedRegionalHint: true,
                statusOnlySucceededWaitsForResult: true, validatedResourceId: agentId, operationIdNotUsedAsAgentId: true,
                nextAndReadyBlockedUntilResult: true, readyActionEnabled: true, navigations,
                sourceReads: 0, additionalCreationPosts: 0, liveWrites: 0, liveNavigation: 0 };
        } finally {
            window.setTimeout = savedSetTimeout; window.fetch = savedFetch;
            if (savedStorage === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, savedStorage);
        }
    })],
    ["Data Agent actual creation rejects unverifiable IDs and never replays an uncertain POST", () => observeWizardContext(async () => {
        const key = "iot-wizard-v4-dataagent-created-id-errors", savedStorage = sessionStorage.getItem(key), savedFetch = window.fetch;
        const savedSetTimeout = window.setTimeout;
        const workspaceId = "11111111-1111-4111-8111-111111111111", agentId = "22222222-2222-4222-8222-222222222222";
        const stateUrl = "https://api.fabric.microsoft.com/v1/operations/33333333-3333-4333-8333-333333333333";
        const source = dataAgentSourceFixture("wot"), observations: Record<string, unknown>[] = [];
        let mode = "", posts = 0, resultReads = 0, polls = 0, released: (() => void) | undefined;
        window.setTimeout = ((callback: TimerHandler, delay?: number, ...args: unknown[]) =>
            savedSetTimeout(callback, delay === 2000 ? 0 : delay, ...args)) as typeof window.setTimeout;
        window.fetch = async (input, init) => {
            const url = String(input);
            if (url.endsWith("/dataAgents") && init?.method === "POST") {
                posts++;
                if (mode === "post-transport") throw new TypeError("Synthetic lost creation response.");
                if (mode === "200" || mode === "201") return new Response(JSON.stringify({ id: agentId, type: "DataAgent", workspaceId }), { status: Number(mode) });
                if (mode === "opaque-create") {
                    const redirect = new Response(null);
                    Object.defineProperties(redirect, { type: { value: "opaqueredirect" }, status: { value: 0 } });
                    return redirect;
                }
                if (mode === "direct-missing-id") return new Response(JSON.stringify({ type: "DataAgent", workspaceId }), { status: 201 });
                return new Response(null, { status: 202, headers: { Location: stateUrl, "Retry-After": "2" } });
            }
            if (url === stateUrl) {
                polls++;
                if (mode === "failed") return new Response(JSON.stringify({ status: "Failed", error: { message: "Synthetic provisioning failure." } }));
                if (mode === "cancelled") return new Response(JSON.stringify({ status: "Cancelled" }));
                if (mode === "timeout") return new Response(JSON.stringify({ status: "Running" }));
                if (mode === "unknown") return new Response(JSON.stringify({ status: "Unknown" }));
                if (mode === "state-transport") throw new TypeError("Synthetic lost operation response.");
                if (mode === "state-denied") return new Response(JSON.stringify({ status: "Failed" }), { status: 403 });
                return new Response(JSON.stringify({ status: "Succeeded", percentComplete: 100 }), { headers: { Location: stateUrl + "/result" } });
            }
            assert(url === stateUrl + "/result" && init?.method === "GET", "Invalid-result test attempted a new lookup or source request.");
            resultReads++;
            if (mode === "unmount") return new Promise(resolve => { released = () => resolve(new Response(JSON.stringify({ id: agentId, type: "DataAgent", workspaceId }))); });
            if (mode === "result-denied") return new Response(JSON.stringify({ errorCode: "Forbidden", message: "Synthetic result denial." }), { status: 403 });
            if (mode === "result-malformed") return new Response("{bad", { status: 200 });
            if (mode === "wrong-workspace") return new Response(JSON.stringify({ id: agentId, workspaceId: "other-workspace", type: "DataAgent" }));
            return new Response(JSON.stringify({ resourceId: `/workspaces/${workspaceId}/dataAgents/${agentId}` }));
        };
        const mountConfigure = async (workspace: string) => {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "configure-agent", completedSteps: ["overview", "data-source"], wizardContext: {
                agentEhId: "eh", agentDbId: "db", agentDbName: source.database, agentTableName: source.table,
                agentName: "Fixture identity", agentTableDocstring: source.docstring, agentColumns: source.columns,
            } }));
            await mount(<IoTSolutionItemDataAgentWizard workloadClient={client}
                item={{ id: "created-id-errors", workspaceId: workspace } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onWizardComplete={() => { throw new Error("No item saves in rejected identity checks."); }}/>);
        };
        try {
            for (const variant of ["200", "201", "direct-missing-id", "result-missing-id", "result-malformed", "wrong-workspace",
                "result-denied", "failed", "cancelled", "timeout", "unknown", "state-transport", "state-denied", "post-transport", "opaque-create"]) {
                mode = variant; posts = 0; resultReads = 0; polls = 0;
                await mountConfigure(workspaceId);
                await click(button("Create Data Agent"));
                if (variant === "200" || variant === "201") {
                    await until(() => !!currentContext.agentCreated, "Normal direct resource response did not bind its validated ID.");
                    assert(currentContext.agentId === agentId && posts === 1 && resultReads === 0 && !next()!.disabled,
                        "Direct creation did not preserve the native result identity.");
                } else {
                    await until(() => !!area.querySelector(".iot-solution-creation-step-label.error"), "Unverified identity did not surface a native creation error.");
                    assert(!currentContext.agentCreated && !currentContext.agentId && next()!.disabled && posts === 1,
                        "Missing/invalid result retained completion or issued another creation POST.");
                    const definitive = variant === "failed" || variant === "cancelled";
                    assert(button("Create Data Agent")!.disabled !== definitive && !area.querySelector(".iot-solution-success"),
                        "A definitive unsuccessful result stayed blocked or an uncertain result enabled duplicate creation.");
                    assert(!!area.textContent?.includes("before creating another") !== definitive,
                        "Creation copy mislabeled definitive failure/cancellation as an uncertain outcome.");
                    if (variant === "timeout") assert(polls === 30, "Timeout did not keep its bounded polling limit.");
                    await settle(); await settle();
                    assert(posts === 1, "An unsuccessful or uncertain operation automatically resent creation.");
                    if (definitive) {
                        assert(resultReads === 0 && polls === 1 &&
                            area.textContent?.includes(variant === "failed" ? "Synthetic provisioning failure." : "cancelled"),
                            "Definitive failure/cancellation lost its real error or fetched a nonexistent created result.");
                        mode = "201";
                        await click(button("Create Data Agent"));
                        await until(() => !!currentContext.agentCreated, "Explicit new creation after confirmed failure did not finish.");
                        assert(posts === 2 && currentContext.agentId === agentId && !next()!.disabled,
                            "A confirmed unsuccessful original operation did not permit exactly one explicit new attempt.");
                    }
                }
                observations.push({ variant, posts, resultReads, polls, completed: !!currentContext.agentCreated,
                    agentId: currentContext.agentId || "", nextDisabled: next()!.disabled });
            }
            await mountConfigure("");
            assert(button("Create Data Agent")!.disabled && !currentContext.agentCreated && next()!.disabled,
                "Missing current workspace allowed creation or success without a destination.");
            mode = "unmount"; posts = 0; resultReads = 0; released = undefined;
            await mountConfigure(workspaceId); await click(button("Create Data Agent"));
            await until(() => !!released, "Pending result did not begin before cancellation.");
            await mount(<div>Another current page</div>);
            const retainedContext = canonicalJson(currentContext);
            released!(); await settle(); await settle();
            assert(canonicalJson(currentContext) === retainedContext && area.textContent === "Another current page" && posts === 1,
                "A cancelled former creation published identity into a newer page or replayed creation.");
            caseObservations = { actualRoot: true, observations, missingWorkspaceDisablesCreate: true,
                oldUnmountedResultCannotPublish: true, noUnknownOutcomePostReplay: true,
                definitiveFailureAndCancellationPermitOnlyExplicitNewAttempt: true, liveWrites: 0 };
        } finally {
            window.setTimeout = savedSetTimeout; window.fetch = savedFetch;
            if (savedStorage === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, savedStorage);
        }
    })],
    ["Data Agent actual Ready opens the current created agent in Fabric without saving or changing other summary content", () => observeWizardContext(async () => {
        const key = "iot-wizard-v4-dataagent-ready-open", saved = sessionStorage.getItem(key), savedFetch = window.fetch;
        const workspaceId = "11111111-1111-4111-8111-111111111111", agentId = "22222222-2222-4222-8222-222222222222";
        const navigations: { target: string; path: string }[] = [];
        let failNavigation = false, finishes = 0, unexpectedFetches = 0;
        const fixture = dataAgentSourceFixture("wot");
        const rootClient: WorkloadClientAPI = { ...client, navigation: { ...client.navigation,
            navigate: async (target, params) => {
                navigations.push({ target, path: params.path });
                if (failNavigation) throw new Error("Synthetic host-navigation failure.");
            },
        } };
        window.fetch = async () => { unexpectedFetches++; throw new Error("Ready navigation cannot dispatch source/resource requests."); };
        const mountReady = async (workspace: string, createdAgent: string) => {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "summary",
                completedSteps: ["overview", "data-source", "configure-agent"], wizardContext: {
                    agentCreated: true, agentId: createdAgent, agentName: "Fixture Ready Agent", agentDbName: fixture.database,
                    agentTableName: fixture.table, agentColumns: fixture.columns, agentTableDocstring: fixture.docstring,
                    workspaceId: "stale-context-workspace-not-current-item",
                } }));
            await mount(<IoTSolutionItemDataAgentWizard workloadClient={rootClient}
                item={{ id: "ready-open", workspaceId: workspace } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onWizardComplete={() => { finishes++; }}/>);
        };
        try {
            const reads = calls.length, auth = frontendScopeRequests.length, consent = consentRequests;
            const imagesBefore = new Set(performance.getEntriesByType("resource").map(entry => entry.name));
            await mountReady(workspaceId, agentId);
            const text = (area.querySelector(".iot-solution-step")!.textContent || "").replace(/\s+/g, " ").trim();
            const usage = area.querySelector(".iot-solution-outcome-list li")!;
            assert((usage.textContent || "").replace(/\s+/g, " ").trim() ===
                "Open your data agent Fixture Ready Agent, and ask a question about your devices or operations." &&
                usage.querySelector("strong")?.textContent === "Fixture Ready Agent",
                "Ready first bullet did not make only the requested opening change and preserve the emphasized actual name.");
            assert(!area.querySelector(".iot-solution-step img") && !text.includes("Open Fabric Copilot") &&
                !performance.getEntriesByType("resource").some(entry => !imagesBefore.has(entry.name) && /\.png(?:\?|$)/i.test(entry.name)),
                "Ready still renders or requests the removed Copilot guide image.");
            assert(area.querySelector("h2")?.textContent?.trim() === "Your IoT Data Agent is ready" &&
                text.includes("The agent is grounded in the real-time telemetry and current device state stored in the modeled data table.") &&
                area.querySelectorAll(".iot-solution-config-card li").length === 5 &&
                area.querySelectorAll(".iot-solution-validation-row").length === 4 &&
                area.querySelectorAll(".iot-solution-architecture-node").length === 4 &&
                text.includes(`${fixture.database} / ${fixture.table}`) && text.includes(`${fixture.columns.length} modeled columns`),
                "Ready changes dropped existing guidance, illustrative questions or actual configuration values.");
            const open = button("Open Data Agent in Fabric")!;
            const actions = area.querySelector<HTMLElement>(".iot-solution-summary-actions")!;
            const hint = actions.querySelector<HTMLElement>(".iot-solution-finish-hint")!;
            assert(open && !open.disabled && open.classList.contains("fui-Button") && actions.children[0] === open &&
                actions.children[1] === hint && (hint.textContent || "").replace(/\s+/g, " ").trim() ===
                    "Select Finish and Close when you are done." && hint.querySelector("strong")?.textContent === "Finish and Close" &&
                getComputedStyle(actions).display === "flex" && getComputedStyle(actions).flexDirection === "column" &&
                getComputedStyle(actions).rowGap === getComputedStyle(actions).getPropertyValue("--spacingVerticalS").trim() &&
                getComputedStyle(open).backgroundColor === getComputedStyle(button("Finish and Close")!).backgroundColor,
                "New action does not match the incumbent primary summary-action style/order or changed the finish hint.");
            const contextBefore = canonicalJson(currentContext);
            assert(navigations.length === 0 && finishes === 0 && unexpectedFetches === 0 &&
                calls.length === reads && frontendScopeRequests.length === auth && consentRequests === consent,
                "Ready performed navigation, auth, source reads or completion on render.");
            await click(open);
            assert(JSON.stringify(navigations) === JSON.stringify([{ target: "host", path: `/groups/${workspaceId}/aiskills/${agentId}` }]) &&
                canonicalJson(currentContext) === contextBefore && finishes === 0 && unexpectedFetches === 0,
                "Button did not navigate only to the current workspace/created agent or changed/saved context.");
            failNavigation = true;
            await click(open);
            await until(() => !!area.querySelector('[role="alert"]')?.textContent?.includes("Unable to open the Data Agent."),
                "Navigation failure was silently swallowed.");
            assert(area.querySelector('[role="alert"]')?.textContent?.includes("Synthetic host-navigation failure.") &&
                canonicalJson(currentContext) === contextBefore && finishes === 0,
                "Navigation failure lost useful error detail or changed completion state.");
            failNavigation = false;
            await click(open);
            assert(!area.querySelector('[role="alert"]'), "Explicit navigation retry retained its stale action error.");
            const navigationCount = navigations.length;
            await mountReady(workspaceId, "");
            assert(button("Open Data Agent in Fabric")?.disabled, "Missing created agent ID left a broken open action enabled.");
            await click(button("Open Data Agent in Fabric"));
            await mountReady("", agentId);
            assert(button("Open Data Agent in Fabric")?.disabled, "Missing current item workspace ID used stale context or enabled a broken route.");
            await click(button("Open Data Agent in Fabric"));
            assert(navigations.length === navigationCount && calls.length === reads && frontendScopeRequests.length === auth &&
                consentRequests === consent && unexpectedFetches === 0 && finishes === 0,
                "Disabled actions navigated or Ready issued unexpected reads/auth/saves.");
            caseObservations = { actualRoot: true, workspaceId, agentId, navigations,
                exactRoute: `/groups/${workspaceId}/aiskills/${agentId}`, openingPhraseAndStrongName: true,
                guideImageAbsentAndNotRequested: true, incumbentPrimaryActionStyleAndOrdering: true,
                finishHintUnchanged: true, missingAgentAndWorkspaceDisabled: true, explicitErrorAndRetry: true,
                sourceReads: 0, authRequests: 0, savesOrCreates: 0, finishes: 0, liveNavigation: 0, liveWrites: 0 };
        } finally {
            window.fetch = savedFetch;
            if (saved === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, saved);
        }
    })],
    ["Modeled-data actual completed success box keeps balanced native one-line and wrapped spacing", () => observeWizardContext(async () => {
        const previousStyle = area.getAttribute("style");
        const expected = "Modeled data created. Select Continue to see what was created and what to do next.";
        try {
            area.style.width = "1440px"; area.style.height = "1000px";
            await wizard(context()); await reachReview();
            const source = (observedWizardProps as React.ComponentProps<typeof HubModelReviewStep>).sourceState;
            const request = { input: hubInput(currentContext, source), coverage: currentContext.hubCoverage,
                names: currentContext.hubOutputNames, allowWaiting: true };
            creationBackend = createCreationMock(request);
            assert(!area.querySelector("#modeled-data-success"), "Success appeared before creation completed.");
            await click(button("Create modeled data"));
            await until(() => currentContext.hubCreation?.phase === "completed" && !!area.querySelector("#modeled-data-success"),
                "Actual root did not reach guarded completed success.");
            const bar = area.querySelector<HTMLElement>("#modeled-data-success")!;
            const body = bar.querySelector<HTMLElement>(".fui-MessageBarBody")!;
            const text = bar.querySelector<HTMLElement>(".iot-solution-success-text")!;
            const icon = bar.querySelector<HTMLElement>(".fui-MessageBar__icon")!;
            const progress = area.querySelector<HTMLElement>(".iot-solution-creation-steps")!;
            assert(body && text && icon && progress, "Completed message did not use the actual native Fluent slots.");
            const callsBefore = calls.length, authBefore = tokenRequests, writesBefore = creationBackend.writes.length;
            const contextBefore = canonicalJson({ creation: currentContext.hubCreation, names: currentContext.hubOutputNames,
                coverage: currentContext.hubCoverage, fields: currentContext.hubFields });
            const metrics: Record<string, unknown>[] = [];
            await document.fonts.ready;
            for (const width of [800, 530]) {
                // Constrain the actual root's mounted native box only, without changing production page sizing.
                bar.style.width = `${width}px`;
                await settle(); await settle();
                const box = bar.getBoundingClientRect(), bodyBox = body.getBoundingClientRect(), iconBox = icon.getBoundingClientRect();
                const style = getComputedStyle(bar), bodyStyle = getComputedStyle(body), textStyle = getComputedStyle(text);
                const lines: { top: number; bottom: number; left: number; right: number }[] = [];
                const walker = document.createTreeWalker(text, NodeFilter.SHOW_TEXT);
                for (let node = walker.nextNode(); node; node = walker.nextNode()) {
                    if (!node.textContent?.trim()) continue;
                    const range = document.createRange(); range.selectNodeContents(node);
                    for (const rect of range.getClientRects()) {
                        if (!rect.width || !rect.height) continue;
                        const line = lines.find(line => Math.abs(line.top - rect.top) < 1);
                        if (line) { line.bottom = Math.max(line.bottom, rect.bottom); line.left = Math.min(line.left, rect.left); line.right = Math.max(line.right, rect.right); }
                        else lines.push({ top: rect.top, bottom: rect.bottom, left: rect.left, right: rect.right });
                    }
                }
                lines.sort((first, second) => first.top - second.top);
                const borderTop = Number.parseFloat(style.borderTopWidth), borderBottom = Number.parseFloat(style.borderBottomWidth);
                const padding = Number.parseFloat(style.getPropertyValue("--spacingVerticalMNudge"));
                const topInset = bodyBox.top - box.top - borderTop, bottomInset = box.bottom - borderBottom - bodyBox.bottom;
                assert(lines.length === (width === 800 ? 1 : 2), "Success did not exercise the required one-line and two-line layouts.");
                assert(Number.parseFloat(style.paddingTop) === padding && Number.parseFloat(style.paddingBottom) === padding &&
                    Math.abs(topInset - padding) < 0.1 && Math.abs(bottomInset - padding) < 0.1 &&
                    Math.abs(topInset - bottomInset) < 0.1,
                    "Native completed box does not have balanced token-based vertical padding.");
                assert(style.alignItems === "start" && Math.abs(iconBox.top - bodyBox.top) < 0.1 &&
                    Number.parseFloat(textStyle.lineHeight) === Number.parseFloat(bodyStyle.lineHeight) &&
                    Math.abs(bodyBox.height - lines.length * Number.parseFloat(textStyle.lineHeight)) < 0.1,
                    "Completed text line height or native first-line icon alignment is inconsistent.");
                assert(bar.scrollWidth <= bar.clientWidth && body.scrollHeight <= body.clientHeight &&
                    lines.every(line => line.top >= box.top + borderTop && line.bottom <= box.bottom - borderBottom &&
                        line.left >= box.left && line.right <= box.right) &&
                    !bar.querySelector(".fui-MessageBar__bottomReflowSpacer"),
                    "Completed message clips/overflows or retains an unnecessary reflow spacer.");
                metrics.push({ width, actualWidth: box.width, height: box.height, lineCount: lines.length,
                    paddingToken: padding,
                    paddingTop: style.paddingTop, paddingBottom: style.paddingBottom, alignItems: style.alignItems,
                    textFontSize: textStyle.fontSize, textLineHeight: textStyle.lineHeight, bodyLineHeight: bodyStyle.lineHeight,
                    bodyHeight: bodyBox.height, bodyTopInset: bodyBox.top - box.top - borderTop,
                    bodyBottomInset: box.bottom - borderBottom - bodyBox.bottom,
                    lineTopInset: lines[0].top - box.top, lineBottomInset: box.bottom - lines[lines.length - 1].bottom,
                    iconCenterOffset: iconBox.top + iconBox.height / 2 - (bodyBox.top + bodyBox.height / 2),
                    iconTopOffset: iconBox.top - bodyBox.top,
                    rootScrollWidth: bar.scrollWidth, rootClientWidth: bar.clientWidth,
                    bodyScrollHeight: body.scrollHeight, bodyClientHeight: body.clientHeight,
                    progressGap: box.top - progress.getBoundingClientRect().bottom,
                    role: bar.getAttribute("role"), background: style.backgroundColor, color: textStyle.color,
                    bottomReflowSpacers: bar.querySelectorAll(".fui-MessageBar__bottomReflowSpacer").length,
                    copy: (bar.textContent || "").replace(/\s+/g, " ").trim(), strong: bar.querySelector("strong")?.textContent,
                    lines: lines.map(line => ({ top: line.top - box.top, bottom: line.bottom - box.top })),
                });
                assert((bar.textContent || "").replace(/\s+/g, " ").trim() === expected &&
                    bar.querySelector("strong")?.textContent === "Continue" && bar.getAttribute("role") === "status",
                    "Success copy, emphasis or status semantics changed.");
            }
            assert(calls.length === callsBefore && tokenRequests === authBefore && creationBackend.writes.length === writesBefore &&
                canonicalJson({ creation: currentContext.hubCreation, names: currentContext.hubOutputNames,
                    coverage: currentContext.hubCoverage, fields: currentContext.hubFields }) === contextBefore,
                "Measuring/resizing success changed creation credit, state, queries, auth or writes.");
            caseObservations = { actualModelWizardRoot: true, exactCopy: expected, metrics,
                completedOnly: true, noMeasurementReadsOrWrites: true, modeledWrites: writesBefore, liveWrites: 0 };
        } finally {
            creationBackend = undefined;
            if (previousStyle === null) area.removeAttribute("style"); else area.setAttribute("style", previousStyle);
        }
    })],
    ["Data Agent actual root preserves same-ID selections and never mounts false table errors before current loads settle", () => observeWizardContext(async () => {
        const key = "iot-wizard-v4-dataagent-source-reselect", savedStorage = sessionStorage.getItem(key), savedFetch = window.fetch;
        const source = dataAgentSourceFixture("wot");
        const requests: string[] = [], transientNotices: string[] = [];
        let deferMetadata = true, deferTables = true, metadataRelease: (() => void) | undefined;
        const tableReads: { signal?: AbortSignal; finish: (kind?: "success" | "empty" | "error") => void }[] = [];
        let noticeMustBeAbsent = true;
        const emptyNotice = "No modeled data tables were found";
        const observer = new MutationObserver(records => {
            if (!noticeMustBeAbsent) return;
            for (const record of records) for (const added of record.addedNodes) {
                const value = added.textContent || "";
                if (value.includes(emptyNotice) || added instanceof Element &&
                    (added.matches('.fui-MessageBar[role="alert"]') || added.querySelector('.fui-MessageBar[role="alert"]'))) {
                    transientNotices.push(value);
                }
            }
        });
        observer.observe(area, { subtree: true, childList: true });
        const counts = () => ({ requests: requests.length, auth: frontendScopeRequests.length, consent: consentRequests });
        const sameCounts = (before: ReturnType<typeof counts>, message: string) =>
            assert(JSON.stringify(counts()) === JSON.stringify(before), message);
        const choose = async (id: string, index: number) => {
            if (area.querySelector(`#${id}`)?.getAttribute("aria-expanded") !== "true") await click(area.querySelector(`#${id}`));
            await click(activeOptions(id)[index]);
        };
        window.fetch = async (input, init) => {
            const address = String(input), body = init?.body ? JSON.parse(String(init.body)) : undefined;
            if (address.includes("/eventhouses/eh-")) {
                requests.push(address.endsWith("eh-1") ? "eventhouse:eh-1" : "eventhouse:eh-2");
                const response = () => new Response(JSON.stringify({ properties: {
                    queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databasesItemIds: ["db-1", "db-2"],
                } }));
                if (deferMetadata) return new Promise(resolve => { metadataRelease = () => resolve(response()); });
                return response();
            }
            if (address.includes("/items") && (!init?.method || init.method === "GET")) {
                requests.push("items:" + (address.includes("Eventhouse") ? "eventhouses" : "databases"));
                return new Response(JSON.stringify({ value: [
                    { id: "eh-1", displayName: "Shared name", type: "Eventhouse" },
                    { id: "eh-2", displayName: "Shared name", type: "Eventhouse" },
                    { id: "db-1", displayName: "Shared database", type: "KQLDatabase" },
                    { id: "db-2", displayName: "Shared database", type: "KQLDatabase" },
                ] }));
            }
            assert(address === "https://mock.kusto.fabric.microsoft.com/v1/rest/mgmt", "Source reselect dispatched an unapproved endpoint.");
            if (body.csl === ".show tables details | project TableName, DocString") {
                requests.push("table-inventory");
                const response = (kind: "success" | "empty" | "error") => kind === "error" ? new Response("Synthetic inventory denial", { status: 403 })
                    : result([["TableName", "string"], ["DocString", "string"]],
                        kind === "empty" ? [] : [source.table, "SecondModeled", "ThirdModeled"].map(name => [name, source.docstring]));
                if (deferTables) return new Promise(resolve => { tableReads.push({
                    signal: init?.signal || undefined, finish: (kind = "success") => resolve(response(kind)),
                }); });
                return response("success");
            }
            assert(body.csl === `.show table ['${source.table}'] schema as json`, "Source reselect issued unexpected KQL.");
            requests.push("native-schema");
            return result([["TableName", "string"], ["Schema", "string"]], [[source.table, JSON.stringify({
                OrderedColumns: source.columns.map(column => ({ Name: column.name, CslType: column.type, DocString: column.docstring })),
            })]]);
        };
        try {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "data-source", completedSteps: ["overview"], wizardContext: {
                agentSuffix: "reselect", agentName: "Operator agent", agentCreated: true, agentId: "existing-agent-credit",
            } }));
            await mount(<IoTSolutionItemDataAgentWizard workloadClient={client}
                item={{ id: "source-reselect", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onWizardComplete={() => { throw new Error("Source reselect cannot save resources."); }}/>);
            await until(() => !!area.querySelector("#agent-eh") && !area.querySelector(".fui-Spinner"), "Initial Eventhouse inventory did not settle.");
            await choose("agent-eh", 0);
            await until(() => !!metadataRelease, "Current database discovery did not start.");
            const pendingEhCounts = counts(), pendingEhContext = JSON.stringify(currentContext);
            await choose("agent-eh", 0);
            sameCounts(pendingEhCounts, "Same Eventhouse restarted discovery/auth while its current load was pending.");
            assert(JSON.stringify(currentContext) === pendingEhContext && !area.querySelector("#agent-db"),
                "Same pending Eventhouse selection invalidated state or prematurely exposed an empty database list.");
            deferMetadata = false; metadataRelease!();
            await until(() => !!area.querySelector("#agent-db") && !area.querySelector(".fui-Spinner"), "Current databases did not appear.");
            const readyEhCounts = counts();
            await choose("agent-eh", 0);
            sameCounts(readyEhCounts, "Same Eventhouse cleared/refetched its database inventory.");
            assert(area.querySelector("#agent-db"), "Same Eventhouse left its database list empty.");
            await click(area.querySelector("#agent-db"));
            assert(activeOptions("agent-db").length === 2, "Same Eventhouse emptied the previously loaded database options.");
            await choose("agent-db", 0);
            await until(() => tableReads.length === 1, "Initial table inventory did not start.");
            assert(area.querySelector(".fui-Spinner") && !area.textContent?.includes(emptyNotice) && transientNotices.length === 0,
                "Initial database selection mounted a false empty/error notice before loading.");
            const pendingDbCounts = counts(), pendingDbContext = JSON.stringify(currentContext);
            await choose("agent-db", 0);
            sameCounts(pendingDbCounts, "Same database started a duplicate inventory/auth request while pending.");
            assert(!tableReads[0].signal?.aborted && JSON.stringify(currentContext) === pendingDbContext,
                "Same database aborted its owned load or invalidated current state.");
            tableReads[0].finish();
            await until(() => !!area.querySelector("#agent-table") && !area.querySelector(".fui-Spinner"), "Initial modeled tables did not appear.");
            await click(area.querySelector("#agent-table"));
            assert(activeOptions("agent-table").length === 3, "Initial inventory lost real modeled tables.");
            await click(activeOptions("agent-table")[0]);
            await until(() => !!currentContext.agentSourceSelected && !next()!.disabled, "Native source inspection did not finish.");
            await act(async () => {
                observedWizardProps!.updateContext("agentName", "Operator agent");
                observedWizardProps!.updateContext("agentCreated", true);
                observedWizardProps!.updateContext("agentId", "existing-agent-credit");
            });
            const selected = JSON.stringify(currentContext), selectedCounts = counts();
            for (let repeat = 0; repeat < 2; repeat++) { await choose("agent-db", 0); await choose("agent-eh", 0); }
            sameCounts(selectedCounts, "Repeated same-ID sources caused discovery/schema/auth calls.");
            assert(JSON.stringify(currentContext) === selected && area.querySelector("#agent-table") && !next()!.disabled &&
                currentContext.agentColumns.length === source.columns.length && currentContext.agentId === "existing-agent-credit" &&
                currentContext.agentName === "Operator agent", "Same source discarded table/docs/columns/credit/names/Continue.");
            await click(next());
            await until(() => !!area.querySelector("#agent-name"), "Preserved source could not reach configuration.");
            assert(area.querySelector<HTMLInputElement>("#agent-name")!.value === "Operator agent" &&
                area.querySelector(".iot-solution-success"), "Same selections lost prior agent completion or entered name.");
            deferTables = false;
            await click(button("Previous"));
            await until(() => !!area.querySelector("#agent-db") && !area.querySelector(".fui-Spinner"), "Returning to source did not settle.");
            deferTables = true;
            const beforeChange = requests.filter(request => request === "table-inventory").length;
            await choose("agent-db", 1);
            await until(() => requests.filter(request => request === "table-inventory").length === beforeChange + 1, "Different database ID with the same name did not reload.");
            assert(currentContext.agentDbId === "db-2" && !currentContext.agentSourceSelected && !currentContext.agentCreated &&
                currentContext.agentColumns.length === 0 && !area.textContent?.includes(emptyNotice),
                "Different database did not retire old credit or showed a false notice.");
            const staleRead = tableReads.at(-1)!;
            await choose("agent-db", 0);
            await until(() => requests.filter(request => request === "table-inventory").length === beforeChange + 2, "Switching back did not admit a new current table load.");
            assert(staleRead.signal?.aborted, "Former database inventory was not retired.");
            staleRead.finish("empty"); await settle();
            assert(!area.textContent?.includes(emptyNotice) && area.querySelector(".fui-Spinner"),
                "Stale former database results replaced the current pending inventory.");
            tableReads.at(-1)!.finish(); await until(() => !!area.querySelector("#agent-table"), "Current tables did not survive stale response.");
            await choose("agent-eh", 1);
            await until(() => currentContext.agentEhId === "eh-2" && !!area.querySelector("#agent-db") && !area.querySelector(".fui-Spinner"),
                "Different Eventhouse ID with the same name did not reload databases.");
            assert(currentContext.agentDbId === "" && !area.querySelector("#agent-table"), "Different Eventhouse retained old dependent source.");
            const beforeEmpty = tableReads.length;
            await choose("agent-db", 0);
            await until(() => tableReads.length === beforeEmpty + 1 && area.querySelector(".fui-Spinner"), "Real empty fixture did not start loading.");
            noticeMustBeAbsent = false;
            tableReads.at(-1)!.finish("empty");
            await until(() => !!area.textContent?.includes(emptyNotice) && !area.querySelector(".fui-Spinner"), "Settled current empty inventory lost its honest notice.");
            const beforeError = tableReads.length;
            await choose("agent-db", 1);
            await until(() => tableReads.length === beforeError + 1 && area.querySelector(".fui-Spinner"), "Changing away from empty did not start current loading.");
            assert(!area.textContent?.includes(emptyNotice), "Previous source's empty warning remained during a new load.");
            tableReads.at(-1)!.finish("error");
            await until(() => [...area.querySelectorAll(".fui-MessageBar")].some(node => node.textContent?.includes("Unable to load available tables")) &&
                !area.querySelector(".fui-Spinner"),
                "Real current inventory failure did not use the existing explicit error box.");
            assert(!area.textContent?.includes(emptyNotice), "A failed inventory was mislabeled as successful empty data.");
            assert(transientNotices.length === 0, "A false notice mounted in an intermediate initial/change/pending render.");
            caseObservations = { requests, samePendingEventhouse: true, samePendingDatabase: true, repeatedSameIdsNoCalls: true,
                selectedNativeProofNamesAndCreditPreserved: true, differentIdsSameDisplayNamesReload: true,
                formerPendingInventoryAbortedAndIgnored: true, falseIntermediateNotices: transientNotices,
                realSettledEmptyNotice: true, realFailureExplicit: true, liveWrites: 0 };
        } finally {
            observer.disconnect(); window.fetch = savedFetch;
            if (savedStorage === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, savedStorage);
        }
    })],
    ["Shared Wizard restores the current step and visits every registered page in both directions", async () => {
        const key = "shared-wizard-default-preservation", saved = sessionStorage.getItem(key);
        const First = () => <div>First fixture page</div>, Middle = () => <div>Middle fixture page</div>;
        const Last = () => <div>Last fixture page</div>;
        try {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "last", completedSteps: ["first", "middle", "last"], wizardContext: {} }));
            await mount(<WizardControl title="Shared Wizard fixture" persistKey={key} steps={[
                { id: "first", title: "First", component: First, validate: () => true },
                { id: "middle", title: "Middle", component: Middle, validate: () => true },
                { id: "last", title: "Last", component: Last, validate: () => true },
            ]} onComplete={() => { throw new Error("No completion in presentation test."); }}/>);
            assert(area.querySelectorAll('[aria-current="step"]').length === 1 &&
                area.querySelector('[aria-current="step"]')?.textContent?.includes("Last"),
                "Persisted navigation did not restore a unique current step.");
            await click(button("Previous"));
            assert(area.textContent?.includes("Middle fixture page"), "Previous skipped a registered page.");
            await click(button("Previous"));
            assert(area.textContent?.includes("First fixture page"), "Previous did not reach the first page.");
            await click(button("Next"));
            assert(area.textContent?.includes("Middle fixture page"), "Next skipped a registered page.");
            await click(button("Next"));
            assert(area.textContent?.includes("Last fixture page"), "Next did not reach the final page.");
            caseObservations = { uniqueRestoredCurrentStep: true, everyPageVisitedInBothDirections: true, liveWrites: 0 };
        } finally {
            if (saved === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, saved);
        }
    }],
    ["Data Agent matched full-root rendered V1 contract states", () => observeWizardContext(async () => {
        const key = "iot-wizard-v4-dataagent-render-contract", savedStorage = sessionStorage.getItem(key), savedFetch = window.fetch;
        const priorStyle = area.getAttribute("style");
        const snapshots: Record<string, unknown>[] = [];
        let loading = false, failInventory = false, deferCreate = false;
        let release: (() => void) | undefined;
        const columns = [{ Name: "deviceId", CslType: "string" }, { Name: "event_time", CslType: "datetime" },
            { Name: "temperature", CslType: "real" }, { Name: "firmware", CslType: "string" }];
        const doc = "Matched fixture device events.\n@iot.meta tableRole=modeled-data\n" +
            '- temperature — "Temperature" (telemetry, double, unit: degC): Declared measurement.\n' +
            '- firmware — "Firmware" (property, string): Declared version.';
        window.fetch = async (input, init) => {
            const address = String(input), body = init?.body ? JSON.parse(String(init.body)) : undefined;
            if (address.includes("/items") && (!init?.method || init.method === "GET") && loading) {
                return new Promise(resolve => { release = () => resolve(new Response(JSON.stringify({ value: [] }))); });
            }
            if (address.includes("/items") && (!init?.method || init.method === "GET") && failInventory) {
                return new Response("Matched inventory failure", { status: 403 });
            }
            if ((address === "https://mock.kusto.fabric.microsoft.com/v1/rest/mgmt" ||
                address === "https://mock.kusto.fabric.microsoft.com/v1/rest/query") && typeof body?.csl === "string" && body.csl.startsWith(".show")) {
                if (body.csl === ".show tables details | project TableName, DocString") return result(
                    [["TableName", "string"], ["DocString", "string"]], [["MatchedData", doc]]);
                assert(body.csl === ".show table ['MatchedData'] schema as json" || body.csl === ".show table MatchedData schema as json",
                    "Render contract attempted an unrelated source read.");
                return result([["TableName", "string"], ["Schema", "string"]], [["MatchedData", JSON.stringify({ OrderedColumns: columns })]]);
            }
            if (address === "https://api.fabric.microsoft.com/v1/workspaces/workspace/dataAgents" && init?.method === "POST") {
                if (deferCreate) return new Promise(resolve => { release = () => resolve(new Response(JSON.stringify({ id: "44444444-4444-4444-8444-444444444444" }), { status: 201 })); });
                return new Response(JSON.stringify({ id: "44444444-4444-4444-8444-444444444444" }), { status: 201 });
            }
            return savedFetch(input, init);
        };
        const capture = async (state: string, selector = ".wizard-control") => {
            const wizard = area.querySelector<HTMLElement>(selector)!;
            assert(wizard, "Actual root Wizard did not render.");
            await new Promise(resolve => setTimeout(resolve, 250));
            const frozenAnimations: string[] = [];
            for (const animation of document.getAnimations()) {
                if (!(animation.effect instanceof KeyframeEffect) || !(animation.effect.target instanceof Element) ||
                    !wizard.contains(animation.effect.target)) continue;
                const timing = animation.effect.getTiming();
                if (timing.iterations === Infinity) {
                    animation.pause(); animation.currentTime = 0;
                    frozenAnimations.push(animation.effect.target.tagName);
                }
            }
            const origin = wizard.getBoundingClientRect();
            const properties = ["display", "position", "width", "height", "minWidth", "maxWidth", "minHeight", "maxHeight",
                "paddingTop", "paddingRight", "paddingBottom", "paddingLeft", "marginTop", "marginRight", "marginBottom", "marginLeft",
                "gap", "rowGap", "columnGap", "fontFamily", "fontSize", "fontWeight", "lineHeight", "letterSpacing",
                "color", "backgroundColor", "borderTopWidth", "borderTopStyle", "borderTopColor", "borderRadius",
                "whiteSpace", "overflowX", "overflowY", "boxSizing", "flexDirection", "alignItems", "justifyContent", "gridTemplateColumns"];
            const entries: Record<string, unknown>[] = [];
            const visit = (node: Node, path: string) => {
                if (node.nodeType === Node.TEXT_NODE) {
                    const text = (node.textContent || "").replace(/\s+/g, " ").trim();
                    if (text) entries.push({ path, node: "#text", text });
                    return;
                }
                if (!(node instanceof Element)) return;
                const style = getComputedStyle(node), rect = node.getBoundingClientRect();
                const attrs = Object.fromEntries([...node.attributes].filter(attribute => !["style"].includes(attribute.name))
                    .map(attribute => [attribute.name, attribute.name === "class" ? attribute.value.split(/\s+/).filter(Boolean).sort().join(" ") : attribute.value]));
                const measured = { x: Math.round((rect.x - origin.x) * 100) / 100, y: Math.round((rect.y - origin.y) * 100) / 100,
                    width: Math.round(rect.width * 100) / 100, height: Math.round(rect.height * 100) / 100 };
                entries.push({ path, node: node.tagName, attrs, style: Object.fromEntries(properties.map(property =>
                    [property, style[property as keyof CSSStyleDeclaration]])), rect: measured,
                    ...(node instanceof HTMLInputElement ? { inputValue: node.value, disabled: node.disabled } : {}) });
                if (node.tagName === "PRE") {
                    entries.push({ path: path + "/generated", node: "#approved-generated-preview", text: "[APPROVED GENERATED CONTENT]" });
                    return;
                }
                [...node.childNodes].forEach((child, index) => visit(child, `${path}/${index}`));
            };
            visit(wizard, "wizard");
            const addedSelectors: string[] = JSON.parse(process.env.AGENT_RENDER_EXTRA_SELECTORS || "[]");
            const matchingAddedSelectors: string[] = [];
            for (const selector of addedSelectors) {
                const potential = selector.replace(/::[\w-]+(?:\([^)]*\))?/g, "")
                    .replace(/:(hover|focus|focus-visible|active|focus-within)\b/g, "");
                if (wizard.matches(potential) || wizard.querySelector(potential)) matchingAddedSelectors.push(selector);
            }
            snapshots.push({ state, rootDimensions: { width: origin.width, height: origin.height }, entries,
                actualText: (wizard.textContent || "").replace(/\s+/g, " ").trim(), generatedPreview: wizard.querySelector("pre")?.textContent || "",
                mode: process.env.AGENT_RENDER_REFERENCE_MODE || "current", frozenInfiniteAnimationTargets: frozenAnimations,
                matchingAddedSelectors });
        };
        const mountState = async (step: string, completed: string[], context: Record<string, unknown> = {}) => {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: step, completedSteps: completed, wizardContext: {
                agentEhId: "eh", agentEhName: "Eventhouse", agentDbId: "db", agentDbName: "Database",
                agentQueryServiceUri: "https://mock.kusto.fabric.microsoft.com", agentTableName: "MatchedData",
                agentTableDocstring: doc, agentColumns: columns.map(column => ({ name: column.Name, type: column.CslType })),
                agentName: "Matched Agent", agentSuffix: "fixed", ...context,
            } }));
            await mount(<IoTSolutionItemDataAgentWizard workloadClient={client}
                item={{ id: "render-contract", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onBack={() => undefined} onWizardComplete={() => { throw new Error("No live saves in render conformance."); }}/>);
        };
        try {
            area.style.width = "1200px"; area.style.height = "900px";
            await document.fonts.ready;
            const editorKey = "iot-wizard-v4-dataagent-entry-render";
            const previousEditorState = sessionStorage.getItem(editorKey);
            sessionStorage.removeItem(editorKey);
            try {
                const rootClient: WorkloadClientAPI = { ...client, itemCrud: {
                    getItem: async () => entryMetadata("entry-render"),
                    getItemDefinition: async () => emptyAppDefinition(),
                    updateItemDefinition: async () => { throw new Error("No item saves in rendered comparison."); },
                } };
                await mount(<FluentProvider theme={webLightTheme} style={{ height: "100%" }}>
                    <App history={createMemoryHistory({ initialEntries: ["/IoTSolutionItem-editor/entry-render"] })} workloadClient={rootClient}/>
                </FluentProvider>, true);
                await until(() => !!area.querySelector('[aria-label="Set up Data Agent"]'), "Actual editor entry card did not load.");
                await capture("editor-entry-card", '[aria-label="Set up Data Agent"]');
                await click(area.querySelector('[aria-label="Set up Data Agent"]'));
                await until(() => !!area.querySelector(".wizard-control"), "Actual editor wrapper did not open Data Agent.");
                await capture("editor-agent-overview");
            } finally {
                if (previousEditorState === null) sessionStorage.removeItem(editorKey); else sessionStorage.setItem(editorKey, previousEditorState);
            }
            await mountState("overview", []); await capture("overview");
            await mountState("data-source", ["overview"], { agentEhId: "", agentEhName: "", agentDbId: "", agentDbName: "",
                agentQueryServiceUri: "", agentTableName: "", agentColumns: [], agentSourceSelected: false });
            await until(() => !area.querySelector(".fui-Spinner"), "Empty source inventory did not settle."); await capture("source-empty");
            loading = true; release = undefined;
            await mountState("data-source", ["overview"], { agentEhId: "", agentEhName: "", agentDbId: "", agentDbName: "",
                agentQueryServiceUri: "", agentTableName: "", agentColumns: [], agentSourceSelected: false });
            await until(() => !!release && !!area.querySelector(".fui-Spinner"), "Loading source state did not render."); await capture("source-loading");
            loading = false; release!(); await settle();
            failInventory = true;
            await mountState("data-source", ["overview"], { agentEhId: "", agentEhName: "", agentDbId: "", agentDbName: "",
                agentQueryServiceUri: "", agentTableName: "", agentColumns: [], agentSourceSelected: false });
            await until(() => !!area.querySelector(".iot-solution-error-text"), "Matched source error did not render."); await capture("source-inventory-error");
            failInventory = false;
            await mountState("data-source", ["overview"]);
            await until(() => !!currentContext.agentSourceSelected && !area.querySelector(".fui-Spinner"), () =>
                "Inspected source did not settle: " + JSON.stringify({ context: currentContext,
                    text: area.querySelector(".iot-solution-step")?.textContent, requests: calls.slice(-8) })); await capture("source-inspected");
            await mountState("configure-agent", ["overview", "data-source"]);
            await until(() => !!area.querySelector("pre"), "Configure content did not render."); await capture("configure-ready");
            deferCreate = true; release = undefined;
            await click(button("Create Data Agent"));
            await until(() => !!release && !!currentContext.agentCreating, "Matched creating state did not render."); await capture("configure-creating");
            deferCreate = false; release!(); await until(() => !!currentContext.agentCreated, "Matched created state did not settle."); await capture("configure-created");
            await mountState("summary", ["overview", "data-source", "configure-agent"], { agentCreated: true, agentId: "synthetic-render-agent" });
            await capture("summary");
            await mountState("configure-agent", ["overview", "data-source", "configure-agent"], { agentCreated: true, agentId: "synthetic-render-agent" });
            await capture("configure-backtracked");
            caseObservations = { mode: process.env.AGENT_RENDER_REFERENCE_MODE || "current", snapshots,
                actualRoot: true, matchedInputs: true, completeDynamicDomAndComputedStyles: true, liveWrites: 0 };
        } finally {
            window.fetch = savedFetch;
            if (priorStyle === null) area.removeAttribute("style"); else area.setAttribute("style", priorStyle);
            if (savedStorage === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, savedStorage);
        }
    })],
    ["Activator actual root uses shared declared model names for all resources and preserves manual names", () => observeWizardContext(async () => {
        const key = "iot-wizard-v4-activator-activator-naming", savedStorage = sessionStorage.getItem(key), savedFetch = window.fetch;
        let source = dataAgentSourceFixture("wot");
        const writes: { resource: string; displayName: string; body: Record<string, any> }[] = [];
        const formats: Record<string, unknown>[] = [];
        window.fetch = async (input, init) => {
            const address = String(input), body = init?.body ? JSON.parse(String(init.body)) : undefined;
            if (address === "https://api.fabric.microsoft.com/v1/workspaces/workspace" && (!init?.method || init.method === "GET")) {
                return new Response(JSON.stringify({ workspaceIdentity: { principalId: "synthetic-existing-workspace-identity" } }));
            }
            if (address === "https://mock.kusto.fabric.microsoft.com/v1/rest/mgmt") {
                if (body.csl === ".show tables details | project TableName, DocString") return result(
                    [["TableName", "string"], ["DocString", "string"]], [[source.table, source.docstring]]);
                assert(body.csl === `.show table ${source.table} schema as json`, "Naming introduced a new source query.");
                return result([["TableName", "string"], ["Schema", "string"]], [[source.table, JSON.stringify({
                    OrderedColumns: source.columns.map(column => ({ Name: column.name, CslType: column.type })),
                })]]);
            }
            if (init?.method === "POST") {
                const resource = address === "https://api.fabric.microsoft.com/v1/connections" ? "connection"
                    : address === "https://api.fabric.microsoft.com/v1/workspaces/workspace/items" && body.type === "Reflex" ? "activator"
                    : address === "https://api.fabric.microsoft.com/v1/workspaces/workspace/eventstreams" ? "eventstream" : undefined;
                assert(resource, "Naming attempted an unexpected resource/provisioning operation.");
                writes.push({ resource: resource!, displayName: body.displayName, body });
                return new Response(JSON.stringify({ id: `synthetic-${resource}-${writes.length}` }), { status: 201 });
            }
            return savedFetch(input, init);
        };
        const mountRoot = async (context: Record<string, unknown>) => {
            sessionStorage.setItem(key, JSON.stringify({ currentStepId: "select-table", completedSteps: ["overview"], wizardContext: {
                actEventhouseId: "eh", actEventhouseName: "Eventhouse", actDatabaseId: "db", actDatabaseName: "Database",
                actQueryServiceUri: "https://mock.kusto.fabric.microsoft.com", actTableName: source.table, actNameSuffix: "review", ...context,
            } }));
            await mount(<IoTSolutionItemActivatorWizard workloadClient={client}
                item={{ id: "activator-naming", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                onWizardComplete={() => { throw new Error("No live item saves are allowed in naming tests."); }}/>);
            await until(() => !!currentContext.actResourceNameBase && !area.querySelector(".fui-Spinner"), "Activator source/name defaults did not settle.");
        };
        const reachResources = async () => {
            await click(next());
            await until(() => !!currentContext.actWorkspaceIdentityReady && !next()!.disabled, "Existing mocked workspace identity did not verify.");
            await click(next());
            await until(() => !!area.querySelector("#act-activator-name"), "Actual resource-name fields did not open.");
        };
        try {
            for (const kind of ["wot", "dtdl"] as const) {
                source = dataAgentSourceFixture(kind);
                const base = kind === "wot" ? "WoT_fixture_device" : "DTDL_fixture_device";
                await mountRoot({});
                assert(currentContext.actResourceNameBase === base &&
                    currentContext.actActivatorName === `Activator_${base}_review` &&
                    currentContext.actEventstreamName === `Activator_Stream_${base}_review`,
                    "Actual root defaults used an ID or changed prefixes/suffix.");
                await reachResources();
                const activator = area.querySelector<HTMLInputElement>("#act-activator-name")!.value;
                const eventstream = area.querySelector<HTMLInputElement>("#act-eventstream-name")!.value;
                assert(activator === `Activator_${base}_review` && eventstream === `Activator_Stream_${base}_review`,
                    "Resource fields did not retain declared-name defaults.");
                const before = writes.length;
                await click(button("Set Up Activator"));
                await until(() => !!currentContext.activatorCreated, "Mocked actual Activator creation did not finish.");
                const created = writes.slice(before);
                assert(JSON.stringify(created.map(({ resource, displayName }) => ({ resource, displayName }))) === JSON.stringify([
                    { resource: "connection", displayName: `KQL-${base}-review` },
                    { resource: "activator", displayName: activator }, { resource: "eventstream", displayName: eventstream },
                ]), "Serialized resource names differ from shared model-name defaults or visible fields.");
                const definition = created.find(write => write.resource === "eventstream")!.body.definition;
                const part = definition.parts.find((part: { path: string }) => part.path === "eventstream.json");
                const topology = JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(part.payload), character => character.charCodeAt(0))));
                const columns = topology.destinations[0].inputSchemas[0].schema.columns;
                assert(columns.some((column: { name: string; type: string }) => column.name === "event_time" && column.type === "DateTime") &&
                    !columns.some((column: { name: string }) => column.name === "enqueuedTime") &&
                    created[0].body.connectionDetails.type === "AzureDataExplorer" &&
                    created[0].body.credentialDetails.credentials.credentialType === "WorkspaceIdentity",
                    "Naming changed the native time/type/connection contract.");
                formats.push({ kind, base, activator, eventstream, connection: created[0].displayName,
                    actualRoot: true, exactSerializedNames: true, eventTimeSchemaUnchanged: true, liveWrites: 0 });
            }
            source = { ...dataAgentSourceFixture("wot"), table: "Selected_Table",
                docstring: '@iot.meta modelSemanticInfo={"id":"urn:fixture:no-declared-name"}\n@iot.meta tableRole=modeled-data' };
            await mountRoot({ actActivatorName: "Operator monitor" });
            assert(currentContext.actActivatorName === "Operator monitor" &&
                currentContext.actEventstreamName === "Activator_Stream_Selected_Table_review" &&
                currentContext.actResourceNameBase === "Selected_Table",
                "Missing sibling default overwrote an entered name or missing model name used the URI.");
            await reachResources();
            await enter("act-activator-name", "Operator edited monitor");
            await enter("act-eventstream-name", "Operator edited stream");
            await click(button("Previous")); await click(button("Previous"));
            await until(() => !!area.querySelector("#act-table-dropdown") && !area.querySelector(".fui-Spinner"), "Source rerender did not settle.");
            await reachResources();
            assert(area.querySelector<HTMLInputElement>("#act-activator-name")!.value === "Operator edited monitor" &&
                area.querySelector<HTMLInputElement>("#act-eventstream-name")!.value === "Operator edited stream",
                "Returning through the source step overwrote manually edited resource names.");
            caseObservations = { formats, actualRoot: true, tableFallback: true, eachManualNameIndependentlyPreserved: true,
                manualNamesSurviveSourceRerender: true, resourcesNamed: ["connection", "activator", "eventstream"],
                liveWrites: 0, identityProvisioning: 0 };
        } finally {
            window.fetch = savedFetch;
            if (savedStorage === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, savedStorage);
        }
    })],
    ["Data Agent V1-style inspected WoT and DTDL previews match saved native instructions and sample pairs", async () => {
        const Agent: TestStep = props => <DataAgentSetupStep {...props} phase="configure"/>;
        const savedFetch = window.fetch;
        const outputs: Record<string, unknown>[] = [];
        let initialShape = "";
        try {
            for (const kind of ["wot", "dtdl"] as const) {
                const source = dataAgentSourceFixture(kind), sourceBefore = JSON.stringify(source);
                const expected = buildDataAgentContext(source), reads = calls.length, auth = tokenRequests;
                const expectedName = kind === "wot" ? "DataAgent_WoT_fixture_device_review" : "DataAgent_DTDL_fixture_device_review";
                let parts: { path: string; payload: string; payloadType: string }[] | undefined;
                window.fetch = async (input, init) => {
                    assert(String(input) === "https://api.fabric.microsoft.com/v1/workspaces/workspace/dataAgents" && init?.method === "POST",
                        "Agent generation attempted an extra lookup, model call, sample query or live operation.");
                    const body = JSON.parse(String(init!.body));
                    assert(body.displayName === expectedName, "Saved agent name used a model ID or differed from the declared-name default.");
                    parts = body.definition.parts;
                    return new Response(JSON.stringify({ id: kind === "wot" ? "55555555-5555-4555-8555-555555555555" :
                        "66666666-6666-4666-8666-666666666666" }), { status: 201 });
                };
                await mount(<Harness component={Agent} initial={{ agentEhId: "eh", agentDbId: "db", agentDbName: source.database,
                    agentTableName: source.table, agentQueryServiceUri: "https://mock.kusto.fabric.microsoft.com",
                    agentSuffix: "review", agentTableDocstring: source.docstring, agentColumns: source.columns }}/>);
                const pre = area.querySelector("pre")!, preview = pre.textContent!;
                const expectedPreview = expected.agentInstructions + "\n\n" + expected.sourceInstructions +
                    `\n## Examples (${expected.examples.length})\n\n` + dataAgentExamplesMarkdown(expected.examples);
                assert(preview === expectedPreview && !button("Create Data Agent")!.disabled &&
                    calls.length === reads && tokenRequests === auth, "Actual configure preview differed or performed extra reads/auth.");
                assert(area.querySelector<HTMLInputElement>("#agent-name")!.value === expectedName &&
                    !preview.includes("Use each configured source's documented model") && !preview.includes("not the IoT Hub enqueue time.") &&
                    !preview.includes("Declared ID:") && preview.includes('deviceId == "YOUR_DEVICE_ID"') &&
                    !preview.includes("<device-id>"), "Requested copy/name/placeholder corrections did not reach actual preview.");
                const rect = (needle: string) => {
                    const walker = document.createTreeWalker(pre, NodeFilter.SHOW_TEXT);
                    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
                        const offset = (node.textContent || "").indexOf(needle);
                        if (offset < 0) continue;
                        const range = document.createRange();
                        range.setStart(node, offset); range.setEnd(node, offset + needle.length);
                        return range.getBoundingClientRect();
                    }
                    throw new Error(`Missing rendered preview text ${needle}.`);
                };
                const question = rect("Q: " + expected.examples[1].question), query = rect(expected.examples[1].query);
                const field = rect("`temperature` (real)"), fact = rect(kind === "dtdl" ? 'DTDL schema: "double"' : 'WoT type: "number"');
                const whiteSpace = getComputedStyle(pre).whiteSpace;
                assert(whiteSpace === "pre-wrap" && query.top > question.bottom && fact.top > field.bottom,
                    "Actual preformatted preview collapsed question/query or field/fact into one rendered line.");
                const shape = JSON.stringify([...area.querySelector(".iot-solution-step")!.querySelectorAll("*")].map(node => ({
                    tag: node.tagName, className: node.getAttribute("class"), style: node.getAttribute("style"), role: node.getAttribute("role"),
                })));
                if (kind === "wot") initialShape = shape;
                else assert(shape === initialShape, "Metadata changed wizard controls/layout/styles rather than just preview content.");
                await click(button("Create Data Agent"));
                await until(() => !!parts && currentContext.agentCreated, "Actual agent configuration was not serialized.");
                const decoded = parts!.map(part => ({ path: part.path,
                    value: JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(part.payload), char => char.charCodeAt(0)))) }));
                const payload = (name: string) => {
                    const part = decoded.find(part => part.path === name);
                    assert(part, `Missing native definition part ${name}.`);
                    return part!.value;
                };
                assert(decoded.length === 8 && new Set(decoded.map(part => part.path)).size === 8, "Unexpected native definition path set.");
                for (const stage of ["draft", "published"]) {
                    const prefix = `Files/Config/${stage}`, datasource = payload(`${prefix}/kusto-${source.database}/datasource.json`);
                    const samples = payload(`${prefix}/kusto-${source.database}/fewshots.json`).fewShots;
                    const group = datasource.elements[0], table = group.children[0];
                    assert(datasource.elements.length === 1 && group.type === "table_grouping" && group.display_name === "Tables" &&
                        group.is_selected === false && group.children.length === 1 && table.type === "kusto.table" &&
                        table.is_selected === true && table.display_name === source.table && group.id !== table.id &&
                        JSON.stringify(table.children.map((column: { id: string; display_name: string; is_selected: boolean }) =>
                            [column.id, column.display_name, column.is_selected])) === JSON.stringify(source.columns.map(column =>
                                [column.name, column.name, true])) &&
                        new Set([group.id, table.id, ...table.children.map((column: { id: string }) => column.id)]).size === source.columns.length + 2,
                        `${stage} datasource retained a phantom root, omitted native Tables membership or changed selected column identities.`);
                    assert(payload(`${prefix}/stage_config.json`).aiInstructions === expected.agentInstructions &&
                        datasource.dataSourceInstructions === expected.sourceInstructions &&
                        table.children.every((column: { description: string | null; type: string; data_type: string; children: unknown[] }, index: number) =>
                            !column.description?.includes("@iot.meta") && column.type === "kusto.column" &&
                            (column.data_type === "System.SByte") === (source.columns[index].type === "bool") && column.children.length === 0) &&
                        JSON.stringify(samples.map(({ question, query }: { question: string; query: string }) => ({ question, query }))) === JSON.stringify(expected.examples) &&
                        samples.every((sample: { id: string }) => typeof sample.id === "string") &&
                        samples.every((sample: { query: string }) => !sample.query.includes("```") && !sample.query.includes("<device-id>")) &&
                        samples[1].query.includes('deviceId == "YOUR_DEVICE_ID"') &&
                        new Set(samples.map((sample: { id: string }) => sample.id)).size === samples.length,
                        `${stage} native metadata or real question/KQL pairs differed from preview.`);
                }
                assert(JSON.stringify(payload(`Files/Config/draft/kusto-${source.database}/datasource.json`)) ===
                    JSON.stringify(payload(`Files/Config/published/kusto-${source.database}/datasource.json`)),
                    "Draft and published native table grouping/selection differ.");
                assert(JSON.stringify(source) === sourceBefore && calls.length === reads, "Agent generation mutated docs or queried source data.");
                outputs.push({ kind, source, preview, context: expected, nativeParts: decoded, nativeSamples: expected.examples,
                    agentName: expectedName, renderedLineBoundaries: { questionBottom: question.bottom, queryTop: query.top,
                        fieldBottom: field.bottom, factTop: fact.top, whiteSpace },
                    sourceReads: 0, modelCalls: 0, liveWrites: 0 });
            }
            const fallback = dataAgentSourceFixture("wot");
            await mount(<Harness component={Agent} initial={{ agentEhId: "eh", agentDbId: "db", agentDbName: fallback.database,
                agentTableName: "Selected_Table", agentSuffix: "review", agentColumns: fallback.columns,
                agentTableDocstring: '@iot.meta modelSemanticInfo={"id":"urn:fixture:no-name"}\n@iot.meta tableRole=modeled-data' }}/>);
            assert(area.querySelector<HTMLInputElement>("#agent-name")!.value === "DataAgent_Selected_Table_review",
                "A missing declared name fell back to the model URI rather than the table.");
            await enter("agent-name", "Operator chosen agent");
            await settle();
            assert(area.querySelector<HTMLInputElement>("#agent-name")!.value === "Operator chosen agent" &&
                currentContext.agentName === "Operator chosen agent", "Default naming overwrote a manually entered name.");
            caseObservations = { formats: outputs, identicalConfigureLayoutAndControls: true, exactPreviewAndSavedInstructions: true,
                bothDraftAndPublishedFewshots: true, nativeTablesGrouping: true, noDuplicateRootTable: true,
                declaredNameDefaults: true, tableNameFallback: true, manualNamePreserved: true,
                actualPreformattedLineBoundaries: true, sourceReads: 0, modelCalls: 0, liveWrites: 0 };
        } finally { window.fetch = savedFetch; }
    }],
    ["Event time consumers require event_time and reject old schemas before Activator serialization", async () => {
        const request = creationRequest(), plan = buildModeledDataPlan(request.input, request.coverage, true, request.names);
        const Agent: TestStep = props => <DataAgentSetupStep {...props} phase="configure"/>;
        const savedFetch = window.fetch;
        const names: string[] = [];
        try {
            for (const time of ["event_time"] as const) {
                const columns = [{ name: "deviceId", type: "string", docstring: plan.documentation.columns.deviceId },
                    { name: time, type: "datetime", docstring: plan.documentation.columns.event_time },
                    { name: "temperature", type: "real", docstring: plan.documentation.columns.temperature }];
                await mount(<Harness component={Agent} initial={{ agentEhId: "eh", agentDbId: "db", agentDbName: "Database", agentTableName: "Inspected",
                    agentQueryServiceUri: "https://mock.kusto.fabric.microsoft.com",
                    agentName: "Inspected Agent", agentTableDocstring: plan.documentation.table, agentColumns: columns }}/>);
                assert(area.textContent?.includes(`arg_max(${time},`) && area.textContent?.includes(`bin(${time}, 5m)`) &&
                    !button("Create Data Agent")!.disabled, "Actual agent configuration did not use its inspected time column.");
                let topology: Record<string, any> | undefined;
                window.fetch = async (input, init) => {
                    const address = String(input);
                    if (address === "https://mock.kusto.fabric.microsoft.com/v1/rest/mgmt") {
                        assert(JSON.parse(String(init?.body)).csl === ".show table Inspected schema as json", "Unexpected Activator schema read.");
                        return result([["TableName", "string"], ["Schema", "string"]], [["Inspected", JSON.stringify({
                            OrderedColumns: columns.map(column => ({ Name: column.name, CslType: column.type })),
                        })]]);
                    }
                    if (address === "https://api.fabric.microsoft.com/v1/workspaces/workspace/eventstreams" && init?.method === "POST") {
                        const body = JSON.parse(String(init.body));
                        const part = body.definition.parts.find((part: { path: string }) => part.path === "eventstream.json");
                        topology = JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(part.payload), char => char.charCodeAt(0))));
                        return new Response(JSON.stringify({ id: "synthetic-inspected-eventstream" }), { status: 201 });
                    }
                    throw new Error("The time-column consumer case permits only its synthetic schema/read and serialization requests.");
                };
                await mount(<Harness component={ActivatorCreateResourcesStep} initial={{ actTableName: "Inspected", actDatabaseName: "Database",
                    actQueryServiceUri: "https://mock.kusto.fabric.microsoft.com", actActivatorName: "Fixture Activator",
                    actEventstreamName: "Fixture Stream", actConnectionId: "fixture-connection", actActivatorId: "fixture-activator" }}/>);
                await click(button("Set Up Activator"));
                await until(() => !!topology && currentContext.activatorCreated, "Activator did not serialize its inspected schema.");
                const serialized = topology!.destinations[0].inputSchemas[0].schema.columns;
                assert(serialized.some((column: { name: string; type: string }) => column.name === time && column.type === "DateTime") &&
                    !serialized.some((column: { name: string }) => column.name === "enqueuedTime"),
                    "Activator invented or renamed an existing table's time column.");
                names.push(time); window.fetch = savedFetch;
            }
            await mount(<Harness component={Agent} initial={{ agentEhId: "eh", agentDbId: "db", agentDbName: "Database", agentTableName: "OldSchema",
                agentQueryServiceUri: "https://mock.kusto.fabric.microsoft.com",
                agentName: "Old Schema Agent", agentTableDocstring: plan.documentation.table, agentSourceSelected: true,
                agentColumns: [{ name: "enqueuedTime", type: "datetime" }] }}/>);
            await until(() => !!area.querySelector(".fui-MessageBar")?.textContent?.includes("event_time:datetime"), "An unsupported old schema did not surface through the existing error box.");
            assert(button("Create Data Agent")!.disabled && currentContext.agentSourceSelected === false, "An old schema silently passed source/configuration guards.");
            let oldEventstreamWrites = 0;
            window.fetch = async (input, init) => {
                if (String(input) === "https://mock.kusto.fabric.microsoft.com/v1/rest/mgmt") return result(
                    [["TableName", "string"], ["Schema", "string"]], [["OldSchema", JSON.stringify({
                        OrderedColumns: [{ Name: "deviceId", CslType: "string" }, { Name: "enqueuedTime", CslType: "datetime" }],
                    })]]);
                oldEventstreamWrites++;
                throw new Error("Old time schemas cannot generate a new Eventstream definition.");
            };
            await mount(<Harness component={ActivatorCreateResourcesStep} initial={{ actTableName: "OldSchema", actDatabaseName: "Database",
                actQueryServiceUri: "https://mock.kusto.fabric.microsoft.com", actActivatorName: "Fixture Activator",
                actEventstreamName: "Fixture Stream", actConnectionId: "fixture-connection", actActivatorId: "fixture-activator" }}/>);
            await click(button("Set Up Activator"));
            await until(() => !!area.textContent?.includes("event_time:datetime"), "Activator accepted an old time schema.");
            assert(!currentContext.activatorCreated && oldEventstreamWrites === 0, "Activator serialized an unsupported old schema.");
            caseObservations = { inspectedTimeColumns: names, actualAgentExamples: true, actualActivatorSerialization: true,
                unsupportedOldSchemaUsesExistingError: true, noOldTimeCompatibility: true, liveWrites: 0 };
        } finally { window.fetch = savedFetch; }
    }],
    ["Model and field semantic uploads persist exact native metadata without rendering raw payloads or crediting prior metadata", () => observeWizardContext(async () => {
        const central = { "@type": ["ModelDefinition", "DeviceModel"], capabilityModel: {
            "@id": "dtmi:fixture:SemanticUpload;1", "@type": "Interface", displayName: { en: "Semantic upload", de: "Semantisches Modell" },
            description: "Actual root model meaning", contents: [
                { "@type": ["Telemetry", "Temperature"], name: "temperature", displayName: "Temperature", schema: "double", unit: "degreeCelsius" },
                { "@type": "Property", name: "state", description: "Reported state", writable: false, schema: "string" },
            ],
        } };
        const observations: Record<string, unknown>[] = [];
        for (const kind of ["wot", "dtdl"] as const) {
            creationBackend = undefined;
            await wizard(context(true));
            await click(area.querySelector(`input[value="${kind}"]`));
            await file("thing-model-upload", JSON.stringify(kind === "wot" ? { ...tm, version: { model: "1.0" } } : central), `${kind}-semantic.json`);
            await until(() => currentContext.hubModel?.kind === kind && !next()!.disabled, "Actual model upload did not preserve valid declared roles.");
            await reachReview();
            const request = { input: hubInput(currentContext), coverage: currentContext.hubCoverage,
                allowWaiting: true, names: currentContext.hubOutputNames };
            const plan = buildModeledDataPlan(request.input, request.coverage, true, request.names);
            assert(plan.documentation.columns.event_time && !plan.documentation.columns.enqueuedTime &&
                plan.commands.filter(command => command.kind === "table" || command.kind === "view").every(command =>
                    "columns" in command && command.columns.some(column => column.name === "event_time" && column.type === "datetime") &&
                    !command.columns.some(column => column.name === "enqueuedTime")),
                "Actual new plans retained an old or duplicate system time column.");
            const modelInfo = JSON.parse(parseIotMeta(plan.documentation.table).modelSemanticInfo);
            const objectDocs = plan.commands.flatMap(command => "docstring" in command ? [command] : []);
            assert(objectDocs.length === 3 && objectDocs.every(command =>
                canonicalJson(JSON.parse(parseIotMeta(command.docstring).modelSemanticInfo)) === canonicalJson(modelInfo) &&
                !command.docstring.includes("@iot.meta deviceModel=") && !command.docstring.includes("@iot.meta modelVersion=")),
                "The three actual object docs did not share the selected canonical model metadata.");
            assert(plan.documentation.table.includes(MODEL_SEMANTIC_INFO_PREFIX) &&
                !/\n(?:title:|description:|@iot.meta (?:deviceModel|modelVersion)=)/.test(plan.documentation.table) &&
                (kind === "wot" ? modelInfo.id === tm.id && modelInfo.version === "1.0" && modelInfo.title === "Device"
                    : modelInfo["@id"] === central.capabilityModel["@id"] && modelInfo.displayName.en === "Semantic upload" &&
                        modelInfo.description === "Actual root model meaning" && modelInfo.version === undefined),
                "Actual model upload did not produce the agreed modelSemanticInfo native object.");
            const reads = calls.length;
            assert(Object.values(plan.documentation.columns).every(doc => doc.startsWith(FIELD_SEMANTIC_INFO_PREFIX) &&
                doc.length <= 1000) && objectDocs.every(command => command.docstring.length <= 1000) &&
                plan.documentation.table.length <= 1000 && !area.textContent?.includes(MODEL_SEMANTIC_INFO_PREFIX) &&
                !area.querySelector("[data-docstring-target]") &&
                !area.querySelector('[aria-label="Exact docstring transport commands"]') && calls.length === reads,
                "The creation page changed bounded native metadata or exposed raw semantic payloads or commands.");
            if (kind === "dtdl") {
                const temperature = JSON.parse(plan.documentation.columns.temperature.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
                const state = JSON.parse(plan.documentation.columns.state.slice(FIELD_SEMANTIC_INFO_PREFIX.length));
                assert(temperature.schema === "double" && temperature.unit === "degreeCelsius" &&
                    JSON.stringify(temperature["@type"]) === '["Telemetry","Temperature"]' && state.writable === false,
                    "The actual IoT Central upload lost native DTDL declarations before Create.");
            }
            creationBackend = createCreationMock(request);
            await click(button("Create modeled data"));
            await until(() => currentContext.hubCreation?.phase === "completed", "Current semantic JSON did not pass actual creator readback.");
            assert(Object.entries(plan.documentation.columns).every(([name, doc]) =>
                creationBackend!.tables.get(request.names.modeled)!.columnDocs[name] === doc) &&
                creationBackend.tables.get(request.names.modeled)!.doc === plan.documentation.table &&
                creationBackend.tables.get(request.names.twinChanges)!.doc === objectDocs.find(command => command.target === request.names.twinChanges)!.docstring &&
                creationBackend.views.get(request.names.twinState)!.doc === objectDocs.find(command => command.target === request.names.twinState)!.docstring &&
                creationBackend.writes.length === 9 && !next()!.disabled,
                "Current-plan bounded column JSON was not the actual written/read-back metadata.");
            await click(next());
            await until(() => !!button("Finish and Close"), "The current semantic JSON creation did not reach Ready.");
            assert(!button("Finish and Close")!.disabled, "Current confirmed JSON outputs could not finish.");
            const completed = currentContext.hubCreation, old = JSON.parse(completed.fingerprint);
            old[3] = "all-object-model-semantic-info-json-v7";
            const before = calls.length;
            await act(async () => observedWizardProps!.updateContext("hubCreation", { ...completed, fingerprint: JSON.stringify(old) }));
            assert(button("Finish and Close")!.disabled && calls.length === before,
                "Old native metadata completion credit passed Ready or triggered automatic replay.");
            await act(async () => observedWizardProps!.updateContext("hubCreation", completed));
            assert(!button("Finish and Close")!.disabled && calls.length === before,
                "Restoring the actual current operation altered native resources or lost current completion.");
            observations.push({ kind, modelInfo, exactThreeObjectAndColumnStoredPayloads: true, nineWrites: true, rawSemanticPayloadsNotRendered: true,
                priorCreditRejectedWithoutReplay: true, actualCloudWrites: 0 });
        }
        caseObservations = { formats: observations };
    })],
    ["Dashboard actual wizard accepts both DO Hub profiles without visual changes and rejects non-DO", () => observeWizardContext(async () => {
        const savedFetch = window.fetch, key = "iot-wizard-v5-dashboard-dashboard-format";
        const savedStorage = sessionStorage.getItem(key);
        const savedResourceResponse = resourceResponse, savedDatabaseItemIds = databaseItemIds;
        const dashboardDatabaseId = "11111111-1111-4111-8111-111111111111";
        databaseItemIds = [dashboardDatabaseId];
        resourceResponse = async () => new Response(JSON.stringify({ value: [
            { id: "eh", displayName: "Eventhouse", type: "Eventhouse", workspaceId: "workspace" },
            { id: dashboardDatabaseId, displayName: "Database", type: "KQLDatabase", workspaceId: "workspace" },
        ] }));
        const inspectionQueries: string[] = [];
        let format: "native" | "cloud" | "empty-native" | "non-do" | "malformed" | "partial" = "cloud";
        let capturedDefinition: Record<string, any> | undefined;
        const rows = () => {
            const profile = format === "native" || format === "empty-native" ? format : "cloud";
            const observation = {
                Telemetry: [dashboardRow({ temperature: 21, state: "ok" }, { do: format !== "non-do", profile }),
                    dashboardRow({ temperature: 23 }, { do: format !== "non-do", profile, component: "motor" })],
                Twins: [dashboardRow({ properties: { reported: { firmware: "1", motor: { __t: "c", target: 7 }, $metadata: {}, $version: 2 } } },
                    { do: format !== "non-do", profile, stream: "twin" })],
            };
            if (format === "malformed") {
                const envelope = JSON.parse(String(observation.Telemetry[0].data)); envelope.schema_version = "99";
                observation.Telemetry[0].data = JSON.stringify(envelope);
            }
            return observation;
        };
        window.fetch = async (input, init) => {
            const address = String(input), body = init?.body ? JSON.parse(String(init.body)) : undefined;
            if (address === "https://api.fabric.microsoft.com/v1/workspaces/workspace/items" && init?.method === "POST") {
                assert(body?.type === "KQLDashboard" && body.definition.parts[0].path === "RealTimeDashboard.json",
                    "Actual dashboard generation did not serialize its incumbent definition part.");
                capturedDefinition = JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(body.definition.parts[0].payload), char => char.charCodeAt(0))));
                return new Response(JSON.stringify({ id: "synthetic-dashboard-no-live-write" }), { status: 201 });
            }
            if (address === "https://mock.kusto.fabric.microsoft.com/v1/rest/query" && typeof body?.csl === "string" &&
                body.csl.startsWith("let dashboard_raw")) {
                inspectionQueries.push(body.csl);
                const values = replayDashboardQuery(body.csl, rows(), {}, dashboardClock, true);
                const columns = body.csl.includes("IsNumericType") ? ["Key", "IsNumericType", "Component", "Leaf", "Error"]
                    : ["Key", "Component", "Leaf", "Error"];
                return new Response(JSON.stringify({ Tables: [
                    { Columns: columns.map(ColumnName => ({ ColumnName })), Rows: values.map(row => columns.map(column => row[column])) },
                    ...(format === "partial" ? [{ TableName: "QueryStatus", Columns: [{ ColumnName: "Severity" }, { ColumnName: "StatusCode" }], Rows: [[2, -1]] }] : []),
                ] }));
            }
            return savedFetch(input, init);
        };
        const signature = () => JSON.stringify([...area.querySelector(".iot-solution-step")!.querySelectorAll("*")].map(node => ({
            tag: node.tagName, className: node.getAttribute("class"), style: node.getAttribute("style"), role: node.getAttribute("role"),
            type: node.getAttribute("type"), text: node.childElementCount ? null : node.textContent, disabled: node.hasAttribute("disabled"),
            checked: node instanceof HTMLInputElement ? node.checked : undefined,
        })));
        let initialSignature = "";
        try {
            for (const kind of ["cloud", "empty-native"] as const) {
                format = kind;
                sessionStorage.setItem(key, JSON.stringify({ currentStepId: "data-source", completedSteps: ["overview"], wizardContext: {
                    dashEhId: "eh", dashEhName: "Eventhouse", dashDbId: dashboardDatabaseId, dashDbName: "Database",
                    dashQueryServiceUri: "https://mock.kusto.fabric.microsoft.com", dashTelemetryTable: "Telemetry", dashPropertiesTable: "Twins",
                } }));
                await mount(<IoTSolutionItemDashboardWizard workloadClient={client}
                    item={{ id: "dashboard-format", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                    onWizardComplete={() => { throw new Error("No real item saves or dashboard writes are permitted."); }}/>);
                await until(() => !!button("Inspect Tables") && !button("Inspect Tables")!.disabled &&
                    !area.querySelector("#dash-telemetry-table")?.hasAttribute("disabled"), "Dashboard source inventory did not settle.");
                const before = inspectionQueries.length;
                await click(button("Inspect Tables"));
                await until(() => !!currentContext.dashInspected, "Actual dashboard inspection failed.");
                assert(inspectionQueries.length === before + 2, "Inspection changed its explicit two-query contract.");
                assert(currentContext.dashTelemetryFields.length === 3 && currentContext.dashPropertyFields.length === 2 &&
                    currentContext.dashTelemetryFields.filter((field: { isNumeric: boolean }) => field.isNumeric).length === 2 &&
                    currentContext.dashTelemetryFields.some((field: { component?: string; leaf?: string }) => field.component === "motor" && field.leaf === "temperature") &&
                    currentContext.dashPropertyFields.some((field: { component?: string; leaf?: string }) => field.component === "motor" && field.leaf === "target"),
                    "DO normalization lost field counts, numeric flags or exact components.");
                assert(!area.querySelector(".iot-solution-field-list")?.textContent?.includes("schema_id") &&
                    area.textContent?.includes("2 time-series (numeric), 1 non-time-series") && !next()!.disabled,
                    "Envelope fields or incorrect numeric badges leaked into the actual wizard.");
                if (kind === "cloud") initialSignature = signature();
                else assert(signature() === initialSignature, "DO Hub transport profile changed DOM structure, styles, badges or controls for equal fields.");
            }
            format = "native";
            await click(button("Re-inspect Tables"));
            await until(() => !!area.querySelector(".fui-MessageBar")?.textContent?.includes("CloudEvents"),
                "Setup certified a native-only identity that generated CloudEvents queries do not consume.");
            assert(!currentContext.dashInspected && next()!.disabled, "An unsupported transport contract passed source inspection.");
            format = "non-do";
            await click(button("Inspect Tables"));
            await until(() => !!area.querySelector(".fui-MessageBar")?.textContent?.includes("Unsupported DO observation"), "Non-DO data was incorrectly accepted.");
            assert(!currentContext.dashInspected && next()!.disabled && !area.querySelector(".iot-solution-field-list"),
                "Unsupported legacy raw data retained successful fields or continuation.");
            format = "malformed";
            await click(button("Inspect Tables"));
            await until(() => !!area.querySelector(".fui-MessageBar")?.textContent?.includes("Unsupported DO observation"), "Malformed DO did not reach the existing error box.");
            assert(!currentContext.dashInspected && next()!.disabled && !area.querySelector(".iot-solution-field-list"),
                "Failed inspection retained successful fields or continuation.");
            format = "partial";
            await click(button("Inspect Tables"));
            await until(() => !!area.querySelector(".fui-MessageBar")?.textContent?.includes("Partial results were discarded"), "Partial native query status became success or empty fields.");
            format = "cloud";
            await click(button("Inspect Tables"));
            await until(() => !!currentContext.dashInspected, "Explicit dashboard inspection retry failed.");
            await click(next());
            await until(() => !!button("Generate Dashboard"), "Actual generation step did not open.");
            await click(button("Generate Dashboard"));
            await until(() => !!capturedDefinition && !!currentContext.dashDashboardCreated, "Mock generation did not serialize the dashboard.");
            const definition = capturedDefinition!;
            assert(definition.tiles.length === 9 && definition.queries.length === 6 && definition.parameters.length === 2 &&
                definition.tiles.filter((tile: { visualType: string }) => tile.visualType === "line").length === 2 &&
                definition.tiles.filter((tile: { visualType: string }) => tile.visualType === "table").length === 1 &&
                definition.tiles.filter((tile: { visualType: string }) => tile.visualType === "card").length === 2,
                "Generated definition changed incumbent data visuals/queries or omitted a selected field.");
            const decorations = definition.tiles.filter((tile: { visualType: string }) => tile.visualType === "designTile");
            assert(definition.schema_version === 88 && definition.embeddedApps.length === 0 &&
                definition.dataSources[0].databaseArtifactId === dashboardDatabaseId &&
                definition.dataSources[0].database === dashboardDatabaseId &&
                decorations.length === 4 && decorations.every((tile: Record<string, any>) =>
                    tile.hideTitle === true && tile.title === "" && tile.pageId === definition.pages[0].id &&
                    !("queryRef" in tile) && !("dataSource" in tile)) &&
                JSON.stringify(decorations.map((tile: Record<string, any>) => [tile.layout, tile.visualOptions])) === JSON.stringify([
                    [{ x: 0, y: 0, width: 18, height: 3 }, { designTile__kind: "textBox",
                        textBox__html: '<div style="font-size: 24px;">Device dashboard</div><div>(Choose a device ID above to get started)</div>' }],
                    [{ x: 0, y: 3, width: 24, height: 1 }, { designTile__kind: "divider", divider__label: "Time-series telemetry" }],
                    [{ x: 0, y: 11, width: 24, height: 1 }, { designTile__kind: "divider", divider__label: "Non-time-series telemetry (last 10 values)" }],
                    [{ x: 0, y: 19, width: 24, height: 1 }, { designTile__kind: "divider", divider__label: "Device properties (last known value)" }],
                ]),
                "Actual generation did not serialize the supported native intro/dividers without decorative queries.");
            assert(definition.queries.every((query: { text: string }) => !query.text.includes("enqueuedTime")) &&
                definition.tiles.filter((tile: { visualType: string }) => tile.visualType === "line").every((tile: { visualOptions: { xColumn: string } }) =>
                    tile.visualOptions.xColumn === "event_time"), "Serialized dashboard time-axis or queries retained an old time alias.");
            assert(definition.queries.every((query: { text: string; usedVariables: string[] }) =>
                query.text.includes("ingestion_time() between (['_startTime'] .. ['_endTime'])") &&
                !/\blet\b|mv-expand|pack_array|dashboard_error|__eventHubProfile|schema_id/.test(query.text) &&
                query.usedVariables.includes("_startTime") && query.usedVariables.includes("_endTime")) &&
                !definition.queries[0].usedVariables.includes("_deviceId"),
                "Serialized hot queries retained validation machinery, an old horizon or cyclic/missing time dependencies.");
            const values = definition.queries.map((query: { text: string }) => replayDashboardQuery(query.text, rows(),
                { _deviceId: "device-1", _startTime: dashboardClock - 86400000, _endTime: dashboardClock }, dashboardClock, true));
            assert(values[0][0].deviceId === "device-1" && values.slice(1).every((rows: Record<string, unknown>[]) => rows.length === 1) &&
                values.slice(1).some((rows: Record<string, unknown>[]) => rows[0].Value === 23) &&
                values.slice(1).some((rows: Record<string, unknown>[]) => rows[0].Value === "ok") &&
                values.slice(1).some((rows: Record<string, unknown>[]) => rows[0].Value === 7),
                "Serialized device/time-series/recent/property queries did not consume actual unwrapped values.");
            caseObservations = { telemetry: currentContext.dashTelemetryFields, properties: currentContext.dashPropertyFields,
                numeric: 2, nonnumeric: 1, equalDoTransportDomStyles: true, inspectionReadsPerClick: 2,
                unsupportedLegacyRejected: true, cloudEventsAndEmptyNative: true, nativeOnlySetupRejected: true,
                existingErrorBoxAndRetry: true, promotedDynamicDateTime: true, queryPathsReplayed: 6,
                nativeDecorations: 4, serializedDefinition: definition, liveWrites: 0 };
        } finally {
            window.fetch = savedFetch;
            resourceResponse = savedResourceResponse; databaseItemIds = savedDatabaseItemIds;
            await act(async () => root?.unmount()); root = undefined;
            if (savedStorage === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, savedStorage);
        }
    })],
    ["Actual root creation keeps raw semantic metadata out of the UI without changing plan reads or state", () => observeWizardContext(async () => {
        await wizard(context(true)); await reachReview();
        const content = area.querySelector(".wizard-control__content-panel")!;
        assert(content.querySelector("h2")?.textContent === "Create modeled data" &&
            !content.querySelector("pre") && !content.textContent?.includes(FIELD_SEMANTIC_INFO_PREFIX) &&
            !content.textContent?.includes(MODEL_SEMANTIC_INFO_PREFIX) &&
            !button("Create modeled data")!.disabled &&
            content.querySelectorAll(".iot-solution-form input[required]").length === 3,
            "The creation page exposed raw metadata or lost its output-name fields or validation gate.");
        const source = (observedWizardProps as React.ComponentProps<typeof HubModelReviewStep>).sourceState;
        const plan = buildModeledDataPlan(hubInput(currentContext, source), currentContext.hubCoverage, true, currentContext.hubOutputNames);
        const commands = plan.commands.filter(command => command.kind === "tableDoc" || command.kind === "viewDoc" || command.kind === "columnDocs");
        const reads = calls.length, auth = tokenRequests, state = canonicalJson({ scope: currentContext.hubScope,
            fields: currentContext.hubFields, creation: currentContext.hubCreation, names: currentContext.hubOutputNames });
        await settle();
        assert(commands.length === 4 && commands.filter(command => "docstring" in command).length === 3 &&
            Object.values(plan.documentation.columns).every(docstring => docstring.startsWith(FIELD_SEMANTIC_INFO_PREFIX) && docstring.length <= 1000) &&
            plan.documentation.table.includes(MODEL_SEMANTIC_INFO_PREFIX) &&
            !content.textContent?.includes(FIELD_SEMANTIC_INFO_PREFIX) && !content.textContent?.includes(MODEL_SEMANTIC_INFO_PREFIX) &&
            calls.length === reads && tokenRequests === auth && state === canonicalJson({ scope: currentContext.hubScope,
                fields: currentContext.hubFields, creation: currentContext.hubCreation, names: currentContext.hubOutputNames }),
            "Rendering changed bounded native plan metadata or caused reads/state changes.");
        caseObservations = { actualRoot: true, rawSemanticPayloadsNotRendered: true, nativePlanMetadataRetained: true,
            noReadsAuthOrStateReset: true, currentCreationStillValid: true, actualWrites: 0 };
    })],
    ["V1 creation presentation uses actual sources counts required single-column outputs and native hints", async () => {
        const request = creationRequest();
        await mount(<Harness component={HubModelReviewStep} initial={alignmentCreationContext(request)}/>);
        const summary = [...area.querySelectorAll(".iot-solution-validation-row")].map(row => row.textContent);
        assert(JSON.stringify(summary) === JSON.stringify(["Telemetry tableTelemetry", "Properties tableTwins",
            "Selected fields1 telemetry, 1 twin property"]) &&
            area.querySelector(".iot-solution-step-description")?.textContent ===
                "Review the resources this step creates in your KQL database. Together, they turn raw JSON into typed columns and add each device's latest state to its telemetry." &&
            area.querySelector("#output-names-help")?.textContent === "You can rename these resources before you create them" &&
            !area.querySelector(".iot-solution-output-card") && !button("Check creation requirements"),
            "V1 creation intro/native summary/default explanation was reconstructed or lost actual values.");
        const fields = [...area.querySelectorAll<HTMLElement>(".iot-solution-form > .iot-solution-field")];
        assert(JSON.stringify(fields.map(field => field.querySelector("label")?.textContent?.replace(/\s*\*$/, ""))) ===
            JSON.stringify(["Modeled data table", "Normalized properties table", "Latest property values (materialized view)"]) &&
            JSON.stringify(fields.map(field => field.querySelector("input")?.id)) ===
                JSON.stringify(["modeled-output-name", "twin-changes-name", "twin-state-name"]) &&
            fields.every(field => field.querySelector("input")?.required && !!field.querySelector(".iot-solution-field-hint")) &&
            fields.every((field, index) => !index || field.getBoundingClientRect().top >= fields[index - 1].getBoundingClientRect().bottom) &&
            area.querySelector("#modeled-output-name-help")?.textContent ===
                "Your typed data: 1 telemetry column and 1 property column, plus device ID and event time, for each event" &&
            !button("Create modeled data")!.disabled && !area.querySelector(".iot-solution-creation-steps"),
            "V1 required name fields/hints/single-column layout or untouched explicit Create gate changed.");
        assert(!area.textContent?.includes(FIELD_SEMANTIC_INFO_PREFIX), "Raw field metadata was rendered in the creation form.");
        const reads = calls.length, auth = tokenRequests;
        await enter("modeled-output-name", "Modeled_Renamed");
        assert(area.querySelector<HTMLInputElement>("#modeled-output-name")!.value === "Modeled_Renamed" &&
            !area.textContent?.includes(MODEL_SEMANTIC_INFO_PREFIX) &&
            calls.length === reads && tokenRequests === auth, "Editing a name exposed raw metadata or triggered an unexpected read.");
        await enter("modeled-output-name", "");
        assert(button("Create modeled data")!.disabled && !!area.querySelector('[role="alert"]') &&
            !area.textContent?.includes(FIELD_SEMANTIC_INFO_PREFIX), "Invalid plan passed the creation guard or exposed raw metadata.");
        const noTwin = creationRequest(false, false);
        await mount(<Harness component={HubModelReviewStep} initial={alignmentCreationContext(noTwin)}/>);
        assert(!!area.querySelector("#twin-changes-name") && !!area.querySelector("#twin-state-name") &&
            !!area.querySelector("#modeled-output-name") &&
            area.querySelector("#twin-state-name-help")?.textContent === "Keeps each device's latest reported value for every property" &&
            buildModeledDataPlan(noTwin.input, noTwin.coverage, true, noTwin.names).commands.some(command => command.csl.includes("30d")),
            "Presentation changed incumbent editable-name validation or omitted actual planned helper outputs.");
        caseObservations = { actualV1Authority: "bdfa06160a64ebc1151ff736351c7c2dc373fd63", nativeSummaryAndHints: true,
            requiredSingleColumnInputs: true, exactCurrentCountsAndSources: true, invalidNameBlocked: true, noUnusedOutputs: true };
    }],
    ["V1 creation presentation preserves honest partial error rows and confirmed effects without success", async () => {
        const request = creationRequest();
        creationBackend = createCreationMock(request);
        creationBackend.options.fail = csl => csl.startsWith(".create materialized-view")
            ? new Error("Synthetic creation failure after the normalized table.") : undefined;
        await mount(<Harness component={HubModelReviewStep} initial={alignmentCreationContext(request)}/>);
        await click(button("Create modeled data"));
        await until(() => currentContext.hubCreation?.phase === "failed", "Guarded fake partial creation did not fail.");
        const result = currentContext.hubCreation;
        assertCreationDisplayVerbs(buildModeledDataPlan(request.input, request.coverage, true, request.names));
        assert(area.querySelectorAll(".iot-solution-creation-step").length === 8 &&
            !!area.querySelector(".iot-solution-creation-step-label.error") &&
            area.querySelector(".iot-solution-creation-step .iot-solution-field-error")?.textContent?.includes("Synthetic creation failure") &&
            !area.querySelector(".iot-solution-success-text") && result.objects.length > 0 &&
            area.textContent?.includes("Nothing was automatically rolled back.") && result.uncertainCommand &&
            !button("Create modeled data")!.disabled,
            "The V1 presentation hid a real operation/error/uncertain effect or fabricated completed success.");
        creationBackend = undefined;
        caseObservations = { logicalRows: 8, underlyingOperationSteps: result.steps.length, failedRowVisible: true, confirmedObjects: result.objects.length,
            uncertainEffectRetained: true, noFalseSuccessOrRollback: true, actualCloudWrites: 0 };
    }],
    ["Creation input order and confirmed display verbs preserve exact command dependencies and shared Continue", async () => {
        caseObservations = await creationPresentationCheck("test");
    }],
    ["Active overview and results diagrams use decorative Fluent arrows with unchanged stage layout", async () => {
        const diagrams = [
            { name: "ingestion-overview", component: <IngestionOverviewStep workloadClient={client}/>, arrows: 2, plus: 0 },
            { name: "dashboard-overview", component: <DashboardOverviewStep/>, arrows: 2, plus: 0 },
            { name: "model-overview", component: <ModelOverviewStep workloadClient={client}/>, arrows: 3, plus: 0 },
            { name: "data-agent-overview", component: <DataAgentOverviewStep/>, arrows: 2, plus: 1 },
            { name: "activator-overview", component: <ActivatorOverviewStep/>, arrows: 3, plus: 0 },
            { name: "ingestion-results", component: <Harness component={IngestionSummaryStep} initial={{}}/>, arrows: 3, plus: 0 },
            { name: "dashboard-results", component: <Harness component={DashboardSummaryStep} initial={{}}/>, arrows: 3, plus: 0 },
            { name: "data-agent-results", component: <Harness component={DataAgentSummaryStep} initial={{}}/>, arrows: 2, plus: 1 },
            { name: "activator-results", component: <Harness component={ActivatorSummaryStep} initial={{}}/>, arrows: 3, plus: 0 },
        ];
        const reads = calls.length + resourceCalls.length, tokens = tokenRequests, observations: object[] = [];
        for (const test of diagrams) {
            await mount(test.component);
            assertDecorativeDiagramArrows(test.arrows, test.plus);
            if (test.name === "ingestion-overview" || test.name === "model-overview") {
                const prerequisites = [...area.querySelectorAll<HTMLElement>(".iot-solution-prerequisites > ul > li")];
                const link = prerequisites[1]?.querySelector<HTMLAnchorElement>("a");
                assert(prerequisites.length === (test.name === "ingestion-overview" ? 4 : 5) &&
                    prerequisites[1].textContent?.replace(/\s+/g, " ").trim() ===
                        (test.name === "ingestion-overview"
                            ? "Optionally, an Azure Device Registry (ADR) namespace linked to your IoT Hub. Linking an ADR namespace to your IoT Hub lets you assign a model to your devices later, when you model your data."
                            : "Optionally, an Azure Device Registry (ADR) namespace linked to your IoT Hub. Linking an ADR namespace to your IoT Hub allows devices to reference your model.") &&
                    prerequisites[1].querySelectorAll("a").length === 1 &&
                    link?.textContent?.trim() === "linked to your IoT Hub" &&
                    link.getAttribute("href") === "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal" &&
                    link.target === "_blank" && link.rel === "noopener noreferrer" && !link.hasAttribute("aria-disabled") &&
                    (test.name === "ingestion-overview"
                        ? prerequisites[0].textContent?.includes("An existing Azure IoT Hub with a system-assigned or user-assigned managed") &&
                            prerequisites[2].textContent?.includes("An existing Eventhouse and KQL Database")
                        : prerequisites[0].querySelector("strong")?.textContent === "Ingest device data from your IoT Hub" &&
                            prerequisites[2].textContent?.includes("A device model:")),
                    "The optional ADR prerequisite lost exact copy, enabled external link, existing emphasis or requested list order.");
            }
            observations.push({ name: test.name, svgArrows: test.arrows, retainedPlusMarkers: test.plus });
        }
        assert(calls.length + resourceCalls.length === reads && tokenRequests === tokens,
            "Rendering diagram icons triggered source reads or authorization.");
        caseObservations = { diagrams: observations, inheritedIconSize: true, decorativeSvg: true,
            responsiveDirectionPreserved: true, sourceReads: 0, authRequests: 0, liveWrites: 0 };
    }],
    ["IoT Hub refresh uses a transparent accessible Fluent button without changing discovery", async () => {
        let requests = 0, defer = false, release: (() => void) | undefined;
        const hub = { id: "/subscriptions/00000000-0000-0000-0000-000000000001/resourceGroups/fixture/providers/Microsoft.Devices/IotHubs/fixture-hub",
            name: "Fixture hub", resourceGroup: "fixture", subscriptionId: "00000000-0000-0000-0000-000000000001", location: "eastus" };
        assignmentResponse = async (input, init) => {
            assert(String(input).includes("/Microsoft.ResourceGraph/") && init?.method === "POST",
                "Refresh dispatched anything other than Hub discovery.");
            requests++;
            if (defer) await new Promise<void>(resolve => { release = resolve; });
            return new Response(JSON.stringify({ data: [hub] }));
        };
        await mount(<Harness component={IoTHubStep} initial={{
            hubName: hub.name, resourceGroup: hub.resourceGroup, subscriptionId: hub.subscriptionId,
        }}/>);
        await until(() => !!button("Refresh IoT Hubs"), "Hub discovery did not expose its refresh action.");
        const refresh = button("Refresh IoT Hubs")!;
        const style = getComputedStyle(refresh), icon = refresh.querySelector("svg")!;
        assert(refresh.classList.contains("fui-Button") && refresh.textContent === "" &&
            style.backgroundColor === "rgba(0, 0, 0, 0)" && style.borderTopColor === "rgba(0, 0, 0, 0)" &&
            Math.abs(icon.getBoundingClientRect().width - 24) <= 1 &&
            Math.abs(icon.getBoundingClientRect().height - 24) <= 1 && requests === 1,
            `The Hub refresh regained a visible box, text, small icon or duplicate discovery: ${JSON.stringify({
                background: style.backgroundColor, border: style.borderTopColor, text: refresh.textContent,
                width: icon.getBoundingClientRect().width, height: icon.getBoundingClientRect().height, requests,
            })}`);
        await act(async () => refresh.focus());
        assert(document.activeElement === refresh, "The Fluent refresh action lost native keyboard focus.");
        defer = true;
        await click(refresh);
        await until(() => !!release, "Refresh did not invoke the existing Hub loader.");
        assert(requests === 2 && !button("Refresh IoT Hubs") && area.textContent?.includes("Loading IoT Hubs..."),
            "An in-flight refresh kept an actionable duplicate control or lost its existing loading feedback.");
        defer = false; await act(async () => release!());
        await until(() => !!button("Refresh IoT Hubs"), "Hub refresh did not recover after the existing loader settled.");
        assert(area.querySelector<HTMLInputElement>("#hub-combobox")?.value === hub.name && requests === 2,
            "Refresh discarded the prior Hub selection or repeated discovery.");
        caseObservations = { transparentFluentButton: true, iconPx: 24, nativeKeyboardFocus: true,
            accessibleName: "Refresh IoT Hubs", initialAndExplicitReadCount: requests, retainedHubSelection: true,
            existingLoadingGatePreserved: true, liveWrites: 0 };
    }],
    ["Active wizard roots use sentence-case titles and navigation while preserving product names", async () => {
        const fixture = { id: "sentence-case", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>;
        const noSave = () => { throw new Error("Copy verification must not save an item."); };
        const cases = [
            { kind: "ingestion", key: "iot-wizard-v5-ingestion-sentence-case",
                component: <IoTSolutionItemDefaultView workloadClient={client} item={fixture} onWizardComplete={noSave}/>,
                title: "Ingest device data from your IoT Hub",
                hero: "Connect your IoT data to Microsoft Fabric",
                steps: ["Overview", "IoT Hub", "Eventhouse", "Raw data tables", "Eventstreams", "IoT Hub routing", "Ready to use"] },
            { kind: "dashboard", key: "iot-wizard-v5-dashboard-sentence-case",
                component: <IoTSolutionItemDashboardWizard workloadClient={client} item={fixture} onWizardComplete={noSave}/>,
                title: "Generate Real-Time Dashboards", hero: "Turn live device data into operational visibility",
                steps: ["Overview", "Data source", "Generate dashboard", "Ready to use"] },
            { kind: "data-agent", key: "iot-wizard-v4-dataagent-sentence-case",
                component: <IoTSolutionItemDataAgentWizard workloadClient={client} item={fixture} onWizardComplete={noSave}/>,
                title: "Set up Data Agent", hero: "Explore device operations using AI",
                steps: ["Overview", "Modeled data source", "Configure Data Agent", "Ready to use"] },
            { kind: "activator", key: "iot-wizard-v4-activator-sentence-case",
                component: <IoTSolutionItemActivatorWizard workloadClient={client} item={fixture} onWizardComplete={noSave}/>,
                title: "Set up Activator rules", hero: "Generate alerts and automated actions based on live IoT data",
                steps: ["Overview", "Monitoring data", "Workspace identity", "Activator resources", "Ready to configure"] },
            { kind: "modeling", key: "",
                component: <IoTSolutionItemModelWizard workloadClient={client} item={fixture}/>,
                title: "Model your device data", hero: "Turn raw IoT data into modeled data",
                steps: ["Get started", "Select data source", "Add device model", "Assign model",
                    "Check model coverage", "Create modeled data", "Review results"] },
        ];
        const reads = calls.length + resourceCalls.length, tokens = tokenRequests, observations: object[] = [];
        for (const test of cases) {
            const previous = test.key ? sessionStorage.getItem(test.key) : null;
            if (test.key) sessionStorage.removeItem(test.key);
            try {
                await mount(test.component);
                assert(area.querySelector(".wizard-control__title")?.textContent === test.title &&
                    area.querySelector("h2")?.textContent?.trim() === test.hero &&
                    JSON.stringify([...area.querySelectorAll(".wizard-control__step-title")].map(node => node.textContent)) ===
                        JSON.stringify(test.steps) && area.querySelectorAll('[aria-current="step"]').length === 1 &&
                    !next()!.disabled, `The ${test.kind} root changed its fixed copy, product casing or initial navigation.`);
                observations.push({ kind: test.kind, title: test.title, steps: test.steps });
            } finally {
                await act(async () => root?.unmount()); root = undefined;
                if (test.key) {
                    if (previous === null) sessionStorage.removeItem(test.key);
                    else sessionStorage.setItem(test.key, previous);
                }
            }
        }
        assert(calls.length + resourceCalls.length === reads && tokenRequests === tokens,
            "Mounting fixed overview copy caused source queries or authorization.");
        caseObservations = { roots: observations, initialNavigationPreserved: true, sourceReads: 0, authRequests: 0, liveWrites: 0 };
    }],
    ["Model sidebar preserves ordered steps numbering and Finish navigation", () => observeWizardContext(async () => {
        const pairs = [
            ["Get started", "See what modeling unlocks"],
            ["Select data source", "Choose your raw IoT tables"],
            ["Add device model", "Upload a model and choose fields"],
            ["Assign model", "Optional: link devices to the model"],
            ["Check model coverage", "Pick devices and confirm coverage"],
            ["Create modeled data", "Build typed telemetry with device state"],
            ["Review results", "See what was created and next steps"],
        ];
        const ids = ["overview", "raw-data", "device-model", "assign-references", "scope-coverage", "review-plan", "configured"];
        let closed = 0;
        await wizard(context(true), () => { closed++; });
        const checkStep = (index: number) => {
            const rows = [...area.querySelectorAll(".wizard-control__step")];
            assert(JSON.stringify(rows.map(row => [row.querySelector(".wizard-control__step-title")?.textContent,
                row.querySelector(".wizard-control__step-description")?.textContent])) === JSON.stringify(pairs) &&
                observedWizardProps?.step.id === ids[index] && observedWizardProps.stepIndex === index &&
                rows[index].getAttribute("aria-current") === "step" &&
                rows.slice(1, 6).every((row, number) => row.querySelector(".wizard-control__step-checkmark") ||
                    row.querySelector(".wizard-control__step-number")?.textContent === String(number + 1)) &&
                !rows[0].querySelector(".wizard-control__step-number") && !rows[6].querySelector(".wizard-control__step-number"),
                `Sidebar copy changed a registered step identity, order, numbering or current position at ${ids[index]}.`);
        };
        checkStep(2);
        await click(step("Get started")); checkStep(0);
        assertDecorativeDiagramArrows(3);
        assert(area.querySelector("h2")?.textContent?.trim() === "Turn raw IoT data into modeled data",
            "A sidebar title replaced the separate modeling overview heading.");
        await click(next());
        await until(tableChoicesReady, "Source metadata did not settle on overview Continue.");
        if (next()!.disabled) await click(sourceValidationAction());
        await until(() => !next()!.disabled, "Current source analysis did not enable Continue.");
        checkStep(1);
        await click(next()); checkStep(2);
        await click(next()); checkStep(3);
        assert(next()!.disabled && button("Skip assignment"), "Sidebar wording bypassed the explicit assignment choice.");
        await click(button("Skip assignment")); checkStep(4);
        await verifyCoverage();
        await until(() => !next()!.disabled, "Current scoped coverage did not enable Continue.");
        await click(next()); checkStep(5);
        const source = (observedWizardProps as React.ComponentProps<typeof HubModelReviewStep>).sourceState;
        const request = { input: hubInput(currentContext, source), coverage: currentContext.hubCoverage,
            allowWaiting: true, names: currentContext.hubOutputNames };
        creationBackend = createCreationMock(request);
        await click(button("Create modeled data"));
        await until(() => currentContext.hubCreation?.phase === "completed", "Mock guarded creation did not complete.");
        await click(next()); checkStep(6);
        assertDecorativeDiagramArrows(3);
        await click(button("Finish and Close"));
        assert(closed === 1 && creationBackend.writes.length === 9, "Sidebar wording changed Finish or guarded creation.");
        caseObservations = { exactOrderedPairs: pairs, stableStepIds: ids, numberedSetupSteps: 5,
            overviewAndReadyUnnumbered: true, guardedFakeWrites: 9, finishBackCalls: closed, liveWrites: 0 };
    })],
    ["Model overview describes modeled data benefits requirements and creation flow", async () => {
        const reads = calls.length + resourceCalls.length, tokens = tokenRequests;
        await mount(<ModelOverviewStep workloadClient={client}/>);
        assert(area.querySelector("h2")?.textContent?.trim() === "Turn raw IoT data into modeled data" &&
            JSON.stringify([...area.querySelectorAll("h3")].map(node => node.textContent)) ===
                JSON.stringify(["What modeling unlocks", "How the modeled dataset is created", "What you\u2019ll need"]) &&
            area.querySelectorAll(".iot-solution-step-description").length === 2 &&
            JSON.stringify([...area.querySelectorAll(".iot-solution-step-description")].map(node => node.textContent?.trim())) ===
                JSON.stringify([
                    "Use a device model to turn the telemetry and property fields you choose into typed, business-friendly columns. Each telemetry event also gets the device's latest state from its twin properties.",
                    "Modeling is optional. Your raw IoT data stays available for analytics and reporting.",
                ]) &&
            area.querySelectorAll(".iot-solution-outcome-list")[0].children.length === 4 &&
            JSON.stringify([...area.querySelectorAll(".iot-solution-outcome-list")[0].children].map(node => node.textContent?.trim())) ===
                JSON.stringify([
                    "Compare measurements with device state, such as temperature by operating mode or vibration by firmware version.",
                    "Query typed columns instead of parsing raw JSON in every query.",
                    "Set more precise alerts that account for device configuration and status.",
                    "Build dashboards, Activator rules, reports, and AI experiences on the same modeled data.",
                ]) &&
            area.querySelectorAll(".iot-solution-prerequisites li").length === 5 &&
            JSON.stringify([...area.querySelectorAll(".iot-solution-prerequisites li")].map(node => node.textContent?.trim().replace(/\s+/g, " "))) ===
                JSON.stringify([
                    "Raw telemetry and properties tables from the Ingest device data from your IoT Hub experience.",
                    "Optionally, an Azure Device Registry (ADR) namespace linked to your IoT Hub. Linking an ADR namespace to your IoT Hub allows devices to reference your model.",
                    "A device model: a Web of Things (WoT) Thing Model (recommended), a Digital Twins Definition Language (DTDL) model, or an IoT Central device template export.",
                    "At least one device event in your raw tables (more events give better suggestions).",
                    "Permission to read raw data and to create tables, update policies, and a materialized view in the KQL database.",
                ]) &&
            area.querySelectorAll(".iot-solution-prerequisites strong").length === 1 &&
            area.querySelector(".iot-solution-prerequisites li strong")?.textContent === "Ingest device data from your IoT Hub" &&
            area.querySelectorAll(".iot-solution-prerequisites a").length === 1 &&
            area.querySelector(".iot-solution-prerequisites li:nth-child(2) a")?.textContent?.trim() === "linked to your IoT Hub" &&
            area.querySelector<HTMLAnchorElement>(".iot-solution-prerequisites a")?.getAttribute("href") ===
                "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal" &&
            area.querySelector<HTMLAnchorElement>(".iot-solution-prerequisites a")?.target === "_blank" &&
            area.querySelector<HTMLAnchorElement>(".iot-solution-prerequisites a")?.rel === "noopener noreferrer",
            "The supplied heading/paragraphs or preserved sections and requirements changed.");
        assert(JSON.stringify([...area.querySelectorAll(".iot-solution-flow-node-title")].map(node => node.textContent)) ===
            JSON.stringify(["Raw IoT tables", "Device model", "Current device state", "Modeled data"]) &&
            JSON.stringify([...area.querySelectorAll(".iot-solution-flow-node-detail")].map(node => node.textContent?.trim())) ===
                JSON.stringify(["Telemetry and property JSON", "Names, types, units, and components",
                    "Latest property values", "Typed telemetry enriched with properties"]) &&
            area.querySelectorAll(".iot-solution-flow-arrow[aria-hidden=true]").length === 3 &&
            !!area.querySelector('.iot-solution-overview-flow[aria-label="Modeled data creation flow"]') &&
            calls.length + resourceCalls.length === reads && tokenRequests === tokens && !button("Cancel"),
            "V1 four-stage presentation was lost or presentation introduced reads/navigation.");
        assertDecorativeDiagramArrows(3);
        caseObservations = { exactUserHeadingAndTwoParagraphs: true, exactFourOrderedBenefits: true,
            latestPropertyValuesCaption: true, fourStages: true,
            exactUserPrerequisitesWithIngestionEmphasis: true, navigationAndRemainingStructureUntouched: true, reads: 0 };
    }],
    ["Thing Model repeated imports surface a bounded expansion error and allow a valid replacement", () => observeWizardContext(async () => {
        const seed = context(true); delete seed.hubModel; seed.hubFields = [];
        await wizard(seed);
        const source = currentContext.hubSourceAnalysis, reads = analysisQueries().length, tokens = tokenRequests;
        const oversized = localRepeatedReferenceModel(7), original = JSON.stringify(oversized);
        assert(new Blob([original]).size < 3000, "The bounded fanout fixture is unexpectedly large.");
        await file("thing-model-upload", original, "bounded-imports.json");
        await until(() => !!area.querySelector("#model-upload-error"), "Expansion-budget failure did not reach the existing upload error surface.");
        assert(area.querySelector("#model-upload-error-message")?.textContent?.includes("expanded Thing Model exceeds the supported node limit (100000)") &&
            !currentContext.hubModel && currentContext.hubFields.length === 0 && next()!.disabled &&
            currentContext.hubSourceAnalysis === source && analysisQueries().length === reads && tokenRequests === tokens,
            "The expansion bound exposed partial fields, discarded source analysis or performed unrelated reads.");
        await file("thing-model-upload", JSON.stringify(localReferenceWorkModel(8000)), "complex-imports.json");
        await until(() => !!area.querySelector("#model-upload-error"), "Reference-heavy work failure did not reach the upload error surface.");
        assert(area.querySelector("#model-upload-error-message")?.textContent ===
            "Thing Model dependency resolution exceeds the supported work limit. Simplify the model or reduce repeated imports." &&
            !currentContext.hubModel && currentContext.hubFields.length === 0 && next()!.disabled,
            "Work-budget guidance exposed a raw counter or partial model state.");
        const largeLiteral = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
            id: "urn:fixture:expanded-byte-guidance", schemaDefinitions: { value: { type: "string", description: "x".repeat(256 * 1024) } },
            properties: Object.fromEntries(Array.from({ length: 10 }, (_, index) =>
                [`field${index}`, { "tm:ref": "#/schemaDefinitions/value", title: `Reading ${index}` }])) };
        await file("thing-model-upload", JSON.stringify(largeLiteral), "large-imports.json");
        await until(() => !!area.querySelector("#model-upload-error"), "Expanded byte failure did not reach the upload error surface.");
        assert(area.querySelector("#model-upload-error-message")?.textContent ===
            "The expanded Thing Model exceeds the supported size (2 MB). Reduce repeated imports or split the model." &&
            !currentContext.hubModel && currentContext.hubFields.length === 0 && next()!.disabled,
            "Expanded-size guidance did not use the uploader's 2 MB units or exposed partial fields.");
        const replacement = localRepeatedReferenceModel(2), replacementText = JSON.stringify(replacement);
        await file("thing-model-upload", replacementText, "small-imports.json");
        await until(() => !!currentContext.hubModel, "A supported replacement could not recover from the bounded resolution error.");
        assert(!area.querySelector("#model-upload-error") && currentContext.hubFields.length === 1 &&
            currentContext.hubFields[0].kustoType === "dynamic" &&
            JSON.stringify(currentContext.hubModel.document) === replacementText &&
            !("resolution" in currentContext.preparationThingModel) &&
            currentContext.hubSourceAnalysis === source && analysisQueries().length === reads && tokenRequests === tokens,
            "Replacement lost whole-object/original-document semantics, persisted diagnostic counters or caused reads.");
        caseObservations = { smallInputBounded: true, explicitNodeLimitError: true, workErrorWithoutRawCounter: true,
            expandedSizeUsesUploadUnits: true, noPartialFields: true,
            validReplacementRecovered: true, sourceAnalysisRetained: true, sourceReads: 0, authRequests: 0, liveWrites: 0 };
    })],
    ["TM ignored parts show one native warning directly under accepted upload and clean replacements clear stale diagnostics", async () => {
        const partial = { ...tm, properties: { ...tm.properties, empty: { type: null },
            "Ignored choice": { oneOf: [{ type: "number" }, { type: "string" }] } },
            events: { controllerOnline: { data: { type: null } } },
            links: [{ rel: "tm:submodel", href: "https://unused.invalid/never-fetch", instanceName: "motor" }] };
        await mount(<Harness component={ModelPreparationStep} initial={context(true)}/>);
        const reads = calls.length + resourceCalls.length, tokens = tokenRequests;
        const tileBefore = area.querySelector<HTMLElement>(".iot-solution-model-upload-area")!;
        const tileStyle = getComputedStyle(tileBefore);
        const identity = [tileStyle.padding, tileStyle.borderRadius, tileStyle.fontFamily, tileStyle.fontSize];
        await file("thing-model-upload", JSON.stringify(partial), "partial.tm.json");
        const tile = area.querySelector<HTMLElement>(".iot-solution-model-upload-area--loaded")!;
        const warning = area.querySelector<HTMLElement>("#model-upload-warnings")!;
        const fieldHeading = area.querySelector("#model-fields-heading")!;
        const intro = "Some parts of your model aren't supported and won't be included:";
        const items = [...warning.querySelectorAll("ul > li")];
        assert(tile && warning && tile.nextElementSibling === warning &&
            !!(warning.compareDocumentPosition(fieldHeading) & Node.DOCUMENT_POSITION_FOLLOWING) &&
            area.querySelectorAll("#model-upload-warnings").length === 1 && !area.querySelector("#model-upload-error") &&
            warning.querySelector(".fui-MessageBarBody")?.firstElementChild?.textContent === intro &&
            warning.textContent?.split(intro).length === 2 &&
            JSON.stringify(items.map(item => item.textContent?.trim())) === JSON.stringify([
                "Fields with no type: property empty, event controllerOnline",
                "Fields that use oneOf: property Ignored choice",
                "Submodels: motor",
            ]) && items.every(item => !item.textContent?.trim().startsWith("Ignored ")) &&
            getComputedStyle(items[0]).fontSize === "14px" && !warning.textContent?.includes("unused.invalid") &&
            !warning.querySelector("button") && area.querySelectorAll("[data-field-id]").length === 2,
            "The warning lost exact placement, brief actual names/reasons, valid siblings or native action-free presentation.");
        const currentStyle = getComputedStyle(tile);
        assert(JSON.stringify([currentStyle.padding, currentStyle.borderRadius, currentStyle.fontFamily, currentStyle.fontSize]) ===
            JSON.stringify(identity) && JSON.stringify(currentContext.hubModel.document) === JSON.stringify(partial),
            "The upload tile changed visually or projection mutated the uploaded original.");
        await file("thing-model-upload", JSON.stringify(tm), "clean.tm.json");
        assert(!area.querySelector("#model-upload-warnings") && !!area.querySelector(".iot-solution-model-upload-area--loaded"),
            "A clean replacement retained another model's warning.");
        for (const boundary of ["format", "source", "unmount"] as const) {
            await mount(<Harness component={ModelPreparationStep} initial={context(true)}/>);
            let release: (value: string) => void = () => undefined;
            const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
            const pending = new File(["{}"], "late-partial.tm.json");
            Object.defineProperty(pending, "text", { value: () => new Promise<string>(resolve => { release = resolve; }) });
            const descriptor = Object.getOwnPropertyDescriptor(input, "files");
            Object.defineProperty(input, "files", { configurable: true, value: [pending] });
            await act(async () => input.dispatchEvent(new Event("change", { bubbles: true })));
            if (boundary === "format") await click(area.querySelector('input[value="dtdl"]'));
            else if (boundary === "source") await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
            else await act(async () => setStepVisible(false));
            await act(async () => release(JSON.stringify(partial)));
            if (boundary === "unmount") await act(async () => setStepVisible(true));
            if (descriptor) Object.defineProperty(input, "files", descriptor); else Reflect.deleteProperty(input, "files");
            assert(!area.querySelector("#model-upload-warnings"), `Late ignored diagnostics crossed the ${boundary} boundary.`);
        }
        await mount(<Harness component={ModelPreparationStep} initial={context()}/>);
        assert(!area.querySelector("#model-upload-warnings"), "DTDL inherited a WoT warning.");
        assert(calls.length + resourceCalls.length === reads && tokenRequests === tokens,
            "Parsing warnings caused data reads or remote dependency/auth discovery.");
        caseObservations = { nativeWarningDirectlyAfterTileBeforeFields: true, validSiblings: 2, groupedNamesAndReasons: true,
            originalAndTileStylesPreserved: true, cleanAndDtdlNoWarning: true, sourceFormatUnmountLateCallbacksIgnored: true, reads: 0 };
    }],
    ["TM warning spacing keeps desktop single and narrow wrapped bullet line boxes inside native block padding", async () => {
        const previousWidth = area.style.width;
        const observations: object[] = [];
        try {
            const fixtures = innerWidth >= 1100 ? [["desktop-single", 1000]] as const : [["narrow-wrapped", 320]] as const;
            for (const [name, width] of fixtures) {
                area.style.width = `${width}px`;
                const document = { ...tm,
                    ...(name === "narrow-wrapped" ? {
                        properties: { ...tm.properties,
                            "Unavailable reading with a deliberately long technical identity": { type: null },
                            "Alternative measurement with another long descriptive identity": { oneOf: [{ type: "number" }, { type: "boolean" }] } },
                        events: { "Controller online status with a long event name": { data: { type: null } } },
                    } : {}),
                    links: [{ rel: "tm:submodel", href: "urn:fixture:unused", instanceName: name === "desktop-single"
                        ? "motor" : "Motor submodel with a long descriptive instance name" }] };
                await mount(<Harness component={ModelPreparationStep} initial={context(true)}/>);
                await file("thing-model-upload", JSON.stringify(document), `${name}.tm.json`);
                await settle();
                const warning = area.querySelector<HTMLElement>("#model-upload-warnings")!;
                const body = warning.querySelector<HTMLElement>(".fui-MessageBarBody")!;
                const intro = body.firstElementChild as HTMLElement;
                const list = body.querySelector<HTMLElement>("ul")!;
                const items = [...list.querySelectorAll<HTMLElement>("li")];
                const last = items[items.length - 1];
                const box = warning.getBoundingClientRect(), bodyBox = body.getBoundingClientRect();
                const introBox = intro.getBoundingClientRect(), lastBox = last.getBoundingClientRect();
                const style = getComputedStyle(warning), listStyle = getComputedStyle(list);
                const topGap = introBox.top - box.top - parseFloat(style.borderTopWidth);
                const bottomGap = box.bottom - parseFloat(style.borderBottomWidth) - lastBox.bottom;
                const range = window.document.createRange(); range.selectNodeContents(last);
                const lines = [...range.getClientRects()];
                assert(items.length === (name === "desktop-single" ? 1 : 3) &&
                    topGap > 0 && Math.abs(topGap - bottomGap) < 1 &&
                    Math.abs(parseFloat(style.paddingTop) - parseFloat(style.paddingBottom)) < 1 &&
                    Math.abs(topGap - parseFloat(style.paddingTop)) < 1 &&
                    Math.abs(bottomGap - parseFloat(style.paddingBottom)) < 1 &&
                    lastBox.bottom <= bodyBox.bottom + 1 && lines.every(line => line.bottom <= box.bottom) &&
                    warning.scrollWidth <= warning.clientWidth + 1 &&
                    (name !== "narrow-wrapped" || items.some(item => item.getBoundingClientRect().height > parseFloat(getComputedStyle(item).lineHeight))),
                    "The focused warning fixture did not wrap honestly or a bullet escaped/clipped its body.");
                observations.push({ name, viewport: innerWidth, containerWidth: width, warningWidth: box.width,
                    warningHeight: box.height, topGap, bottomGap, paddingTop: style.paddingTop, paddingBottom: style.paddingBottom,
                    bodyDisplay: getComputedStyle(body).display, bodyHeight: bodyBox.height,
                    introLineHeight: getComputedStyle(intro).lineHeight, listMarginTop: listStyle.marginTop,
                    listMarginBottom: listStyle.marginBottom, lastLiHeight: lastBox.height,
                    lastLineBottomGap: box.bottom - parseFloat(style.borderBottomWidth) - Math.max(...lines.map(line => line.bottom)),
                    clipped: false, overflow: false });
            }
        } finally { area.style.width = previousWidth; }
        caseObservations = { observations };
    }],
    ["TM ignored-only upload keeps actual wizard Continue blocked and supported remaining fields still require roles", () => observeWizardContext(async () => {
        await wizard(context(true));
        const partial = { ...tm, properties: { remaining: { type: "number" }, unavailable: { type: "null" } }, events: {} };
        await file("thing-model-upload", JSON.stringify(partial), "remaining.tm.json");
        assert(area.querySelectorAll("[data-field-id]").length === 1 && !!area.querySelector("#model-upload-warnings") && next()!.disabled,
            "Ignored fields or sample hints fabricated a role for the remaining supported field.");
        const field = currentContext.hubFields[0] as ModelField;
        await enter(documentElementRole(field.id)!.id, "telemetry");
        assert(!next()!.disabled, "An explicit valid remaining role could not satisfy the model gate.");
        await file("thing-model-upload", JSON.stringify({ ...partial,
            properties: { unavailable: { type: null }, alternative: { oneOf: [{ type: "boolean" }] } } }), "ignored-only.tm.json");
        await until(() => area.querySelector("#model-upload-error-message")?.textContent?.includes(
            "No telemetry or property fields were found in ignored-only.tm.json. Upload a model that describes your device data."),
            "The asynchronous all-ignored file validation did not reach its actual no-fields result.");
        assert(next()!.disabled && !currentContext.hubModel && !area.querySelector(".iot-solution-model-upload-area--loaded") &&
            !area.querySelector("#model-upload-warnings") && !area.querySelector("[data-field-id]") &&
            area.querySelector("#model-upload-error")?.textContent?.includes(
                "No telemetry or property fields were found in ignored-only.tm.json. Upload a model that describes your device data."),
            `All-ignored upload became an accepted empty model or bypassed actual Continue: ${JSON.stringify({
                disabled: next()!.disabled, model: !!currentContext.hubModel,
                loaded: !!area.querySelector(".iot-solution-model-upload-area--loaded"),
                warning: !!area.querySelector("#model-upload-warnings"), fields: area.querySelectorAll("[data-field-id]").length,
                error: area.querySelector("#model-upload-error")?.textContent,
            })}`);
        caseObservations = { remainingFieldManualRoleRequired: true, allIgnoredBlocksActualContinue: true, noSuccessShapedEmptyPlan: true };
    })],
    ["Model Coverage panels require current explicit invocation and retain only legitimate peer or Back evidence", async () => {
        let edit: WizardStepProps["updateContext"] | undefined;
        const View = (props: React.ComponentProps<typeof HubModelScopeStep>) => {
            edit = props.updateContext; return <CoverageView {...props}/>;
        };
        await mount(<Harness component={View} initial={withCurrentHubSummary({ ...context(true), hubScope: undefined })}/>);
        const reads = calls.length, auth = tokenRequests;
        assert(!area.querySelector("[data-coverage-stream], .iot-solution-coverage-assignment-hint"), "Fresh owner exposed unverified result boxes.");
        await click(area.querySelector('input[value="data"]')); await click(area.querySelector('input[value="none"]'));
        assert(!button("Check coverage")!.disabled && !area.querySelector("[data-coverage-stream]") &&
            calls.length === reads && tokenRequests === auth, "Choices alone exposed results or authorized reads.");
        deferredToken = true; tokenReleases = [];
        await click(button("Check coverage"));
        assert(area.querySelectorAll('[data-coverage-stream][aria-busy="true"]').length === 2 &&
            calls.length === reads, "Explicit pending invocation did not reveal honest loading panes.");
        deferredToken = false;
        await act(async () => tokenReleases.splice(0).forEach(release => release()));
        await until(() => !!currentContext.hubCoverage, "Current invocation did not finish.");
        const count = calls.length;
        await act(async () => setStepVisible(false)); await act(async () => setStepVisible(true));
        assert(area.querySelectorAll("[data-coverage-stream]").length === 2 && calls.length === count,
            "Same-owner Back discarded genuine current results or queried automatically.");
        await click(area.querySelector('input[value="specified"]'));
        assert(!area.querySelector("[data-coverage-stream]") && calls.length === count,
            "Choice invalidation borrowed older counts as current results.");
        await verifyCoverage();
        const peer = currentContext.hubStreamChecks.twin.proof, beforeEdit = calls.length;
        await enter("hub-telemetry-scope", "true");
        assert(area.querySelectorAll("[data-coverage-stream]").length === 2 &&
            currentContext.hubStreamChecks.twin.proof === peer && calls.length === beforeEdit,
            "Dirty-only edit discarded its current peer or implicitly reread.");
        coverageReadFailure = "telemetry"; await click(button("Check coverage"));
        await until(() => currentContext.hubStreamChecks.telemetry.phase === "error", "Expected current error did not settle.");
        assert(!button("Retry Telemetry") && !button("Check coverage")!.disabled &&
            !!area.querySelector('.iot-solution-coverage-error[role="alert"]') && !area.querySelector("[data-coverage-stream]"),
            "An invoked error/retry was hidden with the pre-invocation gate.");
        coverageReadFailure = undefined;
        const beforeOwner = calls.length;
        const model = currentContext.hubModel;
        await act(async () => edit!("hubModel", { ...model, description: "Different current model meaning" }));
        await act(async () => edit!("hubModel", model));
        assert(!area.querySelector("[data-coverage-stream]"), "Model ABA resurrected an earlier visible result.");
        await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
        await act(async () => setTestSourceState(sourceFixture));
        assert(!area.querySelector("[data-coverage-stream], .iot-solution-coverage-assignment-hint") && calls.length === beforeOwner,
            "Source ABA resurrected old visible results or dispatched reads.");
        caseObservations = { initialAndChosenNotClickedHidden: true, explicitPendingVisible: true, successAndErrorVisible: true,
            currentPeerAndBackRetained: true, sourceAndChoiceRetirementNoResurrection: true, implicitReads: 0 };
    }],
    ["Model Coverage explicit reference choices preserve exact copy and reach scoped reads and native policies without implicit queries", async () => {
        const initial = withCurrentHubSummary({ ...context(true), hubScope: undefined });
        await mount(<Harness component={CoverageView} initial={initial}/>);
        const requests = calls.length + resourceCalls.length, tokens = tokenRequests, consent = consentRequests;
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" &&
            area.querySelector(".iot-solution-step-description")?.textContent ===
            "Your raw tables can include events from other kinds of devices. Choose which events to include in your modeled data, then check how well your model covers them." &&
            area.querySelector("#model-reference-filter-heading")?.textContent === "1. Filter by model ID" &&
            area.querySelector("#additional-conditions-heading")?.textContent === "2. Add conditions (optional)",
            "Coverage title/intro/section copy differs.");
        const options = [...area.querySelectorAll<HTMLInputElement>('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]')];
        const labels = options.map(option => document.getElementById(option.getAttribute("aria-labelledby") || "")?.textContent);
        assert(options.length === 3 && options.every(option => !option.checked) &&
            JSON.stringify(labels) === JSON.stringify([
                "Don't filter by model ID",
                "Require the model ID",
                "Model ID or additional conditions",
            ]) && button("Check coverage")!.disabled &&
            !button("Check Telemetry") && !button("Check Reported Properties"),
            "Available reference filter has an implicit choice or incorrect native options.");
        for (const option of options) {
            const ids = option.getAttribute("aria-describedby")?.split(" ") || [];
            assert(ids.length >= 1 && ids.every(id => document.getElementById(id)?.textContent?.length),
                "A reference option is not associated with its actual description/example.");
        }
        assert(area.querySelector("#reference-and-description")?.textContent ===
            `Include only events with this model's ID (${tm.id}) that also meet any additional conditions` &&
            area.querySelector("#reference-or-description")?.textContent ===
            "Include events with this model's ID, plus any events that meet the additional conditions" &&
            !area.querySelector('[aria-label="Current device model"]') && !area.querySelector("#reference-blank-description") &&
            !summary("Model reference examples") && !area.querySelector("#hub-telemetry-scope") &&
            area.querySelector("#coverage-results-heading")?.textContent === "3. Check model coverage",
            "Requested canonical reference copy/removals/initial editor hiding are wrong.");
        await click(area.querySelector('input[value="specified"]'));
        const predicate = 'tostring(headers["IoTConnectionDeviceId"]) == "device-1"';
        await enter("hub-telemetry-scope", predicate);
        assert(currentContext.hubScope.mode === "unselected" && button("Check coverage")!.disabled,
            "Editing an additional condition silently selected a reference option.");
        await click(area.querySelector('input[value="data"]'));
        assert(effectiveScope(hubInput(currentContext, sourceFixture).scope, "telemetry") === predicate,
            "No-reference selection retained a model predicate.");
        await click(area.querySelector('input[value="and"]'));
        const reference = effectiveScope({ ...hubInput(currentContext, sourceFixture).scope, additional: "none" }, "telemetry");
        assert(effectiveScope(hubInput(currentContext, sourceFixture).scope, "telemetry") ===
            `${reference} and (${predicate})`, "AND selection did not intersect actual predicates.");
        await click(area.querySelector('input[value="or"]'));
        assert(effectiveScope(hubInput(currentContext, sourceFixture).scope, "twin") === reference &&
            calls.length + resourceCalls.length === requests && tokenRequests === tokens && consentRequests === consent,
            "Blank OR widened a stream to true, or a radio/condition edit implicitly queried.");
        await verifyCoverage();
        const input = hubInput(currentContext, sourceFixture), coverage = currentContext.hubCoverage;
        assert(streamCoverageReady(input, sourceFixture.revision, currentContext.hubStreamChecks, false),
            "Both explicit current checks did not authorize coverage.");
        const plan = buildModeledDataPlan(input, coverage, false, { modeled: "ModelCoverage", twinChanges: "CoverageProperties", twinState: "CoverageState" });
        assert(plan.commands.some(command => command.kind === "policy" && command.source === input.source.telemetryTable &&
            command.query?.includes(`| where ${reference} or (${predicate})`)) &&
            plan.commands.some(command => command.kind === "policy" && command.source === input.source.twinTable &&
                command.query?.includes(`| where ${reference}\n`)),
            "Coverage choice did not reach actual independent generated update policies.");
        const peer = currentContext.hubStreamChecks.twin.proof, before = calls.length;
        await enter("hub-telemetry-scope", "true");
        assert(!currentContext.hubCoverage && !currentContext.hubCreation && currentContext.hubStreamChecks.twin.proof === peer &&
            currentContext.hubStreamChecks.telemetry.phase === "unchecked" && calls.length === before,
            "Additional condition edit kept stale aggregate/Create proof, queried, or discarded current peer evidence.");
        await click(area.querySelector('input[value="and"]')); await click(area.querySelector('input[value="or"]'));
        assert(currentContext.hubStreamChecks.twin.proof === undefined &&
            !streamCoverageReady(hubInput(currentContext, sourceFixture), sourceFixture.revision, currentContext.hubStreamChecks, true),
            "Reference choice ABA restored old processing proof.");
        caseObservations = { noImplicitDefault: true, exactAccessibleThreeOptions: true, radioReads: 0,
            actualPoliciesAndScopedReads: true, blankOrReferenceOnly: true, conditionPeerIsolation: true, choiceAbaRetired: true };
    }],
    ["Model Coverage additional choice keeps inactive drafts out of checks and plans and preserves drafts on Back", async () => {
        const seed = withCurrentHubSummary({ ...context(true), hubScope: { ...context(true).hubScope, additional: "unselected",
            telemetry: 'tostring(headers["IoTConnectionDeviceId"]) == "device-1"', twin: "" }, hubSampleLimit: 1000 });
        await mount(<Harness component={CoverageView} initial={seed}/>);
        const before = calls.length, auth = tokenRequests;
        assert(button("Check coverage")!.disabled && !area.querySelector("#hub-telemetry-scope") &&
            !area.querySelector('[aria-labelledby="additional-conditions-heading"] input:checked') && hubInput(currentContext, sourceFixture).limit === 100,
            "2. Add conditions (optional) defaulted or a hidden/deeper limit bypassed the page choice gate.");
        await click(button("Check coverage"));
        assert(calls.length === before && tokenRequests === auth, "Missing additional choice authorized a read.");
        await click(area.querySelector('input[value="specified"]'));
        await verifyCoverage();
        const checked = currentContext.hubStreamChecks, draft = currentContext.hubScope.telemetry, reads = calls.length;
        await click(area.querySelector('input[value="none"]'));
        assert(!area.querySelector("#hub-telemetry-scope") && currentContext.hubScope.telemetry === draft &&
            hubInput(currentContext, sourceFixture).scope.telemetry === "" &&
            !currentContext.hubCoverage && !currentContext.hubAllowWaiting && calls.length === reads,
            "None lost its draft, retained old credit, queried or left a hidden expression effective.");
        await verifyCoverage();
        const input = hubInput(currentContext, sourceFixture);
        const plan = buildModeledDataPlan(input, currentContext.hubCoverage, false, creationRequest().names);
        assert(calls.slice(reads).filter(query => query.includes("__observation")).every(query => query.includes("| where true\n")) &&
            plan.commands.filter(command => command.kind === "policy").every(command => command.query.includes("| where true\n") &&
                !command.query.includes(draft)),
            "A hidden draft affected processing reads or either final policy.");
        await act(async () => setStepVisible(false)); await act(async () => setStepVisible(true));
        assert(currentContext.hubScope.telemetry === draft && !area.querySelector("#hub-telemetry-scope"), "Same-owner Back lost a hidden draft.");
        const count = calls.length;
        await click(area.querySelector('input[value="specified"]'));
        assert(area.querySelector<HTMLTextAreaElement>("#hub-telemetry-scope")!.value === draft &&
            !currentContext.hubCoverage && calls.length === count &&
            currentContext.hubStreamChecks.twin.proof !== checked.twin.proof,
            "Specify/None/Specify restored stale proof or failed to restore draft text without querying.");
        await enter("hub-telemetry-scope", "true | union Other");
        assert(button("Check coverage")!.disabled, "An invalid active predicate was silently accepted.");
        await click(area.querySelector('input[value="none"]'));
        assert(!button("Check coverage")!.disabled && hubInput(currentContext, sourceFixture).scope.telemetry === "",
            "An invalid inactive draft blocked or contaminated None.");
        caseObservations = { noDefault: true, noChoiceReads: 0, hiddenDraftNeverEffective: true, bothPoliciesUseNone: true,
            sameOwnerBackDraftRetained: true, choiceAbaRetiresProof: true, invalidActiveRejected: true, invalidHiddenInactive: true, coverageLimit: 100 };
    }],
    ["Model Coverage combined Verify checks dirty streams once and reuses current peers on failure retry and edits", async () => {
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        const reads = calls.length, tokens = tokenRequests;
        coverageReadFailure = "telemetry"; coverageReadFailureStatus = 400; deferredToken = true; tokenReleases = [];
        await click(button("Check coverage")); await click(button("Check coverage"));
        await until(() => tokenReleases.length === 1, "Combined Verify did not share one owned stream authorization.");
        assert(tokenRequests === tokens + 1 && button("Check coverage")!.disabled && calls.length === reads &&
            currentContext.hubStreamChecks.telemetry.phase === "checking" && currentContext.hubStreamChecks.twin.phase === "checking",
            "Double submit duplicated coverage or skipped a dirty stream.");
        deferredToken = false;
        await act(async () => tokenReleases.splice(0).forEach(release => release()));
        await until(() => currentContext.hubStreamChecks.telemetry.phase === "error" &&
            currentContext.hubStreamChecks.twin.phase === "checked", "Independent success/failure did not settle.");
        const failure = area.querySelector('.iot-solution-coverage-error[role="alert"]')!;
        const peer = currentContext.hubStreamChecks.twin.proof;
        assert(failure?.classList.contains("fui-MessageBar") && failure.textContent?.includes("Telemetry:") &&
            failure.textContent?.includes("Synthetic current KQL syntax error.") &&
            !area.querySelector("[data-coverage-stream], .iot-solution-coverage-assignment-hint") &&
            !currentContext.hubCoverage && !streamCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks, true) &&
            !button("Retry Reported Properties") && !button("Retry Telemetry") && !button("Check coverage")!.disabled,
            "Current failure did not hide both panels/use the native error box and same primary retry.");
        coverageReadFailure = undefined; coverageReadFailureStatus = 403;
        let before = calls.length;
        await click(button("Check coverage"));
        await until(() => !!currentContext.hubCoverage, "Explicit failed-stream retry did not recover.");
        assert(calls.length === before + 2 && calls.slice(before).every(query => query.includes("['Telemetry']")) &&
            currentContext.hubStreamChecks.twin.proof === peer && !area.querySelector(".iot-solution-coverage-error") &&
            area.querySelectorAll("[data-coverage-stream]").length === 2,
            "Retry reread the current peer or retained a failed result.");
        const twin = currentContext.hubStreamChecks.twin.proof;
        before = calls.length;
        await enter("hub-telemetry-scope", "true");
        assert(calls.length === before && !currentContext.hubCoverage && currentContext.hubStreamChecks.twin.proof === twin &&
            !button("Check coverage")!.disabled, "A filter edit read data, retained assembled proof, or erased the peer.");
        await verifyCoverage();
        assert(calls.length === before + 2 && calls.slice(before).every(query => query.includes("['Telemetry']")) &&
            currentContext.hubStreamChecks.twin.proof === twin && !button("Check coverage")!.disabled,
            "Verify did not reuse still-current peer results.");
        assert(calls.length === before + 2, "Current results were redundantly read.");
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        const consent = consentRequests, auth = tokenRequests, count = calls.length;
        assignmentAuthFailure = true;
        await verifyCoverage();
        assert(consentRequests === consent + 1 && tokenRequests === auth + 2 && calls.length === count + 4,
            "The explicit batch opened parallel consent fallbacks or skipped an independent stream.");
        caseObservations = { twoDirtyStreams: true, doubleSubmitAuthorizations: 1, sharedConsentFallback: 1, failureNotZero: true,
            retryReadsOnlyFailedStream: 2, editedStreamReadsOnly: 2, currentPeerReused: true, implicitReads: 0 };
    }],
    ["Model Coverage current both-stream and sign-in failures keep native detail once and recover through primary Verify", async () => {
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        coverageReadFailure = "both";
        await click(button("Check coverage"));
        await until(() => currentContext.hubStreamChecks.telemetry.phase === "error" &&
            currentContext.hubStreamChecks.twin.phase === "error", "Both current failures did not settle.");
        const error = area.querySelector(".iot-solution-coverage-error")!;
        assert(error.textContent?.includes("Telemetry and Reported Properties:") &&
            error.textContent?.split("Synthetic current coverage denial.").length === 2 &&
            error.querySelector(".fui-MessageBarBody")!.children.length === 1 &&
            !area.querySelector("[data-coverage-stream], .iot-solution-coverage-assignment-hint") &&
            !button("Check coverage")!.disabled, "Equivalent both-stream failures lost context or repeated their boilerplate.");
        coverageReadFailure = undefined;
        const originalConsent = client.auth.acquireAccessToken;
        try {
            await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
            assignmentAuthFailure = true;
            client.auth.acquireAccessToken = async () => { throw { error: 2 }; };
            const reads = calls.length;
            await click(button("Check coverage"));
            await until(() => currentContext.hubStreamChecks.telemetry.phase === "error" &&
                currentContext.hubStreamChecks.twin.phase === "error", "Current authorization failure did not settle.");
            assert(area.querySelector('.iot-solution-coverage-error[role="alert"]')?.textContent?.includes(
                "Entra app is not configured for the scope") &&
                area.querySelector(".iot-solution-coverage-error")?.textContent?.includes("Add the API permission") && calls.length === reads &&
                !area.querySelector("[data-coverage-stream]") && !button("Check coverage")!.disabled,
                "Native authorization detail or explicit primary retry was lost.");
            client.auth.acquireAccessToken = originalConsent; assignmentAuthFailure = false;
            await verifyCoverage();
            assert(!area.querySelector(".iot-solution-coverage-error") && area.querySelectorAll("[data-coverage-stream]").length === 2,
                "Explicit reauthorization did not recover current successful panels.");
        } finally { client.auth.acquireAccessToken = originalConsent; assignmentAuthFailure = false; }
        caseObservations = { bothContextsOneNormalizedMessage: true, nativeErrorMessageBar: true,
            authorizationDetailPreserved: true, noQueryBeforeAuth: true, explicitReauthorizationOnly: true, rowRetryButtons: 0 };
    }],
    ["Model Coverage All fields found is current positive selected evidence only with native green", async () => {
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        assert(!area.querySelector("[data-success=true]"), "Initial unchecked state was shown as successful.");
        await verifyCoverage();
        const statuses = [...area.querySelectorAll<HTMLElement>(".iot-solution-coverage-result-status")];
        const reference = document.createElement("span");
        reference.style.color = getComputedStyle(statuses[0]).getPropertyValue("--colorStatusSuccessForeground1").trim();
        document.body.appendChild(reference); const green = getComputedStyle(reference).color; reference.remove();
        assert(statuses.length === 2 && statuses.every(status => status.textContent === "All fields found" &&
            status.dataset.success === "true" && getComputedStyle(status).color === green) &&
            !area.textContent?.includes("Sample checked."), "Positive complete selected evidence lost exact native-green status.");
        assertCoverageMissingGuidance();
        const reads = calls.length;
        await enter("hub-telemetry-scope", "true");
        assert(!area.querySelector('[data-coverage-stream="telemetry"] [data-success=true]') && calls.length === reads,
            "A dirty stream retained success or implicitly queried.");
        parityPayloads = { telemetry: {}, twin: { properties: { reported: {} } } };
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        await verifyCoverage();
        assert(!area.querySelector("[data-success=true]") && !!area.querySelector("[data-missing-field]"),
            "Missing selected fields were presented as all observed.");
        assertCoverageMissingGuidance();
        emptyProcessing = true;
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        await verifyCoverage();
        assert(!area.querySelector("[data-success=true]") && area.querySelectorAll(".iot-solution-coverage-result-status")
            .length === 2 && !area.querySelector(".iot-solution-coverage-missing-fields"),
            "Zero samples became green or regained a missing-field list.");
        assertCoverageMissingGuidance();
        emptyProcessing = false; parityPayloads = undefined;
        const seed = context(true); seed.hubFields = seed.hubFields.filter((field: ModelField) => field.role === "twin");
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(seed)}/>);
        await verifyCoverage();
        const unselected = area.querySelector<HTMLElement>('[data-coverage-stream="telemetry"] .iot-solution-coverage-result-status')!;
        assert(unselected.textContent === "No fields selected" && unselected.dataset.success === "false",
            "An unselected stream gained vacuous all-observed success.");
        const beforeOwner = calls.length;
        await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
        await act(async () => setTestSourceState(sourceFixture));
        assert(!area.querySelector("[data-coverage-stream], .iot-solution-coverage-error") && calls.length === beforeOwner,
            "Owner ABA resurrected error/success output or queried.");
        assertCoverageMissingGuidance();
        caseObservations = { exactGreenPositiveSelected: true, missingZeroAndNoSelectionNotGreen: true,
            dirtyAndOwnerRetireSuccess: true, noImplicitReads: true };
    }],
    ["Model Coverage explicit reverify refreshes current unchanged predicates and missing-field samples without automatic reads", async () => {
        const request = localFixtureRequest(localSingleModel);
        const initial = withCurrentHubSummary({ ...context(true), hubModel: request.input.model,
            hubFields: request.input.fields, hubScope: request.input.scope });
        parityPayloads = { telemetry: { measurement: localFixtureValues.measurement },
            twin: { properties: { reported: {} } } };
        await mount(<Harness component={CoverageView} initial={initial}/>);
        const start = calls.length;
        await verifyCoverage();
        const previous = currentContext.hubStreamChecks;
        const scope = canonicalJson(hubInput(currentContext, sourceFixture).scope);
        const firstQueries = calls.slice(start);
        const result = (stream: "telemetry" | "twin") => area.querySelector(`[data-coverage-stream="${stream}"]`)!;
        assert(firstQueries.length === 4 && !button("Check coverage")!.disabled &&
            coverageSampleCount(result("telemetry")) === 1 &&
            coverageFieldLabel(result("telemetry")) === "Selected model fields found: 1 of 2 in this sample." &&
            coverageFieldLabel(result("twin")) === "Selected model fields found: 0 of 1 in this sample." &&
            result("telemetry").querySelectorAll("[data-missing-field]").length === 1 &&
            result("twin").querySelectorAll("[data-missing-field]").length === 1,
            "Current results did not expose explicit reverify or truthful initial missing-field evidence.");
        coverageSampleRows = 3;
        parityPayloads = { telemetry: { measurement: localFixtureValues.measurement, alarm: localFixtureValues.alarm },
            twin: { properties: { reported: { history: localFixtureValues.history } } } };
        await settle();
        assert(calls.length === start + 4 &&
            coverageSampleCount(result("telemetry")) === 1,
            "New incoming fixture data triggered implicit reads or changed unverified sample counts.");
        deferredToken = true; tokenReleases = [];
        const tokens = tokenRequests, action = button("Check coverage")!;
        await click(action); await click(action);
        assert(tokenReleases.length === 1 && tokenRequests === tokens + 1 &&
            currentContext.hubStreamChecks.telemetry.phase === "checking" &&
            currentContext.hubStreamChecks.twin.phase === "checking" &&
            !currentContext.hubCoverage && !currentContext.hubAllowWaiting &&
            !streamCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks, true) &&
            coverageSampleCount(result("telemetry")) === undefined &&
            coverageSampleCount(result("twin")) === undefined && calls.length === start + 4,
            "Reverify duplicated authorization or kept old results/proof fresh during its pending check.");
        deferredToken = false;
        await act(async () => tokenReleases.splice(0).forEach(release => release()));
        await until(() => !!currentContext.hubCoverage, "Explicit same-condition reverify did not settle.");
        const queries = calls.slice(start + 4);
        assert(queries.length === 4 && canonicalJson(hubInput(currentContext, sourceFixture).scope) === scope &&
            currentContext.hubStreamChecks.telemetry.proof !== previous.telemetry.proof &&
            currentContext.hubStreamChecks.twin.proof !== previous.twin.proof &&
            (["telemetry", "twin"] as const).every(stream => coverageSampleCount(result(stream)) === 3 &&
                result(stream).querySelectorAll("[data-missing-field]").length === 0) &&
            coverageFieldLabel(result("telemetry")) === "Selected model fields found: 2 of 2 in this sample." &&
            coverageFieldLabel(result("twin")) === "Selected model fields found: 1 of 1 in this sample." &&
            !button("Check coverage")!.disabled, "Reverify did not replace both sample counts/field evidence with fresh results.");
        for (const table of ["Telemetry", "Twins"]) assert(
            JSON.stringify(firstQueries.filter(query => query.includes(`['${table}']`))) ===
            JSON.stringify(queries.filter(query => query.includes(`['${table}']`))),
            "Reverify changed a processing predicate, selected table or bound.");
        caseObservations = { initialCounts: { telemetry: 1, properties: 1 }, freshCounts: { telemetry: 3, properties: 3 },
            initialFieldsObserved: { telemetry: "1/2", properties: "0/1" }, freshFieldsObserved: { telemetry: "2/2", properties: "1/1" },
            samePredicatesAndBounds: true, readsBeforeExplicitReverify: 0, explicitReverifyQueries: 4,
            oneSharedAuthorization: true, pendingOldProofRetired: true, doubleSubmitGuard: true };
    }],
    ["Model Coverage sample results name missing fields and distinguish filtered zero from empty source without historical claims", async () => {
        const seed = context(true);
        seed.hubFields = seed.hubFields.map((field: ModelField) => setFieldRole(field, "telemetry"));
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(seed)}/>);
        assert(coverageSampleCount(area.querySelector('[data-coverage-stream="telemetry"]')) === undefined,
            "Unchecked data was rendered as an observed count.");
        await verifyCoverage();
        const telemetry = area.querySelector('[data-coverage-stream="telemetry"]')!;
        const missing = seed.hubFields.find((field: ModelField) => field.path[0] === "state");
        assert(coverageSampleCount(telemetry) === 1 &&
            coverageFieldLabel(telemetry) === "Selected model fields found: 1 of 2 in this sample." &&
            !area.querySelector("[data-coverage-sample-details]") &&
            telemetry.querySelector("[data-missing-field]")?.getAttribute("data-missing-field") === missing.id &&
            !area.textContent?.includes("Check the field roles and paths in Device Model") &&
            area.querySelector("#coverage-sample-note")?.textContent ===
                "Check up to 100 recent matching events from each table to see which of your selected fields appear",
            "Matching sample size was conflated with selected-field coverage or missing names/recovery/bounds.");
        emptyProcessing = true;
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        await verifyCoverage();
        assert(area.querySelector('[data-coverage-stream="telemetry"]')?.textContent?.includes("No events match your filters. Check your conditions, or wait for devices to send matching data.") &&
            coverageFieldLabel(area.querySelector('[data-coverage-stream="telemetry"]')) === "Selected model fields found: 0 of 1 in this sample." &&
            !currentContext.hubAllowWaiting && !area.querySelector('input[type="checkbox"]') &&
            modelCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks),
            "Verified zero did not permit continuation or fabricated a hidden waiting consent.");
        await verifyCoverage();
        assert(!currentContext.hubAllowWaiting &&
            modelCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks) &&
            coverageSampleCount(area.querySelector('[data-coverage-stream="telemetry"]')) === 0,
            "Reverified zero required a removed consent flag or failed to preserve the genuine zero count.");
        emptyProcessing = false; emptySamples = "both";
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        await verifyCoverage();
        assert(area.querySelector('[data-coverage-stream="telemetry"]')?.textContent?.includes("The source sample was empty.") &&
            !modelCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks),
            "An empty unfiltered source was mistaken for a supported waiting/native-identity proof.");
        emptySamples = "none";
        caseObservations = { matchingVersusSelectedFields: true, namedMissingField: "state", boundedLimit: 100,
            exactSampleCopy: true, obsoleteDetailsAndRecoveryRemoved: true, verifiedZeroNeedsNoOptIn: true, emptySourceNotNativeProof: true };
    }],
    ["Model Coverage native visual rows disclosures paired editors and comparison retain accessible choices without implicit reads", async () => {
        const seed = context(true);
        seed.hubScope = undefined;
        seed.hubFields = seed.hubFields.map((field: ModelField) => setFieldRole(field, "telemetry"));
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(seed)}/>);
        const reads = calls.length, tokens = tokenRequests, consent = consentRequests;
        const options = [...area.querySelectorAll<HTMLInputElement>('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]')];
        assert(options.length === 3 && options.every(option => !option.checked) &&
            area.querySelectorAll(".iot-solution-coverage-choice").length === 5 &&
            !area.querySelector("#reference-blank-description") &&
            !area.querySelector('[aria-labelledby="additional-conditions-heading"] input:checked'),
            "Visual choice rows introduced a default or retained removed reference help.");
        const widths = [...area.querySelectorAll(".iot-solution-coverage-choice")].map(row => row.getBoundingClientRect().width);
        assert(Math.max(...widths) - Math.min(...widths) < 1, "Semantic radio rows have inconsistent whole-row widths.");
        for (const option of options) {
            const label = document.querySelector<HTMLLabelElement>(`label[for="${option.id}"]`);
            assert(label && label.querySelector(".iot-solution-coverage-choice-description"), "Native radio label does not own its row subtitle.");
            await click(label);
            assert(option.checked && area.querySelectorAll('.iot-solution-coverage-choice[data-selected="true"]').length === 1,
                "The full native label did not select exactly one row.");
        }
        await click(area.querySelector('input[value="specified"]'));
        assert(!summary("KQL syntax and available values") && !button("Copy telemetry example") && !button("Copy reported properties example"),
            "Removed syntax/copy controls remain on this page.");
        assert(calls.length === reads && tokenRequests === tokens && consentRequests === consent,
            "A visual choice or disclosure queried/authenticated implicitly.");
        const editorSections = [...area.querySelectorAll<HTMLElement>(".iot-solution-coverage-editors > section")];
        const a = editorSections[0].getBoundingClientRect(), b = editorSections[1].getBoundingClientRect();
        assert(editorSections.length === 2 && editorSections.every(section => !section.querySelector("h4") &&
            !!section.querySelector("label[for]") && !!section.querySelector("textarea")) &&
            (innerWidth > 1100 ? Math.abs(a.top - b.top) < 1 && b.left > a.left : b.top >= a.bottom),
            "Native editors repeated headings or lost their responsive paired/stacked layout.");
        assert(!area.querySelector("[data-coverage-stream]"), "Uninvoked coverage boxes were shown after making choices.");
        await verifyCoverage();
        const resultSections = [...area.querySelectorAll<HTMLElement>(".iot-solution-coverage-streams > section")];
        const first = resultSections[0].getBoundingClientRect(), second = resultSections[1].getBoundingClientRect();
        assert(resultSections.length === 2 && resultSections.every(section => section.querySelectorAll("dl > div").length === 3 &&
            section.querySelector("[data-sample-count]")?.getAttribute("aria-label")?.startsWith("Matching events checked:")) &&
            resultSections[0].querySelector("h4")?.textContent === "Telemetry table (Telemetry)" &&
            resultSections[1].querySelector("h4")?.textContent === "Properties table (Twins)" &&
            (innerWidth > 1100 ? Math.abs(first.top - second.top) < 1 && second.left > first.left : second.top >= first.bottom),
            "The paired result subsections lost actual table names, semantic details, responsiveness or unchecked state.");
        assert(!area.querySelector("[data-coverage-sample-details]") && area.querySelectorAll("[data-missing-field]").length > 0 &&
            [...area.querySelectorAll("[data-missing-field]")].every(field => !field.closest("details")),
            "Missing selected fields were buried inside secondary sample details.");
        caseObservations = { nativeRowClick: true, explicitNoDefault: true, requestedReferenceRemovals: true,
            removedCopyAndSyntax: true, pairedOrStackedEditorsAndResults: true, semanticVerticalDetails: true,
            missingFieldsRemainVisible: true, disclosureAndSelectionReads: 0 };
    }],
    ["Model Coverage compact positive missing tables preserve 14 and 15 identities contrast density and component distinction", async () => {
        const seed = coverageMissingFixture();
        seed.hubFields[0] = { ...seed.hubFields[0],
            title: "Temperature with a deliberately long friendly title that wraps without truncation" };
        parityPayloads = { telemetry: { unmodeledReading: 1 }, twin: { properties: { reported: { unmodeledState: "idle" } } } };
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(seed)}/>);
        await verifyCoverage();
        const measured: object[] = [];
        const luminance = (color: string) => {
            const channels = color.match(/[\d.]+/g)!.slice(0, 3).map(Number).map(value => value / 255)
                .map(value => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4);
            return channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722;
        };
        const contrast = (foreground: string, background: string) => {
            const a = luminance(foreground), b = luminance(background);
            return (Math.max(a, b) + 0.05) / (Math.min(a, b) + 0.05);
        };
        for (const [stream, count] of [["telemetry", 14], ["twin", 15]] as const) {
            const rows = [...area.querySelectorAll<HTMLElement>(`[data-coverage-stream="${stream}"] [data-missing-field]`)];
            const expected = (seed.hubFields as ModelField[]).filter(field => field.included && field.role === stream);
            assert(rows.length === count && JSON.stringify(rows.map(row => row.dataset.missingField)) === JSON.stringify(expected.map(field => field.id)),
                "Missing table changed declared order, count or identity.");
            rows.forEach((row, index) => {
                const primary = row.querySelector<HTMLElement>(".iot-solution-coverage-missing-title")!;
                const secondary = row.querySelector<HTMLElement>(".iot-solution-coverage-missing-name")!;
                assert(row.tagName === "TR" && row.parentElement?.tagName === "TBODY" &&
                    row.closest("table")?.querySelector("caption")?.textContent?.includes(`(${count})`) && !row.closest("details") &&
                    primary.textContent === expected[index].title && secondary.textContent === expected[index].path.join(".") &&
                    getComputedStyle(primary).fontSize === "14px" && getComputedStyle(secondary).fontSize === "14px" &&
                    getComputedStyle(primary).fontWeight === "400" && row.textContent === `${expected[index].title} (${expected[index].path.join(".")})` &&
                    getComputedStyle(secondary).fontStyle === "normal" &&
                    row.children.length === 1 &&
                    [...row.children].every(cell => cell.scrollWidth <= cell.clientWidth + 1),
                    "Compact missing names lost same-row cells, exact identities, native typography or safe wrapping.");
            });
            const row = rows[1], primary = row.querySelector<HTMLElement>(".iot-solution-coverage-missing-title")!;
            const secondary = row.querySelector<HTMLElement>(".iot-solution-coverage-missing-name")!;
            const panel = row.closest<HTMLElement>("[data-coverage-stream]")!;
            const rowHeight = row.getBoundingClientRect().height, background = getComputedStyle(panel).backgroundColor;
            const primaryContrast = contrast(getComputedStyle(primary).color, background);
            const secondaryContrast = contrast(getComputedStyle(secondary).color, background);
            const caption = row.closest("table")!.querySelector("caption")!;
            const guidance = panel.querySelector<HTMLElement>("[data-missing-fields-guidance]")!;
            const captionGap = caption.getBoundingClientRect().top - panel.querySelector("dl")!.getBoundingClientRect().bottom;
            const guidanceGap = guidance.getBoundingClientRect().top - row.closest("table")!.getBoundingClientRect().bottom;
            assert(Math.abs(captionGap - 16) <= 1 && Math.abs(guidanceGap - 8) <= 1 &&
                guidance.previousElementSibling === row.closest("table") &&
                getComputedStyle(caption).fontSize === "14px" &&
                getComputedStyle(panel.querySelector(".iot-solution-coverage-source")!).fontSize === "14px" &&
                [...panel.querySelectorAll("dt,dd")].every(node => getComputedStyle(node).fontSize === "14px"),
                "Result body/table-name typography, list-before-reassurance order or native missing-content gaps changed.");
            assertCoverageMissingGuidance();
            const prior = document.createElement("table");
            prior.style.cssText = `position:absolute;visibility:hidden;width:${row.closest("table")!.getBoundingClientRect().width}px;border-collapse:collapse;table-layout:fixed`;
            const cell = prior.insertRow().insertCell(); cell.style.padding = "8px 0";
            for (const [content, size, line] of [[primary.textContent, "14px", "20px"], [secondary.textContent, "12px", "16px"]]) {
                const text = document.createElement("span"); text.textContent = content;
                Object.assign(text.style, { display: "block", fontSize: size, lineHeight: line,
                    fontFamily: getComputedStyle(primary).fontFamily }); cell.appendChild(text);
            }
            document.body.appendChild(prior); const priorHeight = prior.rows[0].getBoundingClientRect().height; prior.remove();
            assert(rowHeight >= 24 && rowHeight <= 28 && rowHeight <= priorHeight * 0.6 &&
                Math.abs(primary.getBoundingClientRect().top - secondary.getBoundingClientRect().top) < 1 &&
                primaryContrast >= 4.5 && secondaryContrast >= 4.5,
                "The installed extra-small row did not materially reduce measured height or lost same-baseline readable contrast.");
            measured.push({ stream, count, priorStackedRowHeight: priorHeight, compactRowHeight: rowHeight,
                reduction: 1 - rowHeight / priorHeight, primaryContrast, secondaryContrast, background, captionGap, guidanceGap });
        }
        const request = mixedDtdlRequest();
        await mount(<Harness component={CoverageView} initial={{ ...context(), hubModel: request.input.model,
            hubFields: request.input.fields, hubScope: request.input.scope }}/>);
        await verifyCoverage();
        const fields = request.input.fields.filter(field => field.path[0] === "state");
        for (const field of fields) {
            const row = [...area.querySelectorAll<HTMLElement>("[data-missing-field]")].find(node => node.dataset.missingField === field.id)!;
            const technical = row.querySelector(".iot-solution-coverage-missing-name");
            assert(field.component ? technical?.textContent === `${field.component}.state` : !technical,
                "Equal component leaves merged or a duplicate technical line was added.");
        }
        parityPayloads = undefined;
        caseObservations = { counts: { telemetry: 14, properties: 15 }, declaredIdentityOrderPreserved: true,
            normal14InlineParentheses: true, measured, longTitleWraps: true, fullNamesVisible: true,
            componentLeavesDistinct: true, duplicateSecondaryOmitted: true, noDisclosureOrFakeActions: true };
    }],
    ["Model Coverage mixed zero telemetry keeps only positive property missing fields in compact neutral panels", async () => {
        const seed = coverageMissingFixture();
        seed.hubScope = { ...seed.hubScope, mode: "reference", operator: "and" };
        emptyProcessingStream = "telemetry"; coverageSampleRows = 100;
        parityPayloads = { telemetry: { unmodeledReading: 1 }, twin: { properties: { reported: { unmodeledState: "idle" } } } };
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(seed)}/>);
        const reads = calls.length;
        await settle(); assert(calls.length === reads, "Neutral result panels triggered an implicit read.");
        await verifyCoverage();
        const telemetry = area.querySelector<HTMLElement>('[data-coverage-stream="telemetry"]')!;
        const properties = area.querySelector<HTMLElement>('[data-coverage-stream="twin"]')!;
        const hint = area.querySelector<HTMLElement>(".iot-solution-coverage-assignment-hint")!;
        const color = getComputedStyle(hint).backgroundColor.match(/\d+/g)!.slice(0, 3).map(Number);
        assert(coverageSampleCount(telemetry) === 0 && coverageSampleCount(properties) === 100 &&
            !telemetry.querySelector("table, caption, [data-missing-field]") &&
            properties.querySelectorAll("[data-missing-field]").length === 15 &&
            hint.classList.contains("fui-MessageBar") && color[0] === color[1] && color[1] === color[2] &&
            hint.textContent === "New model assignments can take some time to show up in device events" &&
            area.querySelectorAll(".iot-solution-coverage-assignment-hint").length === 1 && !hint.querySelector("button") &&
            modelCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks),
            "Mixed zero/positive results hid the wrong list, lost the native grey hint or reintroduced waiting consent.");
        assertCoverageMissingGuidance();
        const geometry = coverageCardGeometry(true);
        assert(geometry.naturalHeights[0] < geometry.naturalHeights[1],
            "The equal-height regression did not exercise deliberately different zero and missing content.");
        if (innerWidth > 1100) assert(Math.abs(telemetry.getBoundingClientRect().top - properties.getBoundingClientRect().top) < 1,
            "Paired result cards no longer share their native grid row.");
        emptyProcessingStream = undefined; parityPayloads = undefined;
        caseObservations = { counts: [0, 100], telemetryListAbsent: true, propertyMissingFields: 15,
            nativeGreySingleHint: true, equalNaturalTallestHeight: geometry, dynamicallyGrowsWithHelper: true, implicitReads: 0 };
    }],
    ["Model Coverage decoded raw placeholders use smaller normal multiline examples without changing entered queries or reading data", async () => {
        const seed = context(true);
        seed.hubModel = { ...seed.hubModel, id: "urn:coverage:" + "long-identifier-".repeat(24),
            document: { ...seed.hubModel.document, id: "urn:coverage:" + "long-identifier-".repeat(24) } };
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(seed)}/>);
        const identity = area.querySelector<HTMLElement>("#reference-and-description")!;
        assert(identity.textContent?.includes(seed.hubModel.id) &&
            getComputedStyle(identity).overflowWrap === "anywhere" && !area.querySelector('[aria-label="Current device model"]'),
            "The canonical ID was lost, duplicated or unable to wrap.");
        const intro = getComputedStyle(area.querySelector(".iot-solution-step-description")!);
        const note = area.querySelector("#coverage-sample-note")!, style = getComputedStyle(note);
        assert(note.textContent === "Check up to 100 recent matching events from each table to see which of your selected fields appear" &&
            style.fontSize === "14px" && style.fontSize === intro.fontSize && style.lineHeight === intro.lineHeight &&
            style.fontFamily === intro.fontFamily && style.fontStyle === "normal" &&
            area.querySelector("#coverage-predicate-description")?.textContent ===
                "Enter a KQL condition, like the part after where in a query. Only events that match are included.",
            "Exact coverage/instruction copy or matching normal body typography changed.");
        const editorGeometry: object[] = [];
        for (const [stream, table] of [["telemetry", "Telemetry"], ["twin", "Twins"]] as const) {
            const editor = area.querySelector<HTMLTextAreaElement>(`#hub-${stream}-scope`)!;
            const placeholderStyle = getComputedStyle(editor, "::placeholder"), inputStyle = getComputedStyle(editor);
            assert(editor.placeholder.startsWith(`Condition for ${table}\nExample: `) &&
                editor.placeholder.includes("data.payload.value") && !editor.placeholder.includes("parse_json") &&
                !!area.querySelector(`label[for="${editor.id}"]`) &&
                placeholderStyle.fontStyle === "normal" && placeholderStyle.fontSize === "12px" &&
                placeholderStyle.fontFamily === inputStyle.fontFamily && inputStyle.fontStyle === "normal" &&
                inputStyle.fontSize === "14px" && inputStyle.lineHeight === "20px" &&
                placeholderStyle.lineHeight === "22px" && !editor.placeholder.includes("\n\n") && editor.rows === 5 &&
                (stream !== "telemetry" || editor.placeholder.includes("\n\tand ")),
                "The raw-table placeholder lost its name, generated join indentation, smaller native type or unchanged entered-text style.");
            const mirror = document.createElement("div");
            Object.assign(mirror.style, { position: "absolute", visibility: "hidden", left: "0", top: "0",
                width: `${editor.clientWidth - parseFloat(inputStyle.paddingLeft) - parseFloat(inputStyle.paddingRight)}px`,
                fontFamily: placeholderStyle.fontFamily, fontSize: placeholderStyle.fontSize,
                fontStyle: placeholderStyle.fontStyle, fontWeight: placeholderStyle.fontWeight,
                lineHeight: placeholderStyle.lineHeight, whiteSpace: "pre-wrap", tabSize: placeholderStyle.tabSize,
                overflowWrap: inputStyle.overflowWrap, wordBreak: inputStyle.wordBreak });
            mirror.textContent = editor.placeholder; document.body.appendChild(mirror);
            const requiredHeight = mirror.getBoundingClientRect().height;
            const availableHeight = editor.clientHeight - parseFloat(inputStyle.paddingTop) - parseFloat(inputStyle.paddingBottom);
            mirror.remove();
            assert(requiredHeight <= availableHeight + 1 && editor.scrollWidth <= editor.clientWidth + 1,
                "The multiline placeholder does not fit its current desktop/narrow editor.");
            let nativeLineHeightDelta: number | undefined;
            if (stream === "telemetry") {
                const probe = editor.cloneNode(true) as HTMLTextAreaElement;
                probe.id = "scope-placeholder-native-probe"; probe.value = "";
                Object.assign(probe.style, { position: "absolute", visibility: "hidden", width: `${editor.getBoundingClientRect().width}px`,
                    height: "1px", minHeight: "0", maxHeight: "none", flex: "none" });
                const rule = document.createElement("style");
                editor.parentElement!.appendChild(probe); document.head.appendChild(rule);
                try {
                    rule.textContent = "#scope-placeholder-native-probe::placeholder { line-height:20px !important; }";
                    const before = probe.scrollHeight;
                    rule.textContent = "#scope-placeholder-native-probe::placeholder { line-height:22px !important; }";
                    nativeLineHeightDelta = probe.scrollHeight - before;
                    assert(nativeLineHeightDelta > 0, "The browser ignored placeholder-only line-height; computed CSS is not render proof.");
                } finally { probe.remove(); rule.remove(); }
            }
            editorGeometry.push({ stream, viewport: innerWidth, placeholderFont: placeholderStyle.fontSize,
                placeholderLineHeight: placeholderStyle.lineHeight, nativeLineHeightDelta,
                enteredFont: inputStyle.fontSize, enteredLineHeight: inputStyle.lineHeight, requiredHeight, availableHeight, fits: true });
        }
        const component = 'motor and or "quoted"', key = 'sensor and reading or "state"';
        for (const stream of ["telemetry", "twin"] as const) {
            const fields: ModelField[] = [{ ...seed.hubFields[0], role: stream, component, path: [key] }];
            const formatted = coverageConditionExample(fields, stream, true), plain = coverageConditionExample(fields, stream);
            assert(formatted.split("\n\t").join(" ") === plain && formatted.includes(`[${JSON.stringify(key)}]`) &&
                formatted.includes(JSON.stringify(component)) && formatted.includes("\n\tand ") &&
                (stream !== "telemetry" || formatted.includes("\n\tor ")),
                "Generated join formatting split a quoted key/literal or changed the example's expression meaning.");
        }
        const reads = calls.length, tokens = tokenRequests;
        const typed = 'data.payload.value["sensor and reading or state"] == "reading and standby or stop"';
        await enter("hub-telemetry-scope", typed);
        const entered = area.querySelector<HTMLTextAreaElement>("#hub-telemetry-scope")!;
        assert(!area.querySelector(".iot-solution-coverage-example") && !area.querySelector(".iot-solution-coverage-help") &&
            !button("Copy telemetry example") && !button("Copy reported properties example") && !summary("KQL syntax and available values") &&
            entered.value === typed && currentContext.hubScope.telemetry === typed &&
            getComputedStyle(entered).fontSize === "14px" && getComputedStyle(entered).fontStyle === "normal" &&
            getComputedStyle(entered).lineHeight === "20px" &&
            calls.length === reads && tokenRequests === tokens,
            "Entered text was formatted/restyled or a placeholder/condition action implicitly queried.");
        caseObservations = { exactCopy: true, normalBody14: true, realTablePlaceholders: true, automaticDecodeExamples: true,
            placeholderNormal12: true, generatedJoinNewlineTabOnly: true, quotedKeyAndLiteralSafety: true,
            actualTypedValueAndStyleUnchanged: true, editorGeometry,
            associatedLabels: true, copyAndSyntaxRemoved: true, implicitReads: 0, longIdWrap: true };
    }],
    ["Model Coverage assignment delay hint is exclusive to current successful reference zero results", async () => {
        const text = "New model assignments can take some time to show up in device events";
        const hints = () => [...area.querySelectorAll('[role="note"]')].filter(node => node.textContent === text).length;
        let edit: WizardStepProps["updateContext"] | undefined;
        const View = (props: React.ComponentProps<typeof HubModelScopeStep>) => {
            edit = props.updateContext; return <CoverageView {...props}/>;
        };
        emptyProcessing = true;
        for (const mode of ["and", "or", "data", "unavailable"] as const) {
            const seed = context(true);
            seed.hubScope = { ...seed.hubScope, mode: mode === "data" ? "data" : "reference",
                operator: mode === "or" ? "or" : "and" };
            const initial = mode === "unavailable" ? seed : withCurrentHubSummary(seed);
            await mount(<Harness component={View} initial={initial}/>);
            assert(hints() === 0, "Assignment timing was shown before explicit verification.");
            await verifyCoverage();
            const expected = mode === "and" || mode === "or" ? 1 : 0;
            assert(hints() === expected && !area.querySelector(".iot-solution-coverage-missing-fields") &&
                (!expected || area.querySelector(".iot-solution-coverage-assignment-hint")?.classList.contains("fui-MessageBar")) &&
                modelCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks),
                `Current zero result has the wrong reference-only informational hint: ${mode}`);
            assert(!currentContext.hubAssignmentHistory?.length,
                "The delay-hint fixture must independently verify that no assignment receipt/history is required.");
            if (mode !== "and") continue;
            deferredToken = true; tokenReleases = [];
            await click(button("Check coverage"));
            assert(hints() === 0 && !modelCoverageReady(hubInput(currentContext, sourceFixture), 1, currentContext.hubStreamChecks),
                "A pending recheck retained old zero-result advice or continuation credit.");
            deferredToken = false;
            await act(async () => tokenReleases.splice(0).forEach(release => release()));
            await until(() => !!currentContext.hubCoverage, "Current zero recheck did not settle.");
            assert(hints() === 1, "The current successful zero recheck lost its informational hint.");
            await enter("hub-telemetry-scope", "true");
            await enter("hub-telemetry-scope", "");
            assert(hints() === 0 && !currentContext.hubCoverage, "Filter ABA restored assignment advice.");
            coverageReadFailure = "telemetry";
            await click(button("Check coverage"));
            await until(() => currentContext.hubStreamChecks.telemetry.phase === "error", "Expected current read error did not settle.");
            assert(hints() === 0 && !button("Check coverage")!.disabled &&
                !!area.querySelector(".iot-solution-coverage-error") && !area.querySelector("[data-coverage-stream]"),
                "Read failure masqueraded as a successful zero result.");
            coverageReadFailure = undefined;
            await verifyCoverage();
            assert(hints() === 1, "Explicit error recovery did not restore the genuine current zero advice.");
            const model = currentContext.hubModel;
            await act(async () => edit!("hubModel", { ...model, id: "urn:coverage:another-model",
                document: { ...model.document, id: "urn:coverage:another-model" } }));
            await act(async () => edit!("hubModel", model));
            assert(hints() === 0 && !currentContext.hubCoverage, "Model ABA restored old assignment advice.");
            await verifyCoverage();
            assert(hints() === 1, "Current model recheck did not recover its zero-result advice.");
            await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
            assert(hints() === 0, "Source ownership change retained assignment advice.");
            await act(async () => setTestSourceState(sourceFixture));
            assert(hints() === 0, "Source ABA restored assignment advice without a current check.");
        }
        emptyProcessing = false;
        caseObservations = { referenceAndOrOnly: true, priorAssignmentReceiptNotRequired: true,
            oneHintForBothStreams: true, unverifiedLoadingModelFilterSourceAbaAndErrorsExcluded: true,
            ignoreAndUnavailableExcluded: true, explicitRecheckRecovery: true, noTimerPollOrNewGate: true };
    }],
    ["Model Coverage combined dirty checks cancel on unmount without late reads or proof on remount", async () => {
        await mount(<Harness component={CoverageView} initial={withCurrentHubSummary(context(true))}/>);
        deferredToken = true; tokenReleases = [];
        const before = calls.length;
        await click(button("Check coverage"));
        await until(() => tokenReleases.length === 1, "Both dirty checks did not share owned authorization.");
        await act(async () => setStepVisible(false));
        assert(currentContext.hubStreamChecks.telemetry.phase === "unchecked" &&
            currentContext.hubStreamChecks.twin.phase === "unchecked" && !currentContext.hubCoverage,
            "Unmount retained pending stream proof.");
        deferredToken = false;
        await act(async () => tokenReleases.splice(0).forEach(release => release()));
        await settle();
        assert(calls.length === before, "Disposed authorization dispatched coverage reads.");
        await act(async () => setStepVisible(true));
        await verifyCoverage();
        assert(calls.length === before + 4 && !!currentContext.hubCoverage, "Remount failed to explicitly recover both current proofs.");
        caseObservations = { bothDirtyCanceled: true, disposedAuthorizationReads: 0, explicitRemountRecoveryReads: 4 };
    }],
    ["Model Coverage unavailable DTDL unknown pending failed stale and all-unlinked states stay data-only without Hub discovery", async () => {
        await mount(<div className="iot-solution-step iot-solution-model-preparation iot-solution-model-coverage">
            <MessageBar intent="info" layout="multiline" bottomReflowSpacer={null} className="iot-solution-assignment-notice">
                <MessageBarBody>Model reference filtering is unavailable until your IoT Hub links are checked and confirmed. IoT Hub links are checked after choosing Assign model to devices on the Assign model page.</MessageBarBody>
            </MessageBar>
        </div>);
        const baselineNotice = area.querySelector(".iot-solution-assignment-notice")!;
        const baselineType = getComputedStyle(baselineNotice.querySelector(".fui-MessageBarBody")!);
        const baselineDtdl = { family: baselineType.fontFamily, size: baselineType.fontSize, weight: baselineType.fontWeight,
            lineHeight: baselineType.lineHeight, style: baselineType.fontStyle, color: baselineType.color, padding: getComputedStyle(baselineNotice).padding };
        let currentDtdl: ReturnType<typeof assertDtdlFilteringNotice> | undefined;
        let currentSkipped: ReturnType<typeof assertSkippedFilteringNotice> | undefined;
        const base = { ...context(true), hubScope: { mode: "reference", additional: "specified", operator: "or", telemetry: "", twin: "" } };
        const hub = assignmentInventory().hubs[0];
        const skipped = (initial: Record<string, any>) => ({ ...initial,
            hubAssignmentIntent: { owner: assignmentIntentOwner(initial, sourceFixture), mode: "skip" } });
        const variants: [string, Record<string, any>, boolean, string][] = [
            ["dtdl", { ...context(), hubScope: base.hubScope }, false, "DTD L"],
            ["not-checked", { ...base }, false, "checked and confirmed"],
            ["pending", withCurrentHubSummary(base, [{ ...hub, state: "pending" }]), false, "checked and confirmed"],
            ["failed", withCurrentHubSummary(base, [{ ...hub, state: "unverified" }],
                [{ resource: hub.id, message: "Failed current read.", retryable: true }], false), false, "checked and confirmed"],
            ["stale", withCurrentHubSummary(base, [hub], [], true, { ...sourceFixture, revision: 99 }), false, "checked and confirmed"],
            ["all-unlinked", withCurrentHubSummary(base, [{ ...hub, state: "unlinked" }]), false, "None of the IoT Hubs"],
            ["mixed", withCurrentHubSummary(base, [hub, { ...hub, id: hub.id + "-other", state: "unlinked" }]), true, ""],
            ["explicit-skip", skipped(base), false, "model assignment was skipped"],
            ["stale-skip", { ...base, hubAssignmentIntent: { owner: "stale-owner", mode: "skip" } }, false, "checked and confirmed"],
            ["skipped-pending", skipped(withCurrentHubSummary(base, [{ ...hub, state: "pending" }])), false, "checked and confirmed"],
            ["skipped-failed", skipped(withCurrentHubSummary(base, [{ ...hub, state: "unverified" }],
                [{ resource: hub.id, message: "Failed current read.", retryable: true }], false)), false, "checked and confirmed"],
            ["skipped-unlinked", skipped(withCurrentHubSummary(base, [{ ...hub, state: "unlinked" }])), false, "None of the IoT Hubs"],
            ["skipped-linked", skipped(withCurrentHubSummary(base, [hub])), true, ""],
        ];
        for (const [name, initial, available, copy] of variants) {
            const callsBefore = calls.length, tokensBefore = tokenRequests, consentBefore = consentRequests;
            await mount(<Harness component={CoverageView} initial={available ? initial :
                { ...initial, hubScope: { ...initial.hubScope, additional: "unselected" } }}/>);
            assert(!!area.querySelector('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]') === available && !button("Dismiss") &&
                !area.querySelector('[aria-label="Source IoT Hubs"]') && calls.length === callsBefore &&
                tokenRequests === tokensBefore && consentRequests === consentBefore,
                `Availability fabricated evidence or queried Hubs/telemetry on mount: ${name}`);
            if (name === "explicit-skip") {
                currentSkipped = assertSkippedFilteringNotice();
                assert(JSON.stringify(currentSkipped) === JSON.stringify(baselineDtdl),
                    `Skipped wording-only change altered original native type/inheritance: ${JSON.stringify({ baselineDtdl, currentSkipped })}`);
            } else assert(!area.textContent?.includes("Model ID filtering isn't available because model assignment was skipped."),
                `A distinct unavailable/available cause was falsely relabeled as skipped: ${name}`);
            if (name === "dtdl") {
                currentDtdl = assertDtdlFilteringNotice();
                assert(JSON.stringify(currentDtdl) === JSON.stringify(baselineDtdl),
                    `DTDL wording-only change altered original native type/inheritance: ${JSON.stringify({ baselineDtdl, currentDtdl })}`);
            } else assert(!area.textContent?.includes("Model ID filtering isn't available for DTDL models."),
                `A non-DTDL cause was mislabeled as DTDL: ${name}`);
            if (!available) {
                assert(button("Check coverage")!.disabled, "Unavailable reference mode silently chose additional conditions.");
                await click(area.querySelector('input[value="none"]'));
                assert(area.textContent?.includes(name === "dtdl" ? "Model ID filtering isn't available for DTDL models." : copy) &&
                    !button("Check coverage")!.disabled &&
                    hubInput(currentContext, sourceFixture).scope.mode === "data",
                    `Unavailable reference gate blocked optional conditions or retained a reference predicate: ${name}`);
                if (name !== "all-unlinked" && name !== "skipped-unlinked") assert(!area.textContent?.includes("None of the IoT Hubs"),
                    "Unknown/pending/stale Hub evidence was falsely presented as all-unlinked.");
                await verifyCoverage();
                const queries = calls.slice(callsBefore).filter(query => query.includes("__observation"));
                assert(queries.length > 0 && queries.every(query => !query.includes("tostring(parse_json(data)[\"payload\"][\"memberTypeIRI\"]) ==")),
                    `Unavailable mode still filtered processing on a model reference: ${name}`);
            }
        }
        caseObservations = { variants: variants.map(([name]) => name), hubReads: 0, mountQueries: 0,
            dtdlTypography: { baseline: baselineDtdl, current: currentDtdl, unchanged: true },
            skippedTypography: { baseline: baselineDtdl, current: currentSkipped, unchanged: true },
            unavailableConditionsUsable: true, unknownNotUnlinked: true, mixedAvailable: true };
    }],
    ["Model Coverage current Hub context changes retire reference and Create proofs without borrowing unavailable-mode defaults", async () => {
        let edit: WizardStepProps["updateContext"] | undefined;
        const View = (props: React.ComponentProps<typeof HubModelScopeStep>) => {
            edit = props.updateContext; return <CoverageView {...props}/>;
        };
        await mount(<Harness component={View} initial={{ ...context(true), hubScope: undefined }}/>);
        await click(area.querySelector('input[value="specified"]'));
        await enter("hub-telemetry-scope", "tostring(headers[\"IoTConnectionDeviceId\"]) == \"device-1\"");
        const current = withCurrentHubSummary(currentContext);
        const before = calls.length, tokens = tokenRequests;
        await act(async () => edit!("hubAssignmentSource", current.hubAssignmentSource));
        await act(async () => edit!("hubAssignmentHubSummary", current.hubAssignmentHubSummary));
        assert(area.querySelectorAll('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]:checked').length === 0 && button("Check coverage")!.disabled &&
            currentContext.hubScope.mode === "unselected" && calls.length === before && tokenRequests === tokens,
            "A newly available Hub context borrowed the unavailable data-only fallback as a user selection.");
        await click(area.querySelector('input[value="or"]'));
        await verifyCoverage();
        await act(async () => edit!("hubCreation", { operationId: "old-operation", phase: "completed", fingerprint: "old-proof",
            names: creationRequest().names, databaseName: "Mock database", steps: [], objects: [], awaitingData: false }));
        const reads = calls.length;
        await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
        assert(!area.querySelector('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]') && !currentContext.hubCoverage && !currentContext.hubCreation &&
            currentContext.hubStreamChecks.telemetry.phase === "unchecked" && currentContext.hubStreamChecks.twin.phase === "unchecked" &&
            calls.length === reads, "Stale owner evidence retained available reference selection or processing/Create proof.");
        await act(async () => setTestSourceState({ ...sourceFixture, revision: 5 }));
        assert(!area.querySelector('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]') && hubInput(currentContext, { ...sourceFixture, revision: 5 }).scope.mode === "data",
            "Source ABA restored stale positive Hub evidence.");
        const auth = tokenRequests, queries = calls.length;
        await act(async () => setTestSourceState({ ...sourceFixture, revision: 7, ready: false, reason: "Validate the current source tables." }));
        await click(button("Check coverage"));
        assert(button("Check coverage")!.disabled && tokenRequests === auth && calls.length === queries &&
            area.textContent?.includes("Validate the current source tables."),
            "Invalid source readiness dispatched coverage authorization/queries or hid the required recovery.");
        caseObservations = { unknownToAvailableNoDefault: true, ownerProofRetired: true, sourceAbaGuard: true, implicitReads: 0 };
    }],
    ["Model Coverage real wizard Skip path retains no-Hub-read availability through Back and explicit coverage gates", () => observeWizardContext(async () => {
        const fixture = sourceHubFixture("mixed");
        await wizard(context(true)); await click(next());
        const tokens = tokenRequests;
        await click(button("Skip assignment"));
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" &&
            step("Check model coverage")?.textContent?.includes("Pick devices and confirm coverage") &&
            !area.querySelector('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]') && fixture.calls.length === 0 && tokenRequests === tokens &&
            area.textContent?.includes("Model ID filtering isn't available because model assignment was skipped.") && next()!.disabled,
            "Skip introduced Hub discovery, false availability or a completed coverage gate.");
        assertSkippedFilteringNotice();
        const draft = 'isnotempty(tostring(headers["IoTConnectionDeviceId"]))';
        await enter("hub-telemetry-scope", draft);
        await click(area.querySelector('input[value="none"]'));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(area.querySelector("h2")?.textContent === "Assign the model to devices" && fixture.calls.length === 0,
            "Back from scoped conditions changed the optional zero-query assignment path.");
        await click(next());
        assert(!area.querySelector("#hub-telemetry-scope") && currentContext.hubScope.telemetry === draft &&
            currentContext.hubScope.additional === "none" && fixture.calls.length === 0,
            "Actual footer Back lost None or its inactive draft.");
        await click(area.querySelector('input[value="specified"]'));
        assert(area.querySelector<HTMLTextAreaElement>("#hub-telemetry-scope")!.value === draft,
            "Specify after actual Back did not restore the same-owner draft.");
        await verifyCoverage(); await until(() => !next()!.disabled, "Skip independent conditions did not enable review.");
        await click(next()); assert(area.querySelector("h2")?.textContent === "Create modeled data", "Current coverage did not reach review.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" && !next()!.disabled &&
            fixture.calls.length === 0, "Back lost current coverage or dispatched forbidden Hub discovery.");
        await enter("hub-twin-scope", "tostring(headers[\"IoTConnectionDeviceId\"]) == \"device-1\"");
        assert(next()!.disabled && !currentContext.hubCoverage && !currentContext.hubCreation &&
            fixture.calls.length === 0, "2. Add conditions (optional) did not retire review/Create proof on the real wizard path.");
        caseObservations = { exactStepAndPageCopy: true, skipHubReads: 0, noImpossibleChoiceGate: true,
            explicitTwoChecks: true, currentBackRetained: true, conditionsRetireCreate: true };
    })],
    ["Hub Summary starts only after Assign and warns only about known nonlinked Hubs without gating optional continuation", () => observeWizardContext(async () => {
        const observations: object[] = [];
        for (const mode of ["linked", "mixed", "omitted", "null", "pending", "failed"] as const) {
            const fixture = sourceHubFixture(mode);
            await wizard(context(true)); const raw = analysisQueries().length, tokens = tokenRequests;
            await click(next());
            assert(next()!.disabled && !button("Assign model to devices")!.disabled && fixture.calls.length === 0 &&
                tokenRequests === tokens && !area.querySelector('[aria-label="Source IoT Hubs"]') &&
                !area.querySelector('[role="progressbar"]'), "Untouched landing performed assignment discovery or disabled Assign.");
            await click(button("Skip assignment"));
            assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" && fixture.calls.length === 0 &&
                tokenRequests === tokens, "First Skip requested Hub/auth data instead of normal advancement.");
            await click(area.querySelector('[data-testid="wizard-previous-btn"]')); await settle();
            assert(button("Skip assignment")?.getAttribute("aria-pressed") === "true" && fixture.calls.length === 0 &&
                tokenRequests === tokens, "Back with skip auto-navigated or discovered Hubs.");
            await beginAssignmentTargets();
            await until(() => currentContext.hubAssignmentHubSummary?.value.complete === true, "Current Hub links did not complete.");
            await until(() => !area.querySelector('[role="progressbar"]'), "Chosen assignment checks did not settle.");
            const value = currentContext.hubAssignmentHubSummary.value, unavailable = allAssignmentHubsUnavailable(value);
            const box = area.querySelector('[aria-label="Source IoT Hubs"]');
            const headings = [...(box?.querySelectorAll("h3") || [])].map(node => node.textContent);
            assert(!next()!.disabled && currentContext.hubAssignmentIntent.mode === "assign" &&
                !button("Assign model to devices") && !button("Skip assignment") && fixture.patches.length === 0 &&
                analysisQueries().length === raw, "Hub summary performed device/raw reads, wrote or used wrong severity.");
            if (mode === "linked") {
                assert(!box && deviceDiscoveryCalls(fixture.calls).length > 0 &&
                    area.querySelector('[aria-label="Device model assignments"]'),
                    "Linked-only sources retained the removed summary or lost eligible device discovery.");
            } else {
                assert(box?.getAttribute("data-summary-kind") === "warning" && headings.length === 1 &&
                    headings[0] === (unavailable ? "Your IoT hubs aren't linked to Azure Device Registry"
                        : "Some IoT hubs aren't linked to Azure Device Registry") &&
                    box.querySelector(".iot-solution-outcome-section > .fui-Text")?.textContent === (unavailable
                        ? "Model assignment needs hubs linked to an Azure Device Registry (ADR) namespace."
                        : "Devices on these hubs can't be assigned a model until the hubs are linked to an Azure Device Registry (ADR) namespace.") &&
                    JSON.stringify([...box.querySelectorAll("li")].map(node => node.textContent)) ===
                    JSON.stringify(value.hubs.filter((hub: AssignmentInventory["hubs"][number]) =>
                        hub.state === "unlinked" || hub.state === "pending").map((hub: AssignmentInventory["hubs"][number]) => hub.name)) &&
                    !box.textContent?.includes("IoT Hubs Linked to an Azure Device Registry Namespace"),
                    "Warning retained linked information, changed nonlinked content/order or used the wrong native treatment.");
                const link = box.querySelector<HTMLAnchorElement>("a.fui-Link")!;
                const body = box.querySelector(".iot-solution-outcome-section > .fui-Text")!;
                assert(link?.textContent?.trim() === "Learn how to link an IoT hub to ADR" &&
                    link.getAttribute("href") === "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal" &&
                    link.target === "_blank" && link.rel === "noopener noreferrer" && !link.hasAttribute("disabled") &&
                    link.getAttribute("aria-disabled") !== "true" && !link.hasAttribute("aria-describedby") &&
                    !body.contains(link) && body.nextElementSibling?.contains(link) && !box.querySelector("#adr-link-guidance-unavailable"),
                    "Linking guidance lost its exact separate native link line or enabled external-documentation semantics.");
                const beforeLink = fixture.calls.length;
                assertDocumentationClicks([await documentationClick(link,
                    "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal")]);
                assert(fixture.calls.length === beforeLink && fixture.patches.length === 0,
                    "The documentation link invoked assignment discovery or a mutation.");
                if (mode === "mixed") {
                    assert(!area.querySelector("[data-assignment-unavailable-guidance]") &&
                        deviceDiscoveryCalls(fixture.calls).length > 0 && area.querySelector('[aria-label="Device model assignments"]'),
                        "Mixed warning falsely claimed global unavailability or hid linked devices.");
                }
            }
            if (unavailable) {
                assert(headings.length === 1 && headings[0] === "Your IoT hubs aren't linked to Azure Device Registry" &&
                    !box!.textContent?.includes("Select Continue to skip this step.") &&
                    !area.textContent?.includes("could not be verified") && !button("Assign model to devices") &&
                    !button("Skip assignment") && deviceDiscoveryCalls(fixture.calls).length === 0 &&
                    !area.querySelector('[aria-label="Device model assignments"]'),
                    "Known unavailable Hubs used ambiguous copy or demanded an unusable choice.");
                const guidance = area.querySelector("[data-assignment-unavailable-guidance]")!;
                assert(guidance && !box!.contains(guidance) && guidance.previousElementSibling === box &&
                    guidance.textContent?.trim() === "Select Continue to skip this step." &&
                    guidance.querySelectorAll("strong").length === 1 && guidance.querySelector("strong")?.textContent === "Continue",
                    "All-nonlinked guidance is not outside the warning with exactly Continue bold.");
                const link = box!.querySelector<HTMLAnchorElement>("a.fui-Link")!;
                assert(link.getAttribute("href") === "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal" &&
                    link.target === "_blank" && link.rel === "noopener noreferrer" && !link.hasAttribute("disabled") &&
                    !link.hasAttribute("aria-describedby") && !box!.querySelector("#adr-link-guidance-unavailable"),
                    "Unavailable Hub guidance lost the supplied enabled documentation link.");
                const reads = fixture.calls.length;
                await click(next());
                assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" && next()!.disabled,
                    "Unavailable proof auto-navigated or bypassed independent processing checks.");
                await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
                await settle();
                assert(!next()!.disabled && fixture.calls.length === reads, "Same-owner Back reread Hubs or lost explicit assign intent.");
            }
            if (mode !== "linked") {
                const currentBox = area.querySelector<HTMLElement>('[aria-label="Source IoT Hubs"]')!;
                const dismiss = button("Dismiss")!;
                const heading = currentBox.querySelector("h3")!;
                const body = currentBox.querySelector(".fui-MessageBarBody")!;
                assert(currentBox.contains(dismiss) && dismiss.classList.contains("fui-Button") &&
                    getComputedStyle(dismiss).backgroundColor === "rgba(0, 0, 0, 0)" &&
                    getComputedStyle(dismiss).fontSize === "12px" &&
                    Math.abs(dismiss.getBoundingClientRect().right - body.getBoundingClientRect().right) < 1 &&
                    Math.abs(dismiss.getBoundingClientRect().top - heading.getBoundingClientRect().top) < 1 &&
                    currentBox.scrollWidth <= currentBox.clientWidth + 1 &&
                    dismiss.getBoundingClientRect().left >= heading.getBoundingClientRect().right,
                    "Dismiss is not the native small subtle top-right action or collides with wrapped heading.");
                const before = JSON.stringify({ intent: currentContext.hubAssignmentIntent, summary: currentContext.hubAssignmentHubSummary,
                    selected: currentContext.hubAssignmentChoices, history: currentContext.hubAssignmentHistory });
                const reads = fixture.calls.length, requests = tokenRequests, continueDisabled = next()!.disabled;
                await click(dismiss);
                assert(!area.querySelector('[aria-label="Source IoT Hubs"]') && !button("Dismiss") &&
                    area.querySelector('[aria-label="Selected device model"]') &&
                    !!area.querySelector("[data-assignment-unavailable-guidance]") === unavailable &&
                    JSON.stringify({ intent: currentContext.hubAssignmentIntent, summary: currentContext.hubAssignmentHubSummary,
                        selected: currentContext.hubAssignmentChoices, history: currentContext.hubAssignmentHistory }) === before &&
                    fixture.calls.length === reads && tokenRequests === requests && next()!.disabled === continueDisabled &&
                    document.activeElement === (unavailable ? next() : area.querySelector("#assignment-devices-heading")),
                    "Dismiss affected eligibility/queries/state/guidance/model context or lost logical keyboard focus.");
                if (!unavailable) assert(getComputedStyle(area.querySelector('[aria-label="Device model assignments"]')!).borderTopWidth === "0px" &&
                    getComputedStyle(area.querySelector(".iot-solution-config-card")!).borderTopWidth === "1px" &&
                    getComputedStyle(area.querySelector("tbody td")!).borderTopWidth === "1px",
                    "Assignment divider removal affected native card or table-row boundaries.");
            }
            observations.push({ mode, unavailable, summaryComplete: value.complete, requestsBeforeChoiceOrSkip: 0,
                deviceCallsAfterChoice: deviceDiscoveryCalls(fixture.calls).length });
        }
        caseObservations = { observations, rawQueriesAddedBySummary: 0, actualWrites: 0 };
    })],
    ["Hub Summary failed or missing metadata stays visible after Assign without blocking optional Continue or prechoice Skip", () => observeWizardContext(async () => {
        const observations: object[] = [];
        for (const mode of ["authentication", "permission", "network", "malformed", "partial", "truncated", "missing"] as const) {
            const fixture = sourceHubFixture(mode);
            await wizard(context(true)); const raw = analysisQueries().length, tokens = tokenRequests;
            await click(next());
            await click(button("Skip assignment"));
            assert(fixture.calls.length === 0 && tokenRequests === tokens &&
                area.querySelector("h2")?.textContent === "Filter events and check model coverage", "Unknown provenance/error Skip triggered discovery.");
            await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await beginAssignmentTargets();
            await until(() => !!button("Re-check IoT Hubs") && !button("Re-check IoT Hubs")!.disabled,
                "Failed/incomplete Hub check did not expose explicit retry.");
            const linkedHubs = currentContext.hubAssignmentHubSummary?.value.hubs.filter(
                (hub: AssignmentInventory["hubs"][number]) => hub.state === "linked") || [];
            await until(() => !next()!.disabled, "Settled Hub and device checks did not release optional continuation.");
            assert(!next()!.disabled && !allAssignmentHubsUnavailable(currentContext.hubAssignmentHubSummary?.value) &&
                currentContext.hubAssignmentIntent.mode === "assign" && !button("Skip assignment") &&
                !button("Assign model to devices") && (linkedHubs.length > 0 || deviceDiscoveryCalls(fixture.calls).length === 0),
                "Failed/partial Hub metadata blocked optional continuation or granted discovery without a positive Hub.");
            assert(!area.querySelector('[aria-label="Source IoT Hubs"]'),
                "Failed, partial or unknown reads were fabricated into a nonlinked warning.");
            if (mode === "partial") {
                await until(() => !!currentContext.hubAssignmentInventory && !area.querySelector('[role="progressbar"]'),
                    "Verified positive Hub inventory did not settle beside the explicit partial-read error.");
                const targets: AssignmentTarget[] = currentContext.hubAssignmentInventory.value.targets;
                assert(targets.length > 0 && targets.every(target => target.hubIds.every(id =>
                    linkedHubs.some((hub: AssignmentInventory["hubs"][number]) => hub.id === id))),
                    "Partial/unknown Hub evidence granted a device outside current positively verified Hubs.");
            }
            if (mode === "missing") assert(!area.querySelector('[aria-label="Source IoT Hubs"]') && fixture.calls.length === 0,
                "Missing provenance invented source Hubs or triggered auth/discovery.");
            if (mode === "authentication") {
                assert(button("Sign in"), "Hub authorization failure omitted explicit sign-in recovery.");
                fixture.mode = "linked"; await click(button("Sign in"));
                await until(() => currentContext.hubAssignmentHubSummary?.value.complete, "Hub sign-in retry did not recover current metadata.");
                await until(() => !next()!.disabled, "Recovered Hub/device lists did not release optional continuation.");
                assert(!next()!.disabled && !button("Assign model to devices"), "Read retry changed explicit assignment intent.");
            }
            await click(next());
            assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" &&
                fixture.patches.length === 0 && analysisQueries().length === raw,
                "Explicit optional Skip failed, queried devices or changed source samples.");
            observations.push({ mode, requestsBeforeChoiceOrSkip: 0, explicitAssignContinue: true, deviceEligibilityBeforePositiveHub: false });
        }
        caseObservations = { observations, failuresRemainVisible: true, optionalManualTablesPreserved: true };
    })],
    ["Hub Summary pending Assign blocks Continue until positive current links and device discovery settle", () => observeWizardContext(async () => {
        const fixture = sourceHubFixture("mixed");
        const respond = fixture.fetcher;
        let release: (() => void) | undefined;
        assignmentResponse = (url, init) => new URL(String(url)).pathname === fixture.hub.id && !release
            ? new Promise(resolve => { release = () => { void respond(url, init).then(resolve); }; }) : respond(url, init);
        await wizard(context(true)); await click(next());
        const hubReads = () => fixture.calls.filter(call => new URL(call.url).pathname === fixture.hub.id).length;
        assert(hubReads() === 0 && fixture.calls.length === 0 && !button("Assign model to devices")!.disabled,
            "Landing read Hubs or required pre-discovery eligibility.");
        await beginAssignmentTargets();
        await until(() => !!release, "Explicit Assign did not start its owned Hub check.");
        assert(next()!.disabled && currentContext.hubAssignmentIntent.mode === "assign" &&
            !button("Assign model to devices") && !button("Skip assignment") &&
            !area.querySelector('[aria-label="Device model assignments"]') &&
            deviceDiscoveryCalls(fixture.calls).length === 0 &&
            document.activeElement?.textContent === "Assign the model to devices",
            "Pending Hub checks enabled Continue, lost focus or granted premature device discovery.");
        await act(async () => release!());
        await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') &&
            !area.querySelector('[role="progressbar"]'), "Explicit assignment did not load current eligible devices.");
        assert(hubReads() === 1 && deviceDiscoveryCalls(fixture.calls).length > 0 &&
            !area.querySelector('input[aria-label="Find a Target"]') && !button("Assign model to devices") && !button("Skip assignment"),
            "Active inventory duplicated Hub classification, retained search or reintroduced choice buttons.");
        const active = area.querySelector('[aria-label="Device model assignments"]')!;
        const warning = area.querySelector('[aria-label="Source IoT Hubs"]')!;
        const identity = area.querySelector('[aria-label="Selected device model"]')!;
        assert(!active.querySelector("dl.iot-solution-choice-summary") && !active.querySelector("dt") &&
            !active.textContent?.includes("Hubs not linked to ADR have no ADR assignment targets here.") &&
            active.querySelector("h3")?.textContent === "Select devices" &&
            active.querySelector(".iot-solution-outcome-section > .fui-Text")?.textContent ===
                "Choose the devices to assign this model to. Devices that already have a model can't be selected." &&
            warning.getAttribute("data-summary-kind") === "warning" && warning.textContent?.includes("Other source Hub") &&
            !!(identity.compareDocumentPosition(warning) & Node.DOCUMENT_POSITION_FOLLOWING) &&
            identity.textContent?.includes(tm.id) && !identity.textContent?.includes("Source Hubs") &&
            active.querySelector("tbody tr")?.textContent?.includes("Fixture Hub") &&
            active.querySelector("thead")?.textContent?.includes("Current model"),
            "Removed summary/note remains, or the warning, device heading or per-device provenance/reference was lost.");
        const refresh = button("Refresh devices")!;
        assert(refresh.textContent === "" && refresh.getAttribute("aria-label") === "Refresh devices" && refresh.querySelector("svg"),
            "Device refresh is not a native accessible icon-only action.");
        await click(area.querySelector('input[aria-label="Select Fixture device"]'));
        await click(button("Assign to selected devices"));
        const dialog = document.querySelector('[role="dialog"]')!;
        assert(dialog.textContent?.includes(tm.id) && dialog.textContent?.includes("Fixture device") &&
            dialog.textContent?.includes("Fixture Hub"), "Display removal changed explicit confirmation identity/reference details.");
        await click([...dialog.querySelectorAll("button")].find(node => node.textContent === "Cancel"));
        assert(fixture.patches.length === 0, "Choice/selection/cancel dispatched a device write.");
        await click(refresh); await until(() => !area.querySelector('[role="progressbar"]'), "Explicit device refresh did not settle.");
        assert(hubReads() === 2 && fixture.patches.length === 0,
            "Refresh failed to recheck links or caused an automatic write.");
        caseObservations = { hubReadsBeforeAssign: 0, hubReadsOnAssign: 1, duplicateHubReadsOnDeviceDiscovery: 0,
            pendingChecksContinueEnabled: false, noDevicesBeforePositiveEvidence: true,
            onlyExplicitDeviceDiscovery: true, refreshIconAccessible: true, assignmentsWithoutConfirmation: 0 };
    })],
    ["Hub Summary invalid sources and late model or source ABA cannot publish availability or restore intent", () => observeWizardContext(async () => {
        const fixture = sourceHubFixture("omitted");
        const tokens = tokenRequests;
        const unavailable = { ...sourceFixture, ready: false, reason: "Validate the current source." };
        const Blocked = (props: WizardStepProps) => <HubModelAssignmentStep {...props} workloadClient={client}
            sourceState={unavailable} getSourceState={() => unavailable}/>;
        await mount(<Harness component={Blocked} initial={assignmentContext()}/>);
        await settle();
        assert(tokenRequests === tokens && fixture.calls.length === 0 && !area.querySelector('[aria-label="Source IoT Hubs"]'),
            "Invalid source dispatched Hub metadata authentication/read.");
        const pending: (() => void)[] = [], respond = fixture.fetcher;
        assignmentResponse = (url, init) => new URL(String(url)).pathname === fixture.hub.id
            ? new Promise(resolve => pending.push(() => { void respond(url, init).then(resolve); })) : respond(url, init);
        await wizard(context(true)); await click(next());
        assert(pending.length === 0 && next()!.disabled, "Initial page loaded Hubs before choice.");
        await beginAssignmentTargets();
        await until(() => pending.length === 1, "Owned Hub read did not enter its delayed boundary.");
        await act(async () => observedWizardProps!.updateContext("hubModelRevision", (currentContext.hubModelRevision || 0) + 1));
        await until(() => !!button("Assign model to devices"), "Externally retired model epoch retained its loading navigation gate.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await file("thing-model-upload", JSON.stringify(tm), "model-revision.json"); await click(next());
        assert(pending.length === 1 && next()!.disabled && !button("Assign model to devices")!.disabled,
            "New model owner automatically restored assignment discovery.");
        await act(async () => pending[0]()); await settle();
        assert(next()!.disabled && !currentContext.hubAssignmentHubSummary && !currentContext.hubAssignmentIntent,
            "Late old-model Hub read enabled the current gate.");
        await beginAssignmentTargets();
        await until(() => pending.length === 2, "New explicit choice did not start its own Hub read.");
        await act(async () => pending[1]()); await until(() => !!currentContext.hubAssignmentHubSummary, "Current Hub summary did not publish.");
        await click(button("Dismiss"));
        assert(!area.querySelector('[aria-label="Source IoT Hubs"]'), "Current owner warning did not dismiss.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await until(() => tableChoicesReady(), "Source metadata did not settle for ABA.");
        await chooseOption("hub-model-telemetry", "Show All Tables...");
        await chooseOption("hub-model-telemetry", "NewTelemetry"); await chooseOption("hub-model-telemetry", "Telemetry");
        assert(!currentContext.hubAssignmentHubSummary && !currentContext.hubAssignmentIntent, "Source ABA retained cached unavailable proof.");
        await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Source ABA validation unavailable.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Explicit source validation did not complete.");
        await click(next()); await click(next()); await settle();
        assert(next()!.disabled && pending.length === 2 && !button("Assign model to devices")!.disabled,
            "Source ABA restored choice, summary or automatic Hub reads.");
        await beginAssignmentTargets(); await until(() => pending.length === 3, "Explicit current source choice did not recheck Hubs.");
        await act(async () => pending[2]()); await until(() => !!currentContext.hubAssignmentHubSummary, "Current source summary could not recover.");
        assert(area.querySelector('[aria-label="Source IoT Hubs"]') && button("Dismiss"),
            "Source/model ABA inherited an older context's warning dismissal.");
        assert(deviceDiscoveryCalls(fixture.calls).length === 0 && fixture.patches.length === 0,
            "Hub-only owner transitions queried or changed devices.");
        caseObservations = { invalidSourceAuthCalls: 0, lateModelReadIgnored: true,
            modelRevisionBound: true, sourceAbaRetired: true, currentRecovery: true, deviceCalls: 0 };
    })],
    ["Assignment view toggle stays natively disabled in both views during Refresh and Load more without state or read changes", async () => {
        const discover = HubAssignmentApi.prototype.discover, more = HubAssignmentApi.prototype.loadMore;
        const inventory = { ...deviceListInventory(), hasMore: true };
        const pending: { resolve: (value: AssignmentInventory) => void; reject: (error: Error) => void }[] = [];
        let defer = false, loads = 0;
        HubAssignmentApi.prototype.discover = () => {
            loads++;
            return defer ? new Promise((resolve, reject) => pending.push({ resolve, reject })) : Promise.resolve(inventory);
        };
        HubAssignmentApi.prototype.loadMore = async (_inventory, _signal, publication) => {
            loads++;
            const value = await new Promise<AssignmentInventory>((resolve, reject) => pending.push({ resolve, reject }));
            publication?.publish(value); return value;
        };
        try {
            await openAssignmentTargets();
            await until(() => !!button("Show all devices") && !button("Refresh devices")!.disabled, "Initial list did not settle.");
            await click(area.querySelector('input[aria-label="Select Fixture device"]'));
            const toggle = button("Show all devices")!;
            const selection = () => JSON.stringify(currentContext.hubAssignmentChoices.targets.map((target: AssignmentTarget) => target.id));
            const initialSelection = selection(), observations: object[] = [];
            defer = true;
            for (const view of ["unassigned", "all"] as const) {
                if (view === "all") await click(toggle);
                const caption = view === "unassigned" ? "Show all devices" : "Show unassigned devices";
                for (const mode of ["Refresh devices", "Load more"]) {
                    const index = pending.length;
                    await click(button(mode));
                    await until(() => pending.length === index + 1, "Deferred toolbar list did not start.");
                    const reads = calls.length + resourceCalls.length, tokens = tokenRequests, count = loads;
                    const status = area.querySelector("#assignment-current-view")!.textContent;
                    assert(button(caption) === toggle && toggle.tagName === "BUTTON" && toggle.disabled,
                        `${caption} is not the same natively disabled view control during ${mode}.`);
                    area.querySelector<HTMLElement>("h2")!.focus(); toggle.focus();
                    assert(document.activeElement !== toggle, "Disabled view control remained keyboard focusable.");
                    await click(toggle);
                    await act(async () => {
                        for (const key of ["Enter", " "]) {
                            toggle.dispatchEvent(new KeyboardEvent("keydown", { key, bubbles: true }));
                            toggle.dispatchEvent(new KeyboardEvent("keyup", { key, bubbles: true }));
                        }
                    });
                    assert(button(caption) === toggle && area.querySelector("#assignment-current-view")!.textContent === status &&
                        selection() === initialSelection && calls.length + resourceCalls.length === reads &&
                        tokenRequests === tokens && loads === count,
                        "Disabled mouse/keyboard activation changed view/selection or dispatched another read.");
                    await act(async () => pending[index].resolve(inventory));
                    await until(() => !toggle.disabled, "Current list settlement did not re-enable the view control.");
                    assert(button(caption) === toggle && selection() === initialSelection, "Settlement changed view or selection.");
                    observations.push({ view, mode, nativeDisabled: true, attemptedActivationChanges: 0, additionalReads: 0, settledEnabled: true });
                }
            }
            const index = pending.length;
            await click(button("Refresh devices")); await until(() => pending.length === index + 1, "Error fixture did not start.");
            assert(toggle.disabled, "Error-bound refresh did not disable the view control.");
            await act(async () => pending[index].reject(new Error("Synthetic toolbar inventory failure.")));
            await until(() => !toggle.disabled, "Honest current error left the view control disabled.");
            assert(button("Show unassigned devices") === toggle && selection() === initialSelection &&
                area.textContent?.includes("Synthetic toolbar inventory failure."),
                "Error recovery changed the view/selection or hid the existing error.");
            caseObservations = { observations, sameControlBothCaptions: true, selectionPreserved: true,
                nativeDisabledCannotFocus: true, disabledMouseAndKeyboardNoEffect: true,
                extraReadsOnAttemptedActivation: 0, honestErrorReenabled: true, actualWrites: 0 };
        } finally { HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.loadMore = more; }
    }],
    ["Assignment list loading gates forward footer and sidebar across Hub auth read device Refresh and Load more", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover, more = HubAssignmentApi.prototype.loadMore;
        const pending: { resolve: (value: AssignmentInventory) => void; reject: (error: Error) => void }[] = [];
        const fixture = sourceHubFixture("linked"), respond = fixture.fetcher;
        let releaseHub: (() => void) | undefined, watch = true;
        const enabledGaps: boolean[] = [];
        const blocked = async () => {
            assert(next()!.disabled && !area.querySelector<HTMLButtonElement>('[data-testid="wizard-previous-btn"]')!.disabled &&
                step("Add device model")?.getAttribute("aria-disabled") === "false" &&
                step("Check model coverage")?.getAttribute("aria-disabled") === "true",
                "Current list request failed forward gating or blocked backward navigation.");
            await click(next()); await click(step("Check model coverage"));
            await act(async () => { await observedWizardProps!.onNext(); });
            assert(area.querySelector("h2")?.textContent === "Assign the model to devices",
                "Footer/direct-next/sidebar bypassed the current list guard.");
        };
        HubAssignmentApi.prototype.discover = () => new Promise((resolve, reject) => pending.push({ resolve, reject }));
        HubAssignmentApi.prototype.loadMore = async (_inventory, _signal, publication) => {
            const value = await new Promise<AssignmentInventory>((resolve, reject) => pending.push({ resolve, reject }));
            publication?.publish(value); return value;
        };
        assignmentResponse = (url, init) => new URL(String(url)).pathname === fixture.hub.id
            ? new Promise(resolve => { releaseHub = () => { void respond(url, init).then(resolve); }; }) : respond(url, init);
        let observer: MutationObserver | undefined;
        try {
            await wizard(context(true)); await click(next());
            observer = new MutationObserver(() => { if (watch && next() && !next()!.disabled) enabledGaps.push(true); });
            observer.observe(next()!, { attributes: true, attributeFilter: ["disabled"] });
            deferredToken = true; tokenReleases = [];
            await beginAssignmentTargets();
            await until(() => tokenReleases.length === 1, "Hub authentication did not reach its deferred boundary.");
            await blocked();
            deferredToken = false;
            await act(async () => tokenReleases.splice(0).forEach(release => release()));
            await until(() => !!releaseHub, "Hub metadata read did not reach its deferred boundary.");
            await blocked();
            await act(async () => releaseHub!());
            await until(() => pending.length === 1, "Positive Hub result did not start device inventory.");
            await blocked();
            assert(enabledGaps.length === 0, "Continue had an enabled Hub-to-device handoff gap.");
            watch = false;
            const inventory = { ...deviceListInventory(), hasMore: true };
            await act(async () => pending[0].resolve(inventory));
            await until(() => !next()!.disabled, "Settled inventory did not release optional Continue.");
            const before = tokenRequests;
            await click(button("Refresh devices"));
            await until(() => pending.length === 2, "Explicit Refresh did not start inventory.");
            await blocked();
            await act(async () => pending[1].resolve(inventory));
            await until(() => !next()!.disabled, "Refresh settlement did not release Continue.");
            await click(button("Load more"));
            await until(() => pending.length === 3, "Load more did not start its page.");
            await blocked();
            await act(async () => pending[2].resolve({ ...inventory, hasMore: false }));
            await until(() => !next()!.disabled, "Paging settlement did not release Continue.");
            assert(!button("Load more") && fixture.patches.length === 0 && tokenRequests === before,
                "Paging required a write or introduced unrelated authorization.");
            await click(next());
            assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage",
                "Settled optional assignment did not continue without a write.");
            caseObservations = { hubAuthBlocked: true, hubReadBlocked: true, deviceListBlocked: true,
                enabledHandoffGaps: 0, refreshBlocked: true, pagingBlocked: true,
                footerAndDirectNextAndForwardSidebarBlocked: true, backwardNavigationUsable: true, settledWithoutAssignment: true, writes: 0 };
        } finally {
            observer?.disconnect(); deferredToken = false;
            tokenReleases.splice(0).forEach(release => release());
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.loadMore = more;
        }
    })],
    ["Assignment loading permits Previous and backward sidebar cancellation without late completion affecting remount", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover;
        const pending: ((value: AssignmentInventory) => void)[] = [];
        const fixture = sourceHubFixture("linked");
        HubAssignmentApi.prototype.discover = () => new Promise(resolve => pending.push(resolve));
        try {
            await wizard(context(true)); await click(next()); await beginAssignmentTargets();
            await until(() => pending.length === 1, "Deferred list did not start.");
            const first = currentContext.hubAssignmentIntent.listLoading;
            const previous = () => area.querySelector<HTMLButtonElement>('[data-testid="wizard-previous-btn"]')!;
            assert(next()!.disabled && !previous().disabled, "Loading locked Previous or enabled Continue.");
            await click(previous());
            assert(area.querySelector("h2")?.textContent === "Add a device model" && !first.busy,
                "Previous did not leave and retire its owned list.");
            await click(next());
            await until(() => pending.length === 2, "Returning assignment did not start its current list.");
            const second = currentContext.hubAssignmentIntent.listLoading;
            assert(next()!.disabled && !previous().disabled && first !== second, "Remount reused old busy state or blocked backward navigation.");
            await act(async () => pending[0](assignmentInventory())); await settle();
            assert(next()!.disabled && second.busy && !currentContext.hubAssignmentInventory,
                "Late canceled completion cleared the remount gate or published old data.");
            await click(step("Add device model"));
            assert(area.querySelector("h2")?.textContent === "Add a device model" && !second.busy,
                "Backward sidebar navigation did not cancel its owned list.");
            await click(next());
            await until(() => pending.length === 3, "Second remount did not acquire its own list.");
            const third = currentContext.hubAssignmentIntent.listLoading;
            await act(async () => pending[1](assignmentInventory())); await settle();
            assert(next()!.disabled && third.busy && third !== second && !currentContext.hubAssignmentInventory,
                "Late sidebar-canceled completion affected the newer request.");
            await act(async () => pending[2](assignmentInventory()));
            await until(() => !next()!.disabled, "Current list did not release Continue on settlement.");
            await click(next()); await verifyCoverage(); await click(next());
            assert(step("Check model coverage")?.querySelector('[aria-label="Completed"]'),
                "The focused forward-sidebar setup did not mark its prior coverage step completed.");
            const coverage = currentContext.hubCoverage, checks = currentContext.hubStreamChecks;
            await click(step("Assign model"));
            await until(() => pending.length === 4, "Returning from completed coverage did not start its current list.");
            assert(next()!.disabled && !previous().disabled &&
                step("Check model coverage")?.getAttribute("aria-disabled") === "true" &&
                currentContext.hubCoverage === coverage && currentContext.hubStreamChecks === checks,
                "Forward navigation credit was not retired, or loading cleared unrelated coverage evidence.");
            await click(step("Check model coverage"));
            await act(async () => { await observedWizardProps!.onGoToStep("scope-coverage"); });
            assert(area.querySelector("h2")?.textContent === "Assign the model to devices",
                "A previously completed forward sidebar step bypassed the loading gate.");
            await click(previous());
            await act(async () => pending[3](assignmentInventory()));
            assert(!previous().disabled && fixture.patches.length === 0, "Backward navigation or zero-write behavior regressed.");
            caseObservations = { previousUsableWhileLoading: true, backwardSidebarUsable: true,
                disposeOnBack: true, separateRemountLeases: true, lateCompletionIgnored: true,
                currentSettlementReleasesContinue: true, completedForwardSidebarBlocked: true,
                coverageEvidencePreserved: true, writes: 0 };
        } finally { HubAssignmentApi.prototype.discover = discover; }
    })],
    ["Assignment list loading releases settled Hub errors unlinked empty and device errors without requiring writes", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover;
        try {
            for (const kind of ["unlinked", "hub-error", "empty", "device-error"] as const) {
                const fixture = sourceHubFixture(kind === "unlinked" ? "omitted" : kind === "hub-error" ? "network" : "linked");
                let release: (() => void) | undefined;
                HubAssignmentApi.prototype.discover = () => new Promise((resolve, reject) => {
                    release = () => kind === "device-error" ? reject(new Error("Synthetic device listing failed."))
                        : resolve({ ...assignmentInventory(), targets: [] });
                });
                await wizard(context(true)); await click(next()); await beginAssignmentTargets();
                if (kind === "empty" || kind === "device-error") {
                    await until(() => !!release, "Deferred device list did not start.");
                    assert(next()!.disabled, "Deferred empty/error list was not gated.");
                    await act(async () => release!());
                }
                await until(() => !next()!.disabled, `Settled ${kind} request retained a perpetual loading gate.`);
                if (kind === "device-error") assert(area.textContent?.includes("Synthetic device listing failed."), "Device error notification was lost.");
                if (kind === "hub-error") assert(button("Re-check IoT Hubs"), "Hub error recovery was lost.");
                if (kind === "unlinked") assert(area.textContent?.includes("Your IoT hubs aren't linked to Azure Device Registry") &&
                    area.textContent?.includes("Model assignment needs hubs linked to an Azure Device Registry (ADR) namespace."),
                    "Known-unlinked guidance was lost.");
                assert(!currentContext.hubAssignmentBusy && fixture.patches.length === 0 &&
                    !currentContext.hubAssignmentIntent.listLoading.busy, "Settled optional discovery demanded mutation/success.");
                await click(next());
                assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage",
                    `Settled ${kind} branch could not continue.`);
            }
            caseObservations = { settledBranches: ["unlinked", "hub-error", "empty", "device-error"],
                optionalContinuationReleased: true, existingErrorsAndRecoveryRetained: true, writes: 0 };
        } finally { HubAssignmentApi.prototype.discover = discover; }
    })],
    ["Assignment list loading disposal and model ABA cannot clear a newer busy owner or retain old inventory", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover;
        const pending: ((value: AssignmentInventory) => void)[] = [];
        HubAssignmentApi.prototype.discover = () => new Promise(resolve => pending.push(resolve));
        try {
            let previous: Record<string, any> | undefined;
            for (const revision of [1, 2, 3]) {
                const fixture = sourceHubFixture("linked");
                await wizard({ ...context(true), hubModelRevision: revision }); await click(next()); await beginAssignmentTargets();
                await until(() => pending.length === revision, "Current model-revision list did not start.");
                const current = currentContext.hubAssignmentIntent.listLoading;
                assert(next()!.disabled && current.busy, "Current owner was not gated.");
                if (previous) {
                    assert(!previous.loading.busy && previous.loading !== current &&
                        previous.owner !== currentContext.hubAssignmentIntent.owner, "Model ABA reused or retained an old busy owner.");
                    await act(async () => pending[revision - 2](assignmentInventory())); await settle();
                    assert(next()!.disabled && current.busy && currentContext.hubAssignmentIntent.listLoading === current &&
                        !currentContext.hubAssignmentInventory, "Late old inventory/completion cleared or replaced a newer busy owner.");
                }
                assert(fixture.patches.length === 0, "Owner change wrote references.");
                previous = { loading: current, owner: currentContext.hubAssignmentIntent.owner };
            }
            await act(async () => pending[2](assignmentInventory()));
            await until(() => !next()!.disabled, "Current ABA owner did not release after its own list settled.");
            caseObservations = { revisions: [1, 2, 3], disposedBusyRetired: true,
                lateOldCompletionIgnored: true, newOwnerRemainsBlocked: true, currentSettlementReleases: true, writes: 0 };
        } finally { HubAssignmentApi.prototype.discover = discover; }
    })],
    ["Assignment Intent landing requires explicit assign or skip before any Hub or device discovery", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover;
        const discoverHubs = HubAssignmentApi.prototype.discoverHubs;
        let hubChecks = 0;
        HubAssignmentApi.prototype.discoverHubs = async source => {
            hubChecks++;
            return { sourceSignature: canonicalJson(source), hubs: assignmentInventory().hubs, issues: [], complete: true };
        };
        let discoveries = 0, release: ((value: AssignmentInventory) => void) | undefined;
        const requests: string[] = [], writes: string[] = [];
        assignmentResponse = async (url, init) => {
            if (["PATCH", "PUT", "DELETE"].includes(init?.method || "")) writes.push(String(url));
            requests.push(String(url)); throw new Error("No live assignment transport in landing test.");
        };
        HubAssignmentApi.prototype.discover = () => { discoveries++; return new Promise(resolve => { release = resolve; }); };
        try {
            const history = [{ operationId: "past", targetId: "past-device", name: "Past receipt", kind: "device",
                modelId: tm.id, state: "confirmed", message: "Prior verified result." }];
            const seed = { ...context(true), hubAssignmentHistory: history };
            await wizard(seed);
            const rawReads = analysisQueries().length, tokens = tokenRequests;
            await click(next());
            assert(area.querySelector("h2")?.textContent === "Assign the model to devices" &&
                step("Assign model")?.querySelector(".wizard-control__step-number")?.textContent === "3" &&
                step("Assign model")?.textContent?.includes("Optional: link devices to the model") &&
                area.querySelector(".iot-solution-assignment > .iot-solution-step-description")?.textContent?.trim() ===
                "Optional: assign this model to your devices to get these benefits:",
                "Assignment landing title, step text or final selected introduction differs.");
            const benefits = area.querySelectorAll(".iot-solution-assignment > ul.iot-solution-outcome-list");
            assert(benefits.length === 1 && JSON.stringify([...benefits[0].children].map(node =>
                ({ tag: node.tagName, text: node.textContent }))) === JSON.stringify([
                { tag: "LI", text: "Device data is tagged with the model ID automatically" },
                { tag: "LI", text: "Data flows into your modeled data without extra setup" },
                { tag: "LI", text: "Fewer setup steps as you build your IoT solution" },
                { tag: "LI", text: "Easier updates as your devices and data change" },
            ]) && !benefits[0].querySelector("strong,em,a,button"), "Assignment benefits are not the exact four plain native semantic bullets.");
            await settle(); await settle();
            assert(next()!.disabled && button("Assign model to devices") && button("Skip assignment") &&
                !area.querySelector('[aria-label="Device model assignments"]') && discoveries === 0 &&
                requests.length === 0 && hubChecks === 0 && tokenRequests === tokens && !currentContext.hubAssignmentIntent &&
                !area.querySelector('[data-testid="wizard-cancel-btn"]'), "Untouched assignment fabricated a choice, discovery or navigation.");
            await click(step("Check model coverage"));
            assert(area.querySelector("h2")?.textContent === "Assign the model to devices", "Sidebar bypassed required intent.");
            await beginAssignmentTargets();
            await until(() => !!release, "Explicit assign choice did not start existing read-only discovery.");
            assert(currentContext.hubAssignmentIntent.mode === "assign" && next()!.disabled &&
                !currentContext.hubAssignmentChoices?.targets?.length && writes.length === 0 &&
                !button("Assign model to devices") && !button("Skip assignment") &&
                document.activeElement?.textContent === "Select devices",
                "Choosing assign failed to gate its pending inventory, retained initial buttons, lost focus, preselected targets or wrote references.");
            await click(next());
            assert(area.querySelector("h2")?.textContent === "Assign the model to devices", "Loading inventory allowed footer advancement.");
            await act(async () => release!(assignmentInventory())); await settle();
            await until(() => !next()!.disabled, "Settled optional inventory did not release Continue.");
            await click(next());
            assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" && next()!.disabled &&
                !currentContext.hubCoverage,
                "Settled intent choice bypassed scoped checks.");
            await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await until(() => discoveries === 2, "Back with assign intent did not reveal the existing active flow.");
            assert(next()!.disabled, "Remounted inventory read did not block Continue.");
            await act(async () => release!(assignmentInventory()));
            await until(() => !next()!.disabled, "Remounted optional inventory did not release Continue.");
            assert(!button("Assign model to devices") && !button("Skip assignment") && !next()!.disabled &&
                area.querySelector('[aria-label="Device model assignments"]'), "Back reintroduced the initial chooser or replayed assignment.");
            await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await file("thing-model-upload", JSON.stringify(tm), "same-model-new-choice.json"); await click(next());
            const beforeSkip = discoveries;
            await click(button("Skip assignment"));
            assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" &&
                currentContext.hubAssignmentIntent.mode === "skip" &&
                step("Assign model")?.querySelector('[aria-label="Completed"]') &&
                currentContext.hubAssignmentHistory === history && discoveries === beforeSkip && writes.length === 0,
                "First skip click failed normal advancement/completion, discovered targets or changed real history.");
            await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await settle(); await settle();
            assert(button("Skip assignment")?.getAttribute("aria-pressed") === "true" && !next()!.disabled &&
                discoveries === 2 && writes.length === 0 && analysisQueries().length === rawReads &&
                area.querySelector("h2")?.textContent === "Assign the model to devices",
                "Same-owner Back lost skip, auto-advanced in a loop or replayed discovery/writes.");
            caseObservations = { initialIntentRequired: true, untouchedDiscovery: 0, assignChoiceWaitsOnlyForLists: true,
                initialButtonsHiddenAfterAssign: true, activeHeadingFocused: true, skipAutoAdvances: true,
                normalCompletionMarked: true, backNoAutoLoop: true, receiptsPreserved: true,
                sidebarGate: true, scopedChecksIndependent: true, noPreselection: true, assignmentWrites: writes.length, extraRawReads: 0 };
        } finally { HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.discoverHubs = discoverHubs; }
    })],
    ["Assignment DTDL uses footer Continue without inline choices and preserves model fields and zero discovery on Back", () => observeWizardContext(async () => {
        const requests: string[] = [];
        assignmentResponse = async url => { requests.push(String(url)); throw new Error("DTD L must not query ADR."); };
        await wizard();
        const model = currentContext.hubModel, fields = JSON.stringify(currentContext.hubFields), rawReads = analysisQueries().length;
        const tokens = tokenRequests;
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Assign the model to devices" &&
            area.textContent?.includes("Model assignment isn't available for DTDL models.") &&
            area.textContent?.includes("go back and upload a WoT Thing Model") && !button("Assign model to devices") &&
            !area.querySelector('[aria-label="Device model assignments"]') && !next()!.disabled &&
            !button("Skip assignment") && !button("Return to Device Model") &&
            !step("Assign model")?.classList.contains("wizard-control__step--unavailable"),
            "DTD L still skipped the page or mounted WoT assignment inventory.");
        const info = area.querySelector(".fui-MessageBar")!;
        const guidance = area.querySelector("[data-dtdl-assignment-guidance]")!;
        assert(guidance?.textContent === "Select Continue to skip this step." &&
            guidance.previousElementSibling === info && !info.contains(guidance) &&
            guidance.querySelectorAll("strong").length === 1 && guidance.querySelector("strong")?.textContent === "Continue" &&
            !currentContext.hubAssignmentIntent && requests.length === 0 && tokenRequests === tokens,
            "DTD L continuation guidance is not exact/outside/only Continue bold or fabricated assignment state.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(currentContext.hubModel === model && JSON.stringify(currentContext.hubFields) === fields &&
            area.querySelector("#model-fields-heading") && !currentContext.hubAssignmentIntent,
            "Returning to the model changed format/fields or fabricated assignment intent.");
        await click(next()); assert(!next()!.disabled, "Valid DTD L still demands an optional assignment choice.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" &&
            !currentContext.hubAssignmentIntent &&
            step("Assign model")?.querySelector('[aria-label="Completed"]') &&
            step("Check model coverage")?.querySelector(".wizard-control__step-number")?.textContent === "4" && next()!.disabled,
            "DTD L footer Continue did not advance normally or skipped scoped requirements.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await settle(); await settle();
        assert(!button("Skip assignment") && !button("Return to Device Model") && !next()!.disabled &&
            !currentContext.hubAssignmentIntent && area.querySelector("h2")?.textContent === "Assign the model to devices",
            "DTD L Back invented choice state or auto-advanced without another click.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(currentContext.hubModel === model && JSON.stringify(currentContext.hubFields) === fields &&
            requests.length === 0 && tokenRequests === tokens && analysisQueries().length === rawReads,
            "DTD L assignment page changed declared fields or requested assignment auth/data.");
        const dtdlRequests = requests.length;
        await click(area.querySelector('input[value="wot"]'));
        await file("thing-model-upload", JSON.stringify(tm), "wot-after-dtdl.json"); await click(next());
        assert(next()!.disabled && !currentContext.hubAssignmentIntent && button("Assign model to devices"),
            "Format replacement borrowed a DTDL continuation exemption.");
        caseObservations = { dtdlVisitsStep3: true, footerContinueWithoutChoice: true, noInventedIntent: true,
            declaredRolesPreserved: true, sameOwnerBackRetained: true, formatInvalidatesChoice: true, dtdlAdrRequests: dtdlRequests };
    })],
    ["Assignment Intent source and model ABA retire completed sidebar navigation without replaying assignments", () => observeWizardContext(async () => {
        await wizard(context(true));
        await click(next()); await click(button("Skip assignment"));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        const before = currentContext.hubAssignmentIntent.owner;
        await file("thing-model-upload", JSON.stringify({ ...tm, description: "New model declaration" }), "changed.json");
        await file("thing-model-upload", JSON.stringify(tm), "original-again.json");
        assert(!currentContext.hubAssignmentIntent && step("Check model coverage")?.getAttribute("aria-disabled") === "true",
            "Model ABA borrowed completed assignment intent/sidebar access.");
        await click(next()); assert(next()!.disabled, "Model ABA bypassed explicit assignment choice.");
        await click(button("Skip assignment"));
        assert(currentContext.hubAssignmentIntent.owner !== before, "Model ABA reused its old intent owner.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await until(() => tableChoicesReady(), "Source metadata did not settle on Back.");
        await chooseOption("hub-model-telemetry", "Show All Tables...");
        await chooseOption("hub-model-telemetry", "NewTelemetry");
        await chooseOption("hub-model-telemetry", "Telemetry");
        assert(!currentContext.hubAssignmentIntent && step("Check model coverage")?.getAttribute("aria-disabled") === "true",
            "Source ABA retained assignment intent or sidebar completion.");
        await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Source ABA validation unavailable.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Source ABA validation did not complete.");
        await click(next()); await click(next());
        assert(area.querySelector("h2")?.textContent === "Assign the model to devices" && next()!.disabled &&
            !area.querySelector('[aria-label="Device model assignments"]'), "Revalidated source silently restored assign/skip.");
        caseObservations = { sourceAbaInvalidatesChoice: true, modelAbaInvalidatesChoice: true,
            completedSidebarRetired: true, noImplicitInventory: true };
    })],
    ["Assignment Intent busy writes lock choice and navigation while switching preserves confirmed receipts", () => observeWizardContext(async () => {
        const clock = assignmentClock();
        try {
        const reads: string[] = [], state: { patches: { body: unknown; ifMatch: string | null }[]; modelId?: string } = { patches: [] };
        const respond = adrOnlyResponse(reads, state);
        let release: (() => void) | undefined;
        assignmentResponse = (url, init) => init?.method === "PATCH"
            ? new Promise(resolve => { release = () => { void respond(url, init).then(resolve); }; })
            : respond(url, init);
        await wizard(context(true)); await click(next()); await beginAssignmentTargets();
        await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') &&
            !area.querySelector('[role="progressbar"]'), "Explicit assignment did not load authoritative target.");
        await click(area.querySelector('input[aria-label="Select Fixture device"]'));
        assert(state.patches.length === 0 && !next()!.disabled, "Target selection wrote or changed intent gating.");
        await click(button("Assign to selected devices"));
        const dialog = document.querySelector('[role="dialog"]')!;
        await click([...dialog.querySelectorAll("button")].find(node => node.textContent === "Assign"));
        await until(() => !!release && !!currentContext.hubAssignmentBusy, "Confirmed mock write did not enter busy state.");
        assert(next()!.disabled && !button("Assign model to devices") && !button("Skip assignment") &&
            area.querySelector<HTMLButtonElement>('[data-testid="wizard-previous-btn"]')!.disabled &&
            step("Add device model")?.getAttribute("aria-disabled") === "true", "Busy assignment permitted intent or navigation changes.");
        await act(async () => release!());
        await until(() => currentContext.hubAssignmentBusy &&
            currentContext.hubAssignmentHistory?.some((receipt: { state: string }) => receipt.state === "accepted"),
            "Mock accepted assignment did not preserve its unconfirmed receipt.");
        await until(() => clock.pending.size === 1, "The accepted operation did not schedule read-only confirmation.");
        delete state.modelId;
        const readsBeforeCompletion = reads.length;
        await clock.tick();
        await until(() => !currentContext.hubAssignmentBusy &&
            currentContext.hubAssignmentHistory?.some((receipt: { state: string }) => receipt.state === "confirmed"),
            "The original native operation did not confirm completion.");
        assert(reads.slice(readsBeforeCompletion).every(url => url.includes("/asyncOperationStatuses/")) &&
            state.modelId === undefined,
            "Completion fetched mutable device state or depended on another actor retaining the model reference.");
        assertDpsGuidance(true);
        const history = JSON.stringify(currentContext.hubAssignmentHistory);
        await click(next());
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(!button("Assign model to devices") && !button("Skip assignment") &&
            completedAssignmentList() &&
            !area.querySelector("table") && state.patches.length === 1,
            "Back with retained assign intent reintroduced the chooser or replayed writes.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await file("thing-model-upload", JSON.stringify(tm), "same-model-new-decision.json"); await click(next());
        assert(!currentContext.hubAssignmentIntent && !area.textContent?.includes("assigned successfully") &&
            !area.querySelector('[aria-label="Selected device model"]'), "Model revision restored an old success proof.");
        assertDpsGuidance(false);
        await click(button("Skip assignment"));
        assert(JSON.stringify(currentContext.hubAssignmentHistory) === history && state.patches.length === 1 &&
            area.querySelector("h2")?.textContent === "Filter events and check model coverage",
            "Skip undid or replayed a confirmed assignment.");
        caseObservations = { busyChoiceAndNavigationLocked: true, explicitMockWrites: state.patches.length,
            confirmedReceiptRetainedAcrossChoices: true, noReplayedWrite: true, actualCloudWrites: 0 };
        } finally { await mount(<div>Busy fixture closed.</div>); clock.restore(); }
    })],
    ["Assignment waves admit five devices on ten-second gates and keep continuation blocked through all observations", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover, hubs = HubAssignmentApi.prototype.discoverHubs;
        const recheck = HubAssignmentApi.prototype.recheck, clock = assignmentWaveClock(), originalRandom = Math.random;
        const observations: object[] = [];
        Math.random = () => 0.5;
        try {
          for (const outcome of ["confirmed", "mixed", "source-change"] as const) {
            const inventory = assignmentInventory(), writes: string[] = [], polls: string[] = [], reads: string[] = [];
            const active = new Set<string>(), release = new Map<number, () => void>();
            const start = clock.elapsed(), admissions = new Map<string, number>();
            let maxActive = 0, abortedStatusReads = 0;
            inventory.targets = Array.from({ length: 12 }, (_, index) => ({ ...inventory.targets[0],
                id: inventory.targets[0].id + `-${index}`, name: `Wave device ${String(index + 1).padStart(2, "0")}`,
                updatePermission: "granted", writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
                deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: `parallel-${index}` } }));
            HubAssignmentApi.prototype.discoverHubs = async source => ({
                sourceSignature: canonicalJson(source), hubs: inventory.hubs, issues: [], complete: true,
            });
            HubAssignmentApi.prototype.discover = async () => inventory;
            HubAssignmentApi.prototype.recheck = async target => {
                if (!reads.includes(target.id)) {
                    active.add(target.id); maxActive = Math.max(maxActive, active.size);
                    admissions.set(target.id, clock.elapsed() - start);
                }
                reads.push(target.id); return target;
            };
            assignmentResponse = async (input, init) => {
                const url = new URL(String(input));
                if (init?.method === "PATCH") {
                    const target = inventory.targets.find(value => value.id === url.pathname)!;
                    assert(target && new Headers(init.headers).get("If-Match") === target.etag &&
                        JSON.stringify(JSON.parse(String(init.body))) === JSON.stringify({ properties: { deviceTypeRefs: [tm.id] } }),
                        "A parallel assignment lost its current conditional refs-only write.");
                    writes.push(target.id);
                    return new Response(null, { status: 202, headers: { Location:
                        `https://management.azure.com/subscriptions/${target.id.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/${inventory.targets.indexOf(target)}?api-version=2026-11-02-preview`,
                        "Retry-After": "3600" } });
                }
                assert(url.pathname.includes("/operationStatuses/"), "Completion fetched mutable device state.");
                const index = Number(url.pathname.split("/").at(-1));
                polls.push(inventory.targets[index].id);
                init?.signal?.addEventListener("abort", () => { abortedStatusReads++; active.delete(inventory.targets[index].id); }, { once: true });
                return new Promise(resolve => {
                    release.set(index, () => {
                        active.delete(inventory.targets[index].id);
                        resolve(new Response(JSON.stringify({ status: outcome === "mixed" && index === 2 ? "Failed"
                            : outcome === "mixed" && index === 11 ? "Running" : "Succeeded" })));
                    });
                });
            };
            await wizard(context(true)); await click(next()); await beginAssignmentTargets();
            await until(() => !!button("Select all") && !button("Select all")!.disabled, "Parallel devices did not load.");
            const selectedOrder = inventory.targets.map(target => target.name);
            await click(button("Select all")); await click(button("Assign to selected devices"));
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')]
                .find(button => button.textContent === "Assign"));
            const stableRows = () => [...area.querySelectorAll(".iot-solution-creation-step-label")].map(node => node.textContent);
            const assertPending = async () => {
                assert(next()!.disabled && currentContext.hubAssignmentBusy &&
                    JSON.stringify(stableRows()) === JSON.stringify(selectedOrder),
                    "A settled device/wave unlocked Continue or reordered/removed selected progress rows while other waves remained active.");
                await act(async () => { await observedWizardProps!.onNext(); });
                assert(step("Assign model")?.getAttribute("aria-current") === "step",
                    "Programmatic advancement left an active overlapping-wave attempt.");
            };
            await until(() => writes.length === 5, "The first wave did not admit exactly five devices.");
            assert(currentContext.hubAssignmentHistory.filter((receipt: { readbackDeadline?: number }) =>
                receipt.readbackDeadline !== undefined).length === 5, "Waiting wave targets acquired polling deadlines.");
            await assertPending();
            await clock.advance(3000);
            await until(() => release.size === 5, "The five active devices did not poll their own operations independently.");
            await act(async () => release.get(4)!());
            await until(() => area.querySelectorAll('[data-assignment-state="confirmed"]').length === 1, "The first confirmation was not published.");
            assert(writes.length === 5 && polls.length === 5, "One settled device must not start another wave or refill a rolling slot.");
            await assertPending();
            await clock.advance(6999); assert(writes.length === 5, "A new wave started before ten seconds or whole-wave settlement.");
            await clock.advance(1);
            await until(() => writes.length === 10, "The ten-second gate did not start the second wave while older requests remained active.");
            assert(maxActive > 5 && polls.length === 5, "A hidden five-in-flight cap delayed the new wave or older operations were re-polled.");
            await assertPending();
            if (outcome === "source-change") {
                const operationId = currentContext.hubAssignmentIntent.attempt.operationId;
                await act(async () => observedWizardProps!.updateContext("hubModelRevision", (currentContext.hubModelRevision || 0) + 1));
                await until(() => !currentContext.hubAssignmentBusy && !currentContext.hubAssignmentIntent &&
                    clock.pending.size === 0, "A new source/model owner retained the retired queue or operation timers.");
                await settle();
                const receipts = currentContext.hubAssignmentHistory.filter((receipt: { operationId: string }) => receipt.operationId === operationId);
                assert(writes.length === 10 && receipts.length === 12 &&
                    receipts.find((receipt: { targetId: string }) => receipt.targetId === inventory.targets[4].id)?.state === "confirmed" &&
                    inventory.targets.slice(10).every(target => receipts.some((receipt: { targetId: string; state: string; writeDispatched: boolean }) =>
                        receipt.targetId === target.id && receipt.state === "cancelled" && receipt.writeDispatched === false)) &&
                    abortedStatusReads === 4,
                    "Owner disposal dispatched queued work, erased confirmation or left active observations running.");
                const beforeLate = JSON.stringify(currentContext.hubAssignmentHistory);
                await act(async () => release.forEach((complete, index) => { if (index !== 4) complete(); }));
                await clock.advance(60000);
                assert(writes.length === 10 && JSON.stringify(currentContext.hubAssignmentHistory) === beforeLate && clock.pending.size === 0,
                    "Late retired operation results republished or restarted assignments.");
                observations.push({ outcome, admitted: 10, queuedNeverSent: 2, confirmedHistoryRetained: true,
                    abortedStatusReads, noReplay: true });
                await mount(<div>Retired parallel batch closed.</div>);
                continue;
            }
            await clock.advance(1000);
            await act(async () => { for (const index of [0, 1, 2, 3]) release.get(index)!(); });
            await until(() => confirmedDeviceRows() === (outcome === "mixed" ? 4 : 5), "The original wave did not settle.");
            assert(writes.length === 10, "An older wave finishing released the newer wave's gate.");
            await assertPending();
            await clock.advance(2000);
            await until(() => release.size === 10, "The second wave did not poll its own original operations.");
            await act(async () => release.get(5)!()); assert(writes.length === 10, "One member of the current wave released the next gate.");
            await clock.advance(1000);
            await act(async () => { for (const index of [6, 7, 8, 9]) release.get(index)!(); });
            await until(() => writes.length === 12, "The fully settled second wave did not admit the next wave early.");
            assert(inventory.targets.slice(10).every(target => admissions.get(target.id) === 14000),
                "The next gate was not relative to the whole current wave's actual settlement.");
            await assertPending();
            await clock.advance(3000);
            await until(() => release.size === 12, "The last wave did not observe the original accepted operations.");
            await act(async () => release.get(10)!()); await assertPending();
            await act(async () => release.get(11)!());
            if (outcome === "mixed") {
                await assertPending();
                await clock.advance(26999); await assertPending();
                await clock.advance(1);
            }
            await until(() => !currentContext.hubAssignmentBusy && !next()!.disabled && clock.pending.size === 0,
                "The last admitted observation did not settle before Continue was released.");
            assert(maxActive > 5 && writes.length === 12 && polls.length === (outcome === "mixed" ? 13 : 12) &&
                inventory.targets.every(target => reads.filter(id => id === target.id).length === 2),
                "Wave admission, own-operation request counts or per-device pre-write rechecks changed.");
            assert(JSON.stringify(stableRows()) === JSON.stringify(selectedOrder), "Out-of-order completion changed stable selected-device identity.");
            if (outcome === "confirmed") assertCompletedAssignments(selectedOrder);
            else assert(area.querySelectorAll('[data-assignment-state="confirmed"]').length === 10 &&
                area.querySelector('[data-assignment-state="failed"]') && area.querySelector('[data-assignment-state="unknown"]') &&
                area.textContent?.includes(ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE) && !completedAssignmentList() &&
                !area.querySelector('[aria-label="DPS enrollment guidance"]'),
                "Mixed parallel outcomes fabricated success or discarded completed devices.");
            await click(next());
            assert(step("Check model coverage")?.getAttribute("aria-current") === "step", "Settled batch did not permit normal continuation.");
            await verifyCoverage();
            const checks = currentContext.hubStreamChecks;
            await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await until(() => !next()!.disabled, "Settled parallel results did not remain eligible on Back.");
            assert(currentContext.hubStreamChecks === checks && writes.length === 12 && polls.length === (outcome === "mixed" ? 13 : 12) &&
                JSON.stringify(stableRows()) === JSON.stringify(selectedOrder),
                "Remount lost stable results/coverage or replayed parallel assignments/observations.");
            observations.push({ outcome, devices: 12, maxActive, preWriteReads: reads.length, operationStatusReads: polls.length,
                conditionalWrites: writes.length, admissions: [...admissions.values()], noSingleTargetOrOlderWaveEarlyRelease: true,
                wholeWaveSettlesEarly: true, stableSelectedOrder: true,
                postWriteMutableResourceReads: 0, continueGatedUntilWholeBatchSettled: true, backPreservesCoverage: true });
            await mount(<div>Parallel batch fixture closed.</div>);
          }
          caseObservations = { observations, actualCloudWrites: 0 };
        } finally {
            await mount(<div>Parallel batch verification closed.</div>); clock.restore();
            Math.random = originalRandom;
            assignmentResponse = undefined;
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.discoverHubs = hubs;
            HubAssignmentApi.prototype.recheck = recheck;
        }
    })],
    ["Assignment untracked dispatch stays unknown without recovery loops after settlement and remount", () => observeWizardContext(async () => {
        const clock = assignmentClock(), observations: object[] = [];
        const originalNow = Date.now, assignmentServices = HubAssignmentApi.prototype.assignmentServices;
        try {
            for (const outcome of ["500", "503", "network", "missing", "unsafe", "lost"] as const) {
                const guidance = outcome === "missing" ? "The accepted assignment did not expose operation tracking information. Completion is unknown; no assignment will be retried."
                    : outcome === "unsafe" ? "The assignment response contained an unsafe or unsupported operation address."
                    : outcome === "lost" ? "The original assignment operation tracking information is unavailable. Completion is unknown; no assignment will be retried."
                    : "The assignment request may have reached Azure, but its result couldn't be confirmed. Check the device later. No assignment will be retried.";
                const requests: string[] = [], state = { patches: [] as { body: unknown; ifMatch: string | null }[] };
                const respond = adrOnlyResponse(requests, state);
                let writes = 0, afterWriteRequests = 0, dispatched = false;
                HubAssignmentApi.prototype.assignmentServices = function (source) {
                    const services = assignmentServices.call(this, source);
                    return outcome === "lost" ? { ...services, trackingId: () => "retained-handle-no-longer-in-memory" } : services;
                };
                assignmentResponse = async (input, init) => {
                    if (init?.method === "PATCH") {
                        assert(new Headers(init.headers).get("If-Match") === '"fixture-current"' &&
                            JSON.stringify(JSON.parse(String(init.body))) === JSON.stringify({ properties: { deviceTypeRefs: [tm.id] } }),
                            "An uncertain result lost the current conditional refs-only mutation.");
                        writes++; dispatched = true;
                        if (outcome === "network") throw new Error("Synthetic private network detail.");
                        if (outcome === "missing") return new Response(null, { status: 202 });
                        if (outcome === "unsafe") return new Response(null, { status: 202,
                            headers: { Location: "https://unsafe.example.invalid/operation?opaque=fixture" } });
                        if (outcome === "lost") return new Response(null, { status: 202, headers: { Location:
                            `https://management.azure.com/subscriptions/${new URL(String(input)).pathname.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/lost?api-version=2026-11-02-preview` } });
                        return new Response("Synthetic private service body.", { status: Number(outcome) });
                    }
                    if (dispatched) afterWriteRequests++;
                    return respond(input, init);
                };
                await wizard(context(true)); await click(next()); await beginAssignmentTargets();
                await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') &&
                    !area.querySelector('[role="progressbar"]'), "The current uncertain-dispatch candidate did not load.");
                await click(area.querySelector('input[aria-label="Select Fixture device"]'));
                await click(button("Assign to selected devices"));
                await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')]
                    .find(button => button.textContent === "Assign"));
                await until(() => writes === 1 && (outcome === "lost" ? clock.pending.size === 1 : !currentContext.hubAssignmentBusy),
                    "The uncertain dispatch did not reach its observation boundary.");
                if (clock.pending.size) await clock.tick();
                await until(() => area.querySelector('[aria-label="Device assignment progress"]')?.getAttribute("aria-busy") === "false" &&
                    !next()!.disabled, "Settled observation uncertainty did not release Continue.");
                const assertUnknown = () => {
                    const row = area.querySelector<HTMLElement>('[data-assignment-state="unknown"]')!;
                    assert(row?.textContent?.includes(guidance) && row.querySelector('svg[aria-label="Assignment not confirmed"]') &&
                        area.querySelector('[aria-label="Device assignment progress"] h3')?.textContent === "Assigned devices" &&
                        !row.querySelector('[role="progressbar"]') && !area.querySelector('[data-assignment-state="failed"]') &&
                        !button("Refresh results") && !button("Sign in") &&
                        !area.querySelector('[aria-label="DPS enrollment guidance"]') &&
                        !area.textContent?.includes("private") && !area.textContent?.includes("30 seconds"),
                        "A dispatched request without operation tracking became a failed assignment or a pointless recovery loop.");
                    assertAssignmentSummary("incomplete", 1);
                    assertAssignmentSemanticIndicators();
                };
                assertUnknown();
                assert(writes === 1 && afterWriteRequests === 0 && clock.pending.size === 0,
                    "Untracked uncertainty triggered extra status/device reads or a mutation replay.");
                const history = JSON.stringify(currentContext.hubAssignmentHistory);
                if (outcome === "lost") await act(async () => observedWizardProps!.updateContext("hubAssignmentHistory",
                    currentContext.hubAssignmentHistory.map((receipt: AssignmentReceipt) => ({ ...receipt, state: "failed" }))));
                Date.now = () => (currentContext.hubAssignmentHistory[0].readbackDeadline ??
                    currentContext.hubAssignmentHistory[0].preparationDeadline) + 60000;
                await click(next()); await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
                await until(() => !next()!.disabled, "Retained untracked uncertainty did not keep optional Continue eligibility.");
                assertUnknown();
                assert(JSON.stringify(currentContext.hubAssignmentHistory) === history && writes === 1 &&
                    afterWriteRequests === 0 && clock.pending.size === 0,
                    "Remount changed terminal uncertainty, restarted observation or replayed a write.");
                await act(async () => observedWizardProps!.updateContext("hubModelRevision", (currentContext.hubModelRevision || 0) + 1));
                await until(() => !currentContext.hubAssignmentIntent &&
                    !area.querySelector('[aria-label="Device assignment progress"]'), "Owner change retained obsolete uncertain progress.");
                assert(JSON.stringify(currentContext.hubAssignmentHistory) === history && writes === 1 && afterWriteRequests === 0,
                    "Owner retirement erased a truthful uncertain receipt or retried its mutation.");
                observations.push({ outcome, conditionalWrites: writes, afterWriteRequests, terminalUnknown: true,
                    noRefreshLoop: true, stableReceiptOnBackAfterDeadline: true, ownerRetirementKeepsHistory: true });
                await mount(<div>Untracked result closed.</div>);
                Date.now = originalNow;
            }
            caseObservations = { observations, actualCloudWrites: 0 };
        } finally {
            await mount(<div>Untracked verification closed.</div>); clock.restore();
            Date.now = originalNow; HubAssignmentApi.prototype.assignmentServices = assignmentServices;
            assignmentResponse = undefined;
        }
    })],
    ["Assignment confirmation preserves dynamic model identity device provenance singular plural and explicit-only writes", async () => {
        caseObservations = await assignmentConfirmationCheck();
    }],
    ["Assignment timeout keeps customer guidance after batch settlement and remount without extra requests", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover, hubs = HubAssignmentApi.prototype.discoverHubs;
        const recheck = HubAssignmentApi.prototype.recheck, clock = assignmentClock();
        const originalNow = Date.now, originalSet = window.setTimeout, originalClear = window.clearTimeout;
        const budgetTimers = new Map<number, () => void>();
        let sequence = -10000, time = originalNow(), requests = 0, polls = 0, aborted = false;
        const inventory = assignmentInventory();
        inventory.targets[0] = { ...inventory.targets[0], updatePermission: "granted",
            writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
            deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: "timeout-device" } };
        const customerMessage = "The assignment wasn't confirmed within 30 seconds. It may still complete in Azure, so check the device later. No assignment will be retried.";
        Date.now = () => time;
        window.setTimeout = ((callback: TimerHandler, delay?: number, ...args: unknown[]) => {
            if (typeof callback === "function" && delay !== undefined && delay > 11000 && delay <= 30000) {
                const id = sequence--; budgetTimers.set(id, () => callback(...args)); return id;
            }
            return originalSet(callback, delay, ...args);
        }) as typeof window.setTimeout;
        window.clearTimeout = id => {
            if (budgetTimers.has(id)) budgetTimers.delete(id); else originalClear(id);
        };
        try {
            HubAssignmentApi.prototype.discoverHubs = async source => ({
                sourceSignature: canonicalJson(source), hubs: inventory.hubs, issues: [], complete: true,
            });
            HubAssignmentApi.prototype.discover = async () => inventory;
            HubAssignmentApi.prototype.recheck = async target => target;
            assignmentResponse = async (input, init) => {
                if (init?.method === "PATCH") {
                    requests++;
                    return new Response(null, { status: 202, headers: { Location:
                        `https://management.azure.com/subscriptions/${inventory.targets[0].id.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/timeout?api-version=2026-11-02-preview` } });
                }
                assert(String(input).includes("/operationStatuses/timeout"), "Timeout verification queried mutable device state.");
                polls++;
                init?.signal?.addEventListener("abort", () => { aborted = true; }, { once: true });
                return new Promise(() => undefined);
            };
            await wizard(context(true)); await click(next()); await beginAssignmentTargets();
            await until(() => !!button("Select all") && !button("Select all")!.disabled, "Timeout fixture did not load.");
            await click(button("Select all")); await click(button("Assign to selected devices"));
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')]
                .find(button => button.textContent === "Assign"));
            await until(() => requests === 1 && clock.pending.size === 1, "The original mutation did not start operation observation.");
            assert(next()!.disabled && currentContext.hubAssignmentBusy, "An active observation did not gate Continue.");
            time += 3000; await clock.tick();
            await until(() => polls === 1 && budgetTimers.size > 0, "The native status request did not reach its hung boundary.");
            time += 27000;
            await act(async () => {
                const callbacks = [...budgetTimers.values()]; budgetTimers.clear(); callbacks.forEach(callback => callback());
            });
            await until(() => !currentContext.hubAssignmentBusy && !next()!.disabled &&
                area.querySelector('[aria-label="Device assignment progress"]')?.getAttribute("aria-busy") === "false",
                "Observation timeout did not release Continue after the batch settled.");
            const assertMessage = () => {
                const row = area.querySelector<HTMLElement>('[data-assignment-state="unknown"]')!;
                assert(row && row.textContent?.includes(customerMessage) && row.querySelector('svg[aria-label="Assignment not confirmed"]') &&
                    !row.querySelector('[role="progressbar"]') && !area.textContent?.includes("tracking budget") &&
                    !area.textContent?.includes("operation failed") && !area.textContent?.includes("was canceled"),
                    "Elapsed observation did not keep clear customer guidance distinct from remote failure/cancellation.");
                assertAssignmentSummary("incomplete", 1);
                assertAssignmentSemanticIndicators();
            };
            assertMessage();
            assert(aborted && requests === 1 && polls === 1 && clock.pending.size === 0 && !button("Refresh results"),
                "A settled timeout retained a request/timer, offered an implicit new budget or replayed a mutation.");
            const history = JSON.stringify(currentContext.hubAssignmentHistory);
            await click(next()); await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await until(() => !next()!.disabled, "Returning to the completed timeout state did not preserve optional continuation.");
            assertMessage();
            assert(JSON.stringify(currentContext.hubAssignmentHistory) === history && requests === 1 && polls === 1 &&
                clock.pending.size === 0 && budgetTimers.size === 0,
                "Remount rewrote timeout guidance or resumed an expired operation.");
            await act(async () => observedWizardProps!.updateContext("hubAssignmentHistory",
                currentContext.hubAssignmentHistory.map((receipt: object) => ({ ...receipt,
                    message: "The assignment operation tracking budget is missing or expired. Completion is unknown; no assignment will be retried." }))));
            await click(next()); await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await until(() => !next()!.disabled && area.textContent?.includes(customerMessage),
                () => `A retained historical timeout receipt did not normalize its display without resuming: ${JSON.stringify({
                    heading: area.querySelector("h2")?.textContent, history: currentContext.hubAssignmentHistory,
                    progress: area.querySelector('[aria-label="Device assignment progress"]')?.textContent,
                })}`);
            assertMessage();
            assert(requests === 1 && polls === 1 && budgetTimers.size === 0 && clock.pending.size === 0,
                "Historical timeout-copy normalization reactivated expired tracking.");
            caseObservations = { exactRenderedTimeoutGuidance: customerMessage, observationSeconds: 30,
                locallyAbortedStatusWait: true, terminalUnknownNotRemoteFailure: true, continueReleased: true,
                sameMessageAndReceiptOnBack: true, legacyTimeoutCopyNormalizedWithoutResume: true,
                conditionalWrites: requests, statusRequests: polls, noReplayOrResume: true,
                actualCloudWrites: 0 };
        } finally {
            await mount(<div>Timeout fixture closed.</div>);
            Date.now = originalNow; window.setTimeout = originalSet; window.clearTimeout = originalClear; clock.restore();
            budgetTimers.clear(); assignmentResponse = undefined;
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.discoverHubs = hubs;
            HubAssignmentApi.prototype.recheck = recheck;
        }
    })],
    ["Assignment unknown outcomes use native amber warning triangles while failures stay red and confirmations green in both themes", async () => {
        const themes = [];
        for (const dark of [false, true]) themes.push(await assignmentWarningPresentation(dark));
        caseObservations = { themes, nativeStateAndNonretryableUnknownPreserved: true, actualCloudWrites: 0 };
    }],
    ["Assignment Continue waits for the full batch preflight writes and readback then permits settled outcomes", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover, discoverHubs = HubAssignmentApi.prototype.discoverHubs;
        const services = HubAssignmentApi.prototype.assignmentServices, readback = HubAssignmentApi.prototype.readAssignment;
        const clock = assignmentClock(), outcomes: object[] = [];
        const originalNow = Date.now;
        try {
            for (const outcome of ["confirmed", "failed", "timeout"] as const) {
                const inventory = assignmentInventory(), writes: string[] = [], readbacks: string[] = [];
                inventory.targets = [0, 1, 2, 3, 4].map(index => ({ ...inventory.targets[0],
                    id: inventory.targets[0].id + `-${index}`, name: `Batch device ${index + 1}`,
                    writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
                    deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: `batch-${index}` } }));
                let releasePreflight: (() => void) | undefined, releaseWrite: (() => void) | undefined;
                let releaseReadback: (() => void) | undefined;
                HubAssignmentApi.prototype.discoverHubs = async source => ({
                    sourceSignature: canonicalJson(source), hubs: inventory.hubs, issues: [], complete: true,
                });
                HubAssignmentApi.prototype.discover = async () => inventory;
                HubAssignmentApi.prototype.assignmentServices = () => ({
                    read: async target => {
                        if (target.id.endsWith("-0")) await new Promise<void>(resolve => { releasePreflight = resolve; });
                        return target;
                    },
                    write: async (target, _id, _signal, dispatch) => {
                        dispatch(); writes.push(target.id);
                        if (target.id.endsWith("-4")) await new Promise<void>(resolve => { releaseWrite = resolve; });
                        return "accepted";
                    },
                });
                HubAssignmentApi.prototype.readAssignment = async target => {
                    readbacks.push(target.id);
                    if (target.id.endsWith("-4")) {
                        await new Promise<void>(resolve => { releaseReadback = resolve; });
                        if (outcome === "failed") throw new AssignmentApiError("permission", "Owned batch readback denied.");
                        if (outcome === "timeout") return { confirmed: false };
                    }
                    return { target: { ...target, reference: { kind: "assigned", ids: [tm.id] } }, confirmed: true };
                };
                await wizard(context(true)); await click(next()); await beginAssignmentTargets();
                await until(() => !!button("Select all") && !button("Select all")!.disabled, "Batch devices did not load.");
                assert(!next()!.disabled, "Optional assignment blocked continuation before an operation began.");
                await click(button("Select all")); await click(button("Assign to selected devices"));
                await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')]
                    .find(button => button.textContent === "Assign"));
                const assertBlocked = async (phase: string) => {
                    assert(next()!.disabled, `${phase} left Continue enabled while the batch was still active.`);
                    await act(async () => { await observedWizardProps!.onNext(); });
                    assert(step("Assign model")?.getAttribute("aria-current") === "step",
                        `${phase} allowed guarded advancement out of the active batch.`);
                };
                await until(() => !!releasePreflight, "Batch preflight did not start.");
                await assertBlocked("Preflight");
                assert(!writes.includes(inventory.targets[0].id),
                    "The pending device preflight sent its own assignment before its current read settled.");
                await act(async () => releasePreflight!());
                await until(() => !!releaseWrite, "The final batch write did not enter its deferred boundary.");
                await assertBlocked("Native write");
                await act(async () => releaseWrite!());
                await until(() => !currentContext.hubAssignmentBusy && clock.pending.size === 1,
                    "Accepted batch did not enter scheduled read-only verification.");
                await assertBlocked("Scheduled readback");
                await clock.tick();
                await until(() => !!releaseReadback && area.querySelectorAll('[data-assignment-state="confirmed"]').length === 4,
                    "The batch did not reach four confirmed devices and one active readback.");
                await assertBlocked("Four confirmed and one pending");
                assert(writes.length === 5 && readbacks.length === 5, "Batch verification repeated an assignment or read.");
                if (outcome === "timeout") Date.now = () => originalNow() + 31000;
                await act(async () => releaseReadback!());
                await until(() => !next()!.disabled && clock.pending.size === 0, "Settled batch did not restore existing Continue eligibility.");
                Date.now = originalNow;
                if (outcome === "confirmed") assertCompletedAssignments(inventory.targets.map(target => target.name));
                else assert(area.textContent?.includes(outcome === "timeout" ? ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE : "Owned batch readback denied.") &&
                    !completedAssignmentList() && !area.querySelector('[aria-label="DPS enrollment guidance"]') &&
                    area.querySelectorAll('[data-assignment-state="confirmed"]').length === 4,
                    "A terminal failure was hidden, fabricated success or discarded confirmed devices.");
                if (outcome === "timeout") assert(currentContext.hubAssignmentHistory.some((receipt: { state: string; message?: string }) =>
                    receipt.state === "unknown" && receipt.message === ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE),
                    "Observation expiry was mislabeled as remote assignment failure or cancellation.");
                await click(next());
                assert(step("Check model coverage")?.getAttribute("aria-current") === "step",
                    "Settled optional assignment did not permit normal continuation.");
                await verifyCoverage();
                const checks = currentContext.hubStreamChecks;
                await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
                await until(() => !next()!.disabled, "Settled assignment did not remain eligible on Back.");
                assert(currentContext.hubStreamChecks === checks && writes.length === 5 && readbacks.length === 5,
                    "Remounting a settled batch invalidated coverage or replayed assignment/readback.");
                await click(next());
                outcomes.push({ outcome, blockedPhases: ["preflight", "write", "scheduled readback", "four confirmed one pending"],
                    writes: writes.length, readbacks: readbacks.length, settledContinue: true, settledBackPreservesCoverage: true });
                await mount(<div>Batch fixture closed.</div>);
            }
            caseObservations = { outcomes, actualCloudWrites: 0 };
        } finally {
            Date.now = originalNow;
            await mount(<div>Batch verification closed.</div>); clock.restore();
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.discoverHubs = discoverHubs;
            HubAssignmentApi.prototype.assignmentServices = services; HubAssignmentApi.prototype.readAssignment = readback;
        }
    })],
    ["Assignment Intent live source ownership rejects stale actions and late inventory without restoring an ABA choice", async () => {
        const discover = HubAssignmentApi.prototype.discover;
        let release: ((value: AssignmentInventory) => void) | undefined, edit: WizardStepProps["updateContext"] | undefined;
        const Assignment = (props: WizardStepProps & { sourceState?: SourceReadiness; getSourceState?: () => SourceReadiness }) => {
            edit = props.updateContext;
            return <HubModelAssignmentStep {...props} workloadClient={client}/>;
        };
        HubAssignmentApi.prototype.discover = () => new Promise(resolve => { release = resolve; });
        try {
            await mount(<Harness component={Assignment} initial={assignmentContext()}/>);
            await beginAssignmentTargets(); await until(() => !!release, "Owned discovery did not start.");
            await act(async () => setTestSourceState({ ...sourceFixture, key: "another-source", revision: 2 }));
            assert(!currentContext.hubAssignmentIntent && !area.querySelector('[aria-label="Device model assignments"]'),
                "Changed source kept old active intent/inventory.");
            await act(async () => release!(assignmentInventory())); await settle();
            assert(!currentContext.hubAssignmentInventory, "A late owner inventory was published after intent retirement.");
            await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
            assert(!currentContext.hubAssignmentIntent && button("Assign model to devices")?.getAttribute("aria-pressed") !== "true",
                "Source ABA silently restored old assign choice.");
            await click(button("Skip assignment"));
            const owner = currentContext.hubAssignmentIntent.owner;
            await act(async () => edit!("hubModelRevision", 2));
            assert(!currentContext.hubAssignmentIntent && owner !== assignmentIntentOwner(currentContext, { ...sourceFixture, revision: 3 }),
                "Model revision reused old current intent.");
            caseObservations = { currentOwnerRequired: true, sourceAbaRetired: true, lateReadIgnored: true, modelRevisionRetired: true };
        } finally { HubAssignmentApi.prototype.discover = discover; }
    }],
    ["Modeling navigation omits optional Cancel while preserving Previous Continue Finish and shared Cancel support", async () => {
        const original = React.createElement;
        let navigation: React.ComponentProps<typeof WizardControl> | undefined;
        const observer = new Proxy(original, {
            apply(target, thisArg, args) {
                if (args[0] === WizardControl && args[1]?.title === "Model your device data") navigation = args[1];
                return Reflect.apply(target, thisArg, args);
            },
        });
        let back = 0, finished = 0, cancelled = 0;
        Reflect.set(React, "createElement", observer);
        try {
            await mount(<IoTSolutionItemModelWizard workloadClient={client}
                item={{ id: "navigation", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                initialContext={{}} onBack={() => { back++; }}/>);
            assert(navigation && navigation.onCancel === undefined && navigation.onFirstStepBack &&
                navigation.navigationLabels?.previous === "Previous" && navigation.navigationLabels?.next === "Continue" &&
                navigation.navigationLabels?.complete === "Finish and Close" &&
                !Object.prototype.hasOwnProperty.call(navigation.navigationLabels, "cancel") &&
                !area.querySelector('[data-testid="wizard-cancel-btn"]'), "Modeling still supplies or displays global Cancel.");
            const previous = () => area.querySelector<HTMLButtonElement>('[data-testid="wizard-previous-btn"]')!;
            assert(previous().textContent === "Previous" && !previous().disabled && next()?.textContent === "Continue",
                "Normal modeling navigation labels or first-step back changed.");
            await click(next());
            assert(area.querySelector("h2")?.textContent === "Select your data source" &&
                !area.querySelector('[data-testid="wizard-cancel-btn"]'), "Continue lost normal source navigation or introduced Cancel.");
            await click(previous());
            assert(area.querySelector("h2")?.textContent?.trim() === "Turn raw IoT data into modeled data",
                "Previous no longer returns to the modeling overview.");
            await click(previous());
            assert(back === 1 && navigation.canFinish?.("configured", {}) === false,
                "First-step back or verified Finish guard changed.");
        } finally { Reflect.set(React, "createElement", original); }
        const Empty = () => <div>Navigation fixture</div>;
        await mount(<WizardControl title="Shared optional cancellation" navigationLabels={navigation!.navigationLabels}
            steps={[{ id: "first", title: "First", component: Empty, validate: () => true },
                { id: "last", title: "Last", component: Empty, validate: () => true }]}
            onCancel={() => { cancelled++; }} onComplete={() => { finished++; }}/>);
        const cancel = area.querySelector('[data-testid="wizard-cancel-btn"]');
        assert(cancel?.textContent === "Cancel", "Shared optional Cancel support was removed.");
        await click(cancel);
        assert(cancelled === 1, "Shared optional Cancel callback no longer runs.");
        await click(next());
        const finish = area.querySelector<HTMLButtonElement>('[data-testid="wizard-complete-btn"]')!;
        assert(finish.textContent === "Finish and Close" && !finish.disabled,
            "The modeling completion label or shared finish navigation changed.");
        await click(finish);
        assert(finished === 1, "Shared completion callback no longer runs.");
        caseObservations = { modelingCancelAbsent: true, previousContinuePreserved: true,
            firstStepBackCalls: back, verifiedFinishGuardRetained: true, finishLabel: "Finish and Close",
            sharedCancelCalls: cancelled, sharedFinishCalls: finished, actualCloudWrites: 0 };
    }],
    ["Upload tiles share spacing and guarded click keyboard drop behavior in every state", async () => {
        caseObservations = await uploadSpacingCheck();
    }],
    ["Calibration rows use quiet two-level typography and preserve exact model type and unit text", async () => {
        caseObservations = await calibrationRowLayoutCheck();
    }],
    ["Upload click tiles preserve chooser cancellation drop and source stale guards", () => observeWizardContext(async () => {
        const seed = context(true); delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
        await wizard(seed);
        const reads = analysisQueries().length, cache = currentContext.hubSourceAnalysis;
        await click(area.querySelector('input[value="wot"]'));
        const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
        let chooserRequests = 0;
        const observeChooser = (event: Event) => { event.preventDefault(); chooserRequests++; };
        input.addEventListener("click", observeChooser);
        try {
            assert(modelUploadAction("Choose File") && !acceptedModelTile(), "Initial chooser changed.");
            await file("thing-model-upload", JSON.stringify(tm), "thermoProX200.json");
            const tile = acceptedModelTile()!, model = currentContext.hubModel, fields = currentContext.hubFields;
            assert(tile?.classList.contains("iot-solution-drop-zone--loaded") && tile.getAttribute("role") === "button" && tile.tabIndex === 0 &&
                tile.querySelector("#model-upload-file-name")?.textContent === "thermoProX200.json" &&
                tile.querySelector("#model-upload-replace-help")?.textContent === "Click or drop a new file to replace" &&
                tile.querySelector("svg") && !tile.querySelector("button") &&
                !tile.textContent?.includes("fields available") && getComputedStyle(tile).cursor === "pointer",
                "Accepted tile is not one keyboard-accessible clickable replacement surface.");
            const group = area.querySelector('[aria-labelledby="model-fields-group-heading"]')!;
            assertFieldSelectionCopy();
            assert(!group.contains(area.querySelector("#model-fields-heading")) && !area.querySelector("#find-model-field") &&
                group.querySelector('[role="status"]')?.textContent?.trim() === "Select a role for every field to continue" &&
                currentContext.hubFields.length === 2 &&
                tile.querySelector("#model-upload-reference")?.textContent === tm.id &&
                !area.querySelector('[aria-label="Accepted model"]') && !summary("Uploaded Model (JSON)"),
                "External field heading/text, internal status group or inline supplied ID changed.");
            await click(tile.querySelector("#model-upload-file-name"));
            for (const key of ["Enter", " "]) {
                const event = new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true });
                await act(async () => tile.dispatchEvent(event));
            }
            assert(chooserRequests === 3 && currentContext.hubModel === model && currentContext.hubFields === fields,
                "Mouse/Enter/Space double-triggered the chooser or changed accepted data before selection.");
            await act(async () => { input.files = new DataTransfer().files; input.dispatchEvent(new Event("change", { bubbles: true })); });
            assert(currentContext.hubModel === model && input.value === "", "Chooser cancellation discarded the accepted model.");
            await file("thing-model-upload", JSON.stringify(tm), "thermoProX200.json");
            const transfer = new DataTransfer();
            transfer.items.add(new File([JSON.stringify({ ...tm, id: "urn:fixture:dropped-model" })], "replacement.json", { type: "application/json" }));
            await act(async () => acceptedModelTile()!.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: transfer })));
            await until(() => currentContext.hubModel?.id === "urn:fixture:dropped-model", "Drop did not use validated replacement.");
            assert(area.querySelector("#model-upload-file-name")?.textContent === "replacement.json" && chooserRequests === 3,
                "Drop opened another chooser or retained old filename.");
            await file("thing-model-upload", "{bad", "failed.json");
            assert(!acceptedModelTile() && !area.querySelector("#model-upload-reference") && !area.querySelector("#model-fields-heading") &&
                modelUploadAction("Replace Model") && area.querySelector("#model-upload-error") && next()!.disabled,
                "Failed replacement published a success tile, old ID/editor or navigation.");
            await file("thing-model-upload", JSON.stringify(dtdl), "wrong-format.json");
            assert(!acceptedModelTile() && area.querySelector<HTMLInputElement>('input[value="wot"]')!.checked,
                "Tile replacement bypassed selected-format validation.");
            await file("thing-model-upload", JSON.stringify(tm), "recovered.json");
            assert(acceptedModelTile() && !next()!.disabled && currentContext.hubSourceAnalysis === cache &&
                analysisQueries().length === reads, "Replacement recovery changed source analysis or added raw reads.");
            const longName = `${"accepted-model-file-".repeat(12)}.json`, longId = `https://models.example.test/${"canonical-identifier-".repeat(15)}v1`;
            await file("thing-model-upload", JSON.stringify({ ...tm, id: longId }), longName);
            assert(area.querySelector("#model-upload-file-name")?.textContent === longName &&
                area.querySelector("#model-upload-reference")?.textContent === longId &&
                acceptedModelTile()!.scrollWidth <= acceptedModelTile()!.clientWidth + 1,
                "Explicit replacement lost long filename/ID identity or wrapping.");
            caseObservations = { chooserRequests: 3, clickableTileOnly: true, sameFileAndCancel: true,
                validatedDrop: true, failedIdentityAndFieldsHidden: true, paragraphsOutsideGroup: true, extraSourceReads: 0 };
        } finally { input.removeEventListener("click", observeChooser); }
        await mount(<Harness component={ModelPreparationStep} initial={withAnalyzedSamples(context(true))}/>);
        const sourceModel = currentContext.hubModel;
        const guardedInput = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
        const originalFiles = Object.getOwnPropertyDescriptor(guardedInput, "files");
        const pending = new File(["{}"], "pending-source.json");
        let release: (text: string) => void = () => undefined;
        Object.defineProperty(pending, "text", { value: () => new Promise<string>(resolve => { release = resolve; }) });
        Object.defineProperty(guardedInput, "files", { configurable: true, value: [pending] });
        try {
            await act(async () => guardedInput.dispatchEvent(new Event("change", { bubbles: true })));
            assert(controlDisabled(modelUploadAction("Replace Model")) && !acceptedModelTile(), "Pending replacement remained interactive or accepted.");
            await act(async () => setTestSourceState({ ready: false, key: "changed-source", revision: 2, reason: "Validate the changed source." }));
            await act(async () => release(JSON.stringify({ ...tm, id: "urn:fixture:stale-must-not-publish" })));
            assert(currentContext.hubModel === sourceModel && controlDisabled(modelUploadAction("Replace Model")) && guardedInput.disabled,
                "A stale source read published a replacement or enabled the chooser.");
            const transfer = new DataTransfer();
            transfer.items.add(new File([JSON.stringify({ ...tm, id: "urn:fixture:blocked-drop" })], "blocked.json"));
            await act(async () => acceptedModelTile()!.dispatchEvent(new DragEvent("drop", { bubbles: true, cancelable: true, dataTransfer: transfer })));
            assert(currentContext.hubModel === sourceModel, "An unready source accepted a dropped replacement.");
            caseObservations = { ...caseObservations, sourceGuarded: true, staleSourceCompletionDiscarded: true };
        } finally {
            if (originalFiles) Object.defineProperty(guardedInput, "files", originalFiles);
            else Reflect.deleteProperty(guardedInput, "files");
        }
    })],
    ["Inline Model declared metadata uses canonical main and secondary facts and roundtrips selected native column docstrings", () => observeWizardContext(async () => {
        const seed = context(true); delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
        await wizard(seed); const cache = currentContext.hubSourceAnalysis, reads = analysisQueries().length;
        await click(area.querySelector('input[value="wot"]'));
        await file("thing-model-upload", JSON.stringify(metadataThingModel), "declared-metadata.tm.json");
        await until(() => area.querySelectorAll("[data-field-id]").length === 5, "Metadata model did not reach the inline editor.");
        const temperature = currentContext.hubFields.find((field: ModelField) => field.path[0] === "temperature");
        const row = [...area.querySelectorAll<HTMLElement>("[data-field-id]")].find(node => node.dataset.fieldId === temperature.id)!;
        const main = row.querySelector(".iot-solution-model-field-heading")!.textContent!.replace(/\s+/g, " ").trim();
        assert(main === `${temperature.title} (temperature)` &&
            row.querySelector(".iot-solution-model-field-description")?.textContent === temperature.description &&
            row.querySelector(".iot-solution-model-field-metadata")?.textContent === "number \u00b7 degC" &&
            !row.textContent?.includes("→") && currentContext.hubFields.filter((field: ModelField) =>
                ["message", "config", "alarm"].includes(field.path[0])).every((field: ModelField) => !field.role),
            "Field title/name/description or declared metadata changed, or access flags/affordance kind assigned a role.");
        const title = row.querySelector(".iot-solution-model-field-heading > span")!;
        const qualifier = row.querySelector(".iot-solution-field-qualifier")!;
        const metadata = row.querySelector(".iot-solution-model-field-metadata")!;
        const description = row.querySelector(".iot-solution-model-field-description")!;
        assert(getComputedStyle(title).fontWeight === "400" && getComputedStyle(title).fontSize === "14px" &&
            getComputedStyle(qualifier).fontSize === "14px" && getComputedStyle(metadata.firstElementChild!).fontSize === "12px" &&
            metadata.children.length === 1 && metadata.textContent === "number \u00b7 degC" &&
            Math.abs(title.getBoundingClientRect().left - description.getBoundingClientRect().left) < 1 &&
            row.contains(metadata) && !row.querySelector(".iot-solution-model-field-content")!.contains(metadata) &&
            !row.textContent?.includes("Model As") && !summary("Uploaded Model (JSON)") &&
            !area.textContent?.includes("Roles were suggested automatically from the analyzed data.") &&
            !area.textContent?.includes("Your explicit role choices are kept."),
            "V1 regular title/secondary typography, ordered compact metadata or aligned descriptions regressed.");
        const complex = [...area.querySelectorAll<HTMLElement>("[data-field-id]")].find(node => node.dataset.fieldId === "/properties/config")!;
        const complexDetails = complex.querySelector(".iot-solution-model-field-metadata")!.textContent!;
        assert(complexDetails === "object" &&
            !/(^|, )(properties|items|required|oneOf|const):/.test(complexDetails) &&
            currentContext.hubFields.find((field: ModelField) => field.path[0] === "config").declaredMetadata.schema.properties.levels.items.minimum === 0,
            "Complex schema shape was displayed or removed from the canonical stored field.");
        const unobserved = currentContext.hubFields.find((field: ModelField) => !field.role);
        assert(next()!.disabled && documentElementRole(unobserved.id)?.textContent === "Choose a Role" &&
            documentElementRole(unobserved.id)?.getAttribute("aria-invalid") === "true", "The unresolved placeholder passed combined validation.");
        await click(documentElementRole(unobserved.id));
        assert(activeOptions(documentElementRole(unobserved.id)!.id).map(option => option.textContent).join("|") === "Telemetry|Twin Property",
            "Choose a Role appeared as a selectable option.");
        await click(activeOptions(documentElementRole(unobserved.id)!.id)[0]);
        for (const field of currentContext.hubFields as ModelField[]) {
            if (!field.role) await enter(documentElementRole(field.id)!.id, "telemetry");
        }
        assert(!next()!.disabled && currentContext.hubSourceAnalysis === cache && analysisQueries().length === reads,
            "Metadata rendering/matching changed source evidence or combined validation.");
        await reachReview();
        const request = { input: hubInput(currentContext), coverage: currentContext.hubCoverage,
            allowWaiting: true, names: currentContext.hubOutputNames };
        creationBackend = createCreationMock(request);
        await click(button("Create modeled data")); await until(() => currentContext.hubCreation?.phase === "completed", "Metadata native readback did not verify.");
        const modeled = creationBackend.tables.get(request.names.modeled)!;
        for (const field of request.input.fields) {
            const decoded = readColumnSemantics(modeled.columnDocs[field.column])!;
            assert(decoded.version === 3 && modeled.columnDocs[field.column].length <= 1000 &&
                modeled.columnDocs[field.column] === creationBackend.plan.documentation.columns[field.column],
                "The exact bounded declared metadata did not survive selected-column quoting/write/readback.");
        }
        const message = request.input.fields.find(field => field.path[0] === "message")!;
        const full = readColumnSemantics(modeled.columnDocs[message.column])!;
        const config = request.input.fields.find(field => field.path[0] === "config")!;
        const configDoc = modeled.columnDocs[config.column];
        const configMetadata = readColumnSemantics(configDoc)!;
        assert(configMetadata.version === 3 && configDoc.includes('"minimum":0') &&
            configDoc.includes('"maxItems":8') && configDoc.includes('"required":["levels"]'),
            "The bounded outline lost declared nested facts or whole-value structure.");
        assert(full.version === 3 && modeled.doc.includes(metadataThingModel.id) &&
            creationBackend.writes.length === 9 && analysisQueries().length === reads,
            "Bounded metadata changed native processing or source-read ownership.");
        caseObservations = { parserUiDocReadback: true, exactMainLine: true, declaredFalseZeroEmptyNull: true,
            boundedNestedFactsAndAuthoringMetadataRetained: true, noAccessFlagRoleInference: true, extraSourceReads: 0,
            nativeFakeWrites: 9, actualCloudWrites: 0 };
    })],
    ["Inline Model prepares model-free evidence then gates inclusion roles names and guarded creation without a Fields step", () => observeWizardContext(async () => {
        const seed = context(true); delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
        sourceSampleRows = 1000;
        const armReads: string[] = [], armWrites: string[] = [], read = adrOnlyResponse(armReads);
        assignmentResponse = async (url, init) => {
            if (["PATCH", "PUT", "DELETE"].includes(init?.method || "")) { armWrites.push(String(url)); throw new Error("No assignment mutation in this test."); }
            return read(url, init);
        };
        await openAnalysisSource(seed); await until(() => !sourceValidationAction()!.disabled, "Source metadata did not settle.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Source validation did not finish.");
        const cache = currentContext.hubSourceAnalysis, profile = cache.roleEvidence, reads = analysisQueries().length;
        assert(!currentContext.hubModel && profile.bindings.some(binding => binding.path[0] === "temperature" && binding.telemetry === 1000) &&
            profile.bindings.some(binding => binding.path[0] === "state" && binding.twin === 1000) &&
            !currentContext.hubCoverage, "Evidence was not prepared before a model existed or was mistaken for coverage.");
        await click(next());
        assert(!area.querySelector("[data-field-id]") && !step("Fields") &&
            area.querySelector<HTMLInputElement>('input[value="wot"]')?.checked && !currentContext.hubModel && next()!.disabled,
            "The field editor appeared before accepted upload or a standalone Fields step remained.");
        await click(area.querySelector('input[value="wot"]')); await file("thing-model-upload", JSON.stringify(tm), "inline.tm.json");
        await until(() => area.querySelectorAll("[data-field-id]").length === 2 && !next()!.disabled, "Accepted upload did not show current inline roles.");
        const region = area.querySelector<HTMLElement>('[aria-labelledby="model-fields-group-heading"]')!;
        const modelId = area.querySelector<HTMLElement>("#model-upload-reference")!;
        assert(area.querySelector("h2")?.textContent === "Add a device model" &&
            area.querySelector("#model-fields-heading")?.textContent === "3. Select fields" &&
            modelId.textContent === tm.id && acceptedModelTile()!.contains(modelId) &&
            !!(modelId.compareDocumentPosition(region) & Node.DOCUMENT_POSITION_FOLLOWING) && !summary("Uploaded Model (JSON)") &&
            step("Assign model")?.querySelector(".wizard-control__step-number")?.textContent === "3" &&
            step("Check model coverage")?.querySelector(".wizard-control__step-number")?.textContent === "4" &&
            step("Create modeled data")?.querySelector(".wizard-control__step-number")?.textContent === "5" &&
            !region.textContent?.includes("→") && currentContext.hubFields.every((field: ModelField) => field.path.length === 1) &&
            currentContext.hubSourceAnalysis.roleEvidence === profile && analysisQueries().length === reads,
            "Inline location, canonical ID, numbering, stored whole fields or prepared-evidence reuse changed.");
        assert([...region.querySelectorAll(".iot-solution-model-field-card")].every(row =>
            getComputedStyle(row).backgroundColor !== getComputedStyle(region).backgroundColor &&
            getComputedStyle(row).borderTopStyle === "solid") &&
            currentContext.hubFields.length === 2 && currentContext.hubFields.filter((field: ModelField) => field.role === "telemetry").length === 1 &&
            currentContext.hubFields.filter((field: ModelField) => field.role === "twin").length === 1 &&
            region.querySelector('[role="status"]')?.textContent?.trim() === "Select a role for every field to continue" &&
            !region.textContent?.includes("Roles were suggested automatically"), "Native gray/white rows or live top status are missing.");
        await click(button("Deselect all")); assert(next()!.disabled, "Zero included fields passed the combined model gate.");
        await click(button("Select all"));
        const temperature = currentContext.hubFields.find((field: ModelField) => field.path[0] === "temperature");
        await click(documentElementRole(temperature.id));
        assert(activeOptions(documentElementRole(temperature.id)!.id).map(option => option.textContent).join("|") === "Telemetry|Twin Property",
            "The placeholder became a selectable reset instead of actual role choices.");
        await enter(documentElementRole(temperature.id)!.id, "telemetry");
        const longId = `https://example.test/models/${"supplied-model-path-".repeat(18)}version-1`;
        await file("thing-model-upload", JSON.stringify({ ...tm, id: longId }), "replacement.tm.json");
        await until(() => currentContext.hubModel?.id === longId && !next()!.disabled, "Replacement did not retain explicit role intent.");
        assert(currentContext.hubFields.find((field: ModelField) => field.id === temperature.id).roleSource === "manual" &&
            area.querySelector("#model-upload-reference")?.textContent === longId &&
            currentContext.hubSourceAnalysis === cache && analysisQueries().length === reads,
            "Replacement lost manual roles/cache or generated/shortened the supplied ID.");
        await reachScope();
        assert(next()!.disabled && !currentContext.hubCoverage && !step("Fields"), "Prepared unfiltered evidence bypassed processing checks.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(area.querySelector("h2")?.textContent === "Assign the model to devices", "WoT Back did not preserve assignment position.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(area.querySelector("#model-upload-reference")?.textContent === longId && !next()!.disabled, "Back lost accepted model/field intent.");
        await click(button("Deselect all"));
        assert(next()!.disabled && step("Assign model")?.getAttribute("aria-disabled") === "true" &&
            step("Check model coverage")?.getAttribute("aria-disabled") === "true",
            "Previously completed sidebar navigation bypassed the combined invalid field gate.");
        await click(button("Select all"));
        await reachReview();
        const request = { input: hubInput(currentContext), coverage: currentContext.hubCoverage,
            allowWaiting: true, names: currentContext.hubOutputNames };
        creationBackend = createCreationMock(request);
        await click(button("Create modeled data")); await until(() => currentContext.hubCreation?.phase === "completed", "Guarded fake Create did not complete.");
        assert(creationBackend.writes.length === 9 && armWrites.length === 0 && analysisQueries().length === reads,
            "New navigation replayed assignment, changed native output creation or reread source roles.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Your modeled data is ready", "Current verified creation did not reach Ready.");
        caseObservations = { profileBeforeModel: true, sourceRows: [1000, 1000], noStandaloneFields: true,
            numbering: { source: 1, model: 2, assignment: 3, filters: 4, create: 5 }, inlineCombinedGate: true,
            modelIdVerbatim: true, grayGroupWhiteRows: true, noKustoTypeMapping: true, manualReplacementAndBack: true,
            extraRoleReads: 0, assignmentWrites: 0, guardedFakeWrites: 7 };
    })],
    ["Inline Model DTDL inherited component complex fields keep declared roles and canonical ID through assignment guidance", () => observeWizardContext(async () => {
        const seed = context(true); delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
        await wizard(seed);
        const cache = currentContext.hubSourceAnalysis, reads = analysisQueries().length;
        await click(area.querySelector('input[value="dtdl"]')); await file("thing-model-upload", JSON.stringify(mixedDtdlDocument), "complex-dtdl.json");
        const expected = mixedDtdlRequest().input;
        await until(() => currentContext.hubModel?.kind === "dtdl" && !next()!.disabled, "Inline DTDL fields did not validate.");
        assert(area.querySelectorAll("[data-field-id]").length === expected.fields.length &&
            currentContext.hubFields.every((field: ModelField) => field.roleSource === "model" && field.path.length === 1) &&
            !area.querySelector('[role="combobox"][id^="field-role-"]') &&
            currentContext.hubFields.filter((field: ModelField) => field.kustoType === "dynamic").length ===
                expected.fields.filter(field => field.kustoType === "dynamic").length &&
            area.querySelector("#model-upload-reference")?.textContent === expected.model.id &&
            !area.textContent?.includes("DTDL roles are defined by the model.") && !step("Fields"),
            "DTDL IDs/whole types/components/declared roles changed in the shared editor.");
        for (const field of currentContext.hubFields as ModelField[]) {
            if (field.kustoType !== "dynamic") continue;
            const row = [...area.querySelectorAll<HTMLElement>("[data-field-id]")].find(node => node.dataset.fieldId === field.id)!;
            const details = row.querySelector(".iot-solution-model-field-metadata")!.textContent!;
            assert((details.includes("Object") || details.includes("Array")) &&
                !details.includes('"fields"') && !details.includes('"elementSchema"') &&
                JSON.stringify(field.declaredMetadata).includes(field.path[0] === "samples" || field.path[0] === "history" ? "elementSchema" : "fields"),
                "DTD L complex metadata was displayed as a nested shape or altered instead of a concise declared type.");
        }
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Assign the model to devices" && !next()!.disabled &&
            !!area.querySelector("[data-dtdl-assignment-guidance]") && !button("Assign model to devices"),
            "DTDL did not expose its nonmutating assignment guidance and Continue action.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage", "DTDL Continue did not reach filters.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(area.querySelector("h2")?.textContent === "Add a device model" && currentContext.hubSourceAnalysis === cache &&
            analysisQueries().length === reads, "DTD L Back lost fields/cache or reread roles.");
        const noId = { "@type": "Interface", contents: dtdl.contents };
        await file("thing-model-upload", JSON.stringify(noId), "supplied-no-id.json");
        assert(currentContext.hubModel?.id === "" &&
            area.querySelector("#model-upload-reference")?.textContent === "Not provided in the uploaded model." &&
            !next()!.disabled, "A missing parsed DTDL ID was fabricated or became a new parser restriction.");
        const related = localPositiveFixtures[0], original = JSON.stringify(related.document);
        await click(area.querySelector('input[value="wot"]')); await file("thing-model-upload", original, "related-whole.tm.json");
        assert(currentContext.hubModel?.id === related.rootId &&
            area.querySelector("#model-upload-reference")?.textContent === related.rootId &&
            JSON.stringify(currentContext.hubModel.document) === original &&
            area.querySelectorAll("[data-field-id]").length === 3 &&
            currentContext.hubFields.every((field: ModelField) => field.path.length === 1) &&
            currentContext.hubSourceAnalysis === cache && analysisQueries().length === reads,
            "The helper-first related upload displayed the wrong model ID, flattened values or reread source evidence.");
        caseObservations = { declaredDtdlRoles: true, inheritedComponentWholeFields: expected.fields.length,
            dynamicTypesStoredNotPromoted: true, canonicalIdFromParser: true, missingIdHonest: true,
            assignmentExplicitlySkipped: true, extraRoleReads: 0 };
    })],
    ["Inline Model ambiguous and empty evidence require manual roles while format replacement and stale files cannot restore fields", () => observeWizardContext(async () => {
        for (const mode of ["empty", "ambiguous"]) {
            emptySamples = mode === "empty" ? "both" : "none";
            parityPayloads = mode === "ambiguous" ? { telemetry: { temperature: 22, state: "on" },
                twin: { properties: { reported: { temperature: 23, state: "off" } } } } : undefined;
            const seed = context(true); delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
            await wizard(seed); const cache = currentContext.hubSourceAnalysis, reads = analysisQueries().length;
            await click(area.querySelector('input[value="wot"]')); await file("thing-model-upload", JSON.stringify(tm), `${mode}.tm.json`);
            assert(area.querySelectorAll("[data-field-id]").length === 2 && next()!.disabled &&
                currentContext.hubFields.every((field: ModelField) => !field.role) &&
                area.textContent?.includes(mode === "empty" ? "No sample records were found." : "no clear role suggestions"),
                "Unfiltered empty/both-stream evidence fabricated roles or passed combined validation.");
            for (const field of currentContext.hubFields as ModelField[]) await enter(documentElementRole(field.id)!.id,
                field.path[0] === "state" ? "twin" : "telemetry");
            assert(!next()!.disabled, "Explicit manual roles could not resolve the combined model gate.");
            await click(area.querySelector('input[value="dtdl"]'));
            assert(!area.querySelector("[data-field-id]") && !area.querySelector("#model-upload-reference") && next()!.disabled &&
                currentContext.hubSourceAnalysis === cache, "Format change retained accepted ID/fields or cleared prepared source evidence.");
            const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
            const old = Object.getOwnPropertyDescriptor(input, "files");
            let release: (text: string) => void = () => undefined;
            const candidate = new File(["{}"], "late-dtdl.json");
            Object.defineProperty(candidate, "text", { value: () => new Promise<string>(resolve => { release = resolve; }) });
            Object.defineProperty(input, "files", { configurable: true, value: [candidate] });
            try {
                await act(async () => input.dispatchEvent(new Event("change", { bubbles: true })));
                await click(area.querySelector('input[value="wot"]'));
                await act(async () => release(JSON.stringify(dtdl))); await settle();
            } finally { if (old) Object.defineProperty(input, "files", old); else Reflect.deleteProperty(input, "files"); }
            assert(!currentContext.hubModel && !area.querySelector("[data-field-id]") && !area.querySelector("#model-upload-reference") &&
                next()!.disabled && !area.querySelector('[role="alert"]'), "Stale file completion published an ID/editor into the wrong format.");
            await file("thing-model-upload", JSON.stringify(tm), "current.tm.json");
            for (const field of currentContext.hubFields as ModelField[]) await enter(documentElementRole(field.id)!.id,
                field.path[0] === "state" ? "twin" : "telemetry");
            assert(!next()!.disabled && currentContext.hubSourceAnalysis === cache && analysisQueries().length === reads,
                "Explicit current replacement could not reuse prepared evidence.");
            if (mode === "empty") {
                emptySamples = "none"; emptyProcessing = true;
                await reachScope();
                await verifyCoverage();
                await until(() => !next()!.disabled, "Verified zero still required a removed waiting opt-in.");
                await click(next());
                const request = { input: hubInput(currentContext), coverage: currentContext.hubCoverage,
                    allowWaiting: true, names: currentContext.hubOutputNames };
                creationBackend = createCreationMock(request);
                await click(button("Create modeled data")); await until(() => currentContext.hubCreation?.phase === "completed", "Guarded fake waiting Create did not finish.");
                assert(currentContext.hubCreation.awaitingData && creationBackend.writes.length === 9, "Empty source profile became processing adoption.");
                creationBackend = undefined; emptyProcessing = false;
            }
        }
        caseObservations = { honestEmptyAndAmbiguous: true, manualRoleResolution: true, combinedGate: true,
            formatRetiresIdAndEditor: true, staleFileCannotPublish: true, cacheRetainedWithoutReads: true,
            independentScopedWaitingAndGuardedFakeCreate: true };
    })],
    ["Source Validation mirrors Hub action pending success retry and failed revalidation without stale success", () => observeWizardContext(async () => {
        sourceSampleRows = 1000;
        await openAnalysisSource();
        await until(() => !!button("Validate source tables") && !button("Validate source tables")!.disabled, "Initial validation action did not become ready.");
        assert(!button("Validate again") && !area.querySelector(".iot-solution-success") && next()!.disabled,
            "Fresh source showed a prior validation or enabled Continue.");
        deferSourceSamples = true;
        await click(button("Validate source tables"));
        await until(() => sourceSampleReleases.length === 2, "Validation did not reach both bounded source reads.");
        const loading = area.querySelector<HTMLElement>(".iot-solution-loading")!;
        const progress = loading.querySelector<HTMLElement>('[role="progressbar"]')!;
        const label = loading.querySelector<HTMLElement>(".iot-solution-loading-text")!;
        assert(!sourceValidationAction() && label.textContent === "Validating source tables..." &&
            loading.getAttribute("role") === "status" && progress.getAttribute("aria-labelledby") === label.id &&
            getComputedStyle(loading).flexDirection === "column" &&
            label.getBoundingClientRect().top >= progress.getBoundingClientRect().bottom &&
            !area.querySelector(".iot-solution-success") && next()!.disabled,
            "Pending validation retained the action/success or lost the Hub's centered, below-spinner loading.");
        deferSourceSamples = false;
        await act(async () => sourceSampleReleases.splice(0).forEach(release => release()));
        await until(() => !next()!.disabled, "Successful current validation did not enable Continue.");
        const success = area.querySelector<HTMLElement>(".iot-solution-success")!;
        assert(button("Validate again") && !button("Validate source tables") &&
            success.getAttribute("role") === "status" && success.querySelector("svg") &&
            success.querySelector(".iot-solution-success-text")?.textContent === "Source tables validated" &&
            currentContext.hubSourceAnalysis.telemetry.rawRows === 1000 && currentContext.hubSourceAnalysis.twin.rawRows === 1000 &&
            !area.textContent?.includes("Analysis complete") && !area.textContent?.includes("records analyzed"),
            "Successful validation lost the native exact feedback or restored the old count notice.");
        deferredToken = true; tokenReleases = []; sourceSampleStatus = 503;
        await click(button("Validate again"));
        await until(() => tokenReleases.length > 0, "Revalidation did not reach its owned token wait.");
        assert(!sourceValidationAction() && !area.querySelector(".iot-solution-success") &&
            !currentContext.hubSourceAnalysis && next()!.disabled, "A new validation retained old success/readiness.");
        deferredToken = false;
        await act(async () => tokenReleases.splice(0).forEach(release => release()));
        await until(() => !!area.querySelector('[role="alert"]') && !!button("Validate again"), "Failed revalidation did not restore its retry action.");
        assert(!button("Validate source tables") && !area.querySelector(".iot-solution-success") && next()!.disabled,
            "A failed validation retained success or changed the Hub-style re-validation action back to initial.");
        sourceSampleStatus = 0; emptySamples = "both";
        await click(button("Validate again")); await until(() => !next()!.disabled, "Explicit retry with a successful empty sample did not finish.");
        assert(area.querySelector(".iot-solution-success-text")?.textContent === "Source tables validated" &&
            !area.querySelector('[role="alert"]') && currentContext.hubSourceAnalysis.telemetry.rawRows === 0 &&
            currentContext.hubSourceAnalysis.twin.rawRows === 0, "Empty success or failure recovery changed the source validation contract.");
        caseObservations = { initialLabel: true, actionHiddenWhilePending: true, centeredSpinnerLabelBelow: true,
            exactNativeSuccess: true, initialRawRows: [1000, 1000], revalidationClearsSuccessImmediately: true,
            failureRestoresRevalidationAction: true, emptySuccessAllowed: true, oldCountNoticeAbsent: true };
    })],
    ["Source Validation feedback clears on table database Eventhouse changes and cannot return from stale work", () => observeWizardContext(async () => {
        for (const change of ["table", "database", "eventhouse"]) {
            await openAnalysisSource();
            await until(() => !sourceValidationAction()!.disabled, "Reset fixture metadata did not settle.");
            await click(sourceValidationAction()); await until(() => !next()!.disabled, "Reset fixture validation did not finish.");
            assert(area.querySelector(".iot-solution-success"), "The reset fixture had no successful validation feedback.");
            if (change === "table") {
                await chooseOption("hub-model-telemetry", "Show All Tables...");
                await chooseOption("hub-model-telemetry", "NewTelemetry");
            } else await chooseOption(change === "database" ? "hub-model-database" : "hub-model-eventhouse",
                change === "database" ? "Database" : "Eventhouse");
            assert(!area.querySelector(".iot-solution-success") && !currentContext.hubSourceAnalysis &&
                !!button("Validate source tables") && !button("Validate again") && next()!.disabled,
                `${change} selection/reset retained validation feedback or borrowed old readiness.`);
        }
        await openAnalysisSource();
        await until(() => !sourceValidationAction()!.disabled, "ABA feedback metadata did not settle.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "ABA feedback validation did not finish.");
        const signalStart = sourceSampleSignals.length;
        deferSourceSamples = true;
        await click(button("Validate again"));
        await until(() => sourceSampleReleases.length === 2, "Revalidation did not reach its pending raw reads.");
        await chooseOption("hub-model-telemetry", "Show All Tables...");
        await chooseOption("hub-model-telemetry", "NewTelemetry");
        await chooseOption("hub-model-telemetry", "Telemetry");
        deferSourceSamples = false;
        await act(async () => sourceSampleReleases.splice(0).forEach(release => release())); await settle();
        assert(sourceSampleSignals.slice(signalStart).every(signal => signal.aborted) && !currentContext.hubSourceAnalysis &&
            !area.querySelector(".iot-solution-success") && !area.querySelector('[role="alert"]') &&
            !!button("Validate source tables") && next()!.disabled, "A canceled ABA result restored feedback, error or readiness.");
        await click(button("Validate source tables")); await until(() => !next()!.disabled, "Current explicit validation did not recover after ABA.");
        caseObservations = { selectionResets: ["table", "same database", "same Eventhouse"],
            feedbackAndCacheCleared: true, staleAbortedResponseCannotRestoreSuccess: true, explicitCurrentRecovery: true };
    })],
    ["Dashboard raw source pickers share Select a table while keeping distinct labels and prerequisite gates", async () => {
        await mount(<Harness component={DashboardDataStep} initial={{}}/>);
        await until(() => !area.querySelector<HTMLButtonElement>("#dash-eh-select")!.disabled, "Dashboard Eventhouse choices did not settle.");
        const tableIds = ["dash-telemetry-table", "dash-properties-table"];
        const assertTablePlaceholders = (disabled: boolean) => {
            assert(tableIds.every(id => {
                const control = area.querySelector<HTMLButtonElement>(`#${id}`)!;
                return control.textContent?.trim() === "Select a table" && control.disabled === disabled &&
                    !!area.querySelector(`label[for="${id}"]`);
            }) && JSON.stringify([...area.querySelectorAll(".iot-solution-field-label")]
                .map(node => node.textContent?.replace(/\s*\*$/, "").trim())) ===
                JSON.stringify(["Eventhouse", "KQL database", "Raw telemetry table", "Raw properties table"]),
                "Dashboard table placeholders, separate field labels or prerequisite disabling changed.");
        };
        assertTablePlaceholders(true);
        assert(area.querySelector<HTMLButtonElement>("#dash-db-select")!.disabled && button("Inspect Tables")!.disabled,
            "Missing Dashboard source prerequisites enabled database selection or inspection.");
        await chooseOption("dash-eh-select", "Eventhouse");
        await until(() => !area.querySelector<HTMLButtonElement>("#dash-db-select")!.disabled, "Dashboard database choices did not settle.");
        assertTablePlaceholders(true);
        await chooseOption("dash-db-select", "Database");
        await until(() => !area.querySelector<HTMLButtonElement>("#dash-telemetry-table")!.disabled, "Dashboard raw table choices did not settle.");
        assertTablePlaceholders(false);
        await chooseOption("dash-telemetry-table", "Telemetry");
        assert(button("Inspect Tables")!.disabled &&
            area.querySelector("#dash-properties-table")?.textContent?.trim() === "Select a table",
            "Selecting telemetry changed the properties placeholder or bypassed paired table selection.");
        await chooseOption("dash-properties-table", "Twins");
        assert(!button("Inspect Tables")!.disabled && currentContext.dashTelemetryTable === "Telemetry" &&
            currentContext.dashPropertiesTable === "Twins" && !currentContext.dashInspected,
            "Shared placeholders collapsed table identities or implicitly inspected the selected source.");
        caseObservations = { sharedNormalPlaceholder: "Select a table", distinctLabelsAndIds: tableIds,
            prerequisiteDisablingPreserved: true, explicitInspectionStillRequired: true, liveWrites: 0 };
    }],
    ["Source picker placeholders match Dashboard without prerequisite helpers through selection reset and reload", () => observeWizardContext(async () => {
        const analysisReads = analysisQueries().length;
        const initial = context(true);
        for (const key of ["modelEhId", "modelEhName", "modelDbId", "modelDbName", "modelQueryServiceUri",
            "rawTelemetryTable", "rawPropertiesTable"]) initial[key] = "";
        initial.modelRawDataSelected = false;
        const placeholders = [
            ["hub-model-eventhouse", "Select an Eventhouse"],
            ["hub-model-database", "Select a database"],
            ["hub-model-telemetry", "Select a table"],
            ["hub-model-twin", "Select a table"],
        ] as const;
        const checkPlaceholders = (ids: string[]) => {
            for (const [id, text] of placeholders.filter(([id]) => ids.includes(id))) {
                const control = area.querySelector<HTMLButtonElement>(`#${id}`)!;
                assert(control?.textContent?.trim() === text && control.getAttribute("aria-busy") !== "true" &&
                    !control.hasAttribute("aria-describedby") && !!area.querySelector(`label[for="${id}"]`),
                    `The ${id} picker lost its own placeholder, label or settled accessible state.`);
            }
            assert(!area.querySelector("#database-prerequisite, #hub-model-telemetry-prerequisite, #hub-model-twin-prerequisite") &&
                !area.textContent?.includes("Select an Eventhouse first") && !area.textContent?.includes("Select a database first"),
                "Redundant prerequisite helper text or a dangling description remains.");
        };
        await openAnalysisSource(initial);
        await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-eventhouse")!.disabled, "Eventhouse choices did not settle.");
        checkPlaceholders(placeholders.map(([id]) => id));
        assert(placeholders.slice(1).every(([id]) => area.querySelector<HTMLButtonElement>(`#${id}`)!.disabled) &&
            sourceValidationAction()!.disabled && next()!.disabled,
            "Removing prerequisite text enabled unavailable choices or source validation.");
        await chooseOption("hub-model-eventhouse", "Eventhouse");
        await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-database")!.disabled, "Database choices did not settle.");
        checkPlaceholders(["hub-model-database", "hub-model-telemetry", "hub-model-twin"]);
        assert(["hub-model-telemetry", "hub-model-twin"].every(id => area.querySelector<HTMLButtonElement>(`#${id}`)!.disabled),
            "Selecting only an Eventhouse enabled raw table choices.");
        await chooseOption("hub-model-database", "Database");
        await until(tableChoicesReady, "Raw table choices did not settle.");
        checkPlaceholders(["hub-model-telemetry", "hub-model-twin"]);
        await chooseOption("hub-model-telemetry", "Telemetry");
        assert(sourceValidationAction()!.disabled, "One raw table bypassed paired source selection.");
        await chooseOption("hub-model-twin", "Twins");
        assert(!sourceValidationAction()!.disabled && next()!.disabled && analysisQueries().length === analysisReads,
            "Placeholder changes triggered analysis or credited source selection as validation.");
        await chooseOption("hub-model-eventhouse", "Eventhouse");
        await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-database")!.disabled, "Eventhouse reset did not reload its choices.");
        checkPlaceholders(["hub-model-database", "hub-model-telemetry", "hub-model-twin"]);
        assert(!currentContext.modelDbId && !currentContext.rawTelemetryTable && !currentContext.rawPropertiesTable &&
            area.querySelector<HTMLButtonElement>("#hub-model-telemetry")!.disabled && sourceValidationAction()!.disabled,
            "Eventhouse reset retained downstream selections or enabled source validation.");
        let release: (response: Response) => void = () => undefined, pending = false;
        tableMetadataResponse = () => { pending = true; return new Promise(resolve => { release = resolve; }); };
        await chooseOption("hub-model-database", "Database");
        await until(() => pending, "Reload did not reach the pending table request.");
        assert(["hub-model-telemetry", "hub-model-twin"].every(id => {
            const control = area.querySelector<HTMLButtonElement>(`#${id}`)!;
            return control.disabled && control.textContent?.trim() === "Loading..." &&
                control.getAttribute("aria-busy") === "true" && !control.hasAttribute("aria-describedby");
        }), "A genuine table reload lost its loading text, busy state or disabling.");
        await act(async () => release(result([["TableName", "string"], ["DocString", "string"]], [
            ["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw"], ["Twins", "@iot.meta tableRole=iothub-properties-raw"],
        ])));
        await until(tableChoicesReady, "Table reload did not settle.");
        checkPlaceholders(["hub-model-telemetry", "hub-model-twin"]);
        assert(sourceValidationAction()!.disabled && next()!.disabled && analysisQueries().length === analysisReads &&
            !area.querySelector('[role="alert"]'), "Reload fabricated an error, analysis or completed source state.");
        caseObservations = { exactDashboardPlaceholders: placeholders, prerequisiteHelpers: 0, danglingDescriptions: 0,
            initialPrerequisiteDisabling: true, upstreamSelectionAndReset: true, genuineReloadLoadingPreserved: true,
            rawAnalysisReads: analysisQueries().length - analysisReads, liveWrites: 0 };
    })],
    ["Source Reselection opening is inert and same Eventhouse then database restores owned inventory", () => observeWizardContext(async () => {
        await openAnalysisSource();
        await until(() => !sourceValidationAction()!.disabled, "Initial inventory did not settle.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Initial analysis did not finish.");
        const before = JSON.stringify(currentContext), reads = calls.length + resourceCalls.length;
        for (const id of ["hub-model-eventhouse", "hub-model-database"]) {
            const control = area.querySelector<HTMLElement>(`#${id}`)!;
            await click(control);
            await act(async () => control.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true })));
            await settle();
            assert(JSON.stringify(currentContext) === before && calls.length + resourceCalls.length === reads && !next()!.disabled,
                "Opening a source menu changed metadata, selections or analyzed readiness.");
        }
        const ehReads = resourceCalls.filter(url => url.includes("/eventhouses/eh")).length;
        await chooseOption("hub-model-eventhouse", "Eventhouse");
        assert(!currentContext.modelDbId && !currentContext.rawTelemetryTable && !currentContext.rawPropertiesTable &&
            !currentContext.hubSourceAnalysis && next()!.disabled, "Eventhouse reselection did not follow Dashboard's downstream reset.");
        await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-database")!.disabled, "The same Eventhouse did not reload its database choices.");
        await chooseOption("hub-model-database", "Database");
        await settle();
        const state = () => ({ queryUriPresent: !!currentContext.modelQueryServiceUri,
            eventhouseReloads: resourceCalls.filter(url => url.includes("/eventhouses/eh")).length - ehReads,
            tableReads: calls.filter(query => query === ".show tables details | project TableName, DocString").length,
            falseEmpty: area.textContent?.includes("No tables found in Database.") });
        await until(tableChoicesReady, () => `Same Eventhouse/database stranded inventory: ${JSON.stringify(state())}`);
        assert(currentContext.modelQueryServiceUri === "https://mock.kusto.fabric.microsoft.com" && state().eventhouseReloads === 1 &&
            !state().falseEmpty && !currentContext.hubSourceAnalysis && next()!.disabled,
            `Same-identity reload did not restore current metadata or bypassed analysis: ${JSON.stringify(state())}`);
        await chooseOption("hub-model-telemetry", "Telemetry"); await chooseOption("hub-model-twin", "Twins");
        assert(!sourceValidationAction()!.disabled && next()!.disabled, "Inventory recovery implicitly analyzed source data.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Explicit reanalysis could not recover.");
        caseObservations = { openOnlyInert: true, sameEventhouseReloads: 1, queryUriRestored: true,
            tableInventoryRestored: true, matchesDashboardResetReload: true, explicitReanalysisRequired: true };
    })],
    ["Source Reselection same database reloads and cancels a superseded same-identity inventory read", () => observeWizardContext(async () => {
        await openAnalysisSource();
        await until(tableChoicesReady, "Initial database inventory did not settle.");
        let release: (value: Response) => void = () => undefined, pending = false;
        tableMetadataResponse = () => { pending = true; return new Promise(resolve => { release = resolve; }); };
        const before = calls.filter(query => query === ".show tables details | project TableName, DocString").length;
        await chooseOption("hub-model-database", "Database");
        await until(() => pending, "Same database selection cleared the table list without starting an owned reload.");
        assert(area.querySelector("#hub-model-telemetry")?.getAttribute("aria-busy") === "true" &&
            area.querySelector("#hub-model-telemetry")?.textContent?.includes("Loading...") &&
            !area.textContent?.includes("No tables found"), "Pending same-database reload was published as empty.");
        tableMetadataResponse = async () => result([["TableName", "string"], ["DocString", "string"]], [
            ["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw"], ["Twins", "@iot.meta tableRole=iothub-properties-raw"], ["FreshOnly", null],
        ]);
        await chooseOption("hub-model-database", "Database");
        await until(tableChoicesReady, "The current same-database reload did not settle.");
        await act(async () => release(result([["TableName", "string"], ["DocString", "string"]], [["StaleOnly", null]])));
        await settle();
        await chooseOption("hub-model-telemetry", "Show All Tables...");
        assert(activeOptions("hub-model-telemetry").some(option => option.textContent === "FreshOnly") &&
            !activeOptions("hub-model-telemetry").some(option => option.textContent === "StaleOnly") &&
            !area.textContent?.includes("No tables found") && !currentContext.hubSourceAnalysis &&
            calls.filter(query => query === ".show tables details | project TableName, DocString").length === before + 2,
            "A superseded same-identity response replaced the current table inventory or analysis.");
        caseObservations = { sameDatabaseReloads: 2, loadingIsNotEmpty: true, canceledLateResultIgnored: true,
            noImplicitAnalysis: true };
    })],
    ["Source metadata isolates malformed and duplicate table tags without hiding valid roles or manual raw selection", () => observeWizardContext(async () => {
        tableMetadataResponse = async () => result([["TableName", "string"], ["DocString", "string"]], [
            ["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw"],
            ["Twins", "@iot.meta tableRole=iothub-properties-raw"],
            ["Malformed", "@iot.meta invalid"],
            ["Duplicate", "@iot.meta tableRole=iothub-telemetry-raw\n@iot.meta tableRole=modeled-data"],
            ["ManualRaw", null],
        ]);
        await openAnalysisSource();
        await until(tableChoicesReady, "Table metadata did not settle with an unrelated invalid docstring.");
        const warning = area.querySelector('[role="status"] .iot-solution-model-details');
        assert(warning?.textContent?.includes("Malformed") && warning.textContent.includes("Duplicate") &&
            warning.textContent.includes("Malformed line-anchored") && warning.textContent.includes("Duplicate IoT metadata") &&
            !area.querySelector('[role="alert"]') && !area.textContent?.includes("No tables found"),
            "Malformed metadata was hidden, escalated to a table-list failure, or rendered as empty.");
        await click(area.querySelector("#hub-model-telemetry"));
        assert(activeOptions("hub-model-telemetry").map(option => option.textContent).join("|") === "Telemetry|Show All Tables...",
            "Invalid metadata became a telemetry role hint or valid table hints were lost.");
        await chooseOption("hub-model-telemetry", "Show All Tables...");
        assert(["Malformed", "Duplicate", "ManualRaw"].every(name =>
            activeOptions("hub-model-telemetry").some(option => option.textContent === name)),
            "Malformed or unannotated tables disappeared from manual raw selection.");
        await chooseOption("hub-model-telemetry", "Malformed");
        assert(!sourceValidationAction()!.disabled && next()!.disabled, "Manual selection bypassed explicit source analysis.");
        await click(sourceValidationAction());
        await until(() => !next()!.disabled, "Valid raw data could not be analyzed independently of metadata hints.");
        assert(currentContext.hubAssignmentSource?.telemetry.docstring === "" &&
            currentContext.hubAssignmentSource?.twin.docstring === "@iot.meta tableRole=iothub-properties-raw",
            "Invalid metadata leaked into downstream assignment hints or a valid table's metadata was removed.");
        caseObservations = { invalidTables: 2, visibleDiagnostics: true, validRoleHintsRetained: true,
            manualRawSelectionAndExplicitAnalysis: true, malformedAssignmentHintsExcluded: true };
    })],
    ["Source metadata warnings and coverage stay retired when a superseded table response arrives late", () => observeWizardContext(async () => {
        await openAnalysisSource();
        await until(tableChoicesReady, "Initial table metadata did not settle.");
        await click(sourceValidationAction());
        await until(() => !next()!.disabled, "Initial source analysis did not finish.");
        let release: (response: Response) => void = () => undefined, pending = false;
        tableMetadataResponse = () => { pending = true; return new Promise(resolve => { release = resolve; }); };
        await chooseOption("hub-model-database", "Database");
        await until(() => pending, "Superseded table read was not pending.");
        assert(!currentContext.hubSourceAnalysis && !currentContext.hubCoverage && !currentContext.hubStreamChecks &&
            next()!.disabled, "Source reload retained analyzed data or coverage credit.");
        tableMetadataResponse = async () => result([["TableName", "string"], ["DocString", "string"]], [
            ["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw"],
            ["Twins", "@iot.meta tableRole=iothub-properties-raw"],
        ]);
        await chooseOption("hub-model-database", "Database");
        await until(tableChoicesReady, "Current table read did not settle.");
        await act(async () => release(result([["TableName", "string"], ["DocString", "string"]], [
            ["StaleInvalid", "@iot.meta tableRole=a\n@iot.meta tableRole=b"],
        ])));
        await settle();
        assert(!area.querySelector(".iot-solution-model-details") && !area.textContent?.includes("StaleInvalid") &&
            !area.querySelector('[role="alert"]') && !currentContext.hubCoverage && !currentContext.hubStreamChecks &&
            !currentContext.hubSourceAnalysis && next()!.disabled,
            "Late metadata restored a warning, error, source analysis, or scoped coverage.");
        await chooseOption("hub-model-telemetry", "Telemetry");
        await chooseOption("hub-model-twin", "Twins");
        await click(sourceValidationAction());
        await until(() => !next()!.disabled, "Explicit current source analysis did not recover.");
        caseObservations = { supersededMetadataIgnored: true, staleWarningsAndProofCannotPublish: true,
            explicitCurrentRecovery: true };
    })],
    ["Source Reselection metadata loading is in the four visible fields and has no Source Tables heading", async () => {
        const stages: string[] = [];
        for (const stage of ["eventhouse", "database", "tables"]) {
            let pending = false, release: (value: Response) => void = () => undefined;
            const deferred = () => { pending = true; return new Promise<Response>(resolve => { release = resolve; }); };
            resourceResponse = stage === "eventhouse" ? deferred : undefined;
            eventhouseMetadataResponse = stage === "database" ? deferred : undefined;
            tableMetadataResponse = stage === "tables" ? deferred : undefined;
            await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
            await until(() => pending, `The ${stage} loading boundary was not reached.`);
            const ids = stage === "eventhouse" ? ["hub-model-eventhouse"] : stage === "database" ? ["hub-model-database"]
                : ["hub-model-telemetry", "hub-model-twin"];
            assert(["hub-model-eventhouse", "hub-model-database", "hub-model-telemetry", "hub-model-twin"].every(id => area.querySelector(`#${id}`)) &&
                ids.every(id => area.querySelector<HTMLButtonElement>(`#${id}`)!.disabled &&
                    area.querySelector(`#${id}`)!.getAttribute("aria-busy") === "true" &&
                    area.querySelector(`#${id}`)!.textContent?.includes("Loading...")) &&
                !area.textContent?.includes("No tables found") && !area.textContent?.includes("No databases found") &&
                ![...area.querySelectorAll("h3")].some(node => node.textContent === "Source Tables"),
                `The ${stage} field did not follow Dashboard's in-field loading or fabricated empty state.`);
            const response = stage === "eventhouse" ? new Response(JSON.stringify({ value: [
                { id: "eh", type: "Eventhouse", displayName: "Eventhouse" }, { id: "db", type: "KQLDatabase", displayName: "Database" },
            ] })) : stage === "database" ? new Response(JSON.stringify({ properties: {
                queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databasesItemIds: ["db"],
            } })) : result([["TableName", "string"], ["DocString", "string"]], [["Telemetry", null], ["Twins", null]]);
            await act(async () => release(response)); await until(tableChoicesReady, `The ${stage} loading state did not recover.`);
            assert(ids.every(id => area.querySelector(`#${id}`)!.getAttribute("aria-busy") !== "true"),
                "A settled field retained loading presentation.");
            stages.push(stage);
        }
        caseObservations = { inFieldLoadingStages: stages, fourSelectorsAlwaysVisible: true, noGroupHeading: true };
    }],
    ["Source Reselection only a completed current table inventory can publish genuine empty", async () => {
        databaseItemIds = [];
        const before = calls.length;
        await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
        await until(() => area.textContent?.includes("No databases found in Eventhouse."), "Successful no-database metadata did not settle.");
        assert(!area.textContent?.includes("No tables found") &&
            calls.slice(before).every(query => query !== ".show tables details | project TableName, DocString") &&
            area.querySelector<HTMLButtonElement>("#hub-model-telemetry")!.disabled,
            "A skipped, unowned table read was mislabeled as a successful empty inventory.");
        databaseItemIds = ["db"];
        for (const failure of [false, true]) {
            tableMetadataResponse = async () => failure ? new Response("Synthetic failure", { status: 503 })
                : result([["TableName", "string"], ["DocString", "string"]], []);
            await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
            await until(() => failure ? !!area.querySelector('[role="alert"]') : !!area.textContent?.includes("No tables found in Database."),
                "Settled table outcome was not shown.");
            assert(!!area.textContent?.includes("No tables found in Database.") === !failure &&
                !!area.querySelector('[role="alert"]') === failure && !currentContext.hubSourceAnalysis,
                "A real empty success, service failure or cache state was conflated.");
        }
        caseObservations = { skippedReadNotEmpty: true, genuineEmptyShown: true, serviceErrorNotEmpty: true };
    }],
    ["Source Reselection real Eventhouse ABA cancels analysis and rejects late metadata before explicit recovery", () => observeWizardContext(async () => {
        resourceResponse = async () => new Response(JSON.stringify({ value: [
            { id: "eh", type: "Eventhouse", displayName: "Eventhouse" }, { id: "db", type: "KQLDatabase", displayName: "Database" },
            { id: "eh2", type: "Eventhouse", displayName: "Eventhouse Two" }, { id: "db2", type: "KQLDatabase", displayName: "Database Two" },
        ] }));
        await openAnalysisSource(); await until(tableChoicesReady, "ABA initial metadata did not settle.");
        deferSourceSamples = true;
        await click(sourceValidationAction()); await until(() => sourceSampleReleases.length === 2, "Analysis did not reach both pending reads.");
        const revision = currentContext.hubSourceSelectionRevision || 0;
        let pending = false, release: (value: Response) => void = () => undefined;
        eventhouseMetadataResponse = () => { pending = true; return new Promise(resolve => { release = resolve; }); };
        await chooseOption("hub-model-eventhouse", "Eventhouse Two");
        await until(() => pending, "The different Eventhouse read did not start.");
        assert(sourceSampleSignals.every(signal => signal.aborted) && !currentContext.hubSourceAnalysis &&
            !currentContext.hubCoverage && !currentContext.hubStreamChecks && next()!.disabled,
            "A real Eventhouse change did not abort analysis and retire scoped proof.");
        eventhouseMetadataResponse = async () => new Response(JSON.stringify({ properties: {
            queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databasesItemIds: ["db"],
        } }));
        await chooseOption("hub-model-eventhouse", "Eventhouse");
        await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-database")!.disabled, "Returning Eventhouse metadata did not settle.");
        await chooseOption("hub-model-database", "Database"); await until(tableChoicesReady, "Returning database inventory did not settle.");
        deferSourceSamples = false;
        await act(async () => {
            release(new Response(JSON.stringify({ properties: {
                queryServiceUri: "https://stale.kusto.fabric.microsoft.com", databasesItemIds: ["db2"],
            } })));
            sourceSampleReleases.splice(0).forEach(complete => complete());
        });
        await settle();
        assert(currentContext.modelQueryServiceUri === "https://mock.kusto.fabric.microsoft.com" &&
            currentContext.hubSourceSelectionRevision > revision && !currentContext.hubSourceAnalysis && next()!.disabled &&
            !area.querySelector('[role="alert"]'), "An old Eventhouse/sample result revived its URI, error or ready cache after ABA.");
        await chooseOption("hub-model-telemetry", "Telemetry"); await chooseOption("hub-model-twin", "Twins");
        assert(next()!.disabled && !sourceValidationAction()!.disabled, "Metadata recovery silently analyzed changed sources.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Explicit current analysis could not recover after ABA.");
        assert(currentContext.hubSourceAnalysis.key === modelSourceKey("workspace", currentContext), "Analysis was published for an old source owner.");
        caseObservations = { realEventhouseChangeRetiresProof: true, analysisAborted: true,
            lateMetadataAndSamplesRejected: true, awayAndBackRevision: true, explicitCurrentRecovery: true };
    })],
    ["Source Analysis keeps four selectors visible and requires latest1000 raw records per stream before Continue", () => observeWizardContext(async () => {
        const initial = context(true);
        for (const key of ["modelEhId", "modelEhName", "modelDbId", "modelDbName", "modelQueryServiceUri",
            "rawTelemetryTable", "rawPropertiesTable"]) initial[key] = "";
        initial.modelRawDataSelected = false; delete initial.hubModel; delete initial.hubSelectedFormat; initial.hubFields = [];
        sourceSampleRows = 1000;
        await openAnalysisSource(initial);
        assert(["hub-model-eventhouse", "hub-model-database", "hub-model-telemetry", "hub-model-twin"]
            .every(id => !!area.querySelector(`#${id}`)), "The source prerequisites hid a selector.");
        assert(["hub-model-database", "hub-model-telemetry", "hub-model-twin"].every(id =>
            area.querySelector<HTMLButtonElement>(`#${id}`)!.disabled) && sourceValidationAction()?.disabled && next()!.disabled,
            "Unmet prerequisites did not keep the visible controls and Continue disabled.");
        await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-eventhouse")!.disabled, "Eventhouse load did not finish.");
        await chooseOption("hub-model-eventhouse", "Eventhouse");
        await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-database")!.disabled, "Database load did not finish.");
        await chooseOption("hub-model-database", "Database");
        await until(tableChoicesReady, "The raw table inventory did not finish.");
        await chooseOption("hub-model-telemetry", "Telemetry");
        assert(sourceValidationAction()?.disabled && next()!.disabled, "A single selected table completed source analysis.");
        await chooseOption("hub-model-twin", "Twins");
        assert(!sourceValidationAction()!.disabled && next()!.disabled && !currentContext.hubSourceAnalysis,
            "Metadata selection was credited as analyzed data.");
        const before = analysisQueries().length;
        deferSourceSamples = true;
        await click(sourceValidationAction());
        await until(() => sourceSampleReleases.length === 2, "Analysis did not read the two streams independently.");
        assert(area.textContent?.includes("Validating source tables...") && next()!.disabled &&
            !sourceValidationAction() && !currentContext.hubCoverage && !currentContext.hubStreamChecks,
            "Loading analysis enabled Continue or processing coverage.");
        const queries = analysisQueries().slice(before);
        assert(queries.length === 2 && queries.every(query => query.includes("| top 1000 by ingestion_time() desc\n| project data, headers, user_headers") &&
            !query.includes("| where") && !query.includes("memberTypeIRI")), "The actual source queries were not latest1000 unfiltered raw records each.");
        deferSourceSamples = false;
        await act(async () => sourceSampleReleases.splice(0).forEach(release => release()));
        await until(() => !next()!.disabled && !!currentSourceAnalysis("workspace", currentContext), "Successful analysis did not enable Continue.");
        const cached = currentContext.hubSourceAnalysis;
        assert(cached.limit === SOURCE_ANALYSIS_LIMIT && cached.telemetry.rawRows === 1000 && cached.twin.rawRows === 1000 &&
            cached.telemetry.payloads.length === 1000 && cached.twin.payloads.length === 1000 &&
            cached.key === modelSourceKey("workspace", currentContext) && !currentContext.hubCoverage && !currentContext.hubStreamChecks &&
            !currentContext.hubModel, "The cache limit, source binding, model-free behavior or processing separation is wrong.");
        assert(!area.textContent?.includes("Analyzes up to the latest") && !area.textContent?.includes("Analysis complete"),
            "The source step retained the removed analysis helper or success-count notice.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Add a device model", "Analyzed source did not open the existing model step.");
        assert(area.querySelector("#thing-model-upload") && area.querySelector<HTMLInputElement>('input[value="wot"]')?.checked &&
            !currentContext.hubModel && next()!.disabled,
            "The WoT default was missing or source analysis was credited as an accepted model.");
        caseObservations = { visibleSelectors: 4, queryCount: 2, queries, limitPerStream: 1000, actualRawRows: [1000, 1000],
            analysisNeededForContinue: true, modelRequiredForAnalysis: false, unfilteredCacheNotProcessingCoverage: true };
    })],
    ["Source Analysis selection ABA and unmount abort old token or read owners and require explicit reanalysis", () => observeWizardContext(async () => {
        const observations: object[] = [];
        for (const boundary of ["token", "read"] as const) {
            sourceSampleSignals.splice(0);
            await openAnalysisSource();
            await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Source inventory did not settle.");
            const startRevision = currentContext.hubSourceSelectionRevision || 0;
            deferredToken = boundary === "token"; tokenReleases = [];
            deferSourceSamples = boundary === "read";
            await click(sourceValidationAction());
            await until(() => boundary === "token" ? tokenReleases.length > 0 : sourceSampleReleases.length === 2, "Owned analysis did not reach its wait boundary.");
            const telemetryMenu = area.querySelector<HTMLElement>("#hub-model-telemetry")!;
            if (telemetryMenu.getAttribute("aria-expanded") !== "true") await click(telemetryMenu);
            if (activeOptions("hub-model-telemetry").some(option => option.textContent === "Show All Tables..."))
                await chooseOption("hub-model-telemetry", "Show All Tables...");
            await chooseOption("hub-model-telemetry", "NewTelemetry");
            await chooseOption("hub-model-telemetry", "Telemetry");
            assert(next()!.disabled && !currentContext.hubSourceAnalysis &&
                currentContext.hubSourceSelectionRevision > startRevision && !currentContext.hubCoverage && !currentContext.hubStreamChecks,
                "An away-and-back source edit retained analysis or downstream proof.");
            const beforeLate = analysisQueries().length;
            deferredToken = false; deferSourceSamples = false;
            await act(async () => { tokenReleases.splice(0).forEach(release => release()); sourceSampleReleases.splice(0).forEach(release => release()); });
            await settle();
            assert(!currentContext.hubSourceAnalysis && next()!.disabled && analysisQueries().length === beforeLate &&
                !area.querySelector('[role="alert"]'), "A canceled old analysis published readiness, dispatched after a token wait or announced an error.");
            if (boundary === "read") assert(sourceSampleSignals.every(signal => signal.aborted), "Old stream requests were not aborted.");
            await click(sourceValidationAction());
            await until(() => !next()!.disabled && !!currentSourceAnalysis("workspace", currentContext), "Explicit current reanalysis could not recover.");
            observations.push({ boundary, awayAndBackRetiresAnalysis: true, oldWaitCannotPublish: true, currentExplicitRecovery: true });
        }
        await openAnalysisSource();
        await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Unmount fixture did not settle.");
        deferredToken = true; tokenReleases = [];
        await click(sourceValidationAction());
        await until(() => tokenReleases.length > 0, "Disposal did not reach its analysis token wait.");
        await act(async () => setStepVisible(false));
        const reads = analysisQueries().length;
        deferredToken = false;
        await act(async () => tokenReleases.splice(0).forEach(release => release()));
        await settle();
        assert(area.textContent?.includes("Analysis journey closed.") && analysisQueries().length === reads,
            "Unmount allowed an old analysis token to dispatch or reopen the journey.");
        caseObservations = { observations, unmountRetiresPendingAnalysis: true, payloadsLogged: false };
    })],
    ["Source Analysis failed or malformed samples are errors while successful empty evidence can continue without inferred roles", () => observeWizardContext(async () => {
        const initial = context(true); delete initial.hubModel; delete initial.hubSelectedFormat; initial.hubFields = [];
        await openAnalysisSource(initial);
        await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Error fixture source did not settle.");
        for (const failure of ["over-limit", "malformed", "service-error"]) {
            sourceSampleRows = failure === "over-limit" ? 1001 : 1;
            parityPayloads = failure === "malformed" ? { telemetry: [], twin: { properties: { reported: {} } } } : undefined;
            sourceSampleStatus = failure === "service-error" ? 503 : 0;
            await click(sourceValidationAction());
            await until(() => !!area.querySelector('[role="alert"]') && !sourceValidationAction()!.disabled,
                "A failed sample read did not settle as an explicit retryable analysis error.");
            assert(!currentContext.hubSourceAnalysis && next()!.disabled && !currentContext.modelRawDataSelected &&
                !area.textContent?.includes("Analysis complete."), "A failure was credited as successful empty analysis.");
        }
        sourceSampleRows = 1; sourceSampleStatus = 0; parityPayloads = undefined; emptySamples = "both";
        await click(sourceValidationAction());
        await until(() => !next()!.disabled && !!currentContext.hubSourceAnalysis, "Successful empty sample could not complete source analysis.");
        assert(currentContext.hubSourceAnalysis.telemetry.rawRows === 0 && currentContext.hubSourceAnalysis.twin.rawRows === 0 &&
            !area.textContent?.includes("Analysis complete") && !area.querySelector('[role="alert"]'),
            "Successful empty evidence was treated as a malformed failure or restored a removed success notice.");
        const sourceReads = analysisQueries().length;
        await click(next()); await click(area.querySelector('input[value="wot"]'));
        await file("thing-model-upload", JSON.stringify(tm), "empty-evidence.tm.json");
        assert(area.querySelector("h2")?.textContent === "Add a device model" && next()!.disabled &&
            currentContext.hubFields.every((field: ModelField) => !field.role) &&
            area.textContent?.includes("No sample records were found. Choose the roles manually.") &&
            !button("Suggest Roles from Data") && !area.querySelector("#field-hint-depth") && analysisQueries().length === sourceReads,
            "Empty analyzed data inferred roles, restored the redundant workflow or reread sources.");
        const manual = currentContext.hubFields.find((field: ModelField) => field.path[0] === "temperature");
        await enter(documentElementRole(manual.id)!.id, "twin");
        emptySamples = "none";
        parityPayloads = { telemetry: { temperature: 22, state: "ready" },
            twin: { properties: { reported: { temperature: 22, state: "ready" } } } };
        await click(step("Select data source"));
        await until(() => !sourceValidationAction()!.disabled, "Inconclusive reanalysis source did not settle.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Inconclusive source did not complete analysis.");
        await click(next());
        assert(area.textContent?.includes("The analyzed data gives no clear role suggestions.") &&
            currentContext.hubFields.find((field: ModelField) => field.id === manual.id).roleSource === "manual" &&
            currentContext.hubFields.find((field: ModelField) => field.id === manual.id).role === "twin" &&
            currentContext.hubFields.filter((field: ModelField) => field.id !== manual.id).every((field: ModelField) => !field.role) &&
            !area.textContent?.includes("No sample records were found."),
            "Both-stream evidence assigned a role, erased manual intent or retained stale empty feedback.");
        caseObservations = { failures: ["over-limit", "malformed", "service-error"], failuresNeverEmptySuccess: true,
            emptyCanContinue: true, emptyRolesRemainManual: true, inconclusiveRetainsManualAndClearsEmpty: true, extraRoleQueries: 0 };
    })],
    ["Source Analysis cache drives automatic WoT roles and model changes without extra reads while manual and DTDL roles stay explicit", () => observeWizardContext(async () => {
        const initial = context(true); delete initial.hubModel; delete initial.hubSelectedFormat; initial.hubFields = [];
        sourceSampleRows = 1000;
        await openAnalysisSource(initial);
        await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Automatic-role source did not settle.");
        await click(sourceValidationAction());
        await until(() => !next()!.disabled, "Automatic-role source was not analyzed.");
        const cached = currentContext.hubSourceAnalysis, reads = analysisQueries().length;
        await click(next()); await click(area.querySelector('input[value="wot"]'));
        await file("thing-model-upload", JSON.stringify(tm), "automatic-roles.tm.json");
        await until(() => currentContext.hubFields.every((field: ModelField) => !!field.role), "Upload did not apply role suggestions from the source cache.");
        assert(currentContext.hubFields.map((field: ModelField) => field.role).join(",") === "telemetry,twin" &&
            currentContext.hubModel.fields.every((field: ModelField) => !field.role), "WoT affordance kind/flags gained role authority.");
        await click(button("Get Copilot prompt"));
        await until(() => !!area.querySelector("#model-preparation-prompt"), "The optional cached-sample prompt did not appear.");
        assert(!area.querySelector("#authoring-depth") && !button("Inspect Selected Samples") && analysisQueries().length === reads,
            "Optional preparation restored the extra depth/analyze workflow or fetched again.");
        await click(button("Hide prompt"));
        assert(!button("Suggest Roles from Data") && !area.querySelector("#field-hint-depth") &&
            currentContext.hubFields.length === 2 && area.textContent?.includes("Select a role for every field to continue"),
            "Fields still required a second analyze action.");
        const temperature = currentContext.hubFields.find((field: ModelField) => field.path[0] === "temperature");
        await enter(documentElementRole(temperature.id)!.id, "twin");
        await click(step("Add device model"));
        await file("thing-model-upload", JSON.stringify({ ...tm, id: "urn:example:changed-model", title: "Changed customer model" }), "changed-model.tm.json");
        await until(() => currentContext.hubModel?.id === "urn:example:changed-model", "The replacement model did not settle.");
        assert(currentContext.hubFields.find((field: ModelField) => field.id === temperature.id).role === "twin" &&
            currentContext.hubFields.find((field: ModelField) => field.id === temperature.id).roleSource === "manual" &&
            currentContext.hubSourceAnalysis === cached && analysisQueries().length === reads,
            "Reanalysis for a new model lost explicit choices or reread the same source.");
        const reverseDtdl = { "@id": "dtmi:fixture:DeclaredRoles;1", "@type": "Interface", contents: [
            { "@type": "Property", name: "temperature", schema: "double" }, { "@type": "Telemetry", name: "state", schema: "string" },
        ] };
        await click(area.querySelector('input[value="dtdl"]')); await file("thing-model-upload", JSON.stringify(reverseDtdl), "declared-roles.json");
        await until(() => currentContext.hubModel?.kind === "dtdl", "DTDL did not settle.");
        assert(area.querySelector("h2")?.textContent === "Add a device model" &&
            currentContext.hubFields.every((field: ModelField) => field.roleSource === "model") &&
            currentContext.hubFields.find((field: ModelField) => field.path[0] === "temperature").role === "twin" &&
            currentContext.hubFields.find((field: ModelField) => field.path[0] === "state").role === "telemetry" &&
            !area.querySelector('[role="combobox"][id^="field-role-"]') && analysisQueries().length === reads,
            "Cached evidence overrode DTDL-declared roles, entered assignment or fetched roles again.");
        caseObservations = { cacheRawRows: [1000, 1000], automaticWotRoles: true, parserRolesRemainNeutral: true,
            manualRetainedAcrossModelChange: true, dtdlRolesOverrideContraryEvidence: true, additionalRoleOrPreparationReads: 0 };
    })],
    ["Source Analysis Show All is independent and keeps analyzed readiness until a real table selection changes", () => observeWizardContext(async () => {
        const observations: object[] = [];
        for (const keyboard of [false, true]) for (const id of ["hub-model-telemetry", "hub-model-twin"]) {
            await openAnalysisSource();
            await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Show All source did not settle.");
            await click(sourceValidationAction()); await until(() => !next()!.disabled, "Show All source was not analyzed.");
            const cached = currentContext.hubSourceAnalysis, reads = analysisQueries().length;
            const selection = [currentContext.rawTelemetryTable, currentContext.rawPropertiesTable].join("|");
            const control = area.querySelector<HTMLElement>(`#${id}`)!;
            control.focus(); await click(control);
            if (keyboard) for (const [key, code] of [["End", 35], ["Enter", 13]] as const) {
                await act(async () => control.dispatchEvent(new KeyboardEvent("keydown", { key, keyCode: code, which: code, bubbles: true, cancelable: true })));
                await settle();
            } else await chooseOption(id, "Show All Tables...");
            assert(control.getAttribute("aria-expanded") === "true" && activeOptions(id).some(option => option.textContent === "NewTelemetry") &&
                currentContext.hubSourceAnalysis === cached && !next()!.disabled && analysisQueries().length === reads &&
                [currentContext.rawTelemetryTable, currentContext.rawPropertiesTable].join("|") === selection,
                "Show All changed the analyzed selection/readiness or fetched data.");
            await act(async () => control.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape", bubbles: true, cancelable: true })));
            await settle();
            const peer = id === "hub-model-telemetry" ? "hub-model-twin" : "hub-model-telemetry";
            await click(area.querySelector(`#${peer}`));
            assert(activeOptions(peer).some(option => option.textContent === "Show All Tables...") &&
                !activeOptions(peer).some(option => option.textContent === "NewTelemetry"), "One menu's Show All leaked into its peer.");
            await chooseOption(peer, peer === "hub-model-twin" ? "Twins" : "Telemetry");
            const telemetryMenu = area.querySelector<HTMLElement>("#hub-model-telemetry")!;
            if (telemetryMenu.getAttribute("aria-expanded") !== "true") await click(telemetryMenu);
            if (activeOptions("hub-model-telemetry").some(option => option.textContent === "Show All Tables..."))
                await chooseOption("hub-model-telemetry", "Show All Tables...");
            await chooseOption("hub-model-telemetry", "NewTelemetry");
            assert(!currentContext.hubSourceAnalysis && next()!.disabled && analysisQueries().length === reads,
                "A genuine table change retained cached analysis or automatically reanalyzed.");
            await click(sourceValidationAction()); await until(() => !next()!.disabled, "Explicit changed-table analysis did not recover.");
            observations.push({ input: keyboard ? "keyboard" : "mouse", id, showAllDoesNotInvalidateOrRead: true,
                peerIndependent: true, realSelectionRequiresNewAnalysis: true });
        }
        caseObservations = { observations, runtimeExceptions: 0 };
    })],
    ["Source Analysis is not scoped processing coverage and guarded waiting creation still requires both independent checks", () => observeWizardContext(async () => {
        const seedRequest = creationRequest(true);
        await openAnalysisSource(alignmentCreationContext(seedRequest));
        await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Coverage source did not settle.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Coverage source was not analyzed.");
        const sourceReads = analysisQueries().length;
        assert(!currentContext.hubCoverage && !currentContext.hubStreamChecks, "Source samples were credited as processing coverage.");
        await click(next()); await reachScope();
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" && next()!.disabled,
            "Unfiltered source analysis bypassed Data Filters checks.");
        emptyProcessing = true;
        await verifyCoverage();
        assert(!next()!.disabled && !currentContext.hubAllowWaiting && !area.querySelector('input[type="checkbox"]'),
            "Verified zero required a removed waiting opt-in or pretended consent was checked.");
        await click(next());
        const currentRequest = { input: hubInput(currentContext), coverage: currentContext.hubCoverage,
            allowWaiting: true, names: currentContext.hubOutputNames };
        creationBackend = createCreationMock(currentRequest);
        await click(button("Create modeled data"));
        await until(() => currentContext.hubCreation?.phase === "completed", () => `Mock guarded Create failed: ${currentContext.hubCreation?.error}`);
        assert(currentContext.hubCreation.awaitingData && creationBackend.writes.length === 9 &&
            currentContext.hubCreation.objects.length === 3 && currentContext.hubCoverage.telemetry.payloads.length === 0 &&
            currentContext.hubSourceAnalysis.telemetry.payloads.length > 0 && analysisQueries().length === sourceReads,
            "Waiting, source-cache separation or the guarded native creation contract changed.");
        caseObservations = { unfilteredSourceNotCoverage: true, independentlyCheckedStreams: 2,
            verifiedZeroWithoutOptIn: true, fakeWrites: 9, confirmedFakeObjects: 3, sourceRoleReadsAfterAnalysis: 0, actualCloudWrites: 0 };
    })],
    ["Source Analysis model format defaults to WoT and retains explicit DTDL across Back and stale uploads without source rereads", () => observeWizardContext(async () => {
        const initial = context(true);
        delete initial.hubModel; delete initial.hubSelectedFormat; initial.hubFields = []; initial.preparationFileName = "";
        await wizard(initial);
        const intro = "Upload a model that describes your device data. You'll use it to choose which fields become typed columns.";
        const choice = (value: string) => area.querySelector<HTMLInputElement>(`input[value="${value}"]`)!;
        const guide = (format: string) => area.querySelector(`ul[aria-label="Supported ${format} models"]`);
        const cached = currentContext.hubSourceAnalysis, reads = analysisQueries().length;
        assert(area.querySelector(".iot-solution-step-description")?.textContent === intro &&
            area.querySelector("#model-format-heading")?.textContent === "1. Choose a model format" &&
            choice("wot").checked && !choice("dtdl").checked && area.querySelectorAll('input[type="radio"]').length === 2 &&
            area.querySelector("#model-format-wot-title")?.textContent === "WoT Thing Model (recommended)" &&
            guide("WoT") && !guide("DTDL") && area.querySelector("#thing-model-upload") &&
            modelUploadAction("Choose File") && !modelUploadAction("Replace Model") && next()!.disabled,
            "Fresh model preparation did not default to WoT or bypassed actual model validation.");
        const preparation = area.querySelector('[aria-labelledby="model-preparation-heading"]')!, format = area.querySelector("#model-format-heading")!;
        assert(!!(format.compareDocumentPosition(preparation) & Node.DOCUMENT_POSITION_FOLLOWING) &&
            !!(preparation.compareDocumentPosition(guide("WoT")!) & Node.DOCUMENT_POSITION_FOLLOWING),
            "The WoT helper is not between the format picker and supported-model details.");
        await click(choice("wot"));
        assert(guide("WoT")?.querySelectorAll("li").length === 3 && !guide("DTDL") && modelUploadAction("Choose File") &&
            area.querySelector("#thing-model-upload") && guide("WoT")?.textContent?.includes("an array of related Thing Models") &&
            guide("WoT")?.querySelector("li")?.textContent === "Supported model files: a self-contained WoT 1.1 Thing Model object or an array of related Thing Models. Each model needs its own supplied absolute ID." &&
            guide("WoT")?.textContent?.includes("properties and event data using primitive, enum, and complex schemas.") &&
            guide("WoT")?.querySelector("li:last-child")?.textContent === "Not included: actions, submodels, and external definitions." &&
            guide("WoT")?.querySelector("li:last-child strong")?.textContent === "Not included:" &&
            !guide("WoT")?.textContent?.includes("Local definitions:") &&
            !guide("WoT")?.textContent?.includes("Each declared object or array remains one whole value."),
            "The explicit WoT choice did not expose one ordinary uploader and the requested simplified guide.");
        const fixture = localPositiveFixtures[0], original = JSON.stringify(fixture.document);
        await file("thing-model-upload", original, "explicit-wot.json");
        await until(() => currentContext.hubModel?.id === fixture.rootId, "The explicit local WoT upload did not settle.");
        const accepted = currentContext.hubModel;
        for (const field of currentContext.hubFields as ModelField[]) {
            if (!field.role) await enter(documentElementRole(field.id)!.id, field.column === "history" ? "twin" : "telemetry");
        }
        const beforeReselection = JSON.stringify(currentContext);
        await click(choice("wot"));
        assert(JSON.stringify(currentContext) === beforeReselection && currentContext.hubModel === accepted,
            "Selecting the current format reset the model, roles or existing journey state.");
        await click(next());
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(choice("wot").checked && currentContext.hubModel === accepted && acceptedModelTile() &&
            currentContext.hubSourceAnalysis === cached && analysisQueries().length === reads,
            "Back discarded confirmed format/model intent or the source cache.");
        await click(choice("dtdl"));
        assert(!currentContext.hubModel && currentContext.hubFields.length === 0 && !currentContext.hubCoverage &&
            !currentContext.hubStreamChecks && !currentContext.hubCreation && !currentContext.hubAssignmentChoices &&
            currentContext.hubSourceAnalysis === cached && next()!.disabled, "Switch retained model-derived state or discarded source analysis.");
        const expectedDtdl = [
            "Supported model files: a single DTDL interface, an array of related interfaces, or an IoT Central device template export, including inherited capabilities and components.",
            "Supported capabilities: root, inherited, and component telemetry and device-reported properties using primitive, enum, and complex schemas.",
            "Not included: commands and cloud-only properties, because they are not present in device event data.",
        ];
        assert(JSON.stringify([...guide("DTDL")!.querySelectorAll("li")].map(node => node.textContent)) === JSON.stringify(expectedDtdl) &&
            guide("DTDL")!.querySelectorAll("strong").length === 3 && !guide("WoT"),
            "The DTDL guide differs from the exact original three bold-lead-in bullets.");
        await click(choice("wot"));
        let release: (value: string) => void = () => undefined;
        const candidate = new File(["{}"], "pending-wot.json");
        Object.defineProperty(candidate, "text", { value: () => new Promise<string>(resolve => { release = resolve; }) });
        const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
        const previousFiles = Object.getOwnPropertyDescriptor(input, "files");
        Object.defineProperty(input, "files", { configurable: true, value: [candidate] });
        try {
            await act(async () => input.dispatchEvent(new Event("change", { bubbles: true })));
            assert(area.textContent?.includes("Checking pending-wot.json"), "The pending local parse boundary was not reached.");
            await click(choice("dtdl"));
            await act(async () => release(original)); await settle();
        } finally {
            if (previousFiles) Object.defineProperty(input, "files", previousFiles);
            else Reflect.deleteProperty(input, "files");
        }
        assert(choice("dtdl").checked && !currentContext.hubModel && currentContext.hubFields.length === 0 &&
            !area.textContent?.includes("pending-wot.json") && !area.querySelector('[role="alert"]') &&
            currentContext.hubSourceAnalysis === cached, "A stale parse restored its model, title, error or fields after the format switch.");
        await file("thing-model-upload", original, "looks-like-dtdl.json");
        assert(choice("dtdl").checked && !currentContext.hubModel && next()!.disabled &&
            area.querySelector('[role="alert"]')?.textContent?.includes("doesn't match the format you chose"),
            "Filename/content silently changed the explicit format.");
        await file("thing-model-upload", JSON.stringify(dtdl), "device.json");
        assert(currentContext.hubModel?.kind === "dtdl" && currentContext.hubFields.every((field: ModelField) => field.roleSource === "model") &&
            !next()!.disabled && currentContext.hubSourceAnalysis === cached && analysisQueries().length === reads &&
            !area.querySelector('[role="alert"]'), "Explicit DTDL recovery lost declared roles/cache or reread data.");
        caseObservations = { freshWotDefault: true, sameFormatReselectionNoOp: true, exactIntroAndDtdlBullets: true,
            conditionalGuidesAndFileInput: true, originalWotDocumentRetained: JSON.stringify(accepted.document) === original,
            backRetainsModelChoice: true, formatSwitchRetiresPendingParse: true, noFormatGuess: true,
            sourceCacheRetained: true, additionalSourceReads: 0 };
    })],
    ["Source labels use sentence case and automatic source loads have no healthy Refresh controls", () => observeWizardContext(async () => {
        await mount(<IoTSolutionItemModelWizard workloadClient={client}
            item={{ id: "mock", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
            initialContext={context(true)} onBack={() => undefined}/>);
        await until(tableChoicesReady, "The automatically loaded raw source did not settle.");
        assert(area.querySelector("h2")?.textContent === "Select your data source" &&
            area.querySelector(".iot-solution-step-description")?.textContent ===
              "Choose the eventhouse, KQL database, and tables that hold your raw IoT data. Modeling doesn't change these tables. They keep the complete raw JSON.",
              "The raw source heading or exact user paragraph differs.");
        assert(step("Select data source")?.querySelector(".wizard-control__step-number")?.textContent === "1" &&
            step("Get started")?.textContent?.includes("See what modeling unlocks"),
            "Renaming changed the source step number or requested sentence-case overview description.");
        assert([...area.querySelectorAll("label")].map(node => node.textContent?.replace("*", "")).join("|") ===
            "Eventhouse|KQL database|Raw telemetry table|Raw properties table", "Source labels do not use sentence case.");
        assert(!area.querySelector('button[aria-label^="Refresh"]') && !button("Retry") &&
            !area.textContent?.includes("refresh the list") && !currentContext.modelRawDataSelected && next()!.disabled,
            "Healthy source has a removed Refresh affordance, invalid helper or bypassed analysis.");
        await click(sourceValidationAction()); await until(() => !next()!.disabled, "Explicit source analysis did not settle.");
        const before = calls.length + resourceCalls.length;
        await settle(); await settle();
        assert(calls.length + resourceCalls.length === before, "The healthy source started background polling.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Add a device model", "The model title is not capitalized.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await until(tableChoicesReady, "Back did not automatically revalidate the same source.");
        assert(currentContext.rawTelemetryTable === "Telemetry" && currentContext.rawPropertiesTable === "Twins" &&
            step("Select data source")?.getAttribute("aria-current") === "step", "Back/reset or existing source selection changed.");
        caseObservations = { exactUserWording: true, rawStepIdAndNumberPreserved: true, healthyRefreshControls: 0,
            healthyRetryControls: 0, automaticLoadsAndBackReadiness: true, backgroundPolling: false };
    })],
    ["Source Labels database changes retire expanded menus and ignore a canceled late table response", async () => {
        const items = [{ id: "eh", type: "Eventhouse", displayName: "Eventhouse" },
            { id: "db", type: "KQLDatabase", displayName: "Database" }, { id: "db2", type: "KQLDatabase", displayName: "Database Two" }];
        resourceResponse = async () => new Response(JSON.stringify({ value: items }));
        databaseItemIds = ["db", "db2"];
        await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
        await until(tableChoicesReady, "Initial table choices did not settle.");
        await chooseOption("hub-model-telemetry", "Show All Tables...");
        let release: (value: Response) => void = () => undefined, pending = false;
        tableMetadataResponse = () => { pending = true; return new Promise(resolve => { release = resolve; }); };
        await chooseOption("hub-model-database", "Database Two");
        await until(() => pending, "The new database table read did not enter its pending boundary.");
        assert(!tableChoicesReady() && !currentContext.modelRawDataSelected &&
            !currentContext.rawTelemetryTable && !currentContext.rawPropertiesTable &&
            area.querySelector("#hub-model-telemetry")?.getAttribute("aria-expanded") === "false",
            "A source change retained expanded options, selected old tables or old readiness.");
        tableMetadataResponse = async () => result([["TableName", "string"], ["DocString", "string"]], [
            ["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw"],
            ["Twins", "@iot.meta tableRole=iothub-properties-raw"], ["FreshOnly", null],
        ]);
        await chooseOption("hub-model-database", "Database");
        await until(tableChoicesReady, "The current database table list did not settle.");
        await act(async () => release(result([["TableName", "string"], ["DocString", "string"]], [["StaleOnly", null]])));
        await settle();
        await click(area.querySelector("#hub-model-telemetry"));
        assert(activeOptions("hub-model-telemetry").map(option => option.textContent).join("|") === "Telemetry|Show All Tables...",
            "Returning to a source failed to restore the default role-filtered options.");
        await chooseOption("hub-model-telemetry", "Show All Tables...");
        assert(activeOptions("hub-model-telemetry").some(option => option.textContent === "FreshOnly") &&
            !activeOptions("hub-model-telemetry").some(option => option.textContent === "StaleOnly"),
            "A canceled late inventory published across the database owner change.");
        await chooseOption("hub-model-telemetry", "Telemetry");
        await chooseOption("hub-model-twin", "Twins");
        assert(!currentContext.modelRawDataSelected && currentContext.modelDbId === "db" && !sourceValidationAction()!.disabled,
            "Current metadata selection bypassed analysis or did not recover its action.");
        await click(sourceValidationAction());
        await until(() => currentContext.modelRawDataSelected, "Current deliberate analysis could not recover readiness.");
        caseObservations = { expandedMenuRetiredOnSourceChange: true, oldSelectionsCleared: true,
            canceledLateMetadataIgnored: true, currentOptionsOnly: true, deliberateSelectionRestoresReadiness: true };
    }],
    ["Raw source Show All Tables runtime reproduction with mouse and keyboard", async () => {
        const observations: object[] = [];
        for (const keyboard of [false, true]) for (const id of ["hub-model-telemetry", "hub-model-twin"]) {
            await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
            await until(() => !!area.querySelector<HTMLButtonElement>(`#${id}`) &&
                !area.querySelector<HTMLButtonElement>(`#${id}`)!.disabled, "Source inventory did not settle.");
            const before = JSON.stringify({ telemetry: currentContext.rawTelemetryTable, twin: currentContext.rawPropertiesTable,
                owner: currentContext.hubAssignmentSource, ready: currentContext.modelRawDataSelected });
            const reads = calls.length + resourceCalls.length, errorStart = errors.length;
            const control = area.querySelector<HTMLElement>(`#${id}`)!;
            control.focus(); await settle();
            await click(control);
            const sentinel = activeOptions(id).find(option => /show all tables/i.test(option.textContent || ""))!;
            assert(sentinel, "The role-filtered menu did not expose Show All Tables.");
            if (keyboard) {
                for (const [key, code] of [["End", 35], ["Enter", 13]] as const) {
                    await act(async () => control.dispatchEvent(new KeyboardEvent("keydown",
                        { key, code: key, keyCode: code, which: code, bubbles: true, cancelable: true })));
                    await settle();
                }
            } else await click(sentinel);
            const remainedOpen = control.getAttribute("aria-expanded") === "true";
            assert(remainedOpen, "Show All Tables closed the menu during keyboard selection.");
            await act(async () => control.dispatchEvent(new KeyboardEvent("keydown",
                { key: "ArrowDown", code: "ArrowDown", keyCode: 40, which: 40, bubbles: true, cancelable: true })));
            await settle();
            assert(errors.length === errorStart, `Show All Tables emitted a runtime error: ${errors.slice(errorStart).join("\n")}`);
            assert(activeOptions(id).some(option => option.textContent === "NewTelemetry"), "The menu did not reveal the manual table.");
            assert(JSON.stringify({ telemetry: currentContext.rawTelemetryTable, twin: currentContext.rawPropertiesTable,
                owner: currentContext.hubAssignmentSource, ready: currentContext.modelRawDataSelected }) === before,
                "Show All Tables changed the actual selection, owner or readiness.");
            assert(calls.length + resourceCalls.length === reads, "Show All Tables issued a background resource read.");
            await act(async () => control.dispatchEvent(new KeyboardEvent("keydown",
                { key: "Escape", code: "Escape", bubbles: true, cancelable: true })));
            await settle();
            assert(control.getAttribute("aria-expanded") === "false", "The expansion's close guard leaked into the next Escape event.");
            observations.push({ id, input: keyboard ? "keyboard End/Enter" : "mouse click", runtimeErrors: 0,
                remainedOpen, unchangedSelectionOwnerReadiness: true, extraReads: 0 });
        }
        caseObservations = { observations };
    }],
    ["Actual Hub root creation uses full V1 presentation and exact sidebar captions for reference AND and OR with properties", () => observeWizardContext(async () => {
        const discover = HubAssignmentApi.prototype.discover;
        try {
            for (const operator of ["and", "or"] as const) {
                sourceHubFixture("linked");
                HubAssignmentApi.prototype.discover = async () => assignmentInventory();
                await wizard(context(true)); await click(next()); await beginAssignmentTargets();
                await until(() => !!currentContext.hubAssignmentHubSummary?.value.complete && !next()!.disabled,
                    "Actual root assignment/source evidence did not settle.");
                await click(next());
                await click(area.querySelector(`input[value="${operator}"]`));
                await click(area.querySelector('input[value="specified"]'));
                const predicate = 'tostring(headers["IoTConnectionDeviceId"]) == "device-1"';
                await enter("hub-telemetry-scope", predicate);
                await verifyCoverage(); await click(next());
                const content = area.querySelector(".wizard-control__content-panel")!;
                const presentation = content.cloneNode(true) as HTMLElement;
                assert(content.querySelector("h2")?.textContent === "Create modeled data" &&
                    !presentation.textContent?.includes("Processing Scope") && !presentation.textContent?.includes("How Reported Context Is Added") &&
                    !presentation.textContent?.includes("Existing raw history isn't backfilled") &&
                    !presentation.textContent?.includes("latest whole received value") && !content.querySelector(".iot-solution-choice-summary") &&
                    content.querySelectorAll(".iot-solution-validation-row").length === 3 &&
                    content.querySelectorAll(".iot-solution-form input[required]").length === 3 &&
                    !!content.querySelector("#output-names-help") && !button("Create modeled data")!.disabled,
                    "Actual root still composed V2 review-only content around the V1 creation UI.");
                assert(step("Create modeled data")?.textContent?.includes("Build typed telemetry with device state") &&
                    step("Review results")?.textContent?.includes("See what was created and next steps") &&
                    currentContext.hubScope.mode === "reference" && currentContext.hubScope.operator === operator &&
                    currentContext.hubScope.telemetry === predicate && !!currentContext.hubCoverage &&
                    currentContext.hubFields.some((field: ModelField) => field.included && field.role === "twin"),
                    "Root captions or actual preserved scoped/property inputs changed.");
                if (operator === "and") {
                    const source = (observedWizardProps as React.ComponentProps<typeof HubModelReviewStep>).sourceState;
                    creationBackend = createCreationMock({ input: hubInput(currentContext, source), coverage: currentContext.hubCoverage,
                        names: currentContext.hubOutputNames, allowWaiting: true });
                    creationBackend.options.documentationMismatch = true;
                    await click(button("Create modeled data"));
                    await until(() => currentContext.hubCreation?.phase === "failed", "Actual-root metadata readback did not fail.");
                    const metadata = [...content.querySelectorAll<HTMLElement>(".iot-solution-creation-step")]
                        .find(row => row.textContent?.includes(`Set docstring on ${currentContext.hubOutputNames.modeled}`))!;
                    assert(metadata.dataset.status === "failed" && creationBackend.writes.length === 9 &&
                        !content.querySelector("#modeled-data-success") && next()!.disabled &&
                        metadata.textContent?.includes("Native table/column documentation could not be verified"),
                        "Grouped metadata success hid a real column-docstring/readback failure or allowed Ready.");
                    assertCreationDisplayVerbs(buildModeledDataPlan(hubInput(currentContext, source),
                        currentContext.hubCoverage, true, currentContext.hubOutputNames));
                    creationBackend = undefined;
                }
            }
        } finally { HubAssignmentApi.prototype.discover = discover; creationBackend = undefined; }
        caseObservations = { actualRootEntrypoint: true, referenceOperators: ["and", "or"], propertiesIncluded: true,
            obsoleteSectionsAndNarrativeDeleted: true, positiveV1SummaryNamesAndCreate: true, exactTwoSidebarCaptions: true,
            metadataReadbackFailureBlocksReady: true, syntheticWriteCommands: 9, logicalRows: 8, actualWrites: 0 };
    })],
    ["Assignment assigned-other row omits routine read-only prose but keeps checkbox eligibility and write blockers", async () => {
        const discover = HubAssignmentApi.prototype.discover;
        try {
            const value = deviceListInventory();
            HubAssignmentApi.prototype.discover = async () => value;
            await openAssignmentTargets(); await chooseAssignmentView("all");
            const assigned = value.targets.find(target => target.name.startsWith("Existing-reference-device"))!;
            const row = [...area.querySelectorAll("tbody tr")].find(node => node.textContent?.includes(assigned.name))!;
            assert(row && !row.querySelector('input[type="checkbox"]') &&
                !area.textContent?.includes("Existing model assignments are read-only.") &&
                row.textContent?.includes(assigned.reference.kind === "assigned" ? assigned.reference.ids[0] : "") &&
                assignmentBlocker(assigned) === "Existing model assignments are read-only.",
                "Routine prose remains or removing it changed assigned-reference read-only eligibility.");
            await click(button("Select all"));
            assert(currentContext.hubAssignmentChoices.targets.every((target: AssignmentTarget) =>
                target.reference.kind === "unassigned" && !assignmentBlocker(target)),
                "Select all bypassed the existing-reference/unknown write blocker.");
            caseObservations = { sentenceAbsent: true, assignedCheckboxAbsent: true, exactReferenceRetained: true,
                coreWriteBlockerUnchanged: true, eligibleOnlySelection: true, actualWrites: 0 };
        } finally { HubAssignmentApi.prototype.discover = discover; }
    }],
    ["Device Assignment success requires every current attempt receipt and keeps accepted or partial attempts actionable", async () => {
        const discover = HubAssignmentApi.prototype.discover, services = HubAssignmentApi.prototype.assignmentServices,
            readback = HubAssignmentApi.prototype.readAssignment;
        const clock = assignmentClock();
        const observations: object[] = [];
        try {
            for (const outcome of ["completed", "accepted", "partial"] as const) {
                const value = assignmentInventory(), stored = new Set<string>(), writes: string[] = [];
                value.targets = [0, 1, 2].map(index => ({ ...value.targets[0], id: value.targets[0].id + `-${index}`,
                    name: index < 2 ? `Selected device ${index + 1}` :
                        "Selected device 3 with a longer label that wraps naturally without changing the icon and text alignment",
                    writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
                    deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: `device-${index}` } }));
                const polled: string[] = [];
                const pollCounts = new Map<string, number>();
                HubAssignmentApi.prototype.readAssignment = async target => {
                    polled.push(target.id);
                    const count = (pollCounts.get(target.id) || 0) + 1; pollCounts.set(target.id, count);
                    const confirmed = !target.id.endsWith("-2") || count >= 2;
                    return { target: { ...target, reference: confirmed ? { kind: "assigned", ids: [tm.id] } : target.reference }, confirmed };
                };
                HubAssignmentApi.prototype.discover = async () => ({ ...value, targets: value.targets.map(target => ({
                    ...target, reference: stored.has(target.id) ? { kind: "assigned", ids: [tm.id] } : { kind: "unassigned", ids: [] },
                })) });
                HubAssignmentApi.prototype.assignmentServices = () => ({
                    read: async target => ({ ...target, reference: stored.has(target.id)
                        ? { kind: "assigned", ids: [tm.id] } : { kind: "unassigned", ids: [] } }),
                    write: async (target, _id, _signal, dispatch) => {
                        dispatch(); writes.push(target.id);
                        if (outcome === "partial" && target.id.endsWith("-2")) throw Object.assign(new Error("Owned rejected device."), { status: 403 });
                        stored.add(target.id);
                        return outcome === "accepted" ? "accepted" : "completed";
                    },
                });
                const seed = assignmentContext();
                seed.hubAssignmentHistory = value.targets.map(target => ({ operationId: "older-success", targetId: target.id,
                    name: target.name, kind: "device", modelId: tm.id, state: "confirmed" }));
                await openAssignmentTargets(seed);
                await until(() => !!button("Select all") && !button("Select all")!.disabled, "Current operation devices did not load.");
                assert(area.querySelector("#assignment-devices-heading")?.textContent === "Select devices",
                    "The pre-assignment inventory heading is not Select devices.");
                assertDpsGuidance(false);
                assert(!area.textContent?.includes("assigned successfully"), "Old confirmed history satisfied a new attempt.");
                await click(button("Select all")); await click(button("Assign to selected devices"));
                assert(writes.length === 0, "Opening confirmation wrote a reference.");
                await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')]
                    .find(button => button.textContent === "Assign"));
                await until(() => !currentContext.hubAssignmentBusy && writes.length === 3, "The explicit current attempt did not settle.");
                if (outcome === "accepted") {
                    const progress = area.querySelector('[aria-label="Device assignment progress"]')!;
                    const aggregate = progress.querySelector<HTMLElement>('[role="status"]')!;
                    const progressRows = [...progress.querySelectorAll<HTMLElement>(".iot-solution-creation-step")];
                    const progressList = progress.querySelector(".iot-solution-creation-steps");
                    assert(progress.querySelector("h3")?.textContent === "Assigned devices",
                        "Loading/progress changed the section heading.");
                    assert(getComputedStyle(progress).borderTopWidth === "0px", "Progress retained a section divider.");
                    assert(!area.querySelector("table") && !document.querySelector('[role="dialog"]') &&
                        !existingAssignmentsNotice() &&
                        !button("Select all") && !button("Assign to selected devices") &&
                        area.querySelector('[aria-label="Device assignment progress"]') &&
                        currentContext.hubAssignmentHistory.some((receipt: { state: string }) => receipt.state === "accepted") &&
                        !area.textContent?.includes("assigned successfully") &&
                        confirmedDeviceRows() === 0 && polled.length === 0,
                        "Accepted 202 was treated as confirmed assignment success.");
                    assert(aggregate.textContent?.trim() === "Assigning the model to 3 devices..." &&
                        aggregate.getAttribute("aria-hidden") !== "true" &&
                        !progress.querySelector(":scope > .fui-Spinner") && progressRows.length === 3 &&
                        progressRows.every((row, index) => row.textContent === value.targets[index].name &&
                            row.querySelector('[role="progressbar"]')?.getAttribute("aria-label") ===
                            `Checking assignment for ${value.targets[index].name}`),
                        "Routine rows retain status logs or lost accessible spinner/name-only rendering.");
                    assertAssignmentSummary("active", 3);
                    assertDpsGuidance(false);
                    await until(() => clock.pending.size === 1, "The default readback interval was not scheduled.");
                    await clock.tick();
                    await until(() => confirmedDeviceRows() === 2, "Confirmed devices were not shown incrementally.");
                    assertAssignmentSummary("active", 3);
                    assert(area.querySelectorAll('[data-assignment-state="confirmed"] svg[aria-label="Assignment confirmed"]').length === 2 &&
                        clock.pending.size === 1 && polled.length === 3 && !area.textContent?.includes("assigned successfully"),
                        "Incremental native green checks, pending-only scheduling or completion proof changed.");
                    assert([...progress.querySelectorAll(".iot-solution-creation-step")].every((row, index) =>
                        row.textContent === value.targets[index].name) &&
                        !progress.textContent?.includes("Reference confirmed") &&
                        !progress.textContent?.includes("Matching data has not been verified"),
                        "Confirmed rows retained routine receipt paragraphs instead of only check and name.");
                    assertDpsGuidance(false);
                    await clock.tick();
                    await until(() => !!completedAssignmentList(), "Automatic current readback did not confirm every accepted reference.");
                    assert(progress.querySelector(".iot-solution-creation-steps") === progressList &&
                        progressRows.every((row, index) => row === progress.querySelectorAll(".iot-solution-creation-step")[index]),
                        "Completion replaced the mounted device list/rows instead of keeping their native green checks.");
                    assert(polled.length === 4 && pollCounts.get(value.targets[0].id) === 1 &&
                        pollCounts.get(value.targets[1].id) === 1 && clock.pending.size === 0 &&
                        clock.delays.every(delay => delay >= 1900 && delay <= 4000), "Confirmed devices were polled again or cadence changed.");
                }
                if (outcome === "partial") {
                    assert(!area.querySelector("table") && area.querySelector('[aria-label="Device assignment progress"]') &&
                        !existingAssignmentsNotice() &&
                        area.textContent?.includes("Owned rejected device.") && !area.textContent?.includes("assigned successfully"),
                        "One failed current device was hidden behind old or partial success.");
                    assertDpsGuidance(false);
                    assertAssignmentSummary("incomplete", 3);
                } else assert(!area.querySelector("table") && !button("Refresh results") &&
                    !button("Assign to selected devices") && !button("Select all") &&
                    completedAssignmentList(),
                    "Fully confirmed current attempt retained table/controls or lost its exact success count.");
                if (outcome !== "partial") assertCompletedAssignments(value.targets.map(target => target.name));
                assert(writes.length === 3, "Verification replayed a device write.");
                observations.push({ outcome, explicitSyntheticWrites: writes.length, currentAttemptOnly: true,
                    geometry: assignmentProgressGeometry() });
            }
            caseObservations = { observations, actualWrites: 0, oldSuccessIgnored: true };
        } finally {
            await mount(<div>Polling fixture closed.</div>); clock.restore();
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.assignmentServices = services;
            HubAssignmentApi.prototype.readAssignment = readback;
        }
    }],
    ["Assignment read-only Sign in and Refresh results restore an unknown spinner before final outcome", () => observeWizardContext(async () => {
        const clock = assignmentWaveClock(), originalRandom = Math.random, observations: object[] = [];
        Math.random = () => 0.5;
        try {
            for (const kind of ["authentication", "permission"] as const) {
                const requests: string[] = [], assigned = { patches: [] as { body: unknown; ifMatch: string | null }[] };
                const respond = adrOnlyResponse(requests, assigned);
                let writes = 0, statusReads = 0;
                let releaseRecheck: (() => void) | undefined;
                assignmentResponse = async (input, init) => {
                    if (init?.method === "PATCH") {
                        writes++;
                        return new Response(null, { status: 202, headers: { Location:
                            `https://management.azure.com/subscriptions/${new URL(String(input)).pathname.split("/")[2]}/providers/Microsoft.DeviceRegistry/locations/eastus/operationStatuses/recheck-${kind}?api-version=2026-11-02-preview` } });
                    }
                    if (new URL(String(input)).pathname.includes("/operationStatuses/recheck-")) {
                        statusReads++;
                        if (statusReads === 2) return new Promise(resolve => {
                            releaseRecheck = () => resolve(new Response(JSON.stringify({ status: "Running" })));
                        });
                        return statusReads === 1 ? new Response("", { status: kind === "authentication" ? 401 : 403 })
                            : new Response(JSON.stringify({ status: "Running" }));
                    }
                    return respond(input, init);
                };
                await wizard(context(true)); await click(next()); await beginAssignmentTargets();
                await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') &&
                    !area.querySelector('[aria-label="Device model assignments"] [role="progressbar"]'),
                    "The read-only retry candidate did not finish current canonical discovery.");
                await click(area.querySelector('input[aria-label="Select Fixture device"]'));
                await click(button("Assign to selected devices"));
                await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(node => node.textContent === "Assign"));
                await until(() => writes === 1 && clock.pending.size > 0, "Original accepted-operation observation did not start.");
                await clock.advance(3000);
                const actionName = kind === "authentication" ? "Sign in" : "Refresh results";
                await until(() => !!button(actionName) && !button(actionName)!.disabled &&
                    area.querySelector('[aria-label="Device assignment progress"]')?.getAttribute("aria-busy") === "false",
                    "The failed original-operation read did not expose its existing explicit recovery action.");
                const first = currentContext.hubAssignmentHistory.find((receipt: AssignmentReceipt) => receipt.state === "failed") as AssignmentReceipt;
                assert(first?.readbackFailure?.kind === kind && statusReads === 1 && writes === 1,
                    "The retry fixture did not stop on its original typed read failure.");
                const deadline = first.readbackDeadline, beforeTokens = tokenRequests;
                await click(button(actionName));
                const activeRow = () => area.querySelector<HTMLElement>('[data-assignment-state="unknown"]')!;
                const assertChecking = () => {
                    const row = activeRow(), label = row?.querySelector<HTMLElement>(".iot-solution-creation-step-label");
                    assert(row?.querySelector('[role="progressbar"]')?.getAttribute("aria-label") === "Checking assignment for Fixture device" &&
                        !row.querySelector("svg[data-assignment-warning],svg[aria-label='Assignment not confirmed']") &&
                        label?.classList.contains("running") && !label.classList.contains("warning") && !label.classList.contains("error") &&
                        getComputedStyle(label).color === getComputedStyle(
                            area.querySelector(".iot-solution-assignment-summary-label.running")!).color &&
                        row.textContent === "Fixture device" && next()!.disabled &&
                        area.querySelector('[aria-label="Device assignment progress"]')?.getAttribute("aria-busy") === "true",
                        "A rechecking unknown receipt must show the original spinner/running tone without warning or stale error details.");
                    assert(currentContext.hubAssignmentHistory.some((receipt: AssignmentReceipt) =>
                        receipt.operationId === first.operationId && receipt.state === "unknown" &&
                        receipt.readbackDeadline === deadline && receipt.readbackFailure === undefined),
                        "Explicit read recovery changed the operation/deadline instead of clearing only the failed read.");
                    assignmentProgressGeometry();
                };
                assertChecking();
                assert(statusReads === 1 && writes === 1, "Read recovery bypassed the original scheduled wait or replayed the mutation.");
                await clock.advance(3000); assertChecking();
                assert(statusReads === 2 && !!releaseRecheck && tokenRequests > beforeTokens && writes === 1,
                    "Sign in/Refresh did not reauthorize and read only the original operation.");
                await act(async () => releaseRecheck!()); await settle(); assertChecking();
                await clock.advance(deadline! - Date.now() - 1); assertChecking();
                await clock.advance(1);
                await until(() => !next()!.disabled && clock.pending.size === 0, "The original polling deadline was extended by explicit recovery.");
                const finalRow = activeRow(), finalReceipt = currentContext.hubAssignmentHistory.find(
                    (receipt: AssignmentReceipt) => receipt.operationId === first.operationId) as AssignmentReceipt;
                assert(finalRow.querySelector("svg[data-assignment-warning]") && !finalRow.querySelector('[role="progressbar"]') &&
                    finalRow.textContent?.includes(ASSIGNMENT_OBSERVATION_TIMEOUT_MESSAGE) &&
                    finalReceipt.state === "unknown" && finalReceipt.readbackFailure?.retryable === false &&
                    finalReceipt.readbackDeadline === deadline && writes === 1 && !button("Refresh results") && !button("Sign in") &&
                    area.textContent?.includes("Not all assignments are confirmed. Review the device results before continuing.") &&
                    !area.querySelector("[data-assignment-summary]"),
                    "A final unknown result lost yellow warning guidance or gained success/replay/recovery after expiry.");
                assertAssignmentSemanticIndicators();
                observations.push({ kind, originalDeadlineRetained: true, activeUnknownSpinner: true, finalUnknownYellow: true,
                    conditionalWrites: writes, originalStatusReads: statusReads, extraMutationRetries: 0 });
                await mount(<div>Read recovery fixture closed.</div>);
            }
            caseObservations = { observations, actualCloudWrites: 0 };
        } finally {
            await mount(<div>Read recovery closed.</div>); clock.restore();
            Math.random = originalRandom; assignmentResponse = undefined;
        }
    })],
    ["Assignment native PATCH throttling requires a fresh explicit attempt and never retries or observes the rejected write", () => observeWizardContext(async () => {
        const requests: { url: string; method: string }[] = [], armRequests: string[] = [];
        const assigned = { patches: [] as { body: unknown; ifMatch: string | null }[] }, respond = adrOnlyResponse(armRequests, assigned);
        let writes = 0, statusReads = 0, currentEtag = '"fixture-current"';
        const ifMatches: (string | null)[] = [];
        assignmentResponse = async (input, init) => {
            const url = String(input), resource = new URL(url).pathname, method = init?.method || "GET";
            requests.push({ url, method });
            if (method === "PATCH") {
                writes++; ifMatches.push(new Headers(init?.headers).get("If-Match"));
                assert(ifMatches.at(-1) === currentEtag &&
                    JSON.stringify(JSON.parse(String(init?.body))) === JSON.stringify({ properties: { deviceTypeRefs: [tm.id] } }),
                    "A manually reconfirmed assignment lost fresh specific-ETag or refs-only mutation guards.");
                return writes === 1 ? new Response("", { status: 429, headers: { "Retry-After": "3600" } })
                    : new Response(null, { status: 204 });
            }
            if (resource.includes("/operationStatuses/")) {
                statusReads++; throw new Error("A rejected or synchronous assignment has no operation to poll.");
            }
            const result = await respond(input, init);
            if (resource.endsWith("/registryDevices/fixture-device")) {
                const body = await result.json();
                return new Response(JSON.stringify({ ...body, etag: currentEtag }));
            }
            return result;
        };
        try {
            await wizard(context(true)); await click(next()); await beginAssignmentTargets();
            await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') &&
                !area.querySelector('[aria-label="Device model assignments"] [role="progressbar"]'), "The throttled write candidate did not settle.");
            await click(area.querySelector('input[aria-label="Select Fixture device"]'));
            await click(button("Assign to selected devices"));
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(node => node.textContent === "Assign"));
            await until(() => !currentContext.hubAssignmentBusy && !next()!.disabled, "The definite write refusal did not settle.");
            const failed = currentContext.hubAssignmentHistory.at(-1) as AssignmentReceipt;
            assert(failed.state === "failed" && failed.writeDispatched && failed.trackingId === undefined &&
                failed.readbackDeadline === undefined && failed.readbackFailure === undefined &&
                area.querySelector('[data-assignment-state="failed"] .iot-solution-creation-step-label.error') &&
                area.textContent?.includes("The Azure request failed (429).") &&
                !area.querySelector('[data-assignment-state="unknown"],[data-assignment-warning]') &&
                !button("Refresh results") && statusReads === 0 && writes === 1,
                "A definite initial PATCH429 became an unknown operation, acquired a poll or automatic retry.");
            assertAssignmentSemanticIndicators();
            const settledRequests = requests.length;
            await settle(); await settle();
            assert(requests.length === settledRequests && writes === 1, "A rejected mutation was automatically retried.");
            await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
            await file("thing-model-upload", JSON.stringify(tm), "explicit-model-reconfirmation.json");
            await until(() => !!currentContext.hubModel && !next()!.disabled, "The existing explicit model journey did not recover selection.");
            currentEtag = '"fresh-attempt-etag"';
            const beforeFreshDiscovery = requests.length;
            await click(next()); await beginAssignmentTargets();
            await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') &&
                !area.querySelector('[aria-label="Device model assignments"] [role="progressbar"]'),
                "The explicit new attempt did not rediscover a currently unassigned eligible device.");
            assert(writes === 1 && currentContext.hubAssignmentHistory.some((receipt: AssignmentReceipt) =>
                receipt.operationId === failed.operationId && receipt.state === "failed"),
                "Starting fresh selection erased the old refusal or resent it before consent.");
            const fresh = currentContext.hubAssignmentInventory.value.targets.find((target: AssignmentTarget) => target.name === "Fixture device") as AssignmentTarget;
            assert(fresh.reference.kind === "unassigned" && fresh.etag === currentEtag &&
                fresh.referenceEvidence?.currentArmEtag === currentEtag && fresh.updatePermission === "granted" &&
                !!fresh.deviceIdentity && !assignmentBlocker(fresh), "The second selection did not use fresh canonical identity/reference/access evidence.");
            await click(area.querySelector('input[aria-label="Select Fixture device"]')); await click(button("Assign to selected devices"));
            assert(writes === 1, "Selecting a target or opening confirmation resent a previously rejected write.");
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(node => node.textContent === "Cancel"));
            assert(writes === 1, "Cancel consent dispatched an assignment.");
            await click(button("Assign to selected devices"));
            const beforeSecondWrite = requests.length;
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(node => node.textContent === "Assign"));
            await until(() => !!completedAssignmentList() && !next()!.disabled, "The explicitly reconfirmed fresh attempt did not complete.");
            const confirmed = currentContext.hubAssignmentHistory.find((receipt: AssignmentReceipt) => receipt.state === "confirmed") as AssignmentReceipt;
            assert(writes === 2 && statusReads === 0 && confirmed.operationId !== failed.operationId &&
                ifMatches.join("|") === '"fixture-current"|"fresh-attempt-etag"' &&
                currentContext.hubAssignmentHistory.some((receipt: AssignmentReceipt) => receipt.operationId === failed.operationId && receipt.state === "failed"),
                "The explicit second attempt lost fresh If-Match, preserved refusal history or distinct attempt identity.");
            const secondPrechecks = requests.slice(beforeSecondWrite).filter(request => request.method === "GET");
            for (const suffix of ["/IotHubs/fixture-hub", "/namespaces/fixture-namespace", "/registryDevices/fixture-device",
                "/registryDevices/fixture-device/capabilities", "/registryDevices/fixture-device/providers/Microsoft.Authorization/permissions"])
                assert(secondPrechecks.some(request => new URL(request.url).pathname.endsWith(suffix)),
                    `The explicit second mutation skipped a current canonical precheck: ${suffix}`);
            assert(requests.length > beforeFreshDiscovery && writes === 2, "No fresh source checks occurred before new explicit consent.");
            assertCompletedAssignments(["Fixture device"]);
            caseObservations = { initial429State: "failed", actualHttpDispatchRetained: true, rejectedWritePolls: 0,
                automaticMutationRetries: 0, explicitAttempts: 2, distinctOperationIds: true,
                freshCanonicalEtags: ifMatches, oldRefusalHistoryRetained: true, cancelConsentNoWrite: true,
                repeatedSourceIdentityReferencePermissionPreflight: true, actualCloudWrites: 0 };
        } finally { await mount(<div>Explicit throttled attempt closed.</div>); assignmentResponse = undefined; }
    })],
    ["Assignment Polling keeps transient and terminal read failures truthful with no overlap and read-only recovery", async () => {
        const discover = HubAssignmentApi.prototype.discover, services = HubAssignmentApi.prototype.assignmentServices,
            readback = HubAssignmentApi.prototype.readAssignment;
        const clock = assignmentClock();
        const value = assignmentInventory(), calls: string[] = [], writes: string[] = [];
        value.targets = [0, 1].map(index => ({ ...value.targets[0], id: value.targets[0].id + `-${index}`, name: `Device ${index + 1}`,
            writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
            deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: `device-${index}` } }));
        const handles = new Map(value.targets.map(target => [target.id, crypto.randomUUID()]));
        let recovering = false, release: (() => void) | undefined;
        const attempts = new Map<string, number>();
        try {
            HubAssignmentApi.prototype.discover = async () => value;
            HubAssignmentApi.prototype.assignmentServices = () => ({
                read: async target => target, write: async (target, _id, _signal, dispatch) => {
                    dispatch(); writes.push(target.id); return "accepted";
                },
                trackingId: target => handles.get(target.id),
            });
            HubAssignmentApi.prototype.readAssignment = async (target, _modelId, _signal, trackingId) => {
                assert(trackingId === handles.get(target.id), "Read-only recovery changed the original operation handle.");
                calls.push(target.id);
                const count = (attempts.get(target.id) || 0) + 1; attempts.set(target.id, count);
                if (target.id.endsWith("-1") && !recovering) throw new AssignmentApiError("permission", "Owned read access denied.");
                if (target.id.endsWith("-0") && count === 1) throw new AssignmentApiError("request", "Owned transient read failure.", true);
                if (target.id.endsWith("-0")) await new Promise<void>(resolve => { release = resolve; });
                return { target: { ...target, reference: { kind: "assigned", ids: [tm.id] } }, confirmed: true };
            };
            await openAssignmentTargets(); await until(() => !!button("Select all") && !button("Select all")!.disabled, "Failure devices did not load.");
            await click(button("Select all")); await click(button("Assign to selected devices"));
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(button => button.textContent === "Assign"));
            await until(() => clock.pending.size === 1 && writes.length === 2, "Owned accepted attempt did not start result polling.");
            await clock.tick();
            await until(() => area.textContent?.includes("Owned transient read failure.") && area.textContent?.includes("Owned read access denied."),
                "Transient or terminal read failure was hidden.");
            assert(area.querySelectorAll('[data-assignment-state="unknown"] svg[data-assignment-warning]').length === 1 &&
                area.querySelectorAll('[data-assignment-state="failed"] svg[aria-label="Assignment not confirmed"]:not([data-assignment-warning])').length === 1,
                "An uncertain retry outcome lost its native warning or a genuine terminal failure stopped being red.");
            assertAssignmentSemanticIndicators();
            assignmentProgressGeometry();
            assert(clock.pending.size === 1 && calls.length === 2 && !area.querySelector("table") &&
                !area.textContent?.includes("assigned successfully"), "Read error became success or the table returned.");
            assertDpsGuidance(false);
            await clock.tick(); await until(() => !!release, "Second transient verification did not enter its deferred read.");
            assert(clock.pending.size === 0 && calls.length === 3 && button("Refresh results")?.disabled,
                "An overlapping poll or manual request was allowed during the current read.");
            await act(async () => release!());
            await until(() => confirmedDeviceRows() === 1 && !button("Refresh results")!.disabled,
                "Terminal partial result did not preserve confirmed progress or explicit read-only recovery.");
            recovering = true; await click(button("Refresh results"));
            await until(() => clock.pending.size === 1, "Explicit terminal verification retry was not scheduled.");
            await clock.tick();
            await until(() => !!completedAssignmentList(), "Read-only recovery did not confirm the remaining device.");
            assert(writes.length === 2 && calls.length === 4 && attempts.get(value.targets[0].id) === 2 &&
                attempts.get(value.targets[1].id) === 2, "Recovery repeated writes or polled an already confirmed device.");
            assertDpsGuidance(true);
            caseObservations = { defaultInterval: 3000, transientVisible: true, terminalStopped: true, noOverlap: true,
                readOnlyRecovery: true, syntheticWrites: 2, actualWrites: 0 };
        } finally {
            await mount(<div>Failure fixture closed.</div>); clock.restore();
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.assignmentServices = services;
            HubAssignmentApi.prototype.readAssignment = readback;
        }
    }],
    ["Assignment Polling resumes only current accepted reads on remount and discards late source ABA completions", async () => {
        const discover = HubAssignmentApi.prototype.discover, services = HubAssignmentApi.prototype.assignmentServices,
            readback = HubAssignmentApi.prototype.readAssignment;
        const clock = assignmentClock();
        const value = assignmentInventory();
        value.targets[0] = { ...value.targets[0], writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
            deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: "fixture-device" } };
        const trackingId = crypto.randomUUID();
        let writes = 0, discoveries = 0, reads = 0, release: (() => void) | undefined, aborted = false;
        try {
            HubAssignmentApi.prototype.discover = async () => { discoveries++; return value; };
            HubAssignmentApi.prototype.assignmentServices = () => ({
                read: async target => target, write: async (_target, _id, _signal, dispatch) => { dispatch(); writes++; return "accepted"; },
                trackingId: () => trackingId,
            });
            HubAssignmentApi.prototype.readAssignment = (target, _id, signal, handle) => {
                assert(handle === trackingId, "A remount did not retain the original captured operation handle.");
                reads++; signal.addEventListener("abort", () => { aborted = true; }, { once: true });
                return new Promise(resolve => { release = () => resolve({
                    target: { ...target, reference: { kind: "assigned", ids: [tm.id] } }, confirmed: true,
                }); });
            };
            await openAssignmentTargets(); await until(() => !!area.querySelector("input[type='checkbox']"), "Resume device did not load.");
            await click(button("Select all")); await click(button("Assign to selected devices"));
            await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(button => button.textContent === "Assign"));
            await until(() => clock.pending.size === 1 && writes === 1, "Accepted resume receipt did not settle.");
            const attempt = currentContext.hubAssignmentIntent.attempt.operationId;
            await act(async () => setStepVisible(false)); assert(clock.pending.size === 0, "Unmount leaked its pending poll timer.");
            await act(async () => setStepVisible(true));
            await until(() => clock.pending.size === 1 && !!area.querySelector('[aria-label="Device assignment progress"]'), "Current accepted operation did not resume read-only verification.");
            assert(writes === 1 && discoveries === 1 && currentContext.hubAssignmentIntent.attempt.operationId === attempt,
                "Resume replayed assignment/discovery or changed the owned attempt.");
            await clock.tick(); await until(() => !!release, "Resumed read did not enter its deferred boundary.");
            await act(async () => setTestSourceState({ ...sourceFixture, key: "another-source", revision: 2 }));
            assert(aborted && !area.querySelector('[aria-label="Device assignment progress"]') && !currentContext.hubAssignmentIntent,
                "Source change retained the old operation or failed to abort its read.");
            await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
            await act(async () => release!()); await settle();
            assert(clock.pending.size === 0 && writes === 1 && reads === 1 && !area.textContent?.includes("assigned successfully") &&
                !area.querySelector('[aria-label="DPS enrollment guidance"]') && !currentContext.hubAssignmentIntent &&
                currentContext.hubAssignmentHistory.some((receipt: { operationId: string; state: string }) =>
                    receipt.operationId === attempt && receipt.state === "accepted"),
                "Late canceled read restored success/intent or erased the truthful accepted history.");
            caseObservations = { remountReadOnly: true, noReplay: true, noDiscoveryOnResume: true, timerCleared: true,
                sourceAbaRetired: true, lateCompletionIgnored: true, retainedAcceptedHistory: true };
        } finally {
            await mount(<div>Resume fixture closed.</div>); clock.restore();
            HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.assignmentServices = services;
            HubAssignmentApi.prototype.readAssignment = readback;
        }
    }],
    ["Device Assignment initial no-unassigned state requires complete current existing references across any model IDs", async () => {
        const discover = HubAssignmentApi.prototype.discover;
        try {
            const value = assignmentInventory();
            value.targets = [0, 1].map(index => ({ ...value.targets[0], id: value.targets[0].id + `-${index}`,
                name: `Already assigned device ${index + 1}`, writeBlocker: undefined, reference: { kind: "assigned", ids: [tm.id] } }));
            const variants: [string, AssignmentInventory, boolean][] = [
                ["complete", value, true],
                ["mixed-known-unlinked", { ...value, hubs: [...value.hubs,
                    { id: value.hubs[0].id + "-unlinked", name: "Known nonlinked Hub", state: "unlinked" }] }, true],
                ["pending-hub", { ...value, hubs: [...value.hubs,
                    { id: value.hubs[0].id + "-pending", name: "Pending Hub", state: "pending" }] }, false],
                ["unknown-hub", { ...value, hubs: [...value.hubs,
                    { id: value.hubs[0].id + "-unknown", name: "Unknown Hub", state: "unverified" }] }, false],
                ["has-more", { ...value, hasMore: true }, false],
                ["empty", { ...value, targets: [] }, false],
                ["other-model", { ...value, targets: value.targets.map(target => ({
                    ...target, updatePermission: "unknown", reference: { kind: "assigned", ids: ["urn:other"] },
                })) }, true],
                ["mixed-models", { ...value, targets: [value.targets[0],
                    { ...value.targets[1], reference: { kind: "assigned", ids: ["urn:other"] } }] }, true],
                ["unknown", { ...value, targets: [{ ...value.targets[0], reference: { kind: "unknown", reason: "Unknown assignment" } }] }, false],
                ["invalid-reference", { ...value, targets: [{ ...value.targets[0], reference: { kind: "assigned", ids: [""] } }] }, false],
                ["blocked", { ...value, targets: [{ ...value.targets[0], writeBlocker: "No update access.", reference: { kind: "unassigned", ids: [] } }] }, false],
                ["stale", { ...value, targets: [{ ...value.targets[0], referenceEvidence: { source: "adr-point-get", observedAt: "fixture",
                    apiVersion: "2026-11-02-preview", rawCarrier: "present" } }] }, false],
                ["partial", { ...value, issues: [{ resource: "page", message: "Incomplete current read.", retryable: true, scope: "devices" }] }, false],
                ["unlinked", { ...value, hubs: [{ ...value.hubs[0], state: "unlinked" }] }, false],
            ];
            for (const [name, inventory, success] of variants) {
                let discoveries = 0;
                let release: (() => void) | undefined;
                HubAssignmentApi.prototype.discover = () => {
                    discoveries++; return new Promise(resolve => { release = () => resolve(inventory); });
                };
                const seed = assignmentContext();
                seed.hubAssignmentHubSummary.value.hubs = inventory.hubs;
                await openAssignmentTargets(seed);
                const applicable = inventory.hubs.some(hub => hub.state === "linked");
                if (applicable) {
                    await until(() => !!release && !!area.querySelector('[role="progressbar"]'), "Current inventory read did not enter its deferred load.");
                    assert(!existingAssignmentsNotice() && !area.querySelector('[aria-label="DPS enrollment guidance"]') &&
                        !currentContext.hubAssignmentIntent.attempt, "Loading or a retained selection was treated as all-assigned proof.");
                    await act(async () => release!());
                }
                await until(() => !area.querySelector('[role="progressbar"]') &&
                    (!applicable || !!existingAssignmentsNotice() || !!area.querySelector("table")),
                    `Initial matching inventory did not settle: ${name}`);
                assert(!!existingAssignmentsNotice() === success &&
                    !!area.querySelector("table") === (!success && applicable) && !area.textContent?.includes("assigned successfully"),
                    `Initial inventory used filtered/empty/stale evidence as complete success: ${name}`);
                if (!success && applicable) assert(area.querySelector("#assignment-devices-heading")?.textContent === "Select devices",
                    `Incomplete, filtered or unknown inventory falsely used an assigned-devices heading: ${name}`);
                assertDpsGuidance(success);
                if (success) {
                    const notice = existingAssignmentsNotice()!, section = notice.closest("section")!;
                    assert(notice.querySelector(".fui-Text")?.textContent === "All devices already have a model assigned" &&
                        notice.querySelectorAll("svg").length === 1 && notice.getAttribute("role") === "status" &&
                        getComputedStyle(notice.querySelector(".fui-Text")!).fontSize === "14px" &&
                        getComputedStyle(notice.querySelector(".fui-Text")!).fontWeight === "400" &&
                        !section.querySelector(".fui-MessageBar") && !section.querySelector(".iot-solution-creation-steps") &&
                        section.querySelector("h3")?.textContent === "Assigned devices" &&
                        !section.querySelector("[data-assignment-summary]") &&
                        section.querySelector('[aria-label="DPS enrollment guidance"]')?.previousElementSibling === notice &&
                        !button("Select all") && !button("Assign to selected devices") && !button("Show all devices") &&
                        !currentContext.hubAssignmentIntent.attempt && !currentContext.hubAssignmentHistory?.length && discoveries === 1,
                        "Existing assignments were presented as new/current-model assignments, a table/list/box, or caused extra operations.");
                    const beforeLink = { reads: calls.length + resourceCalls.length, tokens: tokenRequests,
                        consent: consentRequests, context: canonicalJson(currentContext), discoveries };
                    const link = section.querySelector<HTMLAnchorElement>('[aria-label="DPS enrollment guidance"] a.fui-Link')!;
                    assertDocumentationClicks([await documentationClick(link, "https://aka.ms/iot-hub-agent-ready-data-doc")]);
                    assert(calls.length + resourceCalls.length === beforeLink.reads && tokenRequests === beforeLink.tokens &&
                        consentRequests === beforeLink.consent && canonicalJson(currentContext) === beforeLink.context &&
                        discoveries === beforeLink.discoveries,
                        "The DPS documentation link performed assignment, enrollment, discovery or authorization work.");
                }
                if (name === "mixed-known-unlinked") assert(area.querySelector('[aria-label="Source IoT Hubs"]')?.textContent?.includes("Known nonlinked Hub"),
                    "Applicable-list success removed the separate known-nonlinked Hub warning.");
            }
            caseObservations = { variants: variants.map(([name]) => name), noWrites: true, loadedInventoryNotFilteredView: true,
                anyAssignedModelQualifies: true, noInventedCurrentOperation: true, noExtraDiscovery: true,
                enabledDpsDocumentationOnly: true, documentationHostNavigationAndProductDefaultPreventionVerified: true };
        } finally { HubAssignmentApi.prototype.discover = discover; }
    }],
    ["Device Table has two semantic columns exact loaded counts eligible toggles paging and honest unknown states", async () => {
        const discover = HubAssignmentApi.prototype.discover, more = HubAssignmentApi.prototype.loadMore;
        const value = deviceListInventory();
        const longReference = value.targets[1].reference.kind === "assigned" ? value.targets[1].reference.ids[0] : "";
        let discoveryCount = 0;
        let finishMore: (() => void) | undefined;
        const statusGeometry: object[] = [];
        const checkStatus = (count: number) => {
            const status = area.querySelector<HTMLElement>("#assignment-device-counts")!;
            const toolbar = status.closest(".iot-solution-assignment-toolbar")!;
            const actions = toolbar.querySelector<HTMLElement>(".iot-solution-model-actions")!;
            const text = status.getBoundingClientRect(), controls = actions.getBoundingClientRect();
            const centerOffset = (text.top + text.bottom - controls.top - controls.bottom) / 2;
            const wrapped = controls.top >= text.bottom + 7;
            assert(status.textContent === `${count} selected` && status.getAttribute("role") === "status" &&
                (wrapped || Math.abs(centerOffset) < 0.5) &&
                toolbar.scrollWidth <= toolbar.clientWidth + 1,
                `Selection status changed copy/live semantics or lost natural vertical centering: ${JSON.stringify({ count, centerOffset, wrapped })}`);
            statusGeometry.push({ viewport: innerWidth, count, text: status.textContent, centerOffset, wrapped });
        };
        try {
            HubAssignmentApi.prototype.discover = async () => { discoveryCount++; return value; };
            HubAssignmentApi.prototype.loadMore = (previous, _signal, publication) => new Promise(resolve => {
                const next = { ...previous, hasMore: false,
                    targets: [...previous.targets, { ...previous.targets[0], id: previous.targets[0].id + "-next", name: "Next-page eligible device" }] };
                finishMore = () => { publication?.publish(next); resolve(next); };
            });
            await openAssignmentTargets();
            await until(() => !!area.querySelector('input[type="checkbox"]'), "Unassigned fixture did not settle.");
            checkStatus(0);
            assert(!area.querySelector('input[aria-label="Find a Target"]') &&
                area.querySelectorAll('tr.iot-solution-assignment-device').length === 2,
                "Device search remains or the current view hid a read-only unassigned device.");
            await click(area.querySelector('input[type="checkbox"]'));
            checkStatus(1);
            const selected = JSON.stringify(currentContext.hubAssignmentChoices);
            const beforeView = discoveryCount;
            await chooseAssignmentView("all");
            const rows = [...area.querySelectorAll<HTMLElement>('tr.iot-solution-assignment-device')];
            const panel = area.querySelector('[aria-labelledby="assignment-devices-heading"]')!;
            const action = button("Assign to selected devices")!;
            assert(panel.contains(action) && !action.closest("table") &&
                [...area.querySelectorAll("button")].filter(button => button.textContent === "Assign to selected devices").length === 1 &&
                action.getBoundingClientRect().top >= panel.querySelector("table")!.getBoundingClientRect().bottom + 7,
                "Primary assignment action is duplicated, outside the panel or not in its compact footer.");
            await click(action);
            const dialog = document.querySelector('[role="dialog"]')!;
            assert(document.querySelectorAll('[role="dialog"]').length === 1 && dialog.textContent?.includes(tm.id),
                "The single panel action lost its exact confirmation trigger.");
            await click([...dialog.querySelectorAll("button")].find(button => button.textContent === "Cancel"));
            assert(document.activeElement === action, "Panel action lost its incumbent return-focus ref.");
            const columns = [...panel.querySelectorAll("th")];
            assert(columns.length === 2 && columns[0].textContent === "Device" && columns[1].textContent === "Current model" &&
                columns.every(column => column.scope === "col") &&
                area.querySelector("#assignment-device-counts")?.textContent === "1 selected" &&
                rows.length === 5 && rows.every(row => getComputedStyle(row.children[0]).backgroundColor !== getComputedStyle(panel).backgroundColor &&
                getComputedStyle(row.children[0]).padding === "6px 12px" && getComputedStyle(row.children[0]).borderTopStyle === "solid" &&
                row.scrollWidth <= row.clientWidth + 1) &&
                area.textContent?.includes("Unknown reference") && area.textContent?.includes("Fixture unknown") &&
                area.textContent?.includes("The source IoT Hub is not linked to an ADR namespace.") &&
                area.textContent?.includes(longReference) &&
                area.querySelectorAll('input[type="checkbox"]').length === 1 &&
                JSON.stringify(currentContext.hubAssignmentChoices) === selected && discoveryCount === beforeView,
                "Native row contrast/wrapping, exact references, readonly eligibility or retained selections changed.");
            const identityLeft = rows.map(row => row.querySelector(".iot-solution-assignment-device-identity")!.children[1].getBoundingClientRect().left);
            assert(identityLeft.every(left => Math.abs(left - identityLeft[0]) < 1), "Checkbox and read-only identity gutters do not align.");
            const labels = columns.map(column => column.querySelector<HTMLElement>(".iot-solution-assignment-column-label")!);
            assert(labels.every(label => getComputedStyle(label).fontWeight === "600" && getComputedStyle(label).fontSize === "14px" &&
                getComputedStyle(label).fontFamily === getComputedStyle(rows[0].querySelector(".fui-Text")!).fontFamily) &&
                rows.every(row => Math.abs(labels[0].getBoundingClientRect().left -
                    row.querySelector(".iot-solution-assignment-device-identity")!.children[1].getBoundingClientRect().left) < 1 &&
                    Math.abs(labels[1].getBoundingClientRect().left - row.children[1].firstElementChild!.getBoundingClientRect().left) < 1),
                "Header weight/family/size or direct name/assignment alignment differs from row content.");
            const toolbar = panel.querySelector<HTMLElement>(".iot-solution-assignment-toolbar")!;
            const actions = toolbar.querySelector<HTMLElement>(".iot-solution-model-actions")!;
            const toolbarButtons = [...actions.querySelectorAll("button")];
            const panelStyle = getComputedStyle(panel);
            assert(!panel.querySelector("select") && !panel.querySelector('label[for="assignment-view"]') &&
                toolbarButtons.map(control => control.textContent || control.getAttribute("aria-label")).join("|") ===
                "Deselect all|Show unassigned devices|Refresh devices" &&
                Math.abs(toolbarButtons[2].getBoundingClientRect().right -
                    (panel.getBoundingClientRect().right - parseFloat(panelStyle.paddingRight) - parseFloat(panelStyle.borderRightWidth))) < 1 &&
                panel.scrollWidth <= panel.clientWidth + 1,
                "Toolbar still has a dropdown or refresh is not the final right-aligned action.");
            await click(button("Show unassigned devices"));
            assert(area.querySelectorAll("tbody tr").length === 2 &&
                JSON.stringify(currentContext.hubAssignmentChoices) === selected && discoveryCount === beforeView &&
                area.querySelector("#assignment-device-counts")?.textContent === "1 selected",
                "View toggle changed loaded counts, selection or issued a discovery request.");
            await click(button("Show all devices"));
            assert(JSON.stringify(currentContext.hubAssignmentChoices) === selected && discoveryCount === beforeView &&
                area.querySelectorAll("tbody tr").length === 5, "Toggling back did not preserve the view and selection.");
            assert(area.textContent?.includes("Assigned to this model") && area.textContent?.includes("Assigned to another model") &&
                area.textContent?.includes("Assignment not confirmed") && !area.textContent?.includes("Showing ") &&
                rows.filter(row => row.children[1].firstElementChild?.textContent === "None").length === 2 &&
                rows.find(row => row.textContent?.includes("Unknown reference"))?.children[1].firstElementChild?.textContent === "Assignment not confirmed" &&
                !area.textContent?.includes("Select Shown Unassigned Targets") && !area.textContent?.includes("Current Model Reference") &&
                button("Deselect all") && !button("Select all"), "Table states, selection toggle or removed labels differ.");
            await click(button("Deselect all")); assert(currentContext.hubAssignmentChoices.targets.length === 0, "Deselect all retained shown selections.");
            checkStatus(0);
            await click(button("Select all")); assert(currentContext.hubAssignmentChoices.targets.length === 1, "Select all granted readonly assignments.");
            checkStatus(1);
            const refresh = button("Refresh devices")!;
            assert(refresh.textContent === "" && refresh.getAttribute("aria-label") === "Refresh devices" && refresh.querySelector("svg"),
                "Refresh arrow lost its exact accessible action.");
            const beforeRefreshSelection = JSON.stringify(currentContext.hubAssignmentChoices);
            let finishRefresh: ((value: AssignmentInventory) => void) | undefined;
            HubAssignmentApi.prototype.discover = () => new Promise(resolve => { finishRefresh = resolve; });
            await click(refresh);
            await until(() => !!finishRefresh && !!area.querySelector('[role="progressbar"]'), "Deferred replacement read did not enter loading.");
            assert(panel.querySelectorAll("tbody tr").length === 0 && panel.querySelectorAll("th[scope='col']").length === 2 &&
                panel.querySelector("table")?.getAttribute("aria-busy") === "true" &&
                panel.contains(button("Refresh devices")!) && button("Refresh devices")!.disabled &&
                button("Assign to selected devices")!.disabled && button("Deselect all")!.disabled &&
                area.querySelector("#assignment-device-counts")?.textContent === "1 selected" &&
                JSON.stringify(currentContext.hubAssignmentChoices) === beforeRefreshSelection &&
                !area.textContent?.includes("No devices found") && !area.textContent?.includes("assigned successfully"),
                "Refresh retained stale rows, removed loading/header controls, cleared selection or fabricated an empty/success state.");
            assertDpsGuidance(false);
            await act(async () => finishRefresh!({ ...value, targets: value.targets.map((target, index) => index === 0
                ? { ...target, name: "Refreshed fixture device" } : target) }));
            await until(() => !area.querySelector('[role="progressbar"]') &&
                !!area.querySelector('input[aria-label="Select Refreshed fixture device"]'), "Current refreshed rows did not return.");
            assert(panel.querySelectorAll("tbody tr").length === 5 && panel.querySelector("table")?.getAttribute("aria-busy") === "false" &&
                currentContext.hubAssignmentChoices.targets[0].name === "Refreshed fixture device" &&
                currentContext.hubAssignmentChoices.targets.length === 1,
                "Refreshed results were not current or normal selection reconciliation was lost.");
            HubAssignmentApi.prototype.discover = async () => ({ ...value, hasMore: true, issues: [{
                resource: "Fixture device page", message: "A device page could not be read.", retryable: true, scope: "devices",
            }] });
            await click(button("Refresh devices"));
            await until(() => !!button("Load more") && !area.querySelector('[role="progressbar"]'), "Partial inventory did not settle.");
            const refreshActions = [...area.querySelectorAll('button[aria-label="Refresh devices"]')];
            assert(refreshActions.length === 2 && refreshActions.every(action => !action.textContent && action.querySelector("svg")),
                "A device warning retained a text refresh action or lost accessible icon recovery.");
            await click(button("Load more"));
            await until(() => !!finishMore && !!area.querySelector('[role="progressbar"]'), "Deferred paging did not start.");
            assert(panel.querySelectorAll("tbody tr").length === 5 && panel.querySelector("table")?.getAttribute("aria-busy") === "true",
                "Load more incorrectly blanked the incumbent loaded rows.");
            await act(async () => finishMore!());
            await until(() => area.textContent?.includes("Next-page eligible device") && !button("Load more"), "Paging lost the next eligible row.");
            await click(button("Select all"));
            assert(currentContext.hubAssignmentChoices.targets.length === 2 &&
                currentContext.hubAssignmentChoices.targets.every((target: { reference: { kind: string }; writeBlocker?: string }) =>
                    target.reference.kind === "unassigned" && !target.writeBlocker), "Bulk selection included unknown/existing/ineligible devices.");
            checkStatus(2);
            HubAssignmentApi.prototype.discover = async () => ({ ...value, targets: [value.targets[0]], hasMore: true });
            await click(button("Refresh devices")); await until(() => !area.querySelector('[role="progressbar"]'), "Reduced loaded view did not settle.");
            await click(button("Deselect all"));
            assert(currentContext.hubAssignmentChoices.targets.length === 1 &&
                currentContext.hubAssignmentChoices.targets[0].name === "Next-page eligible device",
                "Deselect all removed an incumbent selection outside the loaded current view.");
            checkStatus(1);
            await click(button("Select all"));
            checkStatus(2);
            HubAssignmentApi.prototype.discover = async () => { throw new Error("Synthetic refresh failure"); };
            await click(button("Refresh devices"));
            await until(() => area.textContent?.includes("Synthetic refresh failure"), "Refresh failure was hidden.");
            assert(area.textContent?.includes("Fixture device") && button("Assign to selected devices")?.disabled &&
                area.textContent?.includes("List incomplete"), "Read failure discarded loaded rows or enabled assignment.");
            HubAssignmentApi.prototype.discover = async () => ({ ...value, hasMore: false });
            await click(button("Refresh devices"));
            await until(() => !area.querySelector('[role="progressbar"]'), "Inventory refresh did not recover.");
            assert(area.textContent?.includes("Fixture device") && !button("Assign to selected devices")?.disabled, "Recovery lost retained eligible selection.");
            HubAssignmentApi.prototype.discover = async () => ({ hubs: [], targets: [], issues: [], hasMore: false });
            await click(button("Refresh devices"));
            await until(() => area.textContent?.includes("No devices found for these sources"), "All-target empty inventory was misreported.");
            await chooseAssignmentView("unassigned");
            assert(area.textContent?.includes("Use Show all devices to review other assignments."), "Unassigned empty guidance lost its appropriate view recovery.");
            assert(button("Select all")?.disabled && !button("Deselect all") &&
                !area.textContent?.includes("All devices are already assigned"), "Empty eligibility became a deselect action or vacuous success.");
            caseObservations = { statusViews: ["all", "unassigned"], noSearch: true, selectedIdsPreservedAcrossView: true,
                statusGeometry,
                replacementRowsHidden: true, deferredPagingRowsRetained: true, refreshSelectionPreserved: true,
                nativeListContrastAndAlignment: true, exactLongReferenceAndName: true, iconRefreshAccessible: true,
                partialQualified: true, readFailureAndRecoveryDistinct: true, trueEmptyDistinct: true, unknownReferenceReadonly: true };
        } finally { HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.loadMore = more; }
    }],
    ["V1 Ready shows ordered current resources truthful variants enabled Learn links and guarded Finish", async () => {
        const Ready = (props: React.ComponentProps<typeof HubModelConfiguredStep>) => <HubModelConfiguredStep {...props}
            sourceState={{ ...sourceFixture, ...props.sourceState, key: modelSourceKey("workspace", props.wizardContext) }}/>;
        const observations: ReturnType<typeof assertReadyPresentation>[] = [];
        for (const waiting of [false, true]) for (const variant of ["enriched", "telemetry-only", "property-only", "plural"] as const) {
            const request = { ...creationRequest(waiting, variant !== "telemetry-only"), allowWaiting: true };
            if (variant === "property-only") request.input.fields = request.input.fields.map(field =>
                ({ ...field, included: field.role === "twin" }));
            if (variant === "plural") {
                const model = fromThingModel(resolveThingModel({ ...tm, properties: {
                    temperature: { type: "number" }, humidity: { type: "number" }, state: { type: "string" },
                    firmware: { type: "string" }, unusedNoise: { type: "number" },
                } }));
                request.input.model = model;
                request.input.fields = model.fields.map(field => ({ ...setFieldRole(field,
                    ["state", "firmware"].includes(field.path[0]) ? "twin" : "telemetry"), included: field.path[0] !== "unusedNoise" }));
                request.input.scope.reference = { id: model.id! };
            }
            request.input.source.databaseName = "Verified device operations database with a long name";
            request.coverage = { ...request.coverage, fingerprint: modelingFingerprint(request.input),
                fields: inspectFieldEvidence(request.input.fields, request.coverage.telemetry, request.coverage.twin) };
            const backend = createCreationMock(request);
            const creation = await new ModeledDataCreator().run(request, backend.service, true);
            assert(creation.phase === "completed", "Safe Ready fixture did not pass its real creator.");
            const base = { ...alignmentCreationContext(request), modelDbName: request.input.source.databaseName,
                hubCreation: creation, hubAllowWaiting: false };
            const seed = withAnalyzedSamples(withCurrentHubSummary(base, assignmentInventory().hubs, [], true,
                { ...sourceFixture, key: modelSourceKey("workspace", base) }));
            await mount(<Harness component={Ready} initial={seed}/>);
            observations.push(assertReadyPresentation(request.input, creation));
            assert(backend.writes.length === 9 && backend.plan.commands.length === 9 && !summary("Created Outputs"),
                "Ready content work changed the actual planned helper assets or creation sequence.");
            if (variant === "telemetry-only") {
                const retained = { ...creation,
                    plannedOutputs: creation.plannedOutputs!.filter(output => output.name === creation.names.modeled),
                    objects: creation.objects.filter(output => output.name === creation.names.modeled) };
                await mount(<Harness component={Ready} initial={{ ...seed, hubCreation: retained }}/>);
                observations.push(assertReadyPresentation(request.input, retained));
                assert(area.querySelectorAll(".iot-solution-validation-row").length === 1 &&
                    !area.textContent?.includes(creation.names.twinChanges) && !area.textContent?.includes(creation.names.twinState),
                    "A retained modeled-only receipt acquired phantom property helper resources.");
            }
            await mount(<Harness component={Ready} initial={{ ...seed, hubCreation: { ...creation, fingerprint: "stale" } }}/>);
            assert(area.textContent?.includes("Creation needs review") && !area.querySelector(".iot-solution-validation"),
                "Mismatched result exposed verified outputs.");
            await mount(<Harness component={Ready} initial={{ ...seed, hubCreation: { ...creation, objects: [] } }}/>);
            assert(area.textContent?.includes("Not all output definitions are confirmed") && !area.querySelector(".iot-solution-validation"),
                "Missing objects gained a verified output/consumer handoff.");
            await mount(<Harness component={Ready} initial={{ ...seed, hubCreation: { ...creation, phase: "failed" } }}/>);
            assert(area.textContent?.includes("Creation needs review") && !area.querySelector(".iot-solution-validation"),
                "Partial/failed creation became Ready.");
            await mount(<Harness component={Ready} initial={{ ...seed, hubStreamChecks: undefined }}/>);
            assert(area.textContent?.includes("Creation needs review"), "Ready accepted stale/missing coverage credit.");
            assert(button("Return to Create modeled data") && !button("Return to Review"),
                "Unverified results did not name the actual creation-step destination.");
            await click(button("Return to Create modeled data"));
            assert(currentContext.requestedStep === "review-plan", "The recovery label changed its navigation target.");
            for (const invalid of [
                { hubCreation: undefined },
                { hubCreation: { ...creation, databaseName: "A different database" } },
                { hubCreation: { ...creation, names: { ...creation.names, modeled: "Different_Output" } } },
                { hubCreation: { ...creation, plannedOutputs: [] } },
                { hubSourceAnalysis: undefined },
                { hubSourceSelectionRevision: 1 },
                { rawTelemetryTable: "A different telemetry source" },
            ]) {
                await mount(<Harness component={Ready} initial={{ ...seed, ...invalid }}/>);
                assert(area.textContent?.includes("Creation needs review") && !area.querySelector(".iot-solution-model-ready") &&
                    !area.querySelector(".iot-solution-validation") && !area.querySelector('[role="link"]'),
                    "An invalid/current-source-mismatched result exposed resources or next-step placeholders.");
            }
            await mount(<Harness component={Ready} initial={seed}/>);
            await act(async () => setTestSourceState({ ...sourceFixture, ready: false }));
            assert(area.textContent?.includes("Creation needs review"), "Ready accepted an unavailable current source.");
            await act(async () => setTestSourceState(sourceFixture));
            assertReadyPresentation(request.input, creation);
        }
        caseObservations = { waitingVariants: 2, selectionVariants: ["enriched", "telemetry-only", "property-only", "plural"],
            observations, currentVerifiedNamesOnly: true, currentZeroTwinPlanStillNineCommands: true,
            retainedModeledOnlyDoesNotInventHelpers: true, disabledLearnLinks: 2, homeTasksNotCreated: true,
            mismatchSourceDatabaseNamesAndMissingDefinitionsBlocked: true, actualCloudWrites: 0 };
    }],
    ["WoT assignment stays available without a DTDL-only reason or dangling step descriptions", () => observeWizardContext(async () => {
        const requests: string[] = [], state = { patches: [] as { body: unknown; ifMatch: string | null }[] };
        assignmentResponse = adrOnlyResponse(requests, state);
        await wizard(assignmentContext());
        assert(!area.textContent?.includes("Not available for DTDL models") &&
            !step("Assign model")?.hasAttribute("aria-describedby"),
            "An available WoT assignment step displayed or announced the DTDL-only reason.");
        await click(next());
        await beginAssignmentTargets();
        await until(() => currentContext.hubAssignmentInventory?.value?.targets.length === 1 &&
            !area.querySelector('[role="progressbar"]'), "Eligible WoT ADR targets did not settle.");
        assert(area.querySelectorAll('input[type="checkbox"]').length === 1 &&
            !step("Assign model")?.classList.contains("wizard-control__step--unavailable") && !next()!.disabled &&
            !area.textContent?.includes("Not available for DTDL models"),
            "WoT assignment availability or optional continuation changed.");
        await click(area.querySelector('input[type="checkbox"]'));
        assert(!button("Assign to selected devices")?.disabled, "The reason fix changed verified-target selection eligibility.");
        for (const element of area.querySelectorAll("[aria-describedby]")) {
            for (const id of element.getAttribute("aria-describedby")!.split(/\s+/))
                assert(document.getElementById(id), `Dangling accessibility description: ${id}`);
        }
        assert(state.patches.length === 0,
            "Selecting a target dispatched a mutation without confirmation.");
        caseObservations = { availableWotReasonAbsent: true, eligibleAdrTargets: 1, selectionEnabledAssign: true,
            optionalContinue: true, danglingDescriptions: 0, writes: 0 };
    })],
    ...localPositiveFixtures.map((fixture): [string, () => Promise<void>] =>
        [`Local WoT ${fixture.key} upload preserves own root full JSON whole fields roles and guarded creation`, () => observeWizardContext(async () => {
            const original = JSON.stringify(fixture.document), request = localFixtureRequest(fixture.document);
            const scopeStart = frontendScopeRequests.length;
            parityPayloads = wotVariationMockPayloads(request);
            const seed = context(true); delete seed.hubModel; seed.hubFields = [];
            await wizard(seed); await file("thing-model-upload", original, `local-${fixture.key}.json`);
            await until(() => currentContext.hubModel?.id === fixture.rootId, () =>
                `Local upload did not settle: ${area.querySelector("#model-upload-error")?.textContent}`);
            assert(JSON.stringify(currentContext.hubModel.document) === original &&
                currentContext.hubModel.documentVersion === "root-version", "The bundle or selected own root version changed.");
            assert(!area.textContent?.includes("TM array bundles are a workload extension") &&
                !area.textContent?.includes("Selected root:") &&
                area.textContent?.includes("Not included: actions, submodels, and external definitions."),
                "The common WoT upload experience retained a bundle-specific notice or lost the current support guidance.");
            const modelReference = area.querySelector("#model-upload-reference");
            assert(modelReference?.textContent === fixture.rootId &&
                area.querySelector("#model-upload-file-name")?.textContent === currentContext.preparationFileName &&
                acceptedModelTile()!.contains(modelReference),
                "The common model tile did not show the accepted file and selected model's own reference.");
            if (Array.isArray(fixture.document)) {
                const first = fixture.document[0];
                if (first && typeof first === "object" && !Array.isArray(first) && first.id !== fixture.rootId)
                    assert(!modelReference?.textContent?.includes(String(first.id)), "The helper-first definition replaced the model reference.");
            }
            const uploadGeometry = { width: area.clientWidth, scrollWidth: area.scrollWidth };
            assert(!summary("Uploaded Model (JSON)") && JSON.stringify(currentContext.hubRootDocument) === original,
                "Removing the document preview omitted or rewrote retained definitions.");
            assert(area.querySelectorAll("[data-field-id]").length === 3 &&
                currentContext.hubModel.fields.every((field: ModelField) => !field.role),
                "Helper definitions/actions became fields or WoT declarations gained role authority.");
            const fieldsGeometry = { width: area.clientWidth, scrollWidth: area.scrollWidth };
            await until(() => currentContext.hubFields.every((field: ModelField) => !!field.role), "Local-model stream hints did not settle.");
            assert(currentContext.hubFields.map((field: ModelField) => field.role).join(",") === "telemetry,twin,telemetry",
                "Effective whole-field role candidates changed.");
            await click(next());
            assert(area.querySelector("h2")?.textContent === "Assign the model to devices" && next()!.disabled &&
                button("Assign model to devices") && button("Skip assignment"), "The WoT journey lost its explicit assignment choice.");
            await beginAssignmentTargets();
            await until(() => !next()!.disabled, "The explicit optional assignment choice did not settle.");
            assert(area.querySelector('[aria-label="Selected device model"]')?.textContent?.includes(fixture.rootId) &&
                currentContext.hubModel.id === fixture.rootId,
                "Assignment did not use the selected root's own ID.");
            await click(next());
            await verifyCoverage();
            await until(() => !next()!.disabled, "Local-model independent checks did not permit review.");
            await click(next());
            await enter("modeled-output-name", request.names.modeled);
            await enter("twin-changes-name", request.names.twinChanges);
            await enter("twin-state-name", request.names.twinState);
            creationBackend = createCreationMock(request);
            await click(button("Create modeled data"));
            await until(() => currentContext.hubCreation?.phase === "completed", () => `Local creator failed: ${JSON.stringify(currentContext.hubCreation)}`);
            assert(creationBackend.writes.length === 9 && currentContext.hubCreation.objects.length === 3, "Guarded native creation changed.");
            assert(frontendScopeRequests.slice(scopeStart).every(scopes => scopes.every(scope => !/provisioning/i.test(scope))),
                "Local upload/creation requested DPS authorization.");
            caseObservations = { key: fixture.key, ownRootId: currentContext.hubModel.id,
                wholeUploadPreviewUnchanged: true, documentVersion: currentContext.hubModel.documentVersion,
                fields: currentContext.hubFields.map((field: ModelField) => ({ path: field.path, type: field.kustoType, role: field.role })),
                fieldsGeometry, uploadGeometry, bundleSpecificNotice: false,
                modelReferenceText: modelReference?.textContent || null,
                independentStreams: 2, fakeWrites: 9, fakeNativeObjects: 3, cloudWrites: 0, dpsRequests: 0 };
        })]),
    ...localNegativeFixtures.map((fixture): [string, () => Promise<void>] =>
        [`Local WoT ${fixture.key} upload rejects explicitly without retained model or continuation`, () => observeWizardContext(async () => {
            await wizard(context(true));
            await file("thing-model-upload", JSON.stringify(fixture.document), `local-negative-${fixture.key}.json`);
            await until(() => !!area.querySelector("#model-upload-error"), "An invalid local-model upload did not expose its error.");
            const error = area.querySelector("#model-upload-error")!.textContent || "";
            assert(error.includes(fixture.error) && !currentContext.hubModel && next()!.disabled &&
                !area.querySelector(".iot-solution-model-upload-area--loaded"), `An invalid bundle retained a model or hid its cause: ${error}`);
            caseObservations = { key: fixture.key, expectedError: fixture.error, actualError: error,
                modelAvailable: false, continueEnabled: false, cloudWrites: 0 };
        })]),
    ["Mixed source assignment discovers only applicable linked ADR devices and keeps optional Continue without DPS", () => observeWizardContext(async () => {
        const fixture = mixedAssignmentFixture();
        assignmentResponse = fixture.fetcher;
        const seed = assignmentContext(); seed.hubAssignmentSource = mixedAssignmentSource(["u", "a"]);
        const Assign = (props: WizardStepProps) => <HubModelAssignmentStep {...props} workloadClient={client}
            sourceState={sourceFixture} getSourceState={() => sourceFixture}/>;
        const Fields = (props: WizardStepProps) => <HubModelFieldsStep {...props} workloadClient={client}/>;
        const scopeStart = frontendScopeRequests.length;
        await mount(<WizardControl title="Mixed-source test" initialContext={seed} initialStepId="assign"
            steps={[{ id: "assign", title: "Assign", component: Assign, validate: () => true },
                { id: "fields", title: "Fields", component: Fields, validate: () => true }]}
            onWizardComplete={() => { throw new Error("No save."); }}/>);
        await beginAssignmentTargets();
        await until(() => currentContext.hubAssignmentInventory?.value?.targets.length === 2 &&
            !area.querySelector('[role="progressbar"]'), "Mixed-source real adapter inventory did not settle.");
        const inventory = currentContext.hubAssignmentInventory.value as AssignmentInventory;
        assert(inventory.hubs.map(hub => hub.state).join(",") === "unlinked,linked" &&
            inventory.targets.map(target => target.name).join(",") === "a,both", "An unselected/non-IoTHub target was offered.");
        assert(area.querySelectorAll('input[type="checkbox"]').length === 2 && !next()!.disabled &&
            currentContext.hubAssignmentChoices?.targets.length !== 1, "Assignment became mandatory or silently selected a target.");
        assertDpsGuidance(false);
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Select fields", "Optional Continue did not reach the isolated test editor.");
        assertFieldSelectionCopy();
        assert(fixture.calls.every(call => !/provisioning|enrollment/i.test(call.url)) &&
            frontendScopeRequests.slice(scopeStart).every(scopes => scopes.every(scope => scope === "https://management.azure.com/.default")),
            "Mixed-source discovery requested DPS data/token.");
        caseObservations = { hubStates: inventory.hubs, targets: inventory.targets.map(target => target.name),
            optionalContinue: true, selected: 0, dpsRequests: 0, writes: 0 };
    })],
    ["Mixed source assignment explains a verified unlinked Hub rather than displaying only its name", async () => {
        const fixture = mixedAssignmentFixture(), seed = assignmentContext();
        seed.hubAssignmentSource = mixedAssignmentSource(["u", "a"]);
        assignmentResponse = fixture.fetcher;
        await openAssignmentTargets(seed);
        await until(() => currentContext.hubAssignmentInventory?.value?.hubs.length === 2 &&
            !area.querySelector('[role="progressbar"]'), "Mixed linked/unlinked inventory did not settle.");
        const inventory = currentContext.hubAssignmentInventory.value as AssignmentInventory;
        assert(inventory.hubs[0].state === "unlinked", "The underlying API did not verify an unlinked Hub.");
        const summaryText = area.querySelector(".iot-solution-choice-summary")?.textContent || "";
        const warning = area.querySelector('[aria-label="Source IoT Hubs"]')!;
        assert(warning?.querySelector("h3")?.textContent === "Some IoT hubs aren't linked to Azure Device Registry" &&
            warning.textContent?.includes("Devices on these hubs can't be assigned a model until the hubs are linked to an Azure Device Registry (ADR) namespace.") &&
            [...warning.querySelectorAll("li")].some(node => node.textContent === inventory.hubs[0].name),
            `Verified unlinked source has no explanatory status in the rendered assignment UI. Summary=${summaryText}; states=${inventory.hubs.map(hub => hub.state).join(",")}`);
        assert(area.querySelectorAll('input[type="checkbox"]').length === 2 &&
            [...area.querySelectorAll<HTMLInputElement>('input[type="checkbox"]')].some(control => !control.disabled) &&
            currentContext.hubModel.id === seed.hubModel.id,
            "An unlinked source hid verified linked devices or invalidated the accepted model.");
        const discover = HubAssignmentApi.prototype.discover;
        try {
            for (const state of ["pending", "unverified"] as const) {
                HubAssignmentApi.prototype.discover = async () => ({ ...inventory,
                    hubs: [{ ...inventory.hubs[0], state }, inventory.hubs[1]] });
                const Assign = (props: WizardStepProps) => {
                    currentContext = props.wizardContext;
                    return <HubModelAssignmentStep {...props} workloadClient={client}
                        sourceState={sourceFixture} getSourceState={() => sourceFixture}/>;
                };
                const Done = () => <div>Optional assignment skipped</div>;
                const stateSeed = withCurrentHubSummary(seed, [{ ...inventory.hubs[0], state }, inventory.hubs[1]]);
                await mount(<WizardControl title="Mixed link states" initialContext={stateSeed}
                    steps={[{ id: "assign", title: "Assign", component: Assign, validate: () => true },
                        { id: "done", title: "Done", component: Done, validate: () => true }]}
                    onWizardComplete={() => { throw new Error("No save."); }}/>);
                await beginAssignmentTargets();
                await until(() => currentContext.hubAssignmentInventory?.value?.hubs[0]?.state === state &&
                    !area.querySelector('[role="progressbar"]'), "Source-state explanatory fixture did not settle.");
                assert(currentContext.hubAssignmentInventory.value.hubs[0].state === state &&
                    currentContext.hubAssignmentHubSummary.value.hubs[0].state === state &&
                    area.textContent?.includes("Devices from Hubs with pending or unverified ADR links are unavailable.") &&
                    area.textContent?.includes("Verified linked Hubs can still have devices.") &&
                    area.textContent?.includes("You can continue modeling without assignments.") &&
                    (state !== "unverified" || !area.querySelector('[aria-label="Source IoT Hubs"]')),
                    "Pending/unverified source evidence was lost or an unknown link was presented as known unlinked.");
                assert(area.querySelectorAll('input[type="checkbox"]').length === 2 && !next()!.disabled,
                    "A pending/unverified source notice blocked the verified linked targets or optional Continue.");
                await click(area.querySelector('input[type="checkbox"]'));
                assert(!button("Assign to selected devices")?.disabled, "Verified linked targets stopped being selectable.");
                await click(next());
                assert(area.textContent?.includes("Optional assignment skipped"), "The scoped note made assignment mandatory.");
            }
        } finally { HubAssignmentApi.prototype.discover = discover; }
        caseObservations = { hubStates: inventory.hubs, unlinkedStateExplained: true, dataStillModelable: true,
            pendingAndUnverifiedDistinct: true, noNewAssignmentEligibility: true };
    }],
    ["DTDL component upload blocks stale WoT assignment and keeps model-declared whole fields through neutral assignment", () => observeWizardContext(async () => {
        const seed = assignmentContext();
        seed.hubAssignmentChoices = { owner: "stale-wot", targets: assignmentInventory().targets };
        const assignmentCalls: string[] = [];
        assignmentResponse = async url => { assignmentCalls.push(String(url)); throw new Error("DTDL must not discover assignment targets."); };
        await wizard(seed);
        await click(area.querySelector('input[value="dtdl"]'));
        await file("thing-model-upload", JSON.stringify(mixedDtdlDocument), "mixed-components.dtdl.json");
        await until(() => currentContext.hubModel?.kind === "dtdl", "DTDL component upload did not settle.");
        assert(area.querySelector("h2")?.textContent === "Add a device model" &&
            area.querySelectorAll("[data-field-id]").length === 18 && !area.querySelector('[role="combobox"][id^="field-role-"]'),
            "DTDL did not skip WoT assignment or rendered mutable WoT role selectors.");
        assert(currentContext.hubFields.every((field: ModelField) => field.roleSource === "model") &&
            currentContext.hubFields.filter((field: ModelField) => !!field.component).length === 12,
            "Component qualification or declared roles were lost.");
        assert(assignmentCalls.length === 0 && !area.textContent?.includes("Assign to selected devices"), "Stale WoT target choices dispatched assignment discovery.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Assign the model to devices" && !next()!.disabled &&
            area.textContent?.includes("Model assignment isn't available for DTDL models.") &&
            !button("Assign model to devices") && !button("Skip assignment") && assignmentCalls.length === 0,
            "The neutral DTDL assignment step required a WoT decision or dispatched assignment reads.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage", "Neutral DTDL Continue did not reach filters.");
        caseObservations = { fieldCount: 18, componentFieldCount: 12, modelRolesOnly: true,
            roleSelectors: 0, assignmentRequests: 0, originalDocumentUnchanged: JSON.stringify(currentContext.hubModel.document) === JSON.stringify(mixedDtdlDocument) };
    })],
    ["DTDL direct assignment view blocks stale WoT choices before auth or adapter dispatch", async () => {
        const seed = assignmentContext(), request = mixedDtdlRequest();
        seed.hubModel = request.input.model; seed.hubFields = request.input.fields; seed.hubSelectedFormat = "dtdl";
        seed.hubAssignmentChoices = { owner: "stale", targets: assignmentInventory().targets };
        const before = tokenRequests; let reads = 0;
        assignmentResponse = async () => { reads++; throw new Error("No DTDL assignment request allowed."); };
        await openAssignmentTargets(seed);
        assert(area.textContent?.includes("Model assignment isn't available for DTDL models.") &&
            !button("Assign to selected devices") && !area.querySelector('input[type="checkbox"]') && reads === 0 && tokenRequests === before,
            "Direct stale DTDL assignment surface invoked auth/discovery/action.");
        caseObservations = { directGuard: true, staleSelectionsActionable: false, authRequests: 0, adapterRequests: 0 };
    }],
    ...mountedWotVariationKeys.map((key): [string, () => Promise<void>] =>
        [`WoT variation ${key} uploads unchanged hints whole fields checks both streams and creates through the real adapter`, () => observeWizardContext(async () => {
            const variation = wotVariations.find(candidate => candidate.key === key)!;
            const request = wotVariationRequest(variation), original = JSON.stringify(variation.document);
            const scopeStart = frontendScopeRequests.length, requestStart = resourceCalls.length;
            parityPayloads = wotVariationMockPayloads(request);
            const seed = context(true); delete seed.hubModel; seed.hubFields = [];
            await wizard(seed); await file("thing-model-upload", original, `${key}.tm.json`);
            await until(() => currentContext.hubModel?.id === variation.document.id, "The uploaded model did not reach the observed real wizard state.");
            assert(currentContext.hubModel.id === variation.document.id &&
                JSON.stringify(currentContext.hubModel.document) === original, "Upload changed the supplied root ID/document.");
            assert(area.querySelectorAll("[data-field-id]").length === variation.fields!.length &&
                currentContext.hubModel.fields.every((field: ModelField) => !field.role),
                "Declared whole values were flattened or flags/kind gained role authority.");
            await until(() => currentContext.hubFields.every((field: ModelField) => !!field.role), "Stream hints did not resolve declared fields.");
            assert(currentContext.hubFields.every((field: ModelField, index: number) =>
                field.role === request.input.fields[index].role && field.kustoType === request.input.fields[index].kustoType),
                "Hints or whole-value types differed from the declared field matrix.");
            await reachScope();
            await verifyCoverage();
            await until(() => !next()!.disabled, "Both stream checks did not permit review.");
            await click(next());
            await enter("modeled-output-name", request.names.modeled);
            await enter("twin-changes-name", request.names.twinChanges);
            await enter("twin-state-name", request.names.twinState);
            creationBackend = createCreationMock(request);
            await click(button("Create modeled data"));
            await until(() => currentContext.hubCreation?.phase === "completed",
                () => `Real-adapter fixture creation did not complete: ${JSON.stringify(currentContext.hubCreation)}`);
            assert(creationBackend.writes.length === 9 && currentContext.hubCreation.objects.length === 3 &&
                JSON.stringify(currentContext.hubModel.document) === original, "Native output count or immutable upload contract changed.");
            assert(frontendScopeRequests.slice(scopeStart).every(scopes => scopes.every(scope => !/provisioning|dps/i.test(scope))) &&
                resourceCalls.slice(requestStart).every(url => !/enrollment|provisioning/i.test(url)), "A DPS request escaped the variation workflow.");
            caseObservations = { key, sourceId: variation.document.id, originalDocumentUnchanged: true,
                fields: currentContext.hubFields.map((field: ModelField) => ({ path: field.path, type: field.kustoType, role: field.role })),
                checkedStreams: ["telemetry", "twin"], realAdapterFakeNativeWrites: creationBackend.writes.length,
                outputObjects: currentContext.hubCreation.objects.length, actualCloudWrites: 0, dpsRequests: 0 };
        })]),
    ...(["other-reference", "changed-etag", "capability-loss", "permission-loss"] as const)
        .map((boundary): [string, () => Promise<void>] =>
            [`Assignment real ADR adapter rejects ${boundary} at confirmation before dispatch without DPS`, async () => {
                const requests: string[] = [], state: { modelId?: string; patches: { body: unknown; ifMatch: string | null }[] } = { patches: [] };
                const device = assignmentInventory().targets[0], delegate = adrOnlyResponse(requests, state);
                const scopeStart = frontendScopeRequests.length;
                let changed = false, changedReads = 0;
                try {
                    assignmentResponse = async (url, init) => {
                        const response = await delegate(url, init), path = new URL(String(url)).pathname;
                        if (!changed) return response;
                        if (path === device.id && (!init?.method || init.method === "GET") &&
                            (boundary === "other-reference" || boundary === "changed-etag")) {
                            const body = await response.json(); changedReads++;
                            body.etag = '"fixture-concurrent"';
                            if (boundary === "other-reference") body.properties.deviceTypeRefs = ["urn:copilot:test:other-model"];
                            return new Response(JSON.stringify(body), { headers: { "Content-Type": "application/json" } });
                        }
                        if (boundary === "capability-loss" && path === `${device.id}/capabilities` ||
                            boundary === "permission-loss" && path === `${device.id}/providers/Microsoft.Authorization/permissions`) {
                            changedReads++;
                            return new Response(JSON.stringify({ value: [] }), { headers: { "Content-Type": "application/json" } });
                        }
                        return response;
                    };
                    await openAssignmentTargets();
                    await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') &&
                        !area.querySelector('[role="progressbar"]'), "Current authoritative unset candidate did not settle.");
                    await click(area.querySelector('input[aria-label="Select Fixture device"]'));
                    await click(button("Assign to selected devices"));
                    const dialog = document.querySelector('[role="dialog"]')!;
                    assert(dialog?.textContent?.includes(tm.id), "Confirmation omitted the exact current model reference.");
                    changed = true;
                    await click([...dialog.querySelectorAll("button")].find(node => node.textContent === "Assign"));
                    await until(() => currentContext.hubAssignmentHistory?.some((receipt: { state: string }) =>
                        receipt.state === "blocked" || receipt.state === "failed"), "A changed target did not produce an explicit not-changed receipt.");
                    assert(changedReads > 0 && state.patches.length === 0 &&
                        currentContext.hubAssignmentHistory.every((receipt: { state: string }) => !["accepted", "confirmed", "unknown"].includes(receipt.state)),
                        "Confirmation ignored fresh target/source/permission evidence or dispatched an uncertain write.");
                    assert(requests.every(url => !/enrollment|provisioning/i.test(url)) &&
                        frontendScopeRequests.slice(scopeStart).every(scopes => JSON.stringify(scopes) === JSON.stringify(["https://management.azure.com/.default"])),
                        "A negative ADR boundary requested DPS data or tokens.");
                    caseObservations = { boundary, freshChangedReads: changedReads, patchCount: state.patches.length,
                        receipt: currentContext.hubAssignmentHistory.map((receipt: { state: string }) => receipt.state), dpsRequests: 0 };
                } finally { assignmentResponse = undefined; }
            }]),
    ["leaving and returning to Filters retires owned token/read waits while retaining the checked peer in wizard context", async () => {
        for (const stream of ["telemetry", "twin"] as const) {
            for (const boundary of ["token", "read"] as const) {
                const FieldsView = (props: WizardStepProps) => {
                    currentContext = props.wizardContext;
                    return <HubModelFieldsStep {...props} workloadClient={client}/>;
                };
                const FiltersView = (props: WizardStepProps) => {
                    currentContext = props.wizardContext;
                    return <HubModelScopeStep {...props} workloadClient={client}/>;
                };
                await mount(<WizardControl title="Retained-context filter lifecycle" initialContext={context()} initialStepId="filters"
                    steps={[
                        { id: "fields", title: "Fields", component: FieldsView, validate: value => fieldValidation(value.hubFields).length === 0 },
                        { id: "filters", title: "Model Coverage", component: FiltersView,
                            validate: value => streamCoverageReady(hubInput(value), 0, value.hubStreamChecks, !!value.hubAllowWaiting) },
                        { id: "review", title: "Review", component: () => <div>Review</div> },
                    ]}/>);
                const peer = stream === "telemetry" ? "twin" : "telemetry";
                await verifyCoverage();
                await enter(`hub-${stream}-scope`, "tostring(headers[\"IoTConnectionDeviceId\"]) != \"filter-A\"");
                const peerProof = currentContext.hubStreamChecks[peer].proof;
                deferredToken = boundary === "token"; tokenReleases = [];
                deferredScoped = boundary === "read"; releases = [];
                await click(button("Check coverage"));
                await until(() => boundary === "token" ? tokenReleases.length > 0 : releases.length > 0, "Owned check did not reach the deferred boundary.");
                const oldGeneration = currentContext.hubStreamChecks[stream].generation;
                await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
                assert(area.querySelector("h2")?.textContent === "Select fields", "A filter read incorrectly locked ordinary Back in the isolated test wizard.");
                await click(next());
                assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage" &&
                    currentContext.hubStreamChecks[stream].phase === "unchecked" &&
                    currentContext.hubStreamChecks[stream].generation > oldGeneration &&
                    currentContext.hubStreamChecks[peer].proof === peerProof && !button("Check coverage")!.disabled &&
                    next()!.disabled, "Retained context returned stuck checking, lost peer proof, or credited a canceled request.");
                deferredToken = false; deferredScoped = false;
                await click(button("Check coverage"));
                await until(() => currentContext.hubStreamChecks[stream].phase === "checked", "Ordinary explicit recheck was unavailable.");
                const currentProof = currentContext.hubStreamChecks[stream].proof;
                const requestCount = calls.length;
                await act(async () => { tokenReleases.splice(0).forEach(release => release()); releases.splice(0).forEach(release => release()); });
                assert(currentContext.hubStreamChecks[stream].proof === currentProof &&
                    currentContext.hubStreamChecks[peer].proof === peerProof && calls.length === requestCount && !next()!.disabled,
                    "Late disposed callback dispatched again or overwrote newer/peer proof.");
            }
        }
    }],
    ["Customer combined model keeps inline fields and neutral DTDL assignment3 with filters4 across Back", () => observeWizardContext(async () => {
        await wizard();
        assert(area.querySelector("h2")?.textContent === "Add a device model" &&
            step("Select data source")?.querySelector('[aria-label="Completed"]') &&
            step("Add device model")?.querySelector(".wizard-control__step-number")?.textContent === "2" &&
            step("Assign model")?.querySelector(".wizard-control__step-number")?.textContent === "3",
        "Split stage numbering or current model page is wrong.");
        const original = JSON.stringify(currentContext.hubModel);
        assert(area.querySelector("#model-fields-heading")?.textContent === "3. Select fields" &&
            !step("Fields") && step("Check model coverage")?.querySelector(".wizard-control__step-number")?.textContent === "4" &&
            !area.querySelector('[role="combobox"][id^="field-role-"]'), "DTD L model fields lost their read-only declared roles.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Assign the model to devices" && !next()!.disabled &&
            area.textContent?.includes("Model assignment isn't available for DTDL models.") &&
            step("Assign model")?.getAttribute("aria-current") === "step" &&
            !button("Assign model to devices") && !button("Skip assignment"),
            "DTD L assignment lost its neutral unavailable explanation or acquired a WoT-only gate.");
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        await click(area.querySelector('[data-testid="wizard-previous-btn"]'));
        assert(area.querySelector("h2")?.textContent === "Select your data source" && !area.querySelector("#thing-model-upload"), "Source and upload are not separate.");
        await until(() => next() && !next()!.disabled, "Returning source inventories did not become current.");
        await click(next());
        assert(JSON.stringify(currentContext.hubModel) === original && !area.querySelector<HTMLInputElement>("#thing-model-upload")!.disabled &&
            area.querySelectorAll('[aria-current="step"]').length === 1, "Split remount discarded accepted model or readiness/current-step.");
        caseObservations = { dtdlUnavailableInfoPresent: true, neutralContinueWithoutAssignmentReads: true,
            dtdlVisitsAssignment: true, numbering: { model: 2, assignment: 3, filters: 4 },
            completedRawSourceCheckmark: true, retainedModelOnBack: true };
    })],
    ["Compact fields retain whole property/event choices and original JSON without rejected diagnostics", async () => {
        const modelDocument = { ...tm, properties: { measurement: { type: "object", title: "Measurement",
            description: "Whole measurement", properties: { nested: { type: "number" } } } },
            events: { alarm: { title: "Alarm", data: { type: "object", properties: { state: { type: "string" } } } } }, actions: { reboot: {} } };
        parityPayloads = { telemetry: { measurement: { nested: 20 } }, twin: { properties: { reported: { alarm: { state: "idle" } } } } };
        const seed = context(true); delete seed.hubModel; seed.hubFields = [];
        await wizard(seed); await file("thing-model-upload", JSON.stringify(modelDocument), "whole-values.json");
        await until(() => currentContext.preparationFileName === "whole-values.json" &&
            !!currentContext.hubModel && !!acceptedModelTile(), "The whole-value upload did not publish its current accepted model.");
        assert(!summary("Uploaded Model (JSON)") &&
            JSON.stringify(currentContext.hubModel.document) === JSON.stringify(modelDocument), "Original document/actions changed.");
        const rows = [...area.querySelectorAll(".iot-solution-model-field-card")];
        assert(rows.length === 2 &&
            !area.textContent?.includes("Sample evidence:") && !summary("Full model schema"), "Compact whole-value composition is wrong.");
        await until(() => next() && !next()!.disabled, "Reviewable stream hints did not settle.");
        assert([...area.querySelectorAll('[role="combobox"][id^="field-role-"]')].map(control => control.textContent).join(",") === "Telemetry,Twin Property",
            "Property/event kind overrode stream roles.");
        await reachScope();
        await verifyCoverage(); await until(() => !next()!.disabled, "Both checks did not unlock Continue.");
        await click(next());
        assert(!button("Check creation requirements") && button("Create modeled data") &&
            !area.textContent?.includes("Device identity and component checks") && !area.textContent?.includes("ModeledData semantic documentation"),
        "Rejected checking/diagnostic UI remains.");
        parityPayloads = undefined;
    }],
    ["Local filter checks read their own streams and retain unchanged peer proof through a plain-filter edit", () => observeWizardContext(async () => {
        await wizard(); await click(next()); await click(next());
        assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage", "Filters were not reached.");
        let before = calls.length;
        await verifyCoverage();
        await until(() => !next()!.disabled, "Both current checks did not permit continuation.");
        assert(calls.slice(before).length === 4 &&
            calls.slice(before).filter(query => query.includes("['Telemetry']")).length === 2 &&
            calls.slice(before).filter(query => query.includes("['Twins']")).length === 2, "Combined Verify did not independently check both streams.");
        const peer = currentContext.hubStreamChecks.twin.proof;
        await enter("hub-telemetry-scope", "true");
        assert(next()!.disabled && currentContext.hubStreamChecks.twin.proof === peer &&
            currentContext.hubStreamChecks.telemetry.phase === "unchecked" && !currentContext.hubCoverage, "Own-filter edit erased peer or retained assembled proof.");
        before = calls.length; await verifyCoverage(); await until(() => !next()!.disabled, "Edited stream could not be rechecked.");
        assert(calls.slice(before).length === 2 && calls.slice(before).every(query => query.includes("['Telemetry']")) &&
            currentContext.hubStreamChecks.twin.proof === peer, "Combined Verify discarded or reread a valid peer.");
        assert(!area.querySelector('[aria-labelledby="model-reference-filter-heading"] input[type="radio"]') &&
            area.querySelectorAll('[aria-labelledby="additional-conditions-heading"] input[type="radio"]').length === 2 &&
            area.textContent?.includes("Model ID filtering isn't available for DTDL models."),
            "DTD L unexpectedly exposed unsupported reference choices.");
        assert(area.querySelector<HTMLTextAreaElement>("#hub-telemetry-scope")!.placeholder ===
            `Condition for ${currentContext.rawTelemetryTable}\nExample: ${coverageConditionExample(currentContext.hubFields, "telemetry", true)}` &&
            area.querySelector<HTMLTextAreaElement>("#hub-telemetry-scope")!.value === "true" &&
            !summary("Processing expressions") && !summary("Preview existing device data (optional)"), "Filters retained rejected surfaces or a value-shaped placeholder.");
    })],
    ["Single Create performs internal checks then mutations while blocking double submit edits and navigation", () => observeWizardContext(async () => {
        const request = creationRequest();
        const seed = { ...context(true), hubModel: request.input.model, hubFields: request.input.fields, hubScope: request.input.scope,
            hubOutputNames: request.names };
        creationBackend = createCreationMock(request);
        let closed = 0;
        await wizard(seed, () => { closed++; }); await reachReview();
        assert(!button("Check creation requirements") && !button("Create modeled data")!.disabled && creationBackend.writes.length === 0,
            "Create still requires a separate native check or wrote on page entry.");
        const tokens = tokenRequests; deferredToken = true; tokenReleases = [];
        const create = button("Create modeled data")!;
        await click(create); await click(create);
        await until(() => tokenReleases.length > 0, "Create did not reach authentication.");
        assert(tokenRequests === tokens + 1 && area.querySelector<HTMLInputElement>("#modeled-output-name")!.disabled &&
            area.querySelector<HTMLButtonElement>('[data-testid="wizard-previous-btn"]')!.disabled &&
            (await leaveGuard!()).canLeave === false, "Create did not lock the checking boundary.");
        assert([...area.querySelectorAll(".wizard-control__step")].every(element => element.getAttribute("aria-disabled") === "true") &&
            !step("Assign model")?.hasAttribute("aria-describedby") &&
            !area.textContent?.includes("Not available for DTDL models"),
            "Creation navigation lock became a model-unavailability explanation.");
        deferredCreate = true; creationReleases = []; deferredToken = false;
        await act(async () => tokenReleases.splice(0).forEach(release => release()));
        await until(() => creationReleases.length > 0, "Create did not reach its mocked first write.");
        assert(area.querySelectorAll(".iot-solution-creation-step").length === 8 &&
            !!area.querySelector(".iot-solution-creation-steps .fui-Spinner") &&
            !area.querySelector(".iot-solution-success-text"), "Native V1 progress hid current operations or showed premature success.");
        assert(currentContext.hubCreation.steps.filter((step: { status: string }) => step.status === "passed").length >= 8,
            "First mutation preceded complete preflight.");
        deferredCreate = false; await act(async () => creationReleases.splice(0).forEach(release => release()));
        await until(() => currentContext.hubCreation?.phase === "completed", "Single Create did not verify its outputs.");
        assert(!next()!.disabled && creationBackend.writes.filter(query => query.startsWith(".set ['Modeled_Output']")).length === 1,
            "Completion/double-submit contract failed.");
        assert(area.querySelectorAll(".iot-solution-creation-step-label.success").length === 8 &&
            area.querySelector(".iot-solution-success-text strong")?.textContent === "Continue" &&
            area.querySelector(".iot-solution-success-text")?.textContent?.includes("Modeled data created.") &&
            area.querySelectorAll("#modeled-data-success").length === 1 &&
            !summary("Attempt History (1)") && !summary("Processing Scope") && !summary("How Reported Context Is Added") &&
            currentContext.hubCreation.steps.length > 8 && creationBackend.writes.length === 9 &&
            button("Create modeled data")!.disabled,
            "Completed rows were fabricated/suppressed or the V1 success instruction disagreed with the retained footer.");
        await click(next());
        assert(area.querySelector("h2")?.textContent === "Your modeled data is ready" &&
            !area.querySelector<HTMLButtonElement>('[data-testid="wizard-complete-btn"]')!.disabled, "Verified setup did not reach the current final page.");
        await click(button("Finish and Close"));
        assert(closed === 1, "The retained scoped Finish and Close callback did not run exactly once.");
        caseObservations = { creationPreflightAndDoubleSubmitGuard: true, previousAndSidebarLockedWhileCreating: true,
            hostLeaveGuardBlocked: true, availableWotDtdlReasonAbsentWhileLocked: true, verifiedCompletion: true };
        creationBackend = undefined;
    })],
    ["A late local-filter result cannot restore proof after an edit-away-and-back ABA transition", () => observeWizardContext(async () => {
        await wizard(); await click(next()); await click(next());
        await enter("hub-telemetry-scope", "tostring(headers[\"IoTConnectionDeviceId\"]) == \"filter-A\"");
        deferredScoped = true; releases = [];
        await click(button("Check coverage"));
        await until(() => releases.length > 0, "Scoped query was not deferred.");
        const generation = currentContext.hubStreamChecks.telemetry.generation;
        await enter("hub-telemetry-scope", "true");
        await enter("hub-telemetry-scope", "tostring(headers[\"IoTConnectionDeviceId\"]) == \"filter-A\"");
        deferredScoped = false; await act(async () => releases.splice(0).forEach(release => release()));
        assert(currentContext.hubStreamChecks.telemetry.phase === "unchecked" &&
            currentContext.hubStreamChecks.telemetry.generation > generation && !currentContext.hubCoverage && next()!.disabled,
            "ABA late completion restored stale proof.");
        await click(button("Check coverage"));
        await until(() => currentContext.hubStreamChecks.telemetry.phase === "checked", "Current explicit check could not recover.");
    })],
    ["A direct Create preflight failure performs zero writes and each explicit retry rechecks before mutation", async () => {
        const request = creationRequest();
        creationBackend = createCreationMock(request); creationBackend.tables.get("Telemetry")!.streaming = true;
        await mount(<Harness component={HubModelReviewStep} initial={{ ...context(true), hubModel: request.input.model,
            hubFields: request.input.fields, hubScope: request.input.scope, hubCoverage: request.coverage, hubOutputNames: request.names }} />);
        await click(button("Create modeled data"));
        await until(() => currentContext.hubCreation?.phase === "blocked", "Bad prerequisite did not block direct Create.");
        assert(creationBackend.writes.length === 0 && currentContext.hubCreation.objects.length === 0 &&
            !currentContext.hubCreation.uncertainCommand, "Read-only failure claimed a mutation or effects.");
        const priorReads = creationBackend.calls.filter(call => call.csl.includes("principal roles")).length;
        creationBackend.tables.get("Telemetry")!.streaming = false;
        await click(button("Create modeled data")); await until(() => currentContext.hubCreation?.phase === "completed", "Explicit retry did not recheck.");
        assert(creationBackend.calls.filter(call => call.csl.includes("principal roles")).length === priorReads + 1, "Retry reused an old preflight.");
        creationBackend = undefined;
    }],
    ["Current completion is invalidated by changed model meaning while confirmed effects remain historical", async () => {
        const request = creationRequest(), seed = { ...context(true), hubModel: request.input.model, hubFields: request.input.fields,
            hubScope: request.input.scope, hubOutputNames: request.names };
        creationBackend = createCreationMock(request);
        await wizard(seed); await reachReview(); await click(button("Create modeled data"));
        await until(() => currentContext.hubCreation?.phase === "completed", "Original setup did not complete.");
        const writes = creationBackend.writes.length;
        await click(step("Add device model"));
        await file("thing-model-upload", JSON.stringify({ ...request.input.model.document as object, description: "Changed meaning" }), "changed.json");
        await reachReview();
        assert(!currentContext.hubCreation && next()!.disabled &&
            area.textContent?.includes("earlier completed attempt does not verify the current setup"), "Changed meaning borrowed old completion/history.");
        await click(button("Create modeled data")); await until(() => currentContext.hubCreation?.phase === "blocked", "Existing-name collision did not block.");
        assert(creationBackend.writes.length === writes, "Old effects were adopted or overwritten.");
        creationBackend = undefined;
    }],
    ["Disposal during auth or first dispatched Create stops later operations and retains uncertain effects", async () => {
        for (const phase of ["auth", "write"] as const) {
            const request = creationRequest(); creationBackend = createCreationMock(request);
            await mount(<Harness component={HubModelReviewStep} initial={{ ...context(true), hubModel: request.input.model,
                hubFields: request.input.fields, hubScope: request.input.scope, hubCoverage: request.coverage, hubOutputNames: request.names }} />);
            deferredToken = phase === "auth"; tokenReleases = []; deferredCreate = phase === "write"; creationReleases = [];
            await click(button("Create modeled data"));
            await until(() => phase === "auth" ? tokenReleases.length : creationReleases.length, "Mock boundary was not reached.");
            await act(async () => setStepVisible(false));
            deferredToken = false; deferredCreate = false;
            await act(async () => { tokenReleases.splice(0).forEach(release => release()); creationReleases.splice(0).forEach(release => release()); });
            await until(() => !currentContext.hubCreationBusy, "Disposed operation stayed busy.");
            if (phase === "auth") assert(creationBackend.writes.length === 0, "Disposed authentication dispatched mutation.");
            else assert(creationBackend.writes.length === 1 && currentContext.hubCreationHistory?.some((attempt: { uncertainCommand?: string }) => attempt.uncertainCommand),
                "Dispatched uncertain effect was retried or forgotten.");
            creationBackend = undefined;
        }
    }],
    ["Verified zero permits real wizard Continue and guarded Create without consent but invalid source and stale proof stay blocked", () => observeWizardContext(async () => {
        const request = creationRequest(true);
        emptyProcessing = true;
        await wizard({ ...context(true), hubModel: request.input.model, hubFields: request.input.fields, hubScope: request.input.scope, hubOutputNames: request.names });
        await reachScope();
        await verifyCoverage();
        assert(!next()!.disabled && !currentContext.hubAllowWaiting && !area.querySelector('input[type="checkbox"]') &&
            currentContext.hubCoverage.telemetry.payloads.length === 0,
            "Valid verified zero required consent or fabricated observed processing rows.");
        await click(next());
        const currentRequest = { input: hubInput(currentContext), coverage: currentContext.hubCoverage,
            allowWaiting: true, names: currentContext.hubOutputNames };
        creationBackend = createCreationMock(currentRequest);
        assert(!button("Create modeled data")!.disabled && next()!.disabled, "Zero review bypassed actual creation or prevented guarded Create.");
        await click(button("Create modeled data"));
        await until(() => currentContext.hubCreation?.phase === "completed", () => `Verified-zero fake Create failed: ${currentContext.hubCreation?.error}`);
        assert(!next()!.disabled && currentContext.hubCreation.awaitingData && creationBackend.writes.length === 9 &&
            !currentContext.hubAllowWaiting, "Actual root creation credit still used the removed consent flag.");
        await click(next());
        assert(!button("Finish and Close")!.disabled, "Actual root Finish rejected current verified-zero creation.");
        const previous = currentContext.hubCreation;
        await act(async () => observedWizardProps!.updateContext("hubCreation", { ...previous, objects: [] }));
        assert(button("Finish and Close")!.disabled && area.textContent?.includes("Not all output definitions are confirmed"),
            "Actual root Finish or Ready accepted missing confirmed outputs.");
        await act(async () => observedWizardProps!.updateContext("hubCreation", previous));
        assert(!button("Finish and Close")!.disabled && area.querySelector("h2")?.textContent === "Your modeled data is ready",
            "Restored genuine current definitions did not recover the authoritative Ready handoff.");
        await act(async () => observedWizardProps!.updateContext("hubStreamChecks", undefined));
        assert(button("Finish and Close")!.disabled, "Actual root Finish accepted a completed object without current stream proofs.");
        await act(async () => observedWizardProps!.onGoToStep("scope-coverage"));
        emptySamples = "both";
        await verifyCoverage();
        assert(next()!.disabled && !modelCoverageReady(hubInput(currentContext), sourceFixture.revision, currentContext.hubStreamChecks) &&
            !currentContext.hubCreation, "Invalid empty source borrowed prior zero-match/Create proof.");
        await act(async () => observedWizardProps!.updateContext("hubCreation", previous));
        assert(next()!.disabled, "Old completed creation bypassed the invalid-source coverage gate.");
        creationBackend = undefined; emptyProcessing = false; emptySamples = "none";
        caseObservations = { currentZeroContinue: true, noHiddenConsent: true, fakeCreateWrites: 9, actualRootFinishGated: true,
            staleProofBlocksFinish: true, emptySourceBlocksContinue: true, actualCloudWrites: 0 };
    })],
    ["Source-first stale local-file result cannot publish after prerequisite/source revision changes and retains accepted model/manual roles", async () => {
            const seed = context(true);
            await mount(<Harness component={ModelPreparationStep} initial={seed}/>);
            const original = JSON.stringify(seed.hubModel), roles = JSON.stringify(seed.hubFields);
            let release: (value: string) => void = () => undefined;
            const candidate = new File(["{}"], "late-candidate.json");
            Object.defineProperty(candidate, "text", { value: () => new Promise<string>(resolve => { release = resolve; }) });
            const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
            const previousFiles = Object.getOwnPropertyDescriptor(input, "files");
            Object.defineProperty(input, "files", { configurable: true, value: [candidate] });
            try {
                await act(async () => input.dispatchEvent(new Event("change", { bubbles: true })));
                assert(area.textContent?.includes("Checking late-candidate.json"), "Local file read did not start.");
                await act(async () => setTestSourceState({ ready: false, key: "other-source", revision: 2, reason: "Select both source tables before inspecting samples or uploading a model." }));
                await act(async () => release(JSON.stringify({ ...tm, title: "Late candidate must not publish" })));
                assert(JSON.stringify(currentContext.hubModel) === original && JSON.stringify(currentContext.hubFields) === roles &&
                    !area.textContent?.includes("Late candidate must not publish") && input.disabled && !currentContext.hubCoverage, "Late file result published or a source change discarded accepted model/manual intent.");
            } finally {
                if (previousFiles) Object.defineProperty(input, "files", previousFiles);
                else Reflect.deleteProperty(input, "files");
            }
            await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
            await file("thing-model-upload", JSON.stringify({ ...tm, title: "Current valid candidate" }), "current.json");
            await until(() => currentContext.hubModel?.title === "Current valid candidate", "The accepted current upload was not published to its owner.");
            assert(currentContext.hubModel.title === "Current valid candidate", "Current explicit upload could not recover after a blocked stale read.");
        }],
    ["Missing roles keep pagination keyboard selection and included-field continuation gates without a review action", async () => {
            const modelDocument = { ...tm, properties: Object.fromEntries(Array.from({ length: 105 }, (_, index) => [`field${index}`, { type: "number", title: `Field ${index}` }])) };
            const model = fromThingModel(resolveThingModel(modelDocument));
            const fields = model.fields.map((field, index) => index === 104 ? field : setFieldRole(field, "telemetry"));
            await wizard({ ...context(true), hubModel: model, hubFields: fields, preparationFileName: "paged-fields.json" });
            const before = JSON.stringify(currentContext.hubFields);
            assertFieldSelectionCopy();
            assert(next()!.disabled && !button("Review Selections") &&
                !area.querySelector("#find-model-field") && area.querySelectorAll("[data-field-id]").length === 100 &&
                button("Show 100 more fields") &&
                !area.textContent?.includes("Choose Telemetry or Twin property for Field"), "Missing-role recovery is still a per-field wall.");
            await click(button("Show 100 more fields"));
            const target = documentElementRole(fields[104].id);
            assert(target &&
                area.querySelectorAll("[data-field-id]").length === 105 &&
                JSON.stringify(currentContext.hubFields) === before && target.getAttribute("aria-invalid") === "true",
                "Pagination did not reveal the stable missing-role field or silently changed role/inclusion.");
            await click(area.querySelector('input[aria-label="Include Field 104"]'));
            assert(!next()!.disabled && controlDisabled(target) && !currentContext.hubFields[104].role,
                "An excluded field required a role or gained a role as a side effect.");
            await click(area.querySelector('input[aria-label="Include Field 104"]'));
            assert(next()!.disabled && target.getAttribute("aria-invalid") === "true" &&
                JSON.stringify(currentContext.hubFields) === before, "Reincluding an unresolved field did not restore the gate.");
            await act(async () => target.focus());
            assert(document.activeElement === target, "The paged role Dropdown lost keyboard focusability.");
            for (const key of ["Enter", "End", "Enter"]) {
                await act(async () => target!.dispatchEvent(new KeyboardEvent("keydown", { key, code: key, bubbles: true, cancelable: true })));
                await settle();
            }
            assert(!next()!.disabled && !button("Review Selections") && !area.querySelector('[data-selection-needed="true"]'),
                "Resolved roles did not release Continue or retained removed review UI.");
            assert(currentContext.hubFields[104].role === "twin" && currentContext.hubFields[104].roleSource === "manual" &&
                target.textContent === "Twin Property", "Keyboard selection did not choose the actual twin role.");
            await enter(target.id, "telemetry");
            assert(currentContext.hubFields[104].role === "telemetry" && target.textContent === "Telemetry", "Manual role override stopped working.");
            await wizard({ ...context(true), hubModel: model, hubFields: fields, preparationFileName: "paged-fields.json" });
            await click(button("Show 100 more fields"));
            assert(area.querySelectorAll("[data-field-id]").length === 105 && !button("Show 100 more fields") &&
                JSON.stringify(currentContext.hubFields) === before, "Show More lost paged fields or changed roles after search removal.");
        }],
    ["Eventhouse request failure remains distinct from incomplete successful metadata and blocks readiness", async () => {
            const savedFetch = window.fetch;
            try {
                for (const mode of ["forbidden", "network", "incomplete"] as const) {
                    window.fetch = async (input, init) => {
                        if (String(input).includes("/eventhouses/eh")) {
                            if (mode === "network")
                                throw new TypeError("Synthetic Eventhouse network unavailable");
                            return new Response(JSON.stringify(mode === "forbidden"
                                ? { errorCode: "Forbidden", message: "Synthetic Eventhouse access denied" } : {}), { status: mode === "forbidden" ? 403 : 200 });
                        }
                        return savedFetch(input, init);
                    };
                    await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
                    const expected = mode === "forbidden" ? "Synthetic Eventhouse access denied"
                        : mode === "network" ? "Synthetic Eventhouse network unavailable" : "incomplete database information";
                    await until(() => area.textContent?.includes(expected), `Wrong Eventhouse error classification for ${mode}.`);
                    assert(!currentContext.modelRawDataSelected &&
                        (mode === "incomplete" || !area.textContent?.includes("incomplete database information")), "A failed request fabricated malformed-success metadata or readiness.");
                    if (mode === "forbidden")
                        assert(!button("Retry"), "Non-retryable HTTP failure lost typed recovery policy.");
                }
            }
            finally {
                window.fetch = savedFetch;
            }
        }],
    ["Model upload recovery maps the actual DTDL unsupported-file cause and preserves supported-model diagnostics", async () => {
            const wrongKind = { notAModel: true };
            let actual: unknown;
            try {
                parseDtdlModel(wrongKind);
            }
            catch (error: unknown) {
                actual = error;
            }
            assert(actual instanceof UnsupportedDtdlModelFileError &&
                actual.message === "Unsupported model file. Provide a DTDL interface, an array of interfaces, or an IoT Central device template export.", "The fixture does not exercise the actual DTDL parser's unsupported-file path.");
            await wizard();
            await file("thing-model-upload", JSON.stringify(wrongKind), "unsupported-device-model.json");
            assert(area.querySelector("#model-upload-error-message")?.textContent ===
                "This file isn't a supported device model. Upload a WoT Thing Model, a DTDL model, or an IoT Central export." &&
                !area.textContent?.includes("array of interfaces") && next()!.disabled && modelUploadAction("Replace Model"), "The actual parser error bypassed approved customer recovery or exposed its technical inventory.");
            await file("thing-model-upload", JSON.stringify({ ...dtdl, contents: [{ "@type": "Telemetry", name: "missingSchema" }] }), "missing-schema.json");
            assert(area.querySelector("#model-upload-error-message")?.textContent?.includes("Unsupported or missing DTDL schema") &&
                !area.querySelector("#model-upload-error-message")?.textContent?.includes("isn't a supported device model"), "A genuine supported-model failure was flattened into wrong-kind recovery.");
        }],
    ["Model diagnostics preserve active causes absolute IRIs local definitions and native read-only DTDL role information", () => observeWizardContext(async () => {
        const seed = context(true); delete seed.hubModel; seed.hubFields = []; seed.preparationFileName = "";
        await wizard(seed);
        const reads = calls.length + resourceCalls.length, tokens = tokenRequests;
        const { id: _id, ...missingId } = tm;
        const variants = [
            { name: "broken-device-model-with-a-long-filename-for-error-wrapping.json", content: "{",
                expected: "broken-device-model-with-a-long-filename-for-error-wrapping.json isn't valid JSON. Fix the file, then upload it again." },
            { name: "missing-id.json", content: JSON.stringify(missingId),
                expected: "This Thing Model doesn't have an ID. Add an ID, like urn:example:thermostat, then upload the file again." },
            { name: "relative-id.json", content: JSON.stringify({ ...tm, id: "./relative.json" }),
                expected: "The model ID must be an absolute IRI, like urn:example:thermostat. Fix the ID, then upload the file again." },
            { name: "schema.json", content: JSON.stringify({ $schema: "https://json-schema.org/draft/2020-12/schema", type: "object" }),
                expected: "This is a JSON Schema, not a device model. Upload a model file instead." },
            { name: "wrong-format.json", content: JSON.stringify(dtdl),
                expected: "This file doesn't match the format you chose. Change the format or upload a different file." },
            { name: "empty.json", content: JSON.stringify({ ...tm, properties: {} }),
                expected: "No telemetry or property fields were found in empty.json. Upload a model that describes your device data." },
        ];
        for (const value of variants) {
            await file("thing-model-upload", value.content, value.name);
            await until(() => area.querySelector("#model-upload-error-message")?.firstChild?.textContent === value.expected,
                `The active ${value.name} file read/validation did not finish.`);
            assert(area.querySelector("#model-upload-error-file")?.textContent === `${value.name} couldn't be loaded. Try uploading it again` &&
                area.querySelector("#model-upload-error-message")?.firstChild?.textContent === value.expected &&
                !currentContext.hubModel && next()!.disabled,
                `The active ${value.name} recovery lost its exact filename/cause or allowed continuation.`);
        }
        for (const [kind, document] of [
            ["tm:ref", { ...tm, properties: { temperature: { "tm:ref": "https://models.example.test/missing#/properties/value" } } }],
            ["tm:extends", { ...tm, links: [{ rel: "tm:extends", href: "https://models.example.test/missing" }] }],
        ] as const) {
            await file("thing-model-upload", JSON.stringify(document), "external-definition.json");
            await until(() => area.querySelector("#model-upload-error-message")?.textContent?.includes("Required model file missing:"),
                "The local resolver did not finish the unsupported external dependency check.");
            const message = area.querySelector("#model-upload-error-message")?.textContent || "";
            assert(message.includes(kind) && message.includes("External model resolution isn't supported yet.") &&
                message.includes("Required model file missing:") && !currentContext.hubModel && next()!.disabled,
                "External dependency guidance hid the actual missing definition or rejected supported local syntax globally.");
        }
        for (const id of ["urn:example:thermostat", "https://models.example.test/thermostat#v1"]) {
            await file("thing-model-upload", JSON.stringify({ ...tm, id }), "valid-id.json");
            await until(() => currentContext.hubModel?.id === id && !!acceptedModelTile(), "The valid absolute IRI model did not finish loading.");
            assert(currentContext.hubModel?.id === id && !next()!.disabled && !area.querySelector("#model-upload-error"),
                "The URN example accidentally became a URN-only validator or changed the supplied identity.");
        }
        const valid = JSON.stringify(tm), boundary = valid + " ".repeat(2 * 1024 * 1024 - valid.length);
        await file("thing-model-upload", boundary, "exactly-2mb.json");
        await until(() => !!acceptedModelTile() && currentContext.preparationFileName === "exactly-2mb.json",
            "The exact 2 MB model read did not finish.");
        assert(currentContext.hubModel && !next()!.disabled, "The existing exact 2 MB boundary was tightened.");
        await file("thing-model-upload", boundary + " ", "over-2mb.json");
        assert(area.querySelector("#model-upload-error-message")?.textContent === "Upload a file that's 2 MB or smaller" &&
            !currentContext.hubModel && next()!.disabled, "The existing over-2 MB guard was weakened.");
        await file("thing-model-upload", JSON.stringify(diagnosticWarningDocument), "unsupported-parts.json");
        await until(() => !!area.querySelector("#model-upload-warnings") && !!acceptedModelTile(),
            "The supported warning projection did not finish.");
        const warning = area.querySelector("#model-upload-warnings")!;
        assert(warning.querySelector(".fui-Text")?.textContent === "Some parts of your model aren't supported and won't be included:" &&
            JSON.stringify([...warning.querySelectorAll("li")].map(node => node.textContent?.trim())) === JSON.stringify([
                "Fields with no type: property null_sensor_configuration_with_a_long_name_for_wrapping",
                "Fields that use oneOf: property conditional_sensor_value",
                "Submodels: Auxiliary device module with a longer name",
            ]) && currentContext.hubModel && currentContext.hubFields.length === 3 &&
            !area.querySelector("#model-upload-error"), "Unsupported-part copy changed the actual projection or appended periods.");
        await click(area.querySelector('input[value="dtdl"]')); await file("thing-model-upload", JSON.stringify(dtdl), "declared-roles.json");
        await assertDtdlRoleInformation();
        const declared = currentContext.hubFields[0] as ModelField;
        const sameRole = setFieldRole(declared, declared.role);
        assert(sameRole.roleSource === "manual" && sameRole.role === declared.role && declared.roleSource === "model",
            "The informational UI changed baseline helper metadata or mutated the original declared field.");
        let rejected = false;
        try { setFieldRole(declared, declared.role === "telemetry" ? "twin" : "telemetry"); }
        catch (error: unknown) { rejected = error instanceof Error && error.message.includes("declared by your DTDL model"); }
        assert(rejected && !next()!.disabled && calls.length + resourceCalls.length === reads && tokenRequests === tokens,
            "Informational DTDL copy weakened integrity guards or performed remote work.");
        caseObservations = { exactActiveCauses: variants.map(value => value.name), validUrnAndHttps: true,
            exact2MbAcceptedOver2MbRejected: true, unsupportedPartsProjectedNotSilentlyIncluded: true,
            externalOnlyDependencyGuidance: true, declaredRolesReadOnlyAndKeyboardInfo: true, extraReadsOrAuth: 0 };
    })],
    ["Model upload recovery failed candidates stay visible and associated for wrong-format, size and read failures in both formats", async () => {
            for (const kind of ["wot", "dtdl"] as const) {
                await wizard(context(kind === "wot"));
                await reachReview();
                await click(step("Add device model"));
                const assertFailure = (name: string) => {
                    const error = area.querySelector<HTMLElement>("#model-upload-error")!;
                    const identity = document.getElementById(error?.getAttribute("aria-labelledby") || "");
                    assert(error?.getAttribute("role") === "alert" && identity?.textContent === `${name} couldn't be loaded. Try uploading it again` &&
                        document.getElementById(error.getAttribute("aria-describedby") || "")?.id === "model-upload-error-message" &&
                        modelUploadAction("Replace Model")?.getAttribute("aria-describedby") === error.id, "The settled failure did not identify and accessibly associate its actual candidate.");
                    assert(!area.querySelector(".iot-solution-model-upload-area--loaded") && !area.textContent?.includes("fields available") &&
                        !area.textContent?.includes("seed.json") && next()!.disabled &&
                        step("Assign model")?.getAttribute("aria-disabled") === "true", "A rejected candidate retained a loaded file/count or downstream success.");
                };
                const name = `candidate-${kind}-<img onerror=alert(1)>.json`;
                await file("thing-model-upload", JSON.stringify(kind === "wot" ? dtdl : tm), name);
                assertFailure(name);
                assert(!area.querySelector("#model-upload-error img"), "A filename was interpreted as markup.");
                const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
                await act(async () => { input.files = new DataTransfer().files; input.dispatchEvent(new Event("change", { bubbles: true })); });
                assertFailure(name);
                await file("thing-model-upload", JSON.stringify(kind === "wot" ? dtdl : tm), name);
                assertFailure(name);
                assert(input.value === "", "Failed same-file replacement was not re-enabled.");
                await file("thing-model-upload", " ".repeat(2 * 1024 * 1024 + 1), `too-large-${kind}.json`);
                assertFailure(`too-large-${kind}.json`);
                assert(area.querySelector("#model-upload-error-message")?.textContent === "Upload a file that's 2 MB or smaller", "Size recovery changed.");
                await unreadableFile("thing-model-upload", `unreadable-${kind}.json`);
                assertFailure(`unreadable-${kind}.json`);
                assert(area.querySelector("#model-upload-error-message")?.firstChild?.textContent ===
                    "This file couldn't be read. Upload a different model file." &&
                    area.querySelector("#model-upload-error-message details")?.textContent?.includes("Synthetic file read failure"),
                    "The clear read-recovery text lost the actual independent file-read cause.");
                await file("thing-model-upload", JSON.stringify(kind === "wot" ? tm : dtdl), `recovered-${kind}.json`);
                assert(!area.querySelector("#model-upload-error") && area.querySelector(".iot-solution-model-upload-area--loaded") &&
                    !next()!.disabled && (area.textContent?.split(`recovered-${kind}.json`).length || 0) - 1 === 1 &&
                    !acceptedModelTile()?.getAttribute("aria-describedby")?.includes("model-upload-error"), "Recovery retained old failed identity or duplicated the valid loaded filename.");
            }
        }],
    ["Invalid protected metadata rejects a model and valid boolean metadata recovers without rewriting JSON", async () => {
            const seed = context(true);
            delete seed.hubModel;
            seed.hubFields = [];
            seed.preparationFileName = "";
            await wizard(seed);
            const document = { ...tm, "@context": [tm["@context"], {
                        tm: { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true, "@protected": "yes" },
                    }] };
            await file("thing-model-upload", JSON.stringify(document), "invalid-context.json");
            assert(area.querySelector("#model-upload-error-message")?.textContent?.includes("@protected") &&
                area.querySelector("#model-upload-error-file")?.textContent === "invalid-context.json couldn't be loaded. Try uploading it again" &&
                !area.querySelector(".iot-solution-model-upload-area--loaded") && next()!.disabled, "Invalid prefix metadata exposed a loaded/effective model.");
            const valid = { ...document, "@context": [tm["@context"], {
                        tm: { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true, "@protected": false },
                    }] };
            await file("thing-model-upload", JSON.stringify(valid), "valid-context.json");
            await until(() => currentContext.preparationFileName === "valid-context.json" &&
                !!currentContext.hubModel && !!acceptedModelTile(), "The corrected context did not publish its accepted model.");
            assert(!next()!.disabled && !area.querySelector("#model-upload-error"), "Valid boolean metadata did not recover.");
            assert(!summary("Uploaded Model (JSON)") &&
                JSON.stringify(currentContext.hubModel.document) === JSON.stringify(valid), "The accepted original context was rewritten.");
        }],
    ["Model format cards default only new models to WoT and preserve restored choices and same-option state", async () => {
        const fresh = withAnalyzedSamples(context(true));
        delete fresh.hubModel; delete fresh.hubSelectedFormat;
        fresh.hubFields = []; fresh.preparationFileName = "";
        const selectedDtdl = { ...fresh, hubSelectedFormat: "dtdl" };
        const loadedDtdl = { ...withAnalyzedSamples(context()), hubSelectedFormat: "wot" };
        const loadedWot = withAnalyzedSamples(context(true));
        loadedWot.hubCoverage = creationRequest().coverage;
        loadedWot.hubPreview = { marker: "retained-preview" };
        loadedWot.hubModelRevision = 42;
        const reads = calls.length + resourceCalls.length, tokens = tokenRequests, observations: object[] = [];
        for (const [name, seed, selected] of [
            ["new", fresh, "wot"], ["explicit-dtdl", selectedDtdl, "dtdl"],
            ["loaded-dtdl", loadedDtdl, "dtdl"], ["loaded-wot", loadedWot, "wot"],
        ] as const) {
            await mount(<Harness component={ModelPreparationStep} initial={seed}/>);
            const group = area.querySelector<HTMLElement>(".iot-solution-model-format-choices")!;
            const radios = [...group.querySelectorAll<HTMLInputElement>('input[type="radio"]')];
            const cards = [...group.querySelectorAll<HTMLElement>(".iot-solution-coverage-choice")];
            assert(group.getAttribute("role") === "radiogroup" && group.getAttribute("aria-labelledby") === "model-format-heading" &&
                area.querySelector("#model-format-heading")?.textContent === "1. Choose a model format" &&
                radios.length === 2 && radios.every(radio => radio.name && radio.name === radios[0].name && !radio.disabled) &&
                radios.filter(radio => radio.checked).length === 1 && radios.find(radio => radio.checked)?.value === selected &&
                !group.querySelector('input[type="checkbox"]') &&
                JSON.stringify(cards.map(card => card.querySelector(".iot-solution-coverage-choice-title")?.textContent)) ===
                    JSON.stringify(["WoT Thing Model (recommended)", "Digital Twin Definition Language (DTDL)"]) &&
                JSON.stringify(cards.map(card => card.querySelector(".iot-solution-coverage-choice-description")?.textContent)) ===
                    JSON.stringify(["Lets you assign the model to devices.", "Use an existing DTDL model or IoT Central template."]) &&
                radios.every(radio => !!document.getElementById(radio.getAttribute("aria-labelledby")!) &&
                    !!document.getElementById(radio.getAttribute("aria-describedby")!)),
                "Format cards lost their exclusive native radio semantics, exact labels/descriptions or restored choice.");
            const boxes = cards.map(card => card.getBoundingClientRect());
            const labelBoxes = cards.map(card => card.querySelector<HTMLElement>(".fui-Radio__label")!.getBoundingClientRect());
            const titleBoxes = cards.map(card => card.querySelector<HTMLElement>(".iot-solution-coverage-choice-title")!.getBoundingClientRect());
            assert(cards.every(card => getComputedStyle(card).borderTopWidth === "1px" &&
                getComputedStyle(card).borderTopStyle === "solid" && card.scrollWidth <= card.clientWidth + 1) &&
                group.scrollWidth <= group.clientWidth + 1 &&
                (innerWidth > 620
                    ? Math.abs(boxes[0].top - boxes[1].top) < 1 && boxes[0].right < boxes[1].left
                    : boxes[0].bottom < boxes[1].top && Math.abs(boxes[0].left - boxes[1].left) < 1),
                "Coverage-style format cards lost their outlined side-by-side or compact stacked layout.");
            assert((innerWidth <= 620 || Math.abs(boxes[0].height - boxes[1].height) <= 1 &&
                Math.abs(titleBoxes[0].top - titleBoxes[1].top) <= 1) &&
                labelBoxes.every((label, index) => Math.abs(label.top - boxes[index].top - 1) <= 1 &&
                    Math.abs(label.bottom - boxes[index].bottom + 1) <= 1) &&
                cards.every(card => card.querySelectorAll(".fui-Radio__indicator").length === 1),
                `Format cards lost equal desktop heights, top alignment or their full native radio-label hit area: ${JSON.stringify({
                    name, selected, cardHeights: boxes.map(box => box.height), labelHeights: labelBoxes.map(box => box.height),
                })}`);
            const helper = area.querySelector("#model-preparation-heading")?.closest("section");
            const guide = area.querySelector('ul[aria-label^="Supported "]')!;
            const supportedHeading = area.querySelector<HTMLElement>("#model-supported-heading")!;
            assert(supportedHeading.tagName === "SPAN" && supportedHeading.textContent === "What's supported" &&
                supportedHeading.getClientRects().length > 0 && supportedHeading.nextElementSibling === guide &&
                guide.nextElementSibling?.id === "model-file-upload-heading" &&
                (selected === "wot"
                ? helper && group.nextElementSibling === helper && helper.nextElementSibling === supportedHeading &&
                    !helper.querySelector("svg") && !!button("Get Copilot prompt")
                : !helper && !area.querySelector("#model-preparation-content") && !button("Get Copilot prompt") &&
                    group.nextElementSibling === supportedHeading),
                "The helper, subordinate What's supported label, format-specific bullets or upload heading changed order.");
            const headingTypography = assertModelPreparationHeadingTypography();
            const checked = radios.find(radio => radio.checked)!;
            const before = JSON.stringify(currentContext), model = currentContext.hubModel, fields = currentContext.hubFields;
            const coverage = currentContext.hubCoverage, preview = currentContext.hubPreview;
            await act(async () => checked.focus());
            await click(checked);
            assert(document.activeElement === checked && JSON.stringify(currentContext) === before &&
                currentContext.hubModel === model && currentContext.hubFields === fields &&
                currentContext.hubCoverage === coverage && currentContext.hubPreview === preview,
                "Same-format activation lost radio focus or reset a restored model, roles, proof or source ownership.");
            observations.push({ name, selected, cardHeights: boxes.map(box => box.height),
                labelHeights: labelBoxes.map(box => box.height), supportLabelVisible: true, headingTypography,
                sameOptionNoOp: true, nativeRadioFocus: true });
        }
        assert(calls.length + resourceCalls.length === reads && tokenRequests === tokens,
            "Defaulting or rendering model format cards performed source or authorization calls.");
        caseObservations = { cases: observations, exactCardCopy: true, sharedCoverageCardStyles: true,
            responsiveLayout: innerWidth > 620 ? "side-by-side" : "stacked", extraReadsAuthOrWrites: 0 };
    }],
    ["Model Help callout prose matches regular V1 body typography without flattening headings buttons or format choices", async () => {
        const seed = withAnalyzedSamples(context(true));
        delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
        const reads = calls.length, tokens = tokenRequests;
        await mount(<>
            <div id="v1-body-reference"><MessageBar intent="info"><MessageBarBody>
                The normalized table makes property changes easy to query. The materialized view maintains each device's latest property state.
                The modeled data table then combines that state with live telemetry and exposes selected model fields as properly typed columns.
            </MessageBarBody></MessageBar></div>
            <Harness component={ModelPreparationStep} initial={seed}/>
        </>);
        const card = area.querySelector('[aria-labelledby="model-preparation-heading"]')!;
        const prose = [...card.children].find(node => node.textContent ===
            "Copilot can draft a Thing Model from your device data. Get a prompt to copy into Copilot in Fabric.")!;
        const typography = (node: Element) => {
            const style = getComputedStyle(node);
            return { family: style.fontFamily, size: style.fontSize, weight: style.fontWeight, lineHeight: style.lineHeight };
        };
        const reference = typography(area.querySelector("#v1-body-reference .fui-MessageBarBody")!);
        const body = typography(prose), intro = typography(area.querySelector(".iot-solution-step-description")!);
        const heading = typography(card.querySelector("#model-preparation-heading > span")!);
        const action = typography(button("Get Copilot prompt")!);
        const format = typography(area.querySelector("#model-format-heading")!);
        const option = typography(area.querySelector(".fui-Radio__label")!);
        assert(JSON.stringify(body) === JSON.stringify(reference) && JSON.stringify(body) === JSON.stringify(intro) &&
            body.size === "14px" && body.weight === "400" && body.lineHeight === "20px",
            `Help prose differs from incumbent regular body: ${JSON.stringify({ body, reference, intro })}`);
        assert(Number(heading.weight) >= 600 && Number(action.weight) >= 600 && format.size === "16px" &&
            option.size === "14px" && option.weight === "400" &&
            area.querySelector<HTMLInputElement>('input[value="wot"]')?.checked &&
            !area.querySelector<HTMLInputElement>('input[value="dtdl"]')?.checked && !area.querySelector("#model-preparation-prompt") &&
            calls.length === reads && tokenRequests === tokens, "Prose typography altered emphasis, format intent or behavior.");
        caseObservations = { body, reference, intro, heading, action, format, option, extraReadsOrTokens: 0 };
    }],
    ["Model Prompt native callout prepares the exact escaped cached schema and retains fresh WoT and DTDL intent", async () => {
        const disclaimerGaps: number[] = [];
        for (const kind of ["none", "wot", "dtdl"]) {
            const seed = withAnalyzedSamples(context(kind !== "dtdl"));
            if (kind === "none") { delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = ""; }
            seed.hubSourceAnalysis.telemetry = { ...seed.hubSourceAnalysis.telemetry, payloads: [{
                ...seed.hubSourceAnalysis.telemetry.payloads[0],
                value: { temperature: 22, "<img src=x onerror=unsafe>": "private-sample-value-not-in-prompt" },
            }] };
            seed.hubCoverage = creationRequest().coverage; seed.hubPreview = { marker: "retained-preview" };
            seed.hubCreationHistory = [{ phase: "failed", databaseName: "Database", objects: ["Confirmed_Changes"] }];
            await mount(<Harness component={ModelPreparationStep} initial={seed}/>);
            const serialized = JSON.stringify(currentContext), reads = calls.length, tokens = tokenRequests;
            if (kind === "dtdl") {
                assert(area.querySelector<HTMLInputElement>('input[value="dtdl"]')?.checked &&
                    !area.querySelector("#model-preparation-heading") && !area.querySelector("#model-preparation-content") &&
                    !button("Get Copilot prompt") && JSON.stringify(currentContext) === serialized &&
                    calls.length === reads && tokenRequests === tokens,
                    "Restored DTDL intent was overridden or exposed the WoT-only preparation helper.");
                continue;
            }
            const title = area.querySelector<HTMLElement>("#model-preparation-heading")!;
            const card = title.closest(".iot-solution-config-card")!, action = button("Get Copilot prompt")!;
            const description = [...card.querySelectorAll("span")].find(node => node.textContent ===
                "Copilot can draft a Thing Model from your device data. Get a prompt to copy into Copilot in Fabric.")!;
            assert(title.textContent === "Need help creating a Thing Model?" && !card.querySelector("svg") &&
                title.getBoundingClientRect().bottom <= description.getBoundingClientRect().top &&
                description.getBoundingClientRect().bottom <= action.getBoundingClientRect().top &&
                Math.abs(description.getBoundingClientRect().left - action.getBoundingClientRect().left) <= 1 &&
                action.getAttribute("aria-expanded") === "false" && !area.querySelector("#model-preparation-prompt"),
                "The WoT helper lost its title/body/action hierarchy or retained the lightbulb.");
            await act(async () => action.click());
            assert(area.textContent?.includes("Preparing prompt...") && !button("Copy prompt") &&
                !area.querySelector(".iot-solution-model-prompt-empty"),
                "Prompt preparation did not expose its unchanged local loading state.");
            await until(() => !!area.querySelector("#model-preparation-prompt"), "The cached schema prompt did not appear.");
            const cache = seed.hubSourceAnalysis;
            const expected = "Help me create and enrich a valid WoT 1.1 Thing Model with an absolute model ID and a valid JSON-LD context for the following JSON Schema:\n\n" +
                JSON.stringify(inferPayloadSchema([...cache.telemetry.payloads, ...cache.twin.payloads].map(payload => payload.value)), null, 2);
            const prompt = area.querySelector<HTMLElement>("#model-preparation-prompt")!;
            const instructions = area.querySelector("#model-preparation-instructions")!;
            assert(instructions.tagName === "OL" &&
                JSON.stringify([...instructions.querySelectorAll("li")].map(node => node.textContent)) === JSON.stringify([
                    "Copy the prompt. It lists the fields found in your device data",
                    "Paste it into Copilot in Fabric",
                    "Review the Thing Model Copilot drafts, then save it as a .json file",
                    "Upload the file in Upload your model",
                ]) &&
                !instructions.querySelector("a, button, [role='link']") &&
                instructions.lastElementChild?.childNodes.length === 2 &&
                instructions.lastElementChild?.firstChild?.nodeType === Node.TEXT_NODE &&
                instructions.querySelectorAll("strong").length === 1 &&
                instructions.lastElementChild?.lastElementChild?.tagName === "STRONG" &&
                instructions.lastElementChild?.lastElementChild?.textContent === "Upload your model" &&
                area.querySelector("#model-file-upload-heading")?.textContent === "2. Upload your model" &&
                (kind !== "none" || area.querySelector(".iot-solution-model-upload-area")?.textContent?.includes("Drop a model file here")) &&
                area.querySelector("#model-preparation-content")?.textContent?.includes("Copilot can make mistakes. Check the model before you upload it.") &&
                !!(instructions.compareDocumentPosition(prompt) & Node.DOCUMENT_POSITION_FOLLOWING),
                "The selected expanded instructions, emphasis or placement differ from the user's wording.");
            const actions = area.querySelector<HTMLElement>("#model-preparation-content > .iot-solution-model-actions")!;
            const disclaimer = area.querySelector<HTMLElement>(".iot-solution-model-prompt-disclaimer")!;
            const gap = disclaimer.getBoundingClientRect().top - actions.getBoundingClientRect().bottom;
            assert(!area.querySelector(".iot-solution-model-prompt-empty") &&
                getComputedStyle(actions).marginTop === "0px" &&
                disclaimer.previousElementSibling === actions && Math.abs(gap - 12) <= 1 &&
                Math.abs(gap - parseFloat(getComputedStyle(disclaimer).getPropertyValue("--spacingVerticalM"))) <= 1,
                `The actual prompt-action/disclaimer gap must be the 12px Fluent spacingVerticalM token, not ${gap}px.`);
            disclaimerGaps.push(gap);
            assert(prompt.textContent === expected && !prompt.textContent.includes("private-sample-value-not-in-prompt") &&
                !prompt.querySelector("img") && prompt.tabIndex === 0 && button("Copy prompt") && button("Hide prompt") &&
                !area.textContent?.includes("Sample Comparison") && !summary("Preview JSON Schema") &&
                !button("Export JSON Schema (Unavailable)") && JSON.stringify(currentContext) === serialized &&
                calls.length === reads && tokenRequests === tokens, "Prompt shape, escaping, cache reuse or retained journey intent changed.");
            await click(button("Hide prompt"));
            assert(area.querySelector<HTMLDivElement>("#model-preparation-content")!.hidden &&
                document.activeElement === button("Get Copilot prompt") &&
                JSON.stringify(currentContext) === serialized, "Close lost focus or changed model/format/fields/proof.");
            await click(button("Get Copilot prompt"));
            await until(() => !!area.querySelector("#model-preparation-prompt"), "Reopening the local prompt did not finish.");
            assert(area.querySelector("#model-preparation-prompt")!.textContent === expected && calls.length === reads,
                "Reopening changed the prompt or read source data again.");
        }
        caseObservations = { formats: ["fresh WoT default", "WoT", "restored DTDL with no helper"], wotOnlyHelperWithoutLightbulb: true,
            exactPrefixAndPrettySchema: true, escapedPlainText: true, noRawValues: true,
            fourthInstructionBoldPhraseWithoutLink: true, disclaimerGapsPx: disclaimerGaps,
            retainedJourneyState: true, additionalReadsOrTokens: 0, obsoleteUtilityAbsent: true };
    }],
    ["Model Prompt actual root ignores components only for cached preparation and reports component-only evidence honestly", async () => {
        const originalClipboard = Object.getOwnPropertyDescriptor(navigator, "clipboard"), copied: string[] = [];
        Object.defineProperty(navigator, "clipboard", { configurable: true, value: {
            writeText: async (value: string) => { copied.push(value); },
        } });
        const rootValue = { temperature: 22.5, reading: 0, enabled: false, nullable: null, empty: "", values: [] };
        const rootReported = { maintenanceSchedule: { enabled: false, day: 0, nullable: null, note: "" }, tags: ["tag"] };
        const raw = (value: object, stream: "telemetry" | "twin", component = "", module = "") => ({
            data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
                event_kind: stream === "telemetry" ? "measurement" : "lifecycle", event_time: "2026-10-25T00:00:00Z",
                payload: { value } }),
            headers: { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: module, IoTSubject: component }, user_headers: {},
        });
        const outcomes: object[] = [];
        try {
            for (const componentOnly of [false, true]) await observeWizardContext(async () => {
                const seed = withAnalyzedSamples(context(true)), cache = seed.hubSourceAnalysis;
                const componentRows = [
                    raw({ targetTemperature: 25, fanSpeed: 3, compressorLoad: 0.8, reading: "component-value" }, "telemetry", "thermostat"),
                    raw({ powerConsumption: 40, current: 2, voltageRipple: 0.01 }, "telemetry", "powerModule"),
                ];
                cache.telemetry = decodeDoSamples(componentOnly ? componentRows : [
                    raw(rootValue, "telemetry"), ...componentRows, raw({ moduleReading: 1 }, "telemetry", "", "module-1"),
                ], "telemetry", 100);
                cache.twin = decodeDoSamples([raw({ properties: { reported: {
                    ...(componentOnly ? {} : rootReported), $version: 1,
                    thermostat: { __t: "c", targetTemperature: 26 }, powerModule: { __t: "c", current: 3 },
                } } }, "twin")], "twin", 100);
                cache.roleEvidence = prepareRoleEvidence(cache.telemetry, cache.twin);
                const originalAnalysis = JSON.stringify(cache);
                await wizard(seed);
                const serialized = JSON.stringify(currentContext), reads = calls.length, tokens = tokenRequests, resources = resourceCalls.length;
                assert(!area.querySelector("#model-preparation-prompt") && button("Get Copilot prompt")?.getAttribute("aria-expanded") === "false",
                    "Preparing the current cached source was not an explicit disclosure.");
                await click(button("Get Copilot prompt"));
                await until(() => !area.textContent?.includes("Preparing prompt..."), "Root-only cached preparation did not settle.");
                if (componentOnly) {
                    assert(area.querySelector('[role="alert"]')?.textContent ===
                        "The sampled records contain no root-level fields for model preparation." &&
                        !area.querySelector("#model-preparation-prompt") && !button("Copy prompt") &&
                        !area.textContent?.includes("No sample records were found during source validation."),
                        "Component observations were mislabeled as missing raw records or produced a pretend root schema.");
                } else {
                    const schema = inferPayloadSchema([rootValue, { moduleReading: 1 }, rootReported]);
                    const expected = "Help me create and enrich a valid WoT 1.1 Thing Model with an absolute model ID and a valid JSON-LD context for the following JSON Schema:\n\n" + JSON.stringify(schema, null, 2);
                    assert(area.querySelector("#model-preparation-prompt")?.textContent === expected &&
                        schema.properties!.reading.type === "integer" && schema.properties!.moduleReading.type === "integer",
                        "The actual root prompt included component-only fields/types or dropped module-root/ordinary nested data.");
                    for (const key of ["targetTemperature", "fanSpeed", "compressorLoad", "powerConsumption", "current", "voltageRipple", "thermostat", "powerModule"])
                        assert(!Object.prototype.hasOwnProperty.call(schema.properties, key), `Component key ${key} entered the expected root schema.`);
                    await click(button("Copy prompt"));
                    assert(copied.length === 1 && copied[0] === expected && button("Copied")?.querySelector('svg[aria-hidden="true"]') &&
                        !area.querySelector("#model-preparation-copy-status"),
                        "Copy did not use the exact current root-only schema.");
                }
                await click(button("Hide prompt"));
                assert(area.querySelector<HTMLDivElement>("#model-preparation-content")!.hidden &&
                    document.activeElement === button("Get Copilot prompt") &&
                    JSON.stringify(currentContext) === serialized && JSON.stringify(currentContext.hubSourceAnalysis) === originalAnalysis &&
                    calls.length === reads && tokenRequests === tokens && resourceCalls.length === resources,
                    "Root-only preparation mutated source/role/model state, lost disclosure focus, or performed extra reads/auth/writes.");
                outcomes.push({ componentOnly, sourceTelemetryPayloads: cache.telemetry.payloads.length,
                    sourceTwinPayloads: cache.twin.payloads.length, cachedAnalysisUnchanged: true, extraReadsAuthOrWrites: 0 });
            });
            caseObservations = { actualWizardRoot: true, outcomes, exactRootOnlyPromptCopied: copied.length === 1,
                allSixComponentLeavesAbsent: true, wholeMarkedObjectsAbsent: true, moduleRootAndNestedValuesRetained: true };
        } finally {
            if (originalClipboard) Object.defineProperty(navigator, "clipboard", originalClipboard);
            else Reflect.deleteProperty(navigator, "clipboard");
        }
    }],
    ["Model Prompt clipboard uses owned payload success failure and stale-completion guards without system clipboard access", async () => {
        const original = Object.getOwnPropertyDescriptor(navigator, "clipboard"), writes: string[] = [];
        const originalSetTimeout = window.setTimeout, originalClearTimeout = window.clearTimeout;
        const timers: number[] = [], cleared = new Set<number>();
        const observeTimer = new Proxy(originalSetTimeout, { apply(target, thisArg, args) {
            const timer = Reflect.apply(target, thisArg, args);
            if (args[1] === 2000 && typeof timer === "number") timers.push(timer);
            return timer;
        } });
        const observeClear = new Proxy(originalClearTimeout, { apply(target, thisArg, args) {
            if (typeof args[0] === "number" && timers.includes(args[0])) cleared.add(args[0]);
            return Reflect.apply(target, thisArg, args);
        } });
        assert(Reflect.set(window, "setTimeout", observeTimer) && Reflect.set(window, "clearTimeout", observeClear),
            "Owned clipboard timer observers could not be installed.");
        let fail = false, hold = false, release: (rejectCopy?: boolean) => void = () => undefined;
        Object.defineProperty(navigator, "clipboard", { configurable: true, value: { writeText: async (text: string) => {
            writes.push(text);
            if (hold) await new Promise<void>((resolve, reject) => {
                release = rejectCopy => rejectCopy ? reject(new Error("Synthetic late clipboard denial.")) : resolve();
            });
            if (fail) throw new Error("Synthetic clipboard denial with private details.");
        } } });
        const showPrompt = async () => {
            await click(button("Get Copilot prompt"));
            await until(() => !!button("Copy prompt"), "Current cached prompt did not become available.");
        };
        const assertRetired = () => assert(!button("Copied") && !area.textContent?.includes("Copy failed."),
            "A retired clipboard operation published success or failure into another prompt/owner.");
        try {
            await mount(<Harness component={ModelPreparationStep} initial={withAnalyzedSamples(context(true))}/>);
            await showPrompt();
            const expected = area.querySelector("#model-preparation-prompt")!.textContent;
            const reads = calls.length + resourceCalls.length, tokens = tokenRequests, copyButton = button("Copy prompt")!;
            await act(async () => copyButton.focus());
            hold = true; await click(copyButton);
            assert(button("Copying...") === copyButton && copyButton.getAttribute("aria-disabled") === "true" &&
                document.activeElement === copyButton && !button("Copied") && timers.length === 0,
                "Pending copy announced success, changed the button/focus or allowed duplicate submission.");
            await click(copyButton);
            assert(writes.length === 1, "The focusable pending copy button submitted a duplicate clipboard write.");
            await act(async () => release());
            await until(() => !!button("Copied"), "A completed clipboard write did not update its button.");
            assert(button("Copied") === copyButton && document.activeElement === copyButton &&
                copyButton.getAttribute("aria-live") === "polite" && copyButton.querySelector('svg[aria-hidden="true"]') &&
                !area.querySelector("#model-preparation-copy-status") && writes[0] === expected,
                "Real copy success lost the stable accessible button/checkmark or added a separate status line.");
            const firstTimer = timers.at(-1)!;
            await act(async () => { await new Promise(resolve => setTimeout(resolve, 1100)); });
            hold = false; await click(button("Copied"));
            assert(cleared.has(firstTimer) && button("Copied") === copyButton, "Repeat copy did not reset its existing feedback timer.");
            await act(async () => { await new Promise(resolve => setTimeout(resolve, 1100)); });
            assert(button("Copied") === copyButton, "The retired first timer cleared a more recent Copied confirmation.");
            await until(() => !!button("Copy prompt"), "Copied did not reset after the real bounded interval.");
            assert(button("Copy prompt") === copyButton && !copyButton.querySelector("svg"),
                "The feedback reset replaced the button or retained its success icon.");

            fail = true; await click(copyButton);
            assert(area.querySelector('[role="alert"]')?.textContent === "Copy failed. Select and copy the prompt shown above." &&
                !button("Copied") && !area.textContent?.includes("private details"),
                "Clipboard denial was silent, success-shaped or leaked raw details.");
            fail = false; await click(copyButton);
            const hideTimer = timers.at(-1)!, retained = JSON.stringify(currentContext);
            await click(button("Hide prompt"));
            assert(cleared.has(hideTimer) && area.querySelector<HTMLDivElement>("#model-preparation-content")!.hidden &&
                document.activeElement === button("Get Copilot prompt") && JSON.stringify(currentContext) === retained,
                "Hide prompt did not clear feedback while preserving disclosure focus/cache/model state.");
            await showPrompt();
            hold = true; await click(button("Copy prompt"));
            await click(button("Hide prompt")); await act(async () => release());
            assertRetired();
            await showPrompt();
            await click(button("Copy prompt"));
            await act(async () => setTestSourceState({ ready: false, key: "other-source", revision: 2, reason: "Validate the changed source tables." }));
            await act(async () => release(true)); await settle();
            assert(!area.querySelector("#model-preparation-prompt") && !button("Copy prompt"),
                "A retired source retained a copyable prompt.");
            assertRetired();
            hold = false;
            await act(async () => setTestSourceState({ ...sourceFixture, revision: 3 }));
            await showPrompt(); await click(button("Copy prompt"));
            const sourceTimer = timers.at(-1)!;
            await act(async () => setTestSourceState({ ready: false, key: "next-source", revision: 4, reason: "Validate the changed source tables." }));
            assert(cleared.has(sourceTimer), "A source change retained its Copied reset timer.");
            await act(async () => setTestSourceState({ ...sourceFixture, revision: 5 }));
            await showPrompt(); await click(button("Copy prompt"));
            const formatTimer = timers.at(-1)!;
            await click(area.querySelector('input[value="dtdl"]'));
            assert(cleared.has(formatTimer) && !area.querySelector("#model-preparation-content"),
                "A format change retained the WoT helper or its feedback timer.");
            await click(area.querySelector('input[value="wot"]')); await showPrompt();
            hold = true; await click(button("Copy prompt"));
            await click(area.querySelector('input[value="dtdl"]')); await act(async () => release());
            assertRetired();
            await click(area.querySelector('input[value="wot"]')); await showPrompt();
            hold = false; await click(button("Copy prompt"));
            const unmountTimer = timers.at(-1)!;
            await act(async () => setStepVisible(false));
            assert(cleared.has(unmountTimer), "Unmount retained the Copied reset timer.");
            await act(async () => setStepVisible(true)); await showPrompt();
            hold = true; await click(button("Copy prompt"));
            await act(async () => setStepVisible(false)); await act(async () => release(true)); await settle();
            assertRetired();
            assert(area.querySelector('[data-testid="disposed-step"]'), "A retired operation remounted its view.");
            await act(async () => setStepVisible(true)); await showPrompt();
            Object.defineProperty(navigator, "clipboard", { configurable: true, value: undefined });
            await click(button("Copy prompt"));
            assert(area.querySelector('[role="alert"]')?.textContent?.startsWith("Copy failed."),
                "Unavailable clipboard capability did not provide manual-copy recovery.");
            assert(writes.every(text => text === expected) && calls.length + resourceCalls.length === reads && tokenRequests === tokens,
                "Clipboard feedback changed the prompt payload or performed data/auth calls.");
            caseObservations = { clipboardMockOnly: true, exactPayload: true, realTwoSecondResetAndRepeat: true,
                stableFocusableLiveButton: true, decorativeSuccessIconOnly: true, explicitFailure: true,
                timerCleanup: ["repeat", "hide", "source", "format", "unmount"],
                staleCompletionRetired: ["hidden", "source", "format", "unmount"],
                sourceCacheAndPromptPayloadPreserved: true, unavailableCapabilityActionable: true,
                extraReadsOrAuth: 0, realClipboardWrites: 0 };
        } finally {
            hold = false; release();
            await act(async () => setStepVisible(false));
            Reflect.set(window, "setTimeout", originalSetTimeout); Reflect.set(window, "clearTimeout", originalClearTimeout);
            if (original) Object.defineProperty(navigator, "clipboard", original);
            else Reflect.deleteProperty(navigator, "clipboard");
        }
    }],
    ["Model Prompt handles no samples partial samples inference errors and canceled source ownership without new reads", async () => {
        const emptyGeometry: ReturnType<typeof emptyPromptGeometry>[] = [];
        for (const mode of ["both-empty", "telemetry-only", "properties-only", "deep-error", "missing-cache"]) {
            const seed = withAnalyzedSamples(context(true)), cache = seed.hubSourceAnalysis;
            if (mode === "both-empty" || mode === "telemetry-only") cache.twin = { ...cache.twin, payloads: [], rawRows: 0 };
            if (mode === "both-empty" || mode === "properties-only") cache.telemetry = { ...cache.telemetry, payloads: [], rawRows: 0 };
            if (mode === "deep-error") {
                let value: object = { leaf: 1 };
                for (let index = 0; index < 50; index++) value = { nested: value };
                cache.telemetry = { ...cache.telemetry, payloads: [{ ...cache.telemetry.payloads[0], value }] };
            }
            if (mode === "missing-cache") delete seed.hubSourceAnalysis;
            await mount(<Harness component={ModelPreparationStep} initial={seed}/>);
            const reads = calls.length, tokens = tokenRequests;
            await click(button("Get Copilot prompt"));
            await until(() => !area.textContent?.includes("Preparing prompt..."), "Prompt outcome did not settle.");
            if (mode === "both-empty") assert(area.textContent?.includes("No sample records were found during source validation. Choose Validate again in Select data source after data arrives.") &&
                !area.querySelector("#model-preparation-prompt") && !button("Copy prompt") && !area.querySelector('[role="alert"]'),
                "No data fabricated a schema or was mislabeled as failure.");
            else if (mode.endsWith("only")) assert(area.querySelector("#model-preparation-prompt") && button("Copy prompt") &&
                area.textContent?.includes(mode === "telemetry-only" ? "Only telemetry sample records" : "Only reported-property sample records"),
                "Partial evidence lost its prompt or did not identify the available stream.");
            else assert(area.querySelector('[role="alert"]') && !area.querySelector("#model-preparation-prompt") && !button("Copy prompt"),
                "A preparation error was hidden or produced a success-shaped schema.");
            if (mode === "missing-cache") assert(area.querySelector('[role="alert"]')?.textContent ===
                "Choose Validate source tables in Select data source to analyze device data first.",
                "Missing source analysis did not name the current source step and validation action.");
            if (mode === "both-empty") {
                emptyGeometry.push(emptyPromptGeometry());
                const retained = canonicalJson(currentContext), region = area.querySelector<HTMLElement>("#model-preparation-content")!;
                await click(button("Hide prompt"));
                assert(region.hidden && getComputedStyle(region).display === "none" &&
                    document.activeElement === button("Get Copilot prompt") && canonicalJson(currentContext) === retained,
                    "No-sample spacing changed disclosure visibility/focus or retained model/data state.");
                await click(button("Get Copilot prompt"));
                await until(() => !!area.querySelector(".iot-solution-model-prompt-empty"), "The cached no-sample helper did not reopen.");
                emptyGeometry.push(emptyPromptGeometry());
                assert(canonicalJson(currentContext) === retained, "Reopening an empty prompt changed existing model/roles/coverage state.");
            } else assert(!area.querySelector(".iot-solution-model-prompt-empty") &&
                getComputedStyle(area.querySelector("#model-preparation-content .iot-solution-model-actions")!).marginTop === "0px",
                "No-sample-only spacing leaked into a normal, partial or failed prompt.");
            assert(calls.length === reads && tokenRequests === tokens, "Prompt handling performed an additional source read.");
        }
        await mount(<Harness component={ModelPreparationStep} initial={withAnalyzedSamples(context(true))}/>);
        await act(async () => button("Get Copilot prompt")!.click());
        await act(async () => setTestSourceState({ ready: false, key: "retired", revision: 2, reason: "Validate the changed source tables." }));
        await settle(); await settle(); await settle();
        assert(!area.querySelector("#model-preparation-prompt") && !button("Copy prompt") && button("Get Copilot prompt")?.disabled,
            "A canceled source preparation published or enabled a stale prompt.");
        caseObservations = { emptyNoSchema: true, emptyGeometry, hideReopenFocusAndStatePreserved: true,
            partialStreamsExplained: true, errorsExplicit: true,
            sourceChangeRetiresPendingPreparation: true, additionalReadsOrTokens: 0 };
    }],
    ["Focused V1 role-first menus expand independently, remain open and retain manual table selection", async () => {
            await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
            await until(tableChoicesReady, "Table metadata did not settle.");
            await click(sourceValidationAction());
            await until(() => currentContext.modelRawDataSelected, "Initial source analysis did not settle.");
            await click(area.querySelector("#hub-model-telemetry"));
            assert(activeOptions("hub-model-telemetry").map(option => option.textContent).join("|") === "Telemetry|Show All Tables...", "Telemetry did not start with the exact V1 role subset.");
            await chooseOption("hub-model-telemetry", "Show All Tables...");
            assert(area.querySelector("#hub-model-telemetry")?.getAttribute("aria-expanded") === "true" &&
                currentContext.rawTelemetryTable === "Telemetry" && activeOptions("hub-model-telemetry").some(option => option.textContent === "NewTelemetry"), "Show all closed the menu or became a selected source.");
            await chooseOption("hub-model-telemetry", "NewTelemetry");
            assert(currentContext.rawTelemetryTable === "NewTelemetry" &&
                area.querySelector("#hub-model-telemetry")?.getAttribute("aria-expanded") === "false" &&
                !currentContext.modelRawDataSelected && !currentContext.hubSourceAnalysis,
                "A deliberate table change retained old analyzed readiness.");
            await click(area.querySelector("#hub-model-twin"));
            assert(activeOptions("hub-model-twin").map(option => option.textContent).join("|") === "Twins|Show All Tables...", "Telemetry expansion leaked into the independent properties picker.");
            await chooseOption("hub-model-twin", "Show All Tables...");
            assert(area.querySelector("#hub-model-twin")?.getAttribute("aria-expanded") === "true" &&
                currentContext.rawPropertiesTable === "Twins", "Properties expansion closed or changed the selection.");
            await chooseOption("hub-model-twin", "Twins");
            await click(sourceValidationAction());
            await until(() => currentContext.modelRawDataSelected, "Explicit manual-table analysis did not settle.");
            const retained = currentContext;
            await mount(<Harness component={HubModelSourceStep} initial={retained}/>);
            await until(tableChoicesReady, "Automatic inventory reload did not settle on return.");
            assert(currentContext.rawTelemetryTable === "NewTelemetry", "Returning discarded the manually selected table.");
            await click(area.querySelector("#hub-model-telemetry"));
            assert(!activeOptions("hub-model-telemetry").some(option => option.textContent === "NewTelemetry"), "Reload did not restore V1's role-first menu.");
            assert(!area.textContent?.includes("DOObservation") && !area.textContent?.includes("CloudEvents") && !button("Refresh resources"), "Technical inventory or the legacy public escape remains.");
        }],
    ["Focused successful zero-role metadata keeps V1 all-table fallback, including null untagged docstrings", async () => {
            tableMetadataResponse = async () => result([["TableName", "string"], ["DocString", "string"]], [["Telemetry", null], ["Twins", ""], ["NewTelemetry", "A manually created table."]]);
            await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
            await until(() => area.querySelector("#hub-model-telemetry") && tableChoicesReady(), "Untagged metadata did not settle.");
            for (const id of ["hub-model-telemetry", "hub-model-twin"]) {
                await click(area.querySelector(`#${id}`));
                const labels = activeOptions(id).map(option => option.textContent);
                assert(labels.length === 3 && labels.includes("NewTelemetry") && !labels.includes("Show All Tables..."), "A successful zero-role result did not show all tables directly.");
                await chooseOption(id, id === "hub-model-telemetry" ? "Telemetry" : "Twins");
            }
            assert(!currentContext.modelRawDataSelected && !sourceValidationAction()!.disabled, "Untagged sources were ineligible or bypassed analysis.");
            await click(sourceValidationAction());
            await until(() => currentContext.modelRawDataSelected, "Untagged source analysis could not complete.");
        }],
    ["Focused table loading, QueryStatus, metadata errors and settled empty results never become an all-table fallback", async () => {
            let release: (response: Response) => void = () => undefined;
            let pending = false;
            tableMetadataResponse = () => { pending = true; return new Promise(resolve => { release = resolve; }); };
            await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
            await until(() => pending, "Metadata read did not start.");
            assert(!tableChoicesReady() && area.querySelector("#hub-model-telemetry")?.textContent?.includes("Loading...") &&
                !area.textContent?.includes("No tables found") && !area.querySelector('[role="alert"]'), "A pending read was treated as empty or failed.");
            await act(async () => release(result([["TableName", "string"], ["DocString", "string"]], [["Telemetry", null], ["Twins", null]])));
            await until(tableChoicesReady, "Successful pending metadata did not settle.");
            const failures = [
                () => new Response(JSON.stringify({ Tables: [
                        { TableName: "PrimaryResult", Columns: [{ ColumnName: "TableName", ColumnType: "string" }, { ColumnName: "DocString", ColumnType: "string" }],
                            Rows: [["Telemetry", "@iot.meta tableRole=iothub-telemetry-raw"], ["Twins", null]] },
                        { TableName: "QueryStatus", Columns: [{ ColumnName: "Severity" }, { ColumnName: "StatusCode" }], Rows: [[2, -1]] },
                    ] })),
                () => result([["TableName", "string"], ["DocString", "string"]], [["Telemetry", { unexpected: true }]]),
                () => result([["TableName", "string"], ["DocString", "dynamic"]], [["Telemetry", "text"]]),
                () => new Response("Denied", { status: 403 }),
            ];
            for (const response of failures) {
                tableMetadataResponse = async () => response();
                await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
                await until(() => area.querySelector('[role="alert"]')?.textContent?.includes("We couldn't load tables."), "A failed metadata read was hidden.");
                assert(area.querySelector<HTMLButtonElement>("#hub-model-telemetry")?.disabled &&
                    area.querySelector<HTMLButtonElement>("#hub-model-twin")?.disabled && !currentContext.modelRawDataSelected &&
                    !area.textContent?.includes("No tables found"), "Failure was substituted for successful empty/untagged metadata.");
            }
            tableMetadataResponse = async () => result([["TableName", "string"], ["DocString", "string"]], []);
            await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
            await until(() => area.textContent?.includes("No tables found in Database."), "A successful empty read did not produce its settled state.");
            assert(!area.querySelector('[role="alert"]') && !currentContext.modelRawDataSelected, "Empty metadata fabricated an error or eligible source.");
        }],
    ["Raw source errors retain scoped Retry while healthy inventories have no Refresh controls", async () => {
            for (const stage of ["Eventhouses", "databases", "tables"]) {
                resourceResponse = stage === "Eventhouses" ? async () => new Response("Unavailable", { status: 503 }) : undefined;
                eventhouseMetadataResponse = stage === "databases" ? async () => new Response("Unavailable", { status: 503 }) : undefined;
                tableMetadataResponse = stage === "tables" ? async () => new Response("Unavailable", { status: 503 }) : undefined;
                await mount(<Harness component={HubModelSourceStep} initial={context(true)}/>);
                await until(() => area.querySelector('[role="alert"]')?.textContent?.includes(`We couldn't load ${stage}.`) &&
                    !!button("Retry") && !button("Retry")!.disabled, `The ${stage} error did not offer a usable scoped Retry.`);
                assert(!currentContext.modelRawDataSelected && !area.querySelector('button[aria-label^="Refresh"]'),
                    "An error fabricated source readiness or restored healthy Refresh controls.");
                resourceResponse = undefined; eventhouseMetadataResponse = undefined; tableMetadataResponse = undefined;
                await click(button("Retry"));
                await until(tableChoicesReady, `Explicit Retry did not recover ${stage}.`);
                assert(!button("Retry") && !area.querySelector('[role="alert"]') &&
                    !area.querySelector('button[aria-label^="Refresh"]') && !currentContext.modelRawDataSelected &&
                    !sourceValidationAction()!.disabled,
                    "Healthy automatic discovery retained Retry/Refresh or bypassed source analysis.");
                await click(sourceValidationAction());
                await until(() => currentContext.modelRawDataSelected, "Recovered inventory could not complete explicit analysis.");
            }
            caseObservations = { errorStages: ["Eventhouses", "databases", "tables"], explicitRetryRecovery: true,
                healthyRefreshControls: 0, healthyRetryControls: 0 };
        }],
    ["Assignment required TM ID supersedes no-ID validity while preserving original JSON and file interactions", async () => {
            const seed = context(true);
            delete seed.hubModel;
            seed.hubFields = [];
            seed.preparationFileName = "";
            const { id: _id, ...local } = tm;
            const document = { ...local, id: tm.id, "@context": [tm["@context"], {
                        tm: { "@id": "https://www.w3.org/2022/wot/tm#", "@prefix": true },
                        dtdl: "urn:microsoft:dtdl:", sourceKind: "dtdl:sourceKind",
                    }], properties: { ...tm.properties, temperature: { ...tm.properties.temperature, sourceKind: "Telemetry", readOnly: true } } };
            await mount(<Harness component={ModelPreparationStep} initial={seed}/>);
            const before = calls.length;
            await file("thing-model-upload", JSON.stringify(local), "missing-id.json");
            assert(!currentContext.hubModel && !currentContext.hubCoverage && !currentContext.hubStreamChecks &&
                area.textContent?.includes("missing-id.json couldn't be loaded. Try uploading it again") && area.textContent?.includes("doesn't have an ID"),
                "Required-ID rejection retained a usable model/proof or omitted file recovery.");
            await file("thing-model-upload", JSON.stringify(document), "one.json");
            assert(currentContext.hubModel?.id === tm.id && currentContext.hubFields.length === 2 &&
                currentContext.hubFields.every((field: ModelField) => !field.role && !field.component), "The corrected supplied ID or explicit role/component rules changed.");
            assert(currentContext.hubFields.length === 2 && area.textContent?.includes("Select a role for every field to continue") &&
                (area.querySelector(".iot-solution-model-upload-identity")?.textContent?.split("one.json").length || 0) - 1 === 1 &&
                !area.textContent?.includes("Stable ID"), "The primary loaded result is not compact/truthful.");
            assert(!summary("Uploaded Model (JSON)") &&
                JSON.stringify(currentContext.hubModel.document) === JSON.stringify(document), "Original JSON/custom annotations were changed.");
            const model = currentContext.hubModel;
            const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
            await act(async () => { input.files = new DataTransfer().files; input.dispatchEvent(new Event("change", { bubbles: true })); });
            assert(currentContext.hubModel === model && acceptedModelTile(), "Cancelling the chooser discarded the loaded model.");
            assert(input.value === "", "Same-file selection was not re-enabled.");
            await file("thing-model-upload", JSON.stringify(document), "one.json");
            assert(currentContext.hubModel && !area.querySelector('[role="alert"]'), "The same file could not be selected again.");
            const transfer = new DataTransfer();
            transfer.items.add(new File([JSON.stringify(document)], "dropped.json", { type: "application/json" }));
            await act(async () => area.querySelector(".iot-solution-model-upload-area")!.dispatchEvent(new DragEvent("drop", {
                bubbles: true, cancelable: true, dataTransfer: transfer,
            })));
            await until(() => area.textContent?.includes("dropped.json") && !area.querySelector('[aria-busy="true"]'), "Drop did not load the one model file.");
            assert(calls.length === before && !area.querySelector("#thing-model-dependencies"), "Local one-file validation fetched data or exposed a companion picker.");
        }],
    ["Focused file errors report verified causes and accept self-contained TM definitions and DTDL bundles", async () => {
            const seed = context(true);
            delete seed.hubModel;
            seed.hubFields = [];
            seed.preparationFileName = "";
            await wizard(seed);
            for (const [document, message] of [
                [{ $schema: "https://json-schema.org/draft/2020-12/schema", type: "object" }, "This is a JSON Schema, not a device model."],
                [dtdl, "This file doesn't match the format you chose."],
                [{ ...tm, properties: {} }, "No telemetry or property fields were found in model.json."],
                [{ ...tm, properties: { temperature: { type: "float" } } }, "doesn't follow the WoT 1.1 format"],
            ] as const) {
                await file("thing-model-upload", JSON.stringify(document));
                assert(area.querySelector('[role="alert"]')?.textContent?.includes(message) && next()!.disabled &&
                    !area.textContent?.includes("Model loaded"), `Wrong file recovery for ${message}`);
            }
            await file("thing-model-upload", " ".repeat(2 * 1024 * 1024 + 1), "large.json");
            assert(area.textContent?.includes("Upload a file that's 2 MB or smaller") && next()!.disabled, "The file-size fence was lost.");
            const selfContained = { ...tm, schemaDefinitions: { temperature: { type: "number", unit: "degC" } },
                properties: { temperature: { "tm:ref": "#/schemaDefinitions/temperature" }, state: { type: "string" } } };
            await file("thing-model-upload", JSON.stringify(selfContained), "complete.tm.json");
            assert(!next()!.disabled && currentContext.hubFields.length === 2 &&
                area.textContent?.includes("Select a role for every field to continue"), "Self-contained required TM definitions no longer resolve.");
            await click(area.querySelector('input[value="dtdl"]'));
            const base = { "@id": "dtmi:example:Base;1", "@type": "Interface", contents: [{ "@type": "Telemetry", name: "humidity", schema: "double" }] };
            await file("thing-model-upload", JSON.stringify([{ ...dtdl, extends: base["@id"] }, base]), "bundle.json");
            assert(!next()!.disabled && currentContext.hubFields.length === 3 &&
                area.textContent?.includes("Select a role for every field to continue"), "A complete DTDL interface bundle was rejected.");
        }],
    ["Focused model prompt distinguishes both-empty and one-empty samples without resetting model validation", async () => {
            for (const mode of ["both", "telemetry", "twin"] as const) {
                emptySamples = mode;
                const seed = context(true);
                delete seed.hubModel;
                seed.hubFields = [];
                seed.preparationFileName = "";
                await wizard(seed);
                await click(button("Get Copilot prompt"));
                await until(() => !area.textContent?.includes("Preparing prompt..."), "Empty-stream prompt preparation never settled.");
                const message = mode === "both" ? "No sample records were found during source validation."
                    : "This schema reflects the available stream; an empty stream does not prove that a field is unsupported.";
                assert(area.textContent?.includes(message) && next()!.disabled && !area.querySelector('[role="alert"]'), "Empty evidence was hidden, called an error or used to complete the model.");
                if (mode === "both") emptyPromptGeometry();
                else assert(!area.querySelector(".iot-solution-model-prompt-empty"), "Partial samples received the no-sample spacing marker.");
                assert(!button("Export JSON Schema (Unavailable)") && !area.querySelector("a[download]"), "Prompt preparation introduced schema-file delivery.");
            }
        }],
    ["V1-equivalent item loading starts before theme completion and survives resolve, retry and newer events", async () => {
            for (const mode of ["v1", "deferred", "retry", "event-race"] as const) {
                const itemReads: {
                    method: string;
                    params: {
                        itemId: string;
                    };
                }[] = [];
                let themeChanged: ((theme: ThemeConfiguration) => void) | undefined;
                const themeRequests: {
                    resolve: (theme: ThemeConfiguration) => void;
                    reject: (error: Error) => void;
                }[] = [];
                const rootClient: WorkloadClientAPI = {
                    ...client,
                    itemCrud: {
                        getItem: async (params) => {
                            itemReads.push({ method: "metadata", params });
                            return entryMetadata("entry-control");
                        },
                        getItemDefinition: async (params) => {
                            itemReads.push({ method: "definition", params });
                            return emptyAppDefinition();
                        },
                        updateItemDefinition: async () => { throw new Error("Entry checks cannot save an item."); },
                    },
                    theme: {
                        ...client.theme,
                        get: () => new Promise<ThemeConfiguration>((resolve, reject) => { themeRequests.push({ resolve, reject }); }),
                        onChange: callback => { themeChanged = callback; },
                    },
                };
                const history = createMemoryHistory({ initialEntries: ["/IoTSolutionItem-editor/entry-control"] });
                const app = <App history={history} workloadClient={rootClient}/>;
                const resourceStart = resourceCalls.length;
                await mount(mode !== "v1"
                    ? <FabricHostThemeProvider workloadClient={rootClient}>{app}</FabricHostThemeProvider>
                    : <FluentProvider theme={webLightTheme}>{app}</FluentProvider>, true);
                if (mode !== "v1") {
                    const startedBeforeTheme = itemReads.length > 0;
                    if (!startedBeforeTheme) {
                        await act(async () => themeRequests[0].resolve(hostTheme(false)));
                        assert(false, "Required App/item loading waited for the host theme, unlike V1.");
                    }
                }
                await until(() => area.querySelector('[aria-label="Model your device data"]'), "A successful supported item read did not reach Home.");
                const homeBeforeTheme = area.querySelector('[aria-label="Model your device data"]');
                assert(homeBeforeTheme?.querySelector(".iot-solution-home-card-description")?.textContent?.trim() ===
                    "Optionally add semantic context and current device state to live telemetry for richer monitoring, alerts, and AI experiences.",
                    "The Home modeling card description no longer matches the approved image.");
                if (mode === "retry") {
                    await act(async () => themeRequests[0].reject(new Error("Synthetic theme read failure")));
                    assert(button("Try again") && area.querySelector('[role="alert"]')?.textContent?.includes("Unable to load the Fabric theme"), "A failed theme request lost its explicit error/retry state.");
                    await click(button("Try again"));
                    assert(themeRequests.length === 2, "Theme retry did not use the same supported theme.get path.");
                    await act(async () => themeRequests[1].resolve(hostTheme(false)));
                }
                else if (mode === "event-race") {
                    await act(async () => themeChanged!(hostTheme(true)));
                    await act(async () => themeRequests[0].resolve(hostTheme(false)));
                    assert(getComputedStyle(area.querySelector(".fui-FluentProvider")!).getPropertyValue("--colorNeutralBackground1").trim() ===
                        hostTheme(true).tokens.colorNeutralBackground1, "An old theme response overwrote the newer host event.");
                }
                else if (mode === "deferred") {
                    await act(async () => themeRequests[0].resolve(hostTheme(false)));
                }
                assert(area.querySelector('[aria-label="Model your device data"]') === homeBeforeTheme, "Theme resolution/retry remounted the required App/item state.");
                assert(JSON.stringify(itemReads) === JSON.stringify([
                    { method: "metadata", params: { itemId: "entry-control" } },
                    { method: "definition", params: { itemId: "entry-control" } },
                ]), "V1/current loading changed the supported SDK request shape or duplicated item loading.");
                await click(area.querySelector('[aria-label="Model your device data"]'));
                assert(area.querySelector("h2")?.textContent === "Turn raw IoT data into modeled data" &&
                    resourceCalls.length === resourceStart, "Opening Modeling skipped Overview or queried source inventory before entering it.");
                await click(next());
                await until(() => area.querySelector("#hub-model-eventhouse") && !area.querySelector<HTMLButtonElement>("#hub-model-eventhouse")!.disabled, "Valid item/workspace context did not reach source discovery.");
                assert(resourceCalls.slice(resourceStart).some(address => address.includes("/workspaces/workspace/items")) &&
                    !area.textContent?.includes("Open this experience from a Fabric workspace"), "Source discovery lost the workspace returned by the supported item read.");
                const choose = async (id: string, label: string) => {
                    const control = area.querySelector(`#${id}`);
                    await click(control);
                    const listbox = document.getElementById(control!.getAttribute("aria-controls") || "");
                    assert(listbox, `The ${id} dropdown did not expose its active listbox.`);
                    await click([...listbox!.querySelectorAll('[role="option"]')].find(option => option.textContent === label));
                };
                await choose("hub-model-eventhouse", "Eventhouse");
                await until(() => area.querySelector("#hub-model-database") && !area.querySelector<HTMLButtonElement>("#hub-model-database")!.disabled, "The root workspace did not load the chosen Eventhouse's databases.");
                await choose("hub-model-database", "Database");
                await until(() => area.querySelector("#hub-model-telemetry") && tableChoicesReady(), "The root workspace did not load raw table choices.");
                if (innerWidth <= 768)
                    assert(getComputedStyle(area.querySelector(".iot-solution-model-source-grid--tables")!).gridTemplateColumns.split(" ").length === 1, "Narrow item content did not stack the raw table controls at the incumbent breakpoint.");
                await choose("hub-model-telemetry", "Telemetry");
                await choose("hub-model-twin", "Twins");
                assert(next()!.disabled && sourceValidationAction(), "Raw table selection bypassed explicit source validation.");
                await click(sourceValidationAction());
                await until(() => next() && !next()!.disabled, () => "Valid root context and selected tables did not enable Continue: " + JSON.stringify({
                    mode, heading: area.querySelector("h2")?.textContent,
                    selections: ["eventhouse", "database", "telemetry", "twin"].map(name => area.querySelector(`#hub-model-${name}`)?.textContent),
                    alerts: [...area.querySelectorAll('[role="alert"]')].map(alert => alert.textContent),
                    recentQueries: calls.slice(-3),
                }));
                if (mode !== "v1") {
                    const sourceControl = area.querySelector("#hub-model-eventhouse");
                    await act(async () => themeChanged!(hostTheme(true)));
                    assert(area.querySelector("#hub-model-eventhouse") === sourceControl && itemReads.length === 2 &&
                        area.querySelector("#hub-model-telemetry")?.textContent?.includes("Telemetry") &&
                        area.querySelector("#hub-model-twin")?.textContent?.includes("Twins") && !next()!.disabled, "A host-theme change reset the root item/workspace or repeated loading.");
                }
            }
        }],
    ["required metadata/definition failures cannot render loaded Home or source success; explicit retry uses the same SDK path", async () => {
            for (const failure of ["metadata400", "definition403", "missing-metadata", "missing-workspace", "wrong-item",
                "missing-definition", "malformed-definition", "invalid-payload", "null-application"] as const) {
                let fail = true;
                const reads: {
                    method: string;
                    params: {
                        itemId: string;
                    };
                }[] = [];
                const rootClient: WorkloadClientAPI = { ...client, itemCrud: {
                        getItem: async (params) => {
                            reads.push({ method: "metadata", params });
                            if (fail && failure === "metadata400")
                                throw { errorCode: "InvalidItemType",
                                    message: "Requested item type '' is invalid", isRetriable: false, requestId: "mock-request" };
                            if (fail && failure === "missing-metadata")
                                return JSON.parse("{}");
                            if (fail && failure === "missing-workspace")
                                return JSON.parse(JSON.stringify({
                                    item: { ...entryMetadata("load-control").item, workspaceId: null },
                                }));
                            return entryMetadata(fail && failure === "wrong-item" ? "another-item" : "load-control");
                        },
                        getItemDefinition: async (params) => {
                            reads.push({ method: "definition", params });
                            if (fail && failure === "definition403")
                                throw { errorCode: "UnknownError", message: "Definition request rejected", status: 403 };
                            if (fail && failure === "missing-definition")
                                return JSON.parse("null");
                            if (fail && failure === "malformed-definition")
                                return JSON.parse('{"definition":{"parts":"invalid"}}');
                            if (fail && failure === "invalid-payload")
                                return { definition: { format: "", parts: [
                                            { path: "payload.json", payload: "not-base64!", payloadType: PayloadType.InlineBase64 },
                                        ] } };
                            if (fail && failure === "null-application")
                                return { definition: { format: "", parts: [
                                            { path: "payload.json", payload: encodeJsonToBase64(null), payloadType: PayloadType.InlineBase64 },
                                        ] } };
                            return emptyAppDefinition();
                        },
                        updateItemDefinition: async () => { throw new Error("No item writes in load tests."); },
                    } };
                const before = resourceCalls.length;
                await mount(<App history={createMemoryHistory({ initialEntries: ["/IoTSolutionItem-editor/load-control"] })} workloadClient={rootClient}/>);
                await until(() => area.querySelector("h2")?.textContent === "This item could not be loaded", `No root load failure for ${failure}.`);
                assert(!area.querySelector('[aria-label="Model your device data"]') && resourceCalls.length === before &&
                    !area.textContent?.includes("Open this experience from a Fabric workspace"), "A failed required load fabricated Home/workspace/source success.");
                assert(reads.every(read => JSON.stringify(read.params) === '{"itemId":"load-control"}'), "Load recovery invented SDK parameters.");
                if (["metadata400", "missing-metadata", "missing-workspace", "wrong-item"].includes(failure)) {
                    assert(reads.length === 1, "Definition loading continued after required metadata failed.");
                }
                if (failure === "metadata400") {
                    assert(!button("Retry loading item") && area.textContent?.includes("not retryable") &&
                        area.textContent?.includes("InvalidItemType") && area.textContent?.includes("mock-request"), "Non-retryable host failure was hidden or automatically retried.");
                }
                else {
                    const priorReadCount = reads.length;
                    fail = false;
                    await click(button("Retry loading item"));
                    await until(() => area.querySelector('[aria-label="Model your device data"]'), "A successful same-path retry did not load the item.");
                    assert(reads.length === priorReadCount + 2, "Retry did not repeat exactly the ordinary metadata and definition reads.");
                    assert(!area.textContent?.includes("This item could not be loaded"), "Successful platform-only definition stayed in error state.");
                }
            }
        }],
    ["loading waits for actual item and resource responses; a platform-only definition is legitimate empty app configuration", async () => {
            const nativeMetadata = entryMetadata("normalization-control");
            const platformOnly = emptyAppDefinition();
            platformOnly.definition.parts[0].payload = encodeJsonToBase64({ metadata: { type: "", displayName: "Entry control" } });
            const normalized = convertGetItemResultToWorkloadItem(nativeMetadata, platformOnly);
            assert(normalized.id === nativeMetadata.item.id && normalized.workspaceId === nativeMetadata.item.workspaceId &&
                normalized.type === nativeMetadata.item.type && normalized.definition === undefined, "An optional platform part replaced valid required native identity/type or invented app configuration.");
            let resolveMetadata: (value: ReturnType<typeof entryMetadata>) => void = () => undefined;
            let definitions = 0;
            const rootClient: WorkloadClientAPI = { ...client, itemCrud: {
                    getItem: () => new Promise(resolve => { resolveMetadata = resolve; }),
                    getItemDefinition: async () => { definitions++; return emptyAppDefinition(); },
                    updateItemDefinition: async () => { throw new Error("No writes."); },
                } };
            await mount(<App history={createMemoryHistory({ initialEntries: ["/IoTSolutionItem-editor/delayed-control"] })} workloadClient={rootClient}/>);
            assert(area.textContent?.includes("Loading IoT Solution") && definitions === 0 &&
                !area.querySelector('[aria-label="Model your device data"]'), "Pending metadata was treated as an empty loaded item.");
            await act(async () => resolveMetadata(entryMetadata("delayed-control")));
            await until(() => area.querySelector('[aria-label="Model your device data"]'), "Successful platform-only load did not reach Home.");
            let releaseResources: (value: Response) => void = () => undefined;
            let resourcePending = false;
            resourceResponse = () => { resourcePending = true; return new Promise(resolve => { releaseResources = resolve; }); };
            await click(area.querySelector('[aria-label="Model your device data"]'));
            assert(area.querySelector("h2")?.textContent === "Turn raw IoT data into modeled data" && !resourcePending,
                "The empty app configuration skipped Overview or started source discovery before Continue.");
            await click(next());
            await until(() => resourcePending, "Resource discovery did not start.");
            assert(area.querySelector<HTMLButtonElement>("#hub-model-eventhouse")?.disabled && !area.querySelector('[role="alert"]') &&
                !area.textContent?.includes("No Eventhouses found"), "A pending list was diagnosed as a settled error/empty result.");
            resourceResponse = undefined;
            await act(async () => releaseResources(new Response(JSON.stringify({ value: [
                    { id: "eh", type: "Eventhouse", displayName: "Eventhouse" }, { id: "db", type: "KQLDatabase", displayName: "Database" },
                ] }))));
            await until(() => area.querySelector("#hub-model-eventhouse") && !area.querySelector<HTMLButtonElement>("#hub-model-eventhouse")!.disabled, "Completed resource request did not populate source choices.");
            const control = area.querySelector("#hub-model-eventhouse")!;
            await click(control);
            const listbox = document.getElementById(control.getAttribute("aria-controls") || "");
            assert([...listbox!.querySelectorAll('[role="option"]')].some(option => option.textContent === "Eventhouse"), "Visible resource option did not arrive after completion.");
            assert(definitions === 1, "Loading/visibility changes repeated the required definition read.");
        }],
    ["actual model/back/format/invalid/sidebar lifecycle is recoverable and current-step remains unique", async () => {
            await wizard();
            await reachReview();
            for (const title of ["Add device model", "Check model coverage", "Create modeled data"]) {
                await click(title === "Create modeled data" ? next() : step(title));
                assert(area.querySelectorAll('[aria-current="step"]').length === 1 && step(title)?.getAttribute("aria-current") === "step", `Current-step orientation failed for ${title}.`);
            }
            await click(step("Add device model"));
            assert(area.querySelectorAll('[aria-current="step"]').length === 1, "Returned current step was lost.");
            await click(area.querySelector('input[value="wot"]'));
            for (const name of ["Assign model", "Check model coverage", "Create modeled data"]) {
                assert(step(name)?.getAttribute("aria-disabled") === "true", `${name} retained stale navigation.`);
                await click(step(name));
            }
            assert(area.querySelector("h2")?.textContent === "Add a device model", "Stale navigation mounted a missing model.");
            await file("thing-model-upload", "{invalid");
            assert(area.querySelector('[role="alert"]')?.textContent?.includes("isn't valid JSON"), "Replacement error missing.");
            assert(step("Assign model")?.getAttribute("aria-disabled") === "true", "Invalid replacement unlocked downstream navigation.");
            await file("thing-model-upload", JSON.stringify(tm), "valid.json");
            assert(area.textContent?.includes("valid.json") && acceptedModelTile(), "Authoritative filename/Replace state missing.");
            assert(!area.querySelector<HTMLInputElement>("#thing-model-upload")!.offsetWidth, "Native empty-filename input is visible.");
            assert(area.querySelector("#model-fields-heading")?.textContent === "3. Select fields" && !next()!.disabled,
                "Recovery did not restore valid inline fields.");
        }],
    ["failed single-file replacement and missing definitions invalidate navigation before an async read", () => observeWizardContext(async () => {
            await wizard(context(true));
            await reachReview();
            await click(step("Add device model"));
            const input = area.querySelector<HTMLInputElement>("#thing-model-upload")!;
            const transfer = new DataTransfer();
            const pending = new File(["{}"], "waiting.json");
            let release: (text: string) => void = () => undefined;
            Object.defineProperty(pending, "text", { value: () => new Promise<string>(resolve => { release = resolve; }) });
            transfer.items.add(pending);
            // Browser DataTransfer may clone a File, so override the input's file collection for this delayed boundary.
            Object.defineProperty(input, "files", { configurable: true, value: [pending] });
            await act(async () => input.dispatchEvent(new Event("change", { bubbles: true })));
            assert(step("Assign model")?.getAttribute("aria-disabled") === "true" && !acceptedModelTile() &&
                !area.querySelector("#model-upload-reference") && !area.querySelector("#model-fields-heading") &&
                area.querySelector('[aria-busy="true"]'), "Pending replacement retained success identity, fields or navigation.");
            await act(async () => release("{invalid"));
            await settle();
            assert(step("Assign model")?.getAttribute("aria-disabled") === "true", "Failed replacement restored stale completion.");
            await wizard(context(true));
            await reachReview();
            await click(step("Add device model"));
            await file("thing-model-upload", JSON.stringify({ ...tm, links: [{ rel: "tm:extends", href: "./missing.json" }] }));
            await until(() => area.querySelector("#model-upload-error-message")?.textContent?.includes("Required model file missing: ./missing.json."),
                "The actual unresolved local dependency did not finish its diagnostic.");
            assert(area.querySelector("#model-upload-error-message")?.textContent?.startsWith(
                "This model extends another model with tm:extends that isn't included in this file.") &&
                area.textContent?.includes("External model resolution isn't supported yet.") &&
                next()!.disabled && !currentContext.hubModel, "Missing definitions were not recoverable.");
            assert(!area.querySelector("#thing-model-dependencies") && !button("Add related files"), "Recovery names a removed companion-file control.");
        })],
    ["missing-model Fields/Scope/Review mounts render recovery instead of throwing", async () => {
            for (const component of [HubModelFieldsStep, HubModelScopeStep, HubModelReviewStep]) {
                await mount(<Harness component={component} initial={{}}/>);
                await click(button("Return to Device Model"));
                assert(currentContext.requestedStep === "device-model", "Missing model recovery did not route to model.");
            }
        }],
    ["real source selection clears stale hints but retains manual/model roles", async () => {
            const initial = context(true);
            initial.hubFields = [...initial.hubFields.map((field: ModelField, index: number) => ({ ...field, roleSource: index ? "manual" : "hint" })),
                { ...initial.hubFields[0], id: "model-role", roleSource: "model" }];
            initial.hubRoleEvidenceKey = "old-source";
            await mount(<Harness component={HubModelSourceStep} initial={initial}/>);
            await until(() => area.querySelector("#hub-model-telemetry") && tableChoicesReady(), "Sources did not load.");
            await chooseOption("hub-model-telemetry", "Show All Tables...");
            await chooseOption("hub-model-telemetry", "NewTelemetry");
            assert(currentContext.hubFields[0].role === undefined, "Stale automatic role survived source replacement.");
            assert(currentContext.hubFields[1].role && currentContext.hubFields[2].role, "Manual/model choices were discarded.");
            assert(!currentContext.hubRoleEvidenceKey, "Old evidence key survived.");
        }],
    ["Model Prompt assistance preserves the default format without completing a model or implementing file export", () => observeWizardContext(async () => {
            const seed = context();
            delete seed.hubModel;
            delete seed.hubSelectedFormat;
            seed.hubFields = [];
            seed.preparationFileName = "";
            await wizard(seed);
            const disclosure = button("Get Copilot prompt")!;
            assert(modelUploadAction("Choose File") && area.querySelector<HTMLInputElement>('input[value="wot"]')?.checked &&
                area.querySelector<HTMLDivElement>("#model-preparation-content")!.hidden && !currentContext.hubModel,
                "The default uploader or explicit collapsed prompt disclosure is missing.");
            await click(disclosure);
            await until(() => area.querySelector("#model-preparation-prompt"), "Prompt preparation did not complete.");
            assert(next()!.disabled && !currentContext.hubModel && area.querySelector<HTMLInputElement>('input[value="wot"]')?.checked &&
                area.querySelector("#thing-model-upload") && !area.querySelector("a[download]"),
                "Preparing a prompt changed the format, credited an accepted model or implemented file export.");
            await click(button("Hide prompt"));
            await click(area.querySelector('input[value="dtdl"]'));
            await file("thing-model-upload", JSON.stringify(dtdl), "device.json");
            assert(acceptedModelTile() && !next()!.disabled, "Preparing a schema displaced model upload.");
            assert(!button("Get Copilot prompt") && !area.querySelector("#model-preparation-content") &&
                currentContext.hubModel.kind === "dtdl", "A restored DTDL model exposed WoT-only preparation.");
            await click(area.querySelector('input[value="wot"]'));
            await file("thing-model-upload", JSON.stringify(tm), "device.tm.json");
            await until(() => !!acceptedModelTile() && !next()!.disabled, "The explicit WoT model did not become ready.");
            const retained = JSON.stringify(currentContext);
            await click(button("Get Copilot prompt"));
            await until(() => area.querySelector("#model-preparation-prompt"), "Second preparation did not settle.");
            assert(area.querySelector(".iot-solution-model-upload-area--loaded") && !next()!.disabled, "Inspection reset the loaded model.");
            await click(button("Hide prompt"));
            assert(document.activeElement === button("Get Copilot prompt") && acceptedModelTile() && !next()!.disabled &&
                JSON.stringify(currentContext) === retained,
                "Collapsing preparation lost its focus or loaded model.");
        })],
    ["DTDL recovery names the visible chooser or replacement and a complete interface bundle", async () => {
            const seed = context();
            delete seed.hubModel;
            seed.hubFields = [];
            seed.preparationFileName = "";
            await wizard(seed);
            assert(modelUploadAction("Choose File") && area.textContent?.includes("Self-contained JSON or JSON-LD, up to 2 MB.") &&
                area.textContent?.includes("an IoT Central device template export"), "The one-file DTDL action/context is missing.");
            assert(!area.textContent?.includes("Related Thing Model files"), "DTDL recovery names a WoT-only control.");
            await file("thing-model-upload", "{invalid", "invalid.json");
            assert(modelUploadAction("Replace Model") && area.textContent?.includes("invalid.json isn't valid JSON"), "Invalid DTDL recovery names a missing control.");
            await file("thing-model-upload", JSON.stringify({ ...dtdl, extends: "dtmi:example:Missing;1" }), "missing-interface.json");
            assert(area.querySelector('[role="alert"]')?.textContent?.includes("not supplied"), "Missing interface was not surfaced.");
            assert(area.textContent?.includes("Include them in this file, then choose it again."), "Complete-file recovery is missing.");
            assert(!area.textContent?.includes("Related Thing Model files"), "Missing-interface recovery still names a WoT-only control.");
        }],
    ["scoped ribbon Save does not imply modeling persistence", async () => {
            for (const [view, visible] of [["home", true], ["model-wizard", false]]) {
                await mount(<IoTSolutionItemRibbon workloadClient={client} viewContext={{ currentView: view } as ViewContext} isSaveButtonEnabled saveItemCallback={async () => { throw new Error("No writes."); }} openSettingsCallback={async () => undefined}/>);
                assert(!!area.querySelector('[data-testid="ribbon-save-btn"]') === visible, "Save scope is incorrect.");
            }
        }],
    ["supported root host-theme get/change preserves mounted state and custom host tokens", async () => {
            function Stateful() { const [count, setCount] = React.useState(0); return <button id="theme-counter" onClick={() => setCount(count + 1)}>{count}</button>; }
            await mount(<FabricHostThemeProvider workloadClient={client}><Stateful /></FabricHostThemeProvider>, true);
            await until(() => area.querySelector("#theme-counter"), "Host theme did not resolve.");
            await click(area.querySelector("#theme-counter"));
            await act(async () => hostChange!({ ...hostTheme(true), tokens: { ...hostTheme(true).tokens, colorNeutralBackground1: "#101010" } }));
            assert(area.querySelector("#theme-counter")!.textContent === "1", "Theme change reset preparation state.");
            assert(getComputedStyle(area.querySelector(".fui-FluentProvider")!).getPropertyValue("--colorNeutralBackground1").trim() === "#101010", "Host tokens were not applied.");
        }],
    ["Assignment retained no-ID models cannot bypass Fields, Filters or Create boundaries", async () => {
            for (const component of [HubModelFieldsStep, HubModelScopeStep, HubModelReviewStep, HubModelAssignmentStep]) {
                const seed = context(true);
                const { id: _id, ...old } = tm;
                seed.hubModel = { ...seed.hubModel, id: undefined, document: old };
                const before = calls.length;
                await mount(<Harness component={component} initial={seed}/>);
                assert(area.textContent?.includes("doesn't have an ID") && button("Return to Device Model"), "A retained old model bypassed ID recovery.");
                assert(!button("Create modeled data") && !button("Check coverage") && !button("Assign to selected devices") && calls.length === before,
                    "An old no-ID model enabled operations.");
            }
        }],
    ["Assignment source receipts survive real wizard navigation during limited Hub checking", async () => {
            const before = consentRequests;
            const original = HubAssignmentApi.prototype.discoverHubs, observed: unknown[] = [];
            const originalCreate = React.createElement, readiness: unknown[] = [];
            const queryStart = assignmentQueries.length, clock = performance.now();
            const wrappers = new WeakMap<Function, Function>(), snapshots = new WeakMap<Function, Record<string, unknown>>();
            let registeredAssignment: React.ComponentType<WizardStepProps> | undefined;
            let modelAccessor: (() => SourceReadiness) | undefined, assignmentAccessor: (() => SourceReadiness) | undefined;
            const tracedCreate = new Proxy(originalCreate, {
                apply(target, thisArg, args) {
                    const [type, props] = args;
                    if (type === WizardControl && Array.isArray(props?.steps)) {
                        registeredAssignment = props.steps.find((step: { id: string }) => step.id === "assign-references")?.component;
                    }
                    if (type === ModelPreparationStep && props?.wizardContext && typeof props.getSourceState === "function") {
                        modelAccessor = props.getSourceState;
                    }
                    if (type === HubModelAssignmentStep && typeof props?.getSourceState === "function") {
                        const get = props.getSourceState;
                        assignmentAccessor = get;
                        if (props.wizardContext) snapshots.set(get, { snapshot: props.sourceState, metadataKey: props.wizardContext.hubAssignmentSource?.key,
                            modelId: props.wizardContext.hubModel?.id, modelRevision: props.wizardContext.hubModelRevision || 0 });
                        let wrapped = wrappers.get(get);
                        if (!wrapped) {
                            wrapped = function (this: unknown, ...arguments_: unknown[]) {
                                try {
                                    const result = Reflect.apply(get, this, arguments_);
                                    readiness.push({ atMs: performance.now() - clock, state: result, ...snapshots.get(get) });
                                    return result;
                                } catch (error: unknown) {
                                    readiness.push({ atMs: performance.now() - clock, threw: error instanceof Error ? error.message : String(error), ...snapshots.get(get) });
                                    throw error;
                                }
                            };
                            wrappers.set(get, wrapped);
                        }
                        args = [type, { ...props, getSourceState: wrapped }, ...args.slice(2)];
                    }
                    return Reflect.apply(target, thisArg, args);
                },
            });
            try {
                assert(Reflect.set(React, "createElement", tracedCreate), "The real readiness observer could not be installed.");
                HubAssignmentApi.prototype.discoverHubs = function (source, signal, progress) {
                    observed.push(source);
                    return original.call(this, source, signal, progress);
                };
                await wizard(context(true));
                assert(registeredAssignment && modelAccessor, "The real registered factory or live model accessor was not observed.");
                const accessor = modelAccessor!, beforeProbe = accessor();
                assert(beforeProbe.ready, "The real source/model view did not have current source readiness.");
                const probe = Reflect.apply(registeredAssignment!, undefined, []);
                const afterProbe = accessor();
                assert(React.isValidElement(probe) && assignmentAccessor === accessor && beforeProbe === afterProbe &&
                    afterProbe.ready && afterProbe.key === beforeProbe.key && afterProbe.revision === beforeProbe.revision,
                    "A no-props development probe replaced the real live wizard/source owner.");
                await click(next());
                await until(() => area.querySelector("h2")?.textContent === "Assign the model to devices", "Assignment was not reached.");
                assert(observed.length === 0 && next()!.disabled && button("Skip assignment"),
                    "Assignment started discovery or enabled continuation before the explicit WoT decision.");
                await beginAssignmentTargets();
                await until(() => observed.length > 0, () => `Real discovery never entered. readiness=${JSON.stringify(readiness)}; actual=${area.textContent}`);
                await until(() => area.textContent?.includes("source Hub could not be resolved") && !area.querySelector('[role="progressbar"]'),
                    () => `Real discovery did not reach its expected terminal state. args=${JSON.stringify(observed)}; readiness=${JSON.stringify(readiness)}; actual=${area.textContent}`);
                assert(snapshots.get(accessor)?.modelId === tm.id && area.textContent?.includes("source Hub could not be resolved"),
                    `Source metadata or exact model ID did not reach assignment. ExpectedID=${tm.id}; REALdiscoverArgs=${JSON.stringify(observed)}; actual=${area.textContent}`);
                assert(assignmentQueries.slice(queryStart).some(query => query.includes("fixture-hub.azure-devices.net") && query.includes("properties.hostName")),
                    "The real source-scoped ARG query did not use the selected tables' Hub hostname.");
                caseObservations = { noPropsProbePreservedAccessorIdentity: assignmentAccessor === accessor,
                    noPropsProbePreservedStateIdentity: beforeProbe === afterProbe,
                    beforeProbe: { key: beforeProbe.key, ready: beforeProbe.ready, revision: beforeProbe.revision },
                    afterProbe: { key: afterProbe.key, ready: afterProbe.ready, revision: afterProbe.revision },
                    realDiscoveryArgs: observed, actualQueries: assignmentQueries.slice(queryStart), getterTrace: readiness,
                    modelId: tm.id, terminalState: "Source Hub could not be resolved (expected empty synthetic ARG page)" };
                assert(!next()!.disabled && !button("Assign to selected devices") && consentRequests === before &&
                    currentAssignmentIntent(currentContext, accessor())?.mode === "assign",
                    "Unresolved metadata enabled assignment writes, erased the explicit decision or triggered consent.");
                await click(next());
                assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage", "Assignment could not be skipped into filters.");
            } finally { HubAssignmentApi.prototype.discoverHubs = original; Reflect.set(React, "createElement", originalCreate); }
        }],
    ["Assignment ADR-only discovery hides DPS information until success and never requests DPS through refresh or owner changes", async () => {
            const requests: string[] = [];
            const beforeConsent = consentRequests, scopeStart = frontendScopeRequests.length;
            try {
                assignmentResponse = adrOnlyResponse(requests);
                await openAssignmentTargets();
                await until(() => area.textContent?.includes("No confirmed unassigned devices") && !area.querySelector('[role="progressbar"]'),
                    "Real ADR-only discovery did not settle.");
                await chooseAssignmentView("all");
                assert(area.textContent?.includes("Assignment not confirmed") && !area.querySelector('input[type="checkbox"]') &&
                    button("Assign to selected devices")?.disabled, "An unknown ADR reference became selectable.");
                assert(!area.querySelector('[role="tab"]') && !area.querySelector("#assignment-service"), "A retired DPS inventory control remains.");
                assertDpsGuidance(false);
                const beforeRefresh = requests.length;
                await click(button("Refresh devices"));
                await until(() => requests.length > beforeRefresh && !area.querySelector('[role="progressbar"]'), "ADR refresh did not repeat source checks.");
                const beforeOwner = requests.length;
                await act(async () => setTestSourceState({ ...sourceFixture, revision: 2 }));
                assert(!currentContext.hubAssignmentIntent && requests.length === beforeOwner &&
                    !area.querySelector('[aria-label="Device model assignments"]'), "New owner silently resumed assignment discovery.");
                await beginAssignmentTargets();
                await until(() => requests.length > beforeOwner && !area.querySelector('[role="progressbar"]'), "Owner change did not reload current ADR metadata.");
                assert(frontendScopeRequests.slice(scopeStart).every(scopes => JSON.stringify(scopes) === JSON.stringify(["https://management.azure.com/.default"])) &&
                    consentRequests === beforeConsent, "An informational DPS interaction requested a DPS/OBO token.");
                assert(requests.every(request => !/provisioningServices|azure-devices-provisioning|enrollment/i.test(request)),
                    "A DPS request escaped the ADR-only workflow.");
                assertDpsGuidance(false);
                caseObservations = { realAdrDiscovery: true, dpsHiddenUntilSuccess: true, informationalRequests: 0,
                    dpsTokens: 0, dpsDataRequests: 0, ownerReloadArmOnly: true };
            } finally {
                assignmentResponse = undefined;
            }
        }],
    ["Assignment stale enrollment choices cannot enable ADR actions or success-only DPS information", async () => {
            const discover = HubAssignmentApi.prototype.discover;
            try {
                HubAssignmentApi.prototype.discover = async () => assignmentInventory();
                const seed = assignmentContext();
                seed.hubAssignmentChoices = { owner: assignmentOwner(seed.hubAssignmentSource, sourceFixture.revision, seed.hubModel) + ":0",
                    targets: [{ ...assignmentInventory().targets[0], kind: "group", id: "stale-group", reference: { kind: "unassigned", ids: [] },
                        writeBlocker: undefined }] };
                const before = tokenRequests;
                await openAssignmentTargets(seed);
                await until(() => area.textContent?.includes("0 selected") && !area.querySelector('[role="progressbar"]'), "Stale enrollment choice was retained.");
                assert(button("Assign to selected devices")?.disabled && !area.querySelector('input[type="checkbox"]') &&
                    !area.textContent?.includes("stale-group"), "A retained enrollment became an ADR action.");
                assertDpsGuidance(false);
                assert(tokenRequests === before, "Informational guidance requested a token.");
                caseObservations = { staleEnrollmentSelectionRemoved: true, assignmentDisabled: true, dpsHiddenUntilSuccess: true, dpsTokens: 0 };
            } finally {
                HubAssignmentApi.prototype.discover = discover;
            }
        }],
    ["Assignment current canonical references override stale device indexes without querying them", async () => {
        const observations: object[] = [];
        for (const [currentRefs, indexedRefs, indexStatus] of [
            [[], [tm.id], 200],
            [[tm.id], [], 200],
            [["urn:fixture:other"], [tm.id], 200],
            [[], [], 503],
        ] as [string[], string[], number][]) {
            const requests: string[] = [], state = { patches: [] as { body: unknown; ifMatch: string | null }[] };
            const respond = adrOnlyResponse(requests, state);
            let deviceIndexQueries = 0, hubHostnameQueries = 0, pointReads = 0;
            assignmentResponse = async (input, init) => {
                if (String(input).includes("ResourceGraph")) {
                    const query = JSON.parse(String(init?.body)).query as string;
                    if (query.includes("registrydevices")) {
                        deviceIndexQueries++;
                        return new Response(JSON.stringify({ data: [{ id: assignmentInventory().targets[0].id,
                            propertyFieldNames: ["deviceTypeRefs"], reference: indexedRefs }] }), { status: indexStatus });
                    }
                    assert(query.includes("microsoft.devices/iothubs") && query.includes("fixture-hub.azure-devices.net"),
                        "Canonical reference discovery lost hostname-bound Hub resolution.");
                    hubHostnameQueries++;
                }
                const response = await respond(input, init);
                if (new URL(String(input)).pathname === assignmentInventory().targets[0].id) {
                    assert(!init?.method || init.method === "GET", "Reference-only inspection must not send a mutation.");
                    pointReads++;
                    const body = await response.json();
                    body.properties.deviceTypeRefs = currentRefs;
                    return new Response(JSON.stringify(body));
                }
                return response;
            };
            const seed = assignmentContext();
            delete seed.hubAssignmentHubSummary;
            await openAssignmentTargets(seed);
            await until(() => currentContext.hubAssignmentInventory?.value?.targets.length === 1 &&
                !area.querySelector('[role="progressbar"]'), "Canonical reference inventory did not settle.");
            const target = currentContext.hubAssignmentInventory.value.targets[0] as AssignmentTarget;
            assert(JSON.stringify(target.reference) === JSON.stringify(currentRefs.length
                ? { kind: "assigned", ids: currentRefs } : { kind: "unassigned", ids: [] }) &&
                target.referenceEvidence?.source === "adr-point-get" && target.referenceEvidence.rawCarrier === "present" &&
                target.referenceEvidence.apiVersion === "2026-11-02-preview" &&
                target.referenceEvidence.currentArmEtag === target.etag && !!target.deviceIdentity,
                "A stale device index replaced current canonical references or provenance.");
            assert(deviceIndexQueries === 0 && hubHostnameQueries === 1 && pointReads === 1 && state.patches.length === 0,
                "Device references queried ARG, bypassed the exact GET or mutated a device during discovery.");
            if (currentRefs.length) assert(existingAssignmentsNotice() &&
                !area.querySelector('input[type="checkbox"]') && !button("Assign to selected devices"),
                "Just-assigned canonical references remained selectable because the device index lagged.");
            else assert(assignmentBlocker(target) === undefined &&
                !!area.querySelector('input[aria-label="Select Fixture device"]') && !existingAssignmentsNotice(),
                "Just-cleared canonical references stayed read-only because the device index was stale or unavailable.");
            observations.push({ currentRefs, indexedRefs, indexStatus, deviceIndexQueries, hubHostnameQueries, pointReads,
                currentCanonicalAuthority: true, noMutations: true });
        }
        caseObservations = { observations, actualCloudWrites: 0 };
    }],
    ["Assignment authoritative unassigned provenance permits one conditional write and operation-only completion without DPS", async () => {
            const requests: string[] = [], state: { modelId?: string; patches: { body: unknown; ifMatch: string | null }[] } = { patches: [] };
            const scopeStart = frontendScopeRequests.length;
            const clock = assignmentClock();
            try {
                assignmentResponse = adrOnlyResponse(requests, state);
                await openAssignmentTargets();
                await until(() => !!area.querySelector('input[aria-label="Select Fixture device"]') && !area.querySelector('[role="progressbar"]'),
                    "A verified current full GET with genuinely omitted refs did not enable its first-assignment candidate.");
                const target = currentContext.hubAssignmentInventory.value.targets[0];
                assert(target.reference.kind === "unassigned" && target.referenceEvidence.source === "adr-point-get" &&
                    target.referenceEvidence.rawCarrier === "missing" && target.referenceEvidence.apiVersion === "2026-11-02-preview",
                    "Logical unassigned classification lost its authoritative raw-missing provenance.");
                await click(area.querySelector('input[aria-label="Select Fixture device"]'));
                await click(button("Assign to selected devices"));
                let dialog = document.querySelector('[role="dialog"]');
                assert(dialog?.querySelector(".fui-DialogTitle")?.textContent === `Assign ${currentContext.hubModel.title} to 1 device?` &&
                    dialog.querySelector(".fui-Text")?.textContent === `Model ID: ${tm.id}` &&
                    !dialog.textContent?.includes("selected unassigned devices") && !dialog.textContent?.includes("ADR devices:") &&
                    !dialog.textContent?.includes("Existing model references will not be changed.") &&
                    dialog.querySelector("li")?.textContent?.trim() === "Fixture device \u00b7 Fixture Hub" &&
                    !dialog.textContent?.includes("(Devices)") && !dialog.textContent?.includes(target.namespaceId.split("/").at(-1)),
                    "The real first-assignment confirmation lost the exact reference/scope.");
                await click([...dialog!.querySelectorAll("button")].find(node => node.textContent === "Cancel"));
                assert(state.patches.length === 0 && document.activeElement === button("Assign to selected devices"),
                    "Copy-only confirmation edit dispatched on Cancel or lost the return-focus ref.");
                await click(button("Assign to selected devices"));
                dialog = document.querySelector('[role="dialog"]');
                await click([...dialog!.querySelectorAll("button")].find(node => node.textContent === "Assign"));
                await until(() => state.patches.length === 1 &&
                    currentContext.hubAssignmentHistory.some((receipt: { state: string }) => receipt.state === "accepted") &&
                    !document.querySelector('[role="dialog"]') && !area.querySelector("table"),
                    "The actual adapter did not retain a truthful accepted202 progress receipt.");
                await until(() => clock.pending.size === 1, "The real adapter result did not schedule current readback.");
                const beforePoll = requests.length;
                await clock.tick();
                await until(() => !!completedAssignmentList() &&
                    !area.querySelector("table") && !area.querySelector('[role="progressbar"]'),
                    "The original operation did not complete while the ARG carrier remained missing.");
                assert(state.patches.length === 1 && !button("Assign to selected devices") && !button("Refresh results"),
                    "Readback repeated/replaced a stored reference or kept it selectable.");
                assert(requests.length === beforePoll + 1 && requests[beforePoll].includes("/asyncOperationStatuses/fixture-assignment?api-version=2026-11-02-preview") &&
                    clock.delays.every(delay => delay >= 1900 && delay <= 4000),
                    "Completion did not use only the original operation with the first jittered polling interval.");
                assert(requests.every(url => !/provisioningServices|azure-devices-provisioning|enrollment/i.test(url)) &&
                    frontendScopeRequests.slice(scopeStart).every(scopes => JSON.stringify(scopes) === JSON.stringify(["https://management.azure.com/.default"])),
                    "The ADR classifier invoked DPS auth/data transport.");
                caseObservations = { authoritativeUnsetRawCarrier: "missing", argCarrierStillMissing: true, guardedPatchCount: state.patches.length,
                    acceptedThenOperationConfirmed: true, postWriteResourceReads: 0,
                    jitteredInitialCadence: true, finalCheckboxes: 0, dpsRequests: 0 };
            } finally { await mount(<div>Guarded fixture closed.</div>); clock.restore(); assignmentResponse = undefined; }
        }],
    ["Assignment frontend Sign in retries the failed resource through discovery without a backend OBO request or automatic loop", async () => {
            const restore = HubAssignmentApi.prototype.discover;
            try {
                assignmentAuthFailure = true;
                const before = consentRequests;
                const scopeStart = frontendScopeRequests.length;
                await openAssignmentTargets();
                await until(() => !!button("Sign in"), "The token failure did not expose sign-in recovery.");
                const attempts = tokenRequests;
                await settle();
                assert(tokenRequests === attempts && consentRequests === before, "A background authentication retry/OBO loop started.");
                assignmentAuthFailure = false;
                await click(button("Sign in"));
                await until(() => tokenRequests > attempts && area.textContent?.includes("source Hub could not be resolved") &&
                    !area.querySelector('[role="progressbar"]'), "Frontend Sign in did not actually repeat discovery to its expected empty fixture result.");
                assert(consentRequests === before, "Frontend Sign in requested a backend/OBO token.");
                assert(JSON.stringify(frontendScopeRequests.slice(scopeStart)) === JSON.stringify([
                    ["https://management.azure.com/.default"], ["https://management.azure.com/.default"],
                ]), "The failed ARM frontend-resource request was not retried exactly.");
            } finally { assignmentAuthFailure = false; HubAssignmentApi.prototype.discover = restore; }
        }],
    ["Assignment continuation keeps loaded rows and access errors are not empty success", async () => {
            const discover = HubAssignmentApi.prototype.discover, more = HubAssignmentApi.prototype.loadMore;
            try {
                HubAssignmentApi.prototype.discover = async () => ({ ...assignmentInventory(), hasMore: true });
                HubAssignmentApi.prototype.loadMore = async (previous, signal, publication) => {
                    assert(!signal.aborted && publication?.isCurrent(), "A stale continuation reached publication.");
                    const value = { ...previous, hasMore: false,
                        targets: [...previous.targets, { ...previous.targets[0], id: `${previous.targets[0].id}-second`, name: "Second fixture" }] };
                    publication?.publish(value); return value;
                };
                await openAssignmentTargets();
                await until(() => !!button("Load more"), "The continuation was not offered.");
                await chooseAssignmentView("all");
                await click(button("Load more"));
                await until(() => area.textContent?.includes("Second fixture"), "The next page was not retained.");
                assert(area.textContent?.includes("Fixture device") &&
                    area.querySelectorAll("tr.iot-solution-assignment-device").length === 2 &&
                    currentContext.hubAssignmentInventory.value.targets.length === 2 && !button("Load more"),
                    "Loaded rows or final-page state were lost.");
                HubAssignmentApi.prototype.discover = async () => { throw new Error("Synthetic request failure."); };
                await click(button("Refresh devices"));
                await until(() => area.textContent?.includes("Synthetic request failure"), "Refresh failure was hidden.");
                assert(area.textContent?.includes("Second fixture") && button("Assign to selected devices")?.disabled, "A failed refresh discarded prior rows or enabled writes.");
            } finally { HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.loadMore = more; }
        }],
    ["Assignment first-write confirmation preserves Cancel existing references and operation-only success", async () => {
            const discover = HubAssignmentApi.prototype.discover, write = HubAssignmentApi.prototype.patchReference,
                recheck = HubAssignmentApi.prototype.recheck;
            let writes = 0, assigned = false;
            try {
                HubAssignmentApi.prototype.discover = async () => {
                    const inventory = assignmentInventory();
                    inventory.targets[0] = { ...inventory.targets[0], writeBlocker: undefined,
                        reference: assigned ? { kind: "assigned", ids: [tm.id] } : { kind: "unassigned", ids: [] } };
                    inventory.targets.push({ ...inventory.targets[0], id: `${inventory.targets[0].id}-existing`,
                        name: "Existing fixture", reference: { kind: "assigned", ids: ["urn:other:fixture"] } });
                    return inventory;
                };
                HubAssignmentApi.prototype.patchReference = async (target, id, _signal, _source, onDispatch) => {
                    assert(target.reference.kind === "unassigned" && target.etag === '"synthetic-etag"' && id === tm.id,
                        "The operation changed an existing reference or lost its condition.");
                    onDispatch?.(); writes++; assigned = true; return "completed";
                };
                HubAssignmentApi.prototype.recheck = async target => ({ ...target,
                    reference: assigned ? { kind: "assigned", ids: [tm.id] } : { kind: "unassigned", ids: [] } });
                await openAssignmentTargets();
                await until(() => !!area.querySelector('input[type="checkbox"]'), "Verified synthetic target was not offered.");
                await chooseAssignmentView("all");
                assert(area.querySelectorAll('input[type="checkbox"]').length === 1 &&
                    area.textContent?.includes("Existing fixture") && area.textContent?.includes("urn:other:fixture"),
                    "The initial-only confirmation flow made an existing reference selectable.");
                await click(area.querySelector('input[type="checkbox"]'));
                await click(button("Assign to selected devices"));
                let dialog = document.querySelector('[role="dialog"]');
                assert(dialog?.querySelector(".fui-DialogTitle")?.textContent === `Assign ${currentContext.hubModel.title} to 1 device?` &&
                    dialog.querySelector(".fui-Text")?.textContent === `Model ID: ${tm.id}` &&
                    !dialog.textContent?.includes("ADR devices:") && !dialog.textContent?.includes("Existing model references will not be changed."),
                    "Confirmation omitted the exact scope/reference.");
                await click([...dialog!.querySelectorAll("button")].find(node => node.textContent === "Cancel"));
                assert(writes === 0 && document.activeElement === button("Assign to selected devices"), "Cancel dispatched or lost focus.");
                await click(button("Assign to selected devices"));
                dialog = document.querySelector('[role="dialog"]');
                await click([...dialog!.querySelectorAll("button")].find(node => node.textContent === "Assign"));
                await until(() => !!completedAssignmentList(),
                    "Synchronous operation completion did not replace progress with verified success.");
                assert(writes === 1 && !area.querySelector("table") && !button("Assign to selected devices") &&
                    currentContext.hubAssignmentHistory.some((receipt: { state: string; message?: string }) =>
                        receipt.state === "confirmed" && receipt.message ===
                            "The assignment operation completed. Current device state and matching data have not been verified."),
                    "The result claimed adoption, repeated a write or retained assignable controls.");
            } finally {
                HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.patchReference = write;
                HubAssignmentApi.prototype.recheck = recheck;
            }
        }],
    ["Assignment token-wait disposal retires the owned read and late completions cannot replace a remount", async () => {
            const restore = HubAssignmentApi.prototype.discover;
            try {
                deferredToken = true;
                await openAssignmentTargets();
                await until(() => tokenReleases.length > 0, "Assignment token wait was not reached.");
                await act(async () => setStepVisible(false));
                deferredToken = false;
                HubAssignmentApi.prototype.discover = async () => assignmentInventory();
                await act(async () => setStepVisible(true));
                await until(() => area.textContent?.includes("No confirmed unassigned devices"), "Retained context did not recover.");
                const before = area.textContent;
                await act(async () => tokenReleases.splice(0).forEach(release => release()));
                await settle();
                assert(area.textContent === before && !currentContext.hubAssignmentBusy, "A cancelled token result overwrote the current inventory.");
            } finally { deferredToken = false; HubAssignmentApi.prototype.discover = restore; }
        }],
    ["Assignment owned response-wait cancellation preserves the retained context without stale publication", async () => {
            const restore = HubAssignmentApi.prototype.discover;
            let release!: (value: AssignmentInventory) => void;
            try {
                HubAssignmentApi.prototype.discover = () => new Promise(resolve => { release = resolve; });
                await openAssignmentTargets();
                await until(() => !!release, "The scoped read wait was not reached.");
                await act(async () => setStepVisible(false));
                HubAssignmentApi.prototype.discover = async () => assignmentInventory();
                await act(async () => setStepVisible(true));
                await until(() => area.textContent?.includes("No confirmed unassigned devices"), "The current remount did not settle.");
                const old = assignmentInventory(); old.targets[0].name = "Stale result must not appear";
                await act(async () => release(old));
                await chooseAssignmentView("all");
                assert(!area.textContent?.includes("Stale result must not appear") && area.textContent?.includes("Fixture device"),
                    "An old instance published after its owned request was cancelled.");
            } finally { HubAssignmentApi.prototype.discover = restore; }
        }],
    ["Data Agent actual wizard matches six screenshot copy references and creation visibility without changing generated payloads", () => observeWizardContext(async () => {
            const savedFetch = window.fetch, key = "iot-wizard-v4-dataagent-copy-restoration";
            const savedStorage = sessionStorage.getItem(key), source = dataAgentSourceFixture("wot");
            const phases: { phase: string; text: string }[] = [];
            // Independently transcribed surrounding copy from the six user screenshots, not the public source snapshot.
            const reference = {
                sidebar: [
                    ["Overview", "Ask questions grounded in device operations"],
                    ["Modeled data source", "Select the device data that grounds the agent"],
                    ["Configure Data Agent", "Review instructions and create the agent"],
                    ["Ready to use", "Review the agent and example questions"],
                ],
                overview: {
                    heading: "Explore device operations using AI",
                    description: "Create a Fabric Data Agent that translates natural-language questions into queries over your modeled device data, then returns answers and operational insights.",
                    sectionHeadings: ["Why modeled data improves the agent", "How the agent produces answers", "What you will need"],
                    benefits: [
                        "Business-friendly field names, descriptions, types, units, and component context help the agent understand what each measurement represents.",
                        "Each row combines real-time telemetry with the device's latest modeled property values, allowing questions to correlate live behavior with current configuration and operational state.",
                        "A consistent typed schema reduces ambiguity and helps the agent produce more relevant queries, explanations, comparisons, and summaries.",
                    ],
                    flowTitles: ["Modeled device data", "AI instructions", "Fabric Data Agent", "Answers and insights"],
                    flowDetails: ["Real-time telemetry and current device state", "Field meaning, query guidance, and examples",
                        "Interprets questions and generates queries", "Trends, comparisons, state, and anomalies"],
                    requirements: [
                        "A modeled dataset created through the Model your device data experience. This supplies the typed telemetry, current device properties, and model metadata used to ground the agent.",
                        "Device events in the modeled data table so the agent can answer questions using current and historical operational data.",
                        "Permission to read the KQL database and create a Data Agent in this Fabric workspace.",
                    ],
                },
                source: {
                    heading: "Select modeled data source",
                    description: "Select the modeled data table that will ground the agent's answers. The table contains typed real-time device telemetry enriched with current device state from modeled properties.",
                    info: "Modeled data gives the agent consistent field names, data types, descriptions, units, and device-property context, helping it generate more relevant queries and better operational answers.",
                    fieldLabels: ["Eventhouse", "KQL database", "Modeled data table"],
                    validationLabels: ["Selected source", "Modeled fields"],
                },
                configure: {
                    heading: "Configure Data Agent",
                    description: "Name the agent and review the generated guidance that combines your modeled schema, semantic metadata, query rules, and example questions.",
                    validationLabels: ["Grounding data", "Available fields"],
                    nameLabel: "Data Agent name",
                    instructionsLabel: "AI instructions (auto-generated)",
                    hint: "These instructions become part of the Data Agent and teach it how to interpret the modeled fields, distinguish telemetry from device state, apply time and device filters, and generate appropriate KQL. The included example questions and queries provide additional patterns for producing relevant answers.",
                    createdNote: "Default resource names are provided below. You can customize them if needed.",
                    createdMessage: "Data Agent created. Select Next for usage guidance and example operational questions.",
                },
                summary: {
                    heading: "Your IoT Data Agent is ready",
                    description: "The agent is grounded in the real-time telemetry and current device state stored in the modeled data table. Its generated instructions and examples help it interpret the modeled fields and translate natural-language questions into queries.",
                    usageHeading: "How to use your Data Agent",
                    usageTail: ["Include a device, time range, measurement, property, or operational condition when you need a more specific answer.",
                        "Review the generated answer and query, then refine the question or agent instructions if business terminology needs more guidance."],
                    questionsHeading: "Example questions",
                    questions: [
                        "Which devices have reported high temperature while in cooling mode?",
                        "Compare average vibration by firmware version over the last 24 hours.",
                        "Which devices have low battery voltage and are currently active?",
                        "Show the latest measurements and reported state for a specific device.",
                        "Which assets show unusual behavior compared with similar devices?",
                    ],
                    configurationHeading: "Agent configuration",
                },
            };
            let parts: { path: string; payload: string; payloadType: string }[] | undefined;
            const clean = (value: string) => value.replace(/\s+/g, " ").trim();
            const content = (selector: string) => {
                const node = area.querySelector(selector);
                assert(node, `Missing screenshot-reference element ${selector}.`);
                return clean(node!.textContent || "");
            };
            const contents = (selector: string) => [...area.querySelectorAll(selector)].map(node => clean(node.textContent || ""));
            const same = (actual: unknown, expected: unknown, label: string) =>
                assert(JSON.stringify(actual) === JSON.stringify(expected), `${label}: actual ${JSON.stringify(actual)}; reference ${JSON.stringify(expected)}`);
            const common = () => {
                same(contents(".wizard-control__step-title"), reference.sidebar.map(row => row[0]), "Screenshot sidebar titles");
                same(contents(".wizard-control__step-description"), reference.sidebar.map(row => row[1]), "Screenshot sidebar descriptions");
                assert(area.textContent?.includes("Set up Data Agent") && button("Previous"), "Screenshot wrapper/navigation text missing.");
            };
            const text = () => (area.querySelector(".iot-solution-step")!.textContent || "").replace(/\s+/g, " ").trim();
            window.fetch = async (input, init) => {
                const address = String(input), body = init?.body ? JSON.parse(String(init.body)) : undefined;
                if (address === "https://mock.kusto.fabric.microsoft.com/v1/rest/mgmt") {
                    if (body.csl === ".show tables details | project TableName, DocString") return result(
                        [["TableName", "string"], ["DocString", "string"]], [[source.table, source.docstring]]);
                    assert(body.csl === `.show table ['${source.table}'] schema as json`, "Copy restoration changed schema inspection.");
                    return result([["TableName", "string"], ["Schema", "string"]], [[source.table, JSON.stringify({
                        OrderedColumns: source.columns.map(column => ({ Name: column.name, CslType: column.type, DocString: column.docstring })),
                    })]]);
                }
                if (address === "https://api.fabric.microsoft.com/v1/workspaces/workspace/dataAgents" && init?.method === "POST") {
                    assert(body.displayName === "DataAgent_WoT_fixture_device_review", "Restoring copy changed declared-name default.");
                    parts = body.definition.parts;
                    return new Response(JSON.stringify({ id: "77777777-7777-4777-8777-777777777777" }), { status: 201 });
                }
                return savedFetch(input, init);
            };
            try {
                sessionStorage.setItem(key, JSON.stringify({ currentStepId: "overview", completedSteps: [], wizardContext: {
                    agentEhId: "eh", agentEhName: "Eventhouse", agentDbId: "db", agentDbName: "Database",
                    agentTableName: source.table, agentQueryServiceUri: "https://mock.kusto.fabric.microsoft.com",
                    agentTableDocstring: source.docstring, agentSuffix: "review",
                } }));
                await mount(<IoTSolutionItemDataAgentWizard workloadClient={client}
                    item={{ id: "copy-restoration", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                    onWizardComplete={() => { throw new Error("No live/workload saves in copy restoration."); }}/>);
                common();
                same(content(".iot-solution-step h2"), reference.overview.heading, "Screenshot 1 overview heading");
                same(content(".iot-solution-step-description"), reference.overview.description, "Screenshot 1 intro");
                same(contents(".iot-solution-outcome-heading"), reference.overview.sectionHeadings, "Screenshot 1 section headings");
                same(contents(".iot-solution-outcome-section ul li"), reference.overview.benefits, "Screenshot 1 benefits");
                same(contents(".iot-solution-flow-node-title"), reference.overview.flowTitles, "Screenshot 1 flow titles");
                same(contents(".iot-solution-flow-node-detail"), reference.overview.flowDetails, "Screenshot 1 flow details");
                same(contents(".iot-solution-prerequisites li"), reference.overview.requirements, "Screenshot 1 prerequisites");
                same(content(".iot-solution-prerequisites strong"), "Model your device data", "Screenshot 1 strong markup");
                assert(!next()!.disabled, "Screenshot 1 Next must be available.");
                phases.push({ phase: "overview", text: text() });
                await click(next());
                await until(() => !!currentContext.agentSourceSelected && !next()!.disabled && !area.querySelector(".fui-Spinner"),
                    "Actual restored source page did not inspect native metadata.");
                common();
                same(content(".iot-solution-step h2"), reference.source.heading, "Screenshot 2 source heading");
                same(content(".iot-solution-step-description"), reference.source.description, "Screenshot 2 source intro");
                same(content(".fui-MessageBarBody"), reference.source.info, "Screenshot 2 info box");
                same(contents(".iot-solution-field-label").map(label => label.replace(/\s*\*$/, "")), reference.source.fieldLabels, "Screenshot 2 field labels");
                same(contents(".iot-solution-validation-label"), reference.source.validationLabels, "Screenshot 2 summary labels");
                same(contents(".iot-solution-validation-value"), [`Database / ${source.table}`, `${source.columns.length} typed columns available to the agent`],
                    "Screenshot 2 real source/count substitutions");
                phases.push({ phase: "source", text: text() });
                await click(next());
                await until(() => !!area.querySelector("pre"), "Actual configure preview did not appear.");
                common();
                same(content(".iot-solution-step h2"), reference.configure.heading, "Screenshot 3 configure heading");
                same(content(".iot-solution-step-description"), reference.configure.description, "Screenshots 3/4 configure intro");
                same(contents(".iot-solution-validation-label"), reference.configure.validationLabels, "Screenshot 3 grounding labels");
                same(contents(".iot-solution-field-label").map(label => label.replace(/\s*\*$/, "")),
                    [reference.configure.nameLabel, reference.configure.instructionsLabel], "Screenshots 3/4 name/instructions labels");
                same(content(".iot-solution-field-hint"), reference.configure.hint, "Screenshots 3/4 instructions help");
                assert(!area.querySelector(".iot-solution-resource-name-note"), "Screenshots 3/4 have no default-name note before creation.");
                assert(!area.querySelector<HTMLInputElement>("#agent-name")!.disabled && next()!.disabled && button("Create Data Agent"),
                    "Screenshots 3/4 pre-create name/control visibility differs.");
                const expected = buildDataAgentContext({ database: "Database", table: source.table, docstring: source.docstring,
                    columns: currentContext.agentColumns });
                const preview = expected.agentInstructions + "\n\n" + expected.sourceInstructions +
                    `\n## Examples (${expected.examples.length})\n\n` + dataAgentExamplesMarkdown(expected.examples);
                assert(area.querySelector("pre")!.textContent === preview && preview.includes('deviceId == "YOUR_DEVICE_ID"'),
                    "Copy restoration changed approved generated instructions/examples.");
                same(area.querySelector<HTMLInputElement>("#agent-name")!.value, "DataAgent_WoT_fixture_device_review", "Approved model-name substitution");
                phases.push({ phase: "configure", text: text().replace(preview.replace(/\s+/g, " ").trim(), "[GENERATED_PREVIEW]") });
                await click(button("Create Data Agent"));
                await until(() => !!parts && currentContext.agentCreated, "Actual mocked native save did not complete.");
                common();
                same(content(".iot-solution-resource-name-note"), reference.configure.createdNote, "Screenshot 5 created-only name note");
                same(content(".iot-solution-success-text"), reference.configure.createdMessage, "Screenshot 5 success message");
                same(contents(".iot-solution-creation-step-label"),
                    ["Build Data Agent definition", 'Create Data Agent "DataAgent_WoT_fixture_device_review"'], "Screenshot 5 dynamic creation labels");
                assert(area.querySelector<HTMLInputElement>("#agent-name")!.disabled && !area.querySelector("pre") &&
                    !area.querySelector(".iot-solution-field-hint") && !button("Create Data Agent") && !next()!.disabled,
                    "Screenshot 5 created state did not hide preview/button and retain the disabled name/available Next.");
                assert(area.querySelector(".iot-solution-success") && area.querySelectorAll(".iot-solution-creation-step-label.success").length === 2,
                    "Screenshot 5 successful native status markers missing.");
                phases.push({ phase: "configure-created", text: text() });
                const decoded = parts!.map(part => ({ path: part.path,
                    value: JSON.parse(new TextDecoder().decode(Uint8Array.from(atob(part.payload), character => character.charCodeAt(0)))) }));
                for (const stage of ["draft", "published"]) {
                    const prefix = `Files/Config/${stage}`, folder = `${prefix}/kusto-Database`;
                    const get = (path: string) => decoded.find(part => part.path === path)!.value;
                    assert(get(`${prefix}/stage_config.json`).aiInstructions === expected.agentInstructions &&
                        get(`${folder}/datasource.json`).dataSourceInstructions === expected.sourceInstructions &&
                        get(`${folder}/datasource.json`).elements[0].type === "table_grouping" &&
                        get(`${folder}/datasource.json`).elements[0].children[0].type === "kusto.table" &&
                        JSON.stringify(get(`${folder}/fewshots.json`).fewShots.map(({ question, query }: { question: string; query: string }) =>
                            ({ question, query }))) === JSON.stringify(expected.examples),
                        "Restoring surrounding copy altered native generation or saved fewshots.");
                }
                await click(next());
                await until(() => !!button("Finish and Close"), "Actual summary did not open.");
                const questions = reference.summary.questions;
                common();
                same(content(".iot-solution-step h2"), reference.summary.heading, "Screenshot 6 summary heading");
                same(content(".iot-solution-step-description"), reference.summary.description, "Screenshot 6 summary intro");
                same(content(".iot-solution-outcome-heading"), reference.summary.usageHeading, "Screenshot 6 usage heading");
                const usage = contents(".iot-solution-outcome-section > .iot-solution-outcome-list li");
                same(usage.slice(1), reference.summary.usageTail, "Screenshot 6 usage instructions");
                // User-authorized departure from screenshot 6: open the agent directly, remove the obsolete Copilot guide, and add the Fabric action.
                same(usage[0], "Open your data agent DataAgent_WoT_fixture_device_review, and ask a question about your devices or operations.",
                    "Approved Ready-page opening with the actual agent name");
                same(content(".iot-solution-config-card > strong"), reference.summary.questionsHeading, "Screenshot 6 questions heading");
                assert(contents(".iot-solution-outcome-heading").includes(reference.summary.configurationHeading) &&
                    !area.querySelector<HTMLImageElement>('img[alt="Fabric Copilot chat with the configured Data Agent"]') &&
                    !!button("Open Data Agent in Fabric"),
                    "Ready-page authorized guide removal/action or retained screenshot configuration heading differs.");
                assert(area.querySelector("h2")?.textContent?.trim() === "Your IoT Data Agent is ready" &&
                    text().includes("The agent is grounded in the real-time telemetry and current device state stored in the modeled data table. Its generated instructions and examples help it interpret the modeled fields and translate natural-language questions into queries.") &&
                    text().includes("Telemetry with device state") &&
                    JSON.stringify([...area.querySelector(".iot-solution-config-card")!.querySelectorAll("li")].map(node => node.textContent)) === JSON.stringify(questions) &&
                    text().includes("DataAgent_WoT_fixture_device_review") && text().includes(`Database / ${source.table}`) &&
                    text().includes(`${source.columns.length} modeled columns`),
                    "Actual summary did not restore V1 illustrative copy or lost real displayed context.");
                phases.push({ phase: "summary", text: text() });
                await click(button("Previous"));
                const currentStep = area.querySelector('[aria-current="step"]');
                assert(currentStep?.textContent?.includes("Configure Data Agent") &&
                    currentStep.querySelector(".wizard-control__step-circle--current") &&
                    currentStep.querySelector('[aria-label="Step 2"]'),
                    "Data Agent no longer uses the explicitly accepted shared current-step presentation.");
                caseObservations = { phases, screenshotReferenceChecklist: reference, allSixScreenshotReferencesCovered: true,
                    coveredReferences: ["1-overview", "2-source", "3-configure-before-top", "4-configure-before-bottom",
                        "5-configure-created", "6-summary"], createdNoteOnlyAfterCreation: true, actualRootAllFourPhases: true,
                    summaryScreenshotAuthorizedExceptions: ["Open your data agent opening", "obsolete Copilot guide removed", "Open Data Agent in Fabric action"],
                    actualGeneratedPreview: preview, nativeParts: decoded, exactGeneratedPreviewAndPayloadPreserved: true,
                    summaryIllustrativeQuestions: questions, generatedSamplesRemainSeparateAndSaved: true,
                    acceptedSharedCurrentStepPresentation: true, liveWrites: 0 };
            } finally {
                window.fetch = savedFetch;
                if (savedStorage === null) sessionStorage.removeItem(key); else sessionStorage.setItem(key, savedStorage);
            }
        })]
];
const originalFetch = window.fetch;
export async function runModelingRegressionChecks() {
    const results: {
        name: string;
        passed: boolean;
        error?: string;
        durationMs?: number;
        observations?: Record<string, unknown>;
    }[] = [];
    const parameters = new URLSearchParams(location.search);
    const filter = parameters.get("filter"), exact = parameters.getAll("exact-name");
    const selected = exact.length ? tests.filter(([name]) => exact.includes(name)) : tests.filter(([name]) => !filter || name.startsWith(filter));
    const reportSelection = document.getElementById("test-results")!;
    const selectionError = exact.some(name => !name) || new Set(exact).size !== exact.length ||
        exact.some(name => !tests.some(([registered]) => registered === name)) || !selected.length;
    if (selectionError) {
        results.push({ name: "Exact case selection", passed: false, error: "Unknown, duplicate, empty or zero-match test selection. No cases executed." });
        reportSelection.textContent = JSON.stringify(results);
        reportSelection.dataset.complete = "true";
        return results;
    }
    if (exact.length || parameters.has("selectionAck")) {
        reportSelection.dataset.selectedNames = JSON.stringify(selected.map(([name]) => name));
        reportSelection.dataset.selectionPending = "true";
        const deadline = performance.now() + 2000;
        while (reportSelection.dataset.selectionPending === "true" && performance.now() < deadline) await pause();
        if (reportSelection.dataset.selectionPending === "true") {
            results.push({ name: "Exact case selection", passed: false, error: "Exact selection was not acknowledged. No cases executed." });
            reportSelection.textContent = JSON.stringify(results); reportSelection.dataset.complete = "true"; return results;
        }
    }
    window.fetch = mockFetch;
    for (const [name, run] of selected) {
        document.getElementById("test-results")!.dataset.currentTest = name;
        document.getElementById("test-results")!.dataset.currentTestStartedAt = new Date().toISOString();
        document.getElementById("test-results")!.textContent = JSON.stringify(results);
        const startedAt = performance.now();
        caseObservations = undefined;
        try {
            await observeWizardContext(run);
            results.push({ name, passed: true, durationMs: performance.now() - startedAt, observations: caseObservations });
        }
        catch (error: unknown) {
            results.push({ name, passed: false, error: error instanceof Error ? error.stack || error.message : String(error),
                durationMs: performance.now() - startedAt, observations: caseObservations });
            if (new URLSearchParams(location.search).has("captureFailures")) {
                const report = document.getElementById("test-results")!;
                report.textContent = JSON.stringify(results);
                report.dataset.failurePending = "true";
                const deadline = performance.now() + 2000;
                while (report.dataset.failurePending === "true" && performance.now() < deadline) await pause();
            }
        }
        document.getElementById("test-results")!.textContent = JSON.stringify(results);
        deferredScoped = false;
        releases.splice(0).forEach(release => release());
        sampleState = "normal";
        deferredCreate = false;
        creationReleases.splice(0).forEach(release => release());
        creationBackend = undefined;
        deferredToken = false;
        tokenReleases.splice(0).forEach(release => release());
        emptyProcessing = false;
        emptyProcessingStream = undefined;
        coverageReadFailure = undefined;
        coverageSampleRows = 1;
        coverageReadFailureStatus = 403;
        managementStatusFailure = false;
        resourceResponse = undefined;
        tableMetadataResponse = undefined;
        eventhouseMetadataResponse = undefined;
        databaseItemIds = ["db"];
        emptySamples = "none";
        sourceSampleRows = 1; sourceSampleStatus = 0; deferSourceSamples = false;
        sourceSampleReleases.splice(0).forEach(release => release());
        sourceSampleSignals.splice(0);
        parityPayloads = undefined;
        assignmentResponse = undefined;
        assignmentAuthFailure = false;
        assignmentQueries.splice(0);
    }
    await act(async () => root?.unmount());
    root = undefined;
    window.fetch = originalFetch;
    if (errors.length)
        results.push({ name: "No uncaught component errors", passed: false, error: errors.join("; ") });
    const report = document.getElementById("test-results")!;
    report.textContent = JSON.stringify(results, null, 2);
    report.dataset.complete = "true";
    report.dataset.passed = String(results.every(result => result.passed));
    return results;
}
const run = document.createElement("button");
run.textContent = "Run isolated component regressions";
run.onclick = () => void runModelingRegressionChecks();
document.body.prepend(run);
const report = document.createElement("pre");
report.id = "test-results";
document.body.appendChild(report);
if (new URLSearchParams(location.search).get("run") === "1")
    void runModelingRegressionChecks();
const preview = new URLSearchParams(location.search).get("preview");
if (preview) {
    run.hidden = true;
    report.hidden = true;
    document.body.style.margin = "0";
    const notice = document.querySelector<HTMLParagraphElement>("body > p")!;
    notice.style.cssText = "margin:0;padding:12px;font:14px Segoe UI,sans-serif;";
    area.style.height = "calc(100vh - 44px)";
    window.fetch = mockFetch;
    void (async () => {
        if (preview.startsWith("assignment-warning-")) {
            const proof = await assignmentWarningPresentation(preview.endsWith("-dark"));
            document.documentElement.dataset.assignmentWarningProof = JSON.stringify(proof);
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("model-coverage-followup-")) {
            const proof = await coverageResultFollowupPresentation(preview === "model-coverage-followup-reference-results"
                ? "reference" : preview === "model-coverage-followup-skipped-results" ? "skipped"
                    : preview === "model-coverage-followup-dtdl-results" ? "dtdl" : "data");
            document.documentElement.dataset.coverageFollowupProof = JSON.stringify(proof);
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("assignment-spacing-")) {
            const outcome = preview.slice("assignment-spacing-".length);
            assert(["active", "confirmed", "partial", "unknown"].includes(outcome), "Unexpected assignment spacing preview.");
            const proof = await assignmentSpacingPresentation(outcome as "active" | "confirmed" | "partial" | "unknown");
            document.documentElement.dataset.assignmentSpacingProof = JSON.stringify(proof);
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("model-headings-")) {
            await observeWizardContext(async () => {
                await wizard(context(preview.endsWith("-wot")));
                if (preview.endsWith("-wot")) {
                    await file("thing-model-upload", JSON.stringify(headingPresentationDocument), "model-heading-and-role-example.json");
                    await until(() => currentContext.hubModel?.fields.length === 4 &&
                        area.querySelectorAll(".iot-solution-selection-cue").length === 2,
                        "The missing-role presentation fixture did not finish its local upload.");
                }
                await until(() => !!area.querySelector("#model-fields-heading"), "The model heading preview did not settle.");
                const before = canonicalJson(currentContext), reads = calls.length + resourceCalls.length, tokens = tokenRequests;
                const typography = assertModelPreparationHeadingTypography();
                await settle();
                assert(canonicalJson(currentContext) === before && calls.length + resourceCalls.length === reads && tokenRequests === tokens,
                    "Heading/spacing presentation changed format/model/fields or dispatched source/auth work.");
                document.documentElement.dataset.modelHeadingTypography = JSON.stringify(typography);
            });
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("model-creation-")) {
            const proof = await creationPresentationCheck(preview === "model-creation-complete" ? "complete" : "form");
            document.documentElement.dataset.creationPresentationProof = JSON.stringify(proof);
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview === "model-ready") {
            const proof = await creationPresentationCheck("ready");
            document.documentElement.dataset.readyPresentationProof = JSON.stringify(proof);
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("model-diagnostic-")) {
            await observeWizardContext(async () => {
                const seed = context(true); delete seed.hubModel; seed.hubFields = []; seed.preparationFileName = "";
                await wizard(seed);
                const reads = calls.length + resourceCalls.length, tokens = tokenRequests;
                if (preview === "model-diagnostic-json") {
                    await file("thing-model-upload", "{", "device-model-with-a-long-filename-for-readable-JSON-error-wrapping.json");
                } else if (preview === "model-diagnostic-missing-id") {
                    const { id: _id, ...document } = tm;
                    await file("thing-model-upload", JSON.stringify(document), "model-without-an-identifier.json");
                } else if (preview === "model-diagnostic-warnings") {
                    await file("thing-model-upload", JSON.stringify(diagnosticWarningDocument), "unsupported-model-parts.json");
                    await until(() => !!area.querySelector("#model-upload-warnings") && !!acceptedModelTile(),
                        "The diagnostic warning preview did not finish loading its accepted fields.");
                } else {
                    await click(area.querySelector('input[value="dtdl"]'));
                    await file("thing-model-upload", JSON.stringify(dtdl), "read-only-declared-roles.json");
                    await assertDtdlRoleInformation(true);
                }
                if (["model-diagnostic-json", "model-diagnostic-missing-id"].includes(preview)) await until(() =>
                    !!area.querySelector("#model-upload-error-file") && !!area.querySelector("#model-upload-error-message"),
                    "The diagnostic error preview did not finish its actual file read.");
                assert(calls.length + resourceCalls.length === reads && tokenRequests === tokens,
                    "Local diagnostic presentation started source/auth work.");
            });
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("model-coverage-")) {
            const discover = HubAssignmentApi.prototype.discover;
            try {
                sourceHubFixture(preview === "model-coverage-unlinked" ? "omitted" : "linked");
                HubAssignmentApi.prototype.discover = async () => assignmentInventory();
                await observeWizardContext(async () => {
                    const seed = preview === "model-coverage-positive-results" ? context(true) : coverageMissingFixture();
                    seed.hubScope = undefined;
                    if (["model-coverage-positive-results", "model-coverage-results", "model-coverage-mixed-results", "model-coverage-dense-results"].includes(preview)) {
                        seed.rawTelemetryTable = coveragePreviewTelemetryTable;
                        seed.rawPropertiesTable = coveragePreviewPropertiesTable;
                        tableMetadataResponse = () => result([["TableName", "string"], ["DocString", "string"]], [
                            [coveragePreviewTelemetryTable, "@iot.meta tableRole=iothub-telemetry-raw\n@iot.meta source=fixture-hub.azure-devices.net"],
                            [coveragePreviewPropertiesTable, "@iot.meta tableRole=iothub-properties-raw\n@iot.meta source=fixture-hub.azure-devices.net"],
                        ]);
                        seed.hubFields[0] = { ...seed.hubFields[0],
                            title: "Temperature with a deliberately long friendly title that wraps without truncation" };
                    }
                    await wizard(seed); await click(next());
                    await beginAssignmentTargets();
                    await until(() => !!currentContext.hubAssignmentHubSummary?.value.complete, "Coverage preview Hub context did not settle.");
                    await until(() => !next()!.disabled, "Coverage preview device inventory did not settle.");
                    await click(next());
                    assert(area.querySelector("h2")?.textContent === "Filter events and check model coverage",
                        "Coverage preview did not open the actual step.");
                    if (["model-coverage-conditions", "model-coverage-positive-results", "model-coverage-results",
                        "model-coverage-mixed-results", "model-coverage-dense-results"].includes(preview)) {
                        await click(area.querySelector('input[value="and"]'));
                        await click(area.querySelector('input[value="specified"]'));
                    }
                    if (["model-coverage-positive-results", "model-coverage-results", "model-coverage-mixed-results", "model-coverage-dense-results"].includes(preview)) {
                        if (preview === "model-coverage-results") emptyProcessing = true;
                        else if (preview !== "model-coverage-positive-results") {
                            coverageSampleRows = 4;
                            emptyProcessingStream = preview === "model-coverage-mixed-results" ? "telemetry" : undefined;
                            parityPayloads = { telemetry: { unmodeledReading: 1 },
                                twin: { properties: { reported: { unmodeledState: "idle" } } } };
                        }
                        await verifyCoverage();
                        assertCoverageMissingGuidance();
                        document.documentElement.dataset.coverageCardGeometry = JSON.stringify(coverageCardGeometry());
                    }
                });
            } finally { HubAssignmentApi.prototype.discover = discover; }
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            Reflect.set(window, "__coveragePreviewCounts", () => ({ queries: calls.length, tokens: tokenRequests,
                resources: resourceCalls.length, consent: consentRequests, modelId: currentContext.hubModel.id }));
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview === "assignment-confirmation") {
            await assignmentConfirmationCheck(true);
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("assignment-poll-")) {
            const discover = HubAssignmentApi.prototype.discover, services = HubAssignmentApi.prototype.assignmentServices,
                readback = HubAssignmentApi.prototype.readAssignment;
            const clock = assignmentClock();
            try {
                sourceHubFixture("linked");
                const value = assignmentInventory();
                value.targets = [0, 1, 2].map(index => ({ ...value.targets[0], id: value.targets[0].id + `-${index}`,
                    name: ["ThermoPro device 01", "ThermoPro device 02", "ThermoPro device with a longer identity 03"][index],
                    writeBlocker: undefined, reference: { kind: "unassigned", ids: [] },
                    deviceIdentity: { uuid: "11111111-1111-4111-8111-111111111111", externalDeviceId: `device-${index}` } }));
                HubAssignmentApi.prototype.discover = async () => value;
                HubAssignmentApi.prototype.assignmentServices = () => ({
                    read: async target => target, write: async (_target, _id, _signal, dispatch) => { dispatch(); return "accepted"; },
                });
                let complete = false;
                HubAssignmentApi.prototype.readAssignment = async target => {
                    const confirmed = complete || target.id.endsWith("-0");
                    return { target: { ...target, reference: confirmed ? { kind: "assigned", ids: [tm.id] } : target.reference }, confirmed };
                };
                await observeWizardContext(async () => {
                    await wizard(context(true)); await click(next()); await beginAssignmentTargets();
                    await until(() => !!button("Select all") && !button("Select all")!.disabled, "Progress preview devices did not settle.");
                    await click(button("Select all")); await click(button("Assign to selected devices"));
                    await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')].find(button => button.textContent === "Assign"));
                    await until(() => clock.pending.size === 1, "Progress preview result timer was not scheduled.");
                    await clock.tick();
                    await until(() => confirmedDeviceRows() === 1, "Incremental progress preview did not settle.");
                    assertAssignmentSummary("active", 3);
                    if (preview === "assignment-poll-success") {
                        complete = true; await clock.tick();
                        await until(() => !!completedAssignmentList(),
                            "Readback preview did not reach truthful success.");
                    }
                });
            } finally {
                clock.restore(); HubAssignmentApi.prototype.discover = discover;
                HubAssignmentApi.prototype.assignmentServices = services; HubAssignmentApi.prototype.readAssignment = readback;
            }
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("device-table-")) {
            const discover = HubAssignmentApi.prototype.discover, services = HubAssignmentApi.prototype.assignmentServices;
            try {
                sourceHubFixture(preview === "device-table-mixed" ? "mixed" : "linked");
                const value = preview === "device-table-mixed" ? deviceListInventory() : assignmentInventory();
                if (preview !== "device-table-mixed") value.targets = [0, ...(preview === "device-table-confirmed" ? [1, 2] : [])]
                    .map(index => ({ ...value.targets[0], id: value.targets[0].id + `-${index}`,
                        name: `Fixture device ${index + 1}`, writeBlocker: undefined,
                        reference: preview === "device-table-already" ? { kind: "assigned", ids: [tm.id] } : { kind: "unassigned", ids: [] } }));
                const stored = new Set<string>();
                HubAssignmentApi.prototype.discover = async () => value;
                HubAssignmentApi.prototype.assignmentServices = () => ({
                    read: async target => ({ ...target, reference: stored.has(target.id)
                        ? { kind: "assigned", ids: [tm.id] } : target.reference }),
                    write: async (target, _id, _signal, dispatch) => { dispatch(); stored.add(target.id); return "completed"; },
                });
                await observeWizardContext(async () => {
                    await wizard(context(true)); await click(next()); await beginAssignmentTargets();
                    await until(() => !area.querySelector('[role="progressbar"]') &&
                        (!!area.querySelector("table") || !!completedAssignmentList() || !!existingAssignmentsNotice()),
                        "Device-table preview did not settle.");
                    if (preview === "device-table-mixed") await chooseAssignmentView("all");
                    if (preview === "device-table-confirmed") {
                        await click(button("Select all")); await click(button("Assign to selected devices"));
                        await click([...document.querySelectorAll<HTMLButtonElement>('[role="dialog"] button')]
                            .find(button => button.textContent === "Assign"));
                        await until(() => !!completedAssignmentList(),
                            "Confirmed synthetic preview did not produce current-operation success.");
                    }
                });
            } finally { HubAssignmentApi.prototype.discover = discover; HubAssignmentApi.prototype.assignmentServices = services; }
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("hub-summary-")) {
            const discover = HubAssignmentApi.prototype.discover;
            try {
                const fixture = sourceHubFixture(preview === "hub-summary-unlinked" ? "omitted" :
                    preview === "hub-summary-linked" ? "linked" : "mixed");
                let pendingHub = false;
                if (preview === "hub-summary-loading") assignmentResponse = (url, init) =>
                    new URL(String(url)).pathname === fixture.hub.id ? new Promise(() => { pendingHub = true; }) : fixture.fetcher(url, init);
                if (preview === "hub-summary-devices") HubAssignmentApi.prototype.discover = async () => deviceListInventory();
                await observeWizardContext(async () => {
                    await wizard(context(true)); await click(next());
                    if (preview !== "hub-summary-initial") {
                        await beginAssignmentTargets();
                        if (preview === "hub-summary-loading") await until(() => pendingHub &&
                            !!area.querySelector('[role="progressbar"]'), "Chosen Hub-loading preview did not start.");
                        else await until(() => !!currentContext.hubAssignmentHubSummary?.value.complete,
                            "Chosen Hub summary preview did not complete its limited reads.");
                    }
                    if (preview === "hub-summary-devices") {
                        await until(() => !!area.querySelector('tr.iot-solution-assignment-device') &&
                            !area.querySelector('[role="progressbar"]'), "Device list preview did not settle.");
                        await chooseAssignmentView("all");
                    }
                });
            } finally { HubAssignmentApi.prototype.discover = discover; }
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("assignment-intent-")) {
            const discover = HubAssignmentApi.prototype.discover;
            try {
                sourceHubFixture("linked");
                if (preview === "assignment-intent-active") HubAssignmentApi.prototype.discover = async () => {
                    const value = assignmentInventory();
                    value.targets[0] = { ...value.targets[0], reference: { kind: "unassigned", ids: [] }, writeBlocker: undefined };
                    return value;
                };
                await observeWizardContext(async () => {
                    await wizard(context(preview !== "assignment-intent-dtdl"));
                    await click(next());
                    assert(area.querySelector("h2")?.textContent === "Assign the model to devices", "Assignment preview missed its normal step.");
                    if (preview === "assignment-intent-active") {
                        await beginAssignmentTargets();
                        await until(() => !!currentContext.hubAssignmentInventory && !area.querySelector('[role="progressbar"]'),
                            "Active assignment preview did not settle.");
                    }
                });
            } finally { HubAssignmentApi.prototype.discover = discover; }
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("upload-spacing-")) {
            document.documentElement.dataset.uploadSpacing = JSON.stringify(await uploadSpacingCheck(preview === "upload-spacing-accepted"));
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview === "calibration-layout") {
            const result = await calibrationRowLayoutCheck();
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.calibrationLayout = JSON.stringify(result);
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview === "accepted-model" || preview === "field-complex" || preview.startsWith("field-presentation-") || preview.startsWith("inline-model-")) {
            await observeWizardContext(async () => {
                const seed = context(true); delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
                if (preview === "inline-model-ambiguous") parityPayloads = { telemetry: { temperature: 22, state: "on" },
                    twin: { properties: { reported: { temperature: 23, state: "off" } } } };
                await wizard(seed);
                const dtdlFormat = preview === "inline-model-dtdl";
                await click(area.querySelector(`input[value="${dtdlFormat ? "dtdl" : "wot"}"]`));
                const document = dtdlFormat ? mixedDtdlDocument : preview === "inline-model-ambiguous" ? tm : preview === "accepted-model" ? metadataThingModel : {
                    ...metadataThingModel, id: `https://models.example.test/${"supplied-model-identifier-".repeat(7)}version-1`,
                };
                await file("thing-model-upload", JSON.stringify(document),
                    preview === "accepted-model" || preview.startsWith("field-presentation-") ? "thermoProX200.json" : "inline-model.json");
                await until(() => !!area.querySelector("#model-fields-heading"), "The inline field selection preview did not appear.");
                if (preview === "inline-model-failed") await file("thing-model-upload", "{bad-json", "failed-replacement.json");
                else if (preview !== "inline-model-ambiguous") {
                    for (const field of currentContext.hubFields as ModelField[]) {
                        if (!field.role) await enter(documentElementRole(field.id)!.id, "telemetry");
                    }
                }
            });
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("model-prompt-")) {
            await observeWizardContext(async () => {
                const seed = withAnalyzedSamples(context(true)), cache = seed.hubSourceAnalysis;
                delete seed.hubModel; delete seed.hubSelectedFormat; seed.hubFields = []; seed.preparationFileName = "";
                if (preview === "model-prompt-empty" || preview === "model-prompt-partial")
                    cache.twin = { ...cache.twin, payloads: [], rawRows: 0 };
                if (preview === "model-prompt-empty") cache.telemetry = { ...cache.telemetry, payloads: [], rawRows: 0 };
                if (preview === "model-prompt-error") {
                    let value: object = { reading: 1 };
                    for (let index = 0; index < 50; index++) value = { nested: value };
                    cache.telemetry = { ...cache.telemetry, payloads: [{ ...cache.telemetry.payloads[0], value }] };
                }
                await openAnalysisSource(seed);
                await until(() => !next()!.disabled, "The cached preview source did not become ready.");
                await click(next());
                if (preview !== "model-prompt-collapsed") {
                    await click(button("Get Copilot prompt"));
                    await until(() => !area.textContent?.includes("Preparing prompt..."), "The model prompt preview did not settle.");
                    if (preview === "model-prompt-copy-failed" || preview === "model-prompt-copied") {
                        Object.defineProperty(navigator, "clipboard", { configurable: true, value: {
                            writeText: async () => {
                                if (preview === "model-prompt-copy-failed") throw new Error("Synthetic owned clipboard denial.");
                            },
                        } });
                        await click(button("Copy prompt"));
                    }
                }
                if (preview === "model-prompt-empty") {
                    document.documentElement.dataset.emptyPromptGeometry = JSON.stringify(emptyPromptGeometry());
                }
            });
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            Reflect.set(window, "__modelFormatPreviewCounts", () => ({ queries: calls.length, tokens: tokenRequests,
                resources: resourceCalls.length, assignmentQueries: assignmentQueries.length }));
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("source-validation-")) {
            await observeWizardContext(async () => {
                const initial = context(true);
                if (preview === "source-validation-model-intro") {
                    delete initial.hubModel; delete initial.hubSelectedFormat; initial.hubFields = []; initial.preparationFileName = "";
                }
                await openAnalysisSource(initial);
                await until(() => !!sourceValidationAction() && !sourceValidationAction()!.disabled, "Source validation preview metadata did not settle.");
                if (preview === "source-validation-initial") return;
                deferSourceSamples = preview === "source-validation-pending";
                sourceSampleStatus = preview === "source-validation-failed" ? 503 : 0;
                await click(sourceValidationAction());
                if (deferSourceSamples) {
                    await until(() => sourceSampleReleases.length === 2, "The validation preview did not reach two pending reads.");
                    return;
                }
                if (sourceSampleStatus) {
                    await until(() => !!area.querySelector('[role="alert"]') && !!button("Validate again"), "The failed validation preview did not finish.");
                    return;
                }
                await until(() => !next()!.disabled, "The validation preview did not finish successfully.");
                if (preview === "source-validation-selection-reset") {
                    await chooseOption("hub-model-telemetry", "Show All Tables...");
                    await chooseOption("hub-model-telemetry", "NewTelemetry");
                } else if (preview === "source-validation-model-intro") await click(next());
            });
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("source-reselection-")) {
            await observeWizardContext(async () => {
                let pending = false;
                const deferred = () => { pending = true; return new Promise<Response>(() => undefined); };
                if (preview === "source-reselection-eventhouses") resourceResponse = deferred;
                await openAnalysisSource();
                if (preview === "source-reselection-eventhouses") {
                    await until(() => pending, "The Eventhouse field did not reach its loading boundary.");
                    return;
                }
                await until(tableChoicesReady, "The source preview inventory did not settle.");
                if (preview === "source-reselection-databases") {
                    eventhouseMetadataResponse = deferred;
                    await chooseOption("hub-model-eventhouse", "Eventhouse");
                    await until(() => pending, "The database field did not reach its loading boundary.");
                } else if (preview === "source-reselection-tables") {
                    tableMetadataResponse = deferred;
                    await chooseOption("hub-model-database", "Database");
                    await until(() => pending, "The table fields did not reach their loading boundary.");
                } else {
                    await chooseOption("hub-model-eventhouse", "Eventhouse");
                    await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-database")!.disabled, "The reselected Eventhouse did not reload.");
                    await chooseOption("hub-model-database", "Database"); await until(tableChoicesReady, "The reselected database did not reload.");
                    await chooseOption("hub-model-telemetry", "Telemetry"); await chooseOption("hub-model-twin", "Twins");
                    assert(!sourceValidationAction()!.disabled && next()!.disabled && !area.textContent?.includes("No tables found"),
                        "Recovered source silently analyzed or lost its inventory.");
                }
            });
            area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")!.parentElement!.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("source-labels")) {
            await observeWizardContext(async () => {
                if (preview === "source-labels-model") await wizard(context(true));
                else {
                    await mount(<IoTSolutionItemModelWizard workloadClient={client}
                        item={{ id: "mock", workspaceId: "workspace" } as ItemWithDefinition<IoTSolutionItemDefinition>}
                        initialContext={context(true)} onBack={() => undefined}/>);
                    await until(tableChoicesReady, "Source screenshot inventory did not settle.");
                    const id = preview === "source-labels-telemetry" ? "hub-model-telemetry"
                        : preview === "source-labels-properties" ? "hub-model-twin" : "";
                    if (id) {
                        const control = area.querySelector<HTMLElement>(`#${id}`)!;
                        control.focus(); await click(control);
                        if (id === "hub-model-twin") {
                            for (const [key, code] of [["End", 35], ["Enter", 13]] as const) {
                                await act(async () => control.dispatchEvent(new KeyboardEvent("keydown",
                                    { key, keyCode: code, which: code, bubbles: true, cancelable: true })));
                                await settle();
                            }
                        } else await chooseOption(id, "Show All Tables...");
                        assert(control.getAttribute("aria-expanded") === "true" &&
                            activeOptions(id).some(option => option.textContent === "NewTelemetry"), "Show All preview did not retain the expanded menu.");
                    }
                }
            });
            const frame = area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")?.parentElement;
            assert(frame, "The source-labels preview frame is missing."); frame.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("source-analysis-")) {
            await observeWizardContext(async () => {
                const initial = context(true);
                delete initial.hubModel; delete initial.hubSelectedFormat; initial.hubFields = []; initial.preparationFileName = "";
                if (preview === "source-analysis-prerequisites") {
                    for (const key of ["modelEhId", "modelEhName", "modelDbId", "modelDbName", "modelQueryServiceUri",
                        "rawTelemetryTable", "rawPropertiesTable"]) initial[key] = "";
                }
                sourceSampleRows = 1000;
                await openAnalysisSource(initial);
                await until(() => !area.querySelector<HTMLButtonElement>("#hub-model-eventhouse")!.disabled, "Preview resources did not settle.");
                if (preview === "source-analysis-prerequisites") return;
                await until(() => !sourceValidationAction()!.disabled, "Preview source did not settle.");
                if (preview === "source-analysis-selected") return;
                deferSourceSamples = preview === "source-analysis-loading";
                await click(sourceValidationAction());
                if (deferSourceSamples) {
                    await until(() => sourceSampleReleases.length === 2, "Preview did not reach two owned pending reads.");
                    return;
                }
                await until(() => !next()!.disabled, "Preview source analysis did not complete.");
                if (preview === "source-analysis-analyzed") return;
                if (preview === "source-analysis-menu") {
                    await chooseOption("hub-model-telemetry", "Show All Tables...");
                    assert(area.querySelector("#hub-model-telemetry")!.getAttribute("aria-expanded") === "true",
                        "Preview Show All menu did not remain open.");
                    return;
                }
                await click(next());
                if (preview === "source-analysis-model") return;
                await click(area.querySelector(`input[value="${preview === "source-analysis-dtdl" ? "dtdl" : "wot"}"]`));
                if (preview === "source-analysis-fields") {
                    await file("thing-model-upload", JSON.stringify(tm), "automatic-role-preview.tm.json");
                }
            });
            const frame = area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")?.parentElement;
            assert(frame, "The source-analysis preview frame is missing."); frame.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        if (preview.startsWith("alignment-")) {
            await observeWizardContext(async () => {
                const request = preview === "alignment-fields" ? mixedDtdlRequest() : creationRequest();
                request.input.source.telemetryTable = "Telemetry"; request.input.source.twinTable = "Twins";
                if (preview === "alignment-review") request.input.scope = { mode: "reference", additional: "specified", operator: "or",
                    reference: { id: request.input.model.id! }, telemetry: "tostring(headers[\"IoTConnectionDeviceId\"]) == \"device-1\"", twin: "" };
                request.coverage.fingerprint = modelingFingerprint(request.input);
                const initial = alignmentCreationContext(request);
                parityPayloads = { telemetry: request.coverage.telemetry.payloads[0].value,
                    twin: { properties: { reported: request.coverage.twin.payloads[0].reported } } };
                await wizard(initial); await reachReview();
                if (preview === "alignment-fields") await click(step("Add device model"));
                else {
                    await enter("modeled-output-name", request.names.modeled);
                    await enter("twin-changes-name", request.names.twinChanges);
                    await enter("twin-state-name", request.names.twinState);
                    if (preview === "alignment-ready") {
                        creationBackend = createCreationMock(request);
                        await click(button("Create modeled data"));
                        await until(() => currentContext.hubCreation?.phase === "completed",
                            () => `Mock preview creation did not complete: ${JSON.stringify(currentContext.hubCreation)}`);
                        await click(next());
                    }
                }
            });
            const frame = area.querySelector<HTMLElement>(".iot-solution-hub-model-wizard")?.parentElement;
            assert(frame, "The existing wizard preview frame is missing.");
            frame.style.height = "100%";
            document.documentElement.dataset.previewKind = preview;
            document.documentElement.dataset.previewReady = "true";
            return;
        }
        const seed = context();
        if (["raw", "model", "loaded", "schema", "dtdl", "error-wot", "error-dtdl", "error-size", "error-read"].includes(preview)) {
            delete seed.hubModel;
            seed.hubFields = [];
            seed.preparationFileName = "";
            seed.hubSelectedFormat = preview === "dtdl" || preview === "error-dtdl" ? "dtdl" : "wot";
            seed.modelRawDataSelected = preview !== "raw";
            seed.modelEhName = "iot-eh";
            seed.modelDbName = "iot-eh";
            seed.rawTelemetryTable = "TelemetryRawIoTHub_from_canary";
            seed.rawPropertiesTable = "PropertiesRawIoTHub_from_canary";
            tableMetadataResponse = async () => result([["TableName", "string"], ["DocString", "string"]], [
                [seed.rawTelemetryTable, "@iot.meta tableRole=iothub-telemetry-raw"],
                [seed.rawPropertiesTable, "@iot.meta tableRole=iothub-properties-raw"], ["ManualSample", null],
            ]);
        }
        if (preview.startsWith("parity-")) {
            const document = { ...tm, title: "Declared whole values", properties: {
                    measurement: { type: "object", title: "Measurement", description: "Whole measured value.",
                        properties: { nested: { type: "object", properties: { reading: { type: "number", unit: "degC" } } } } },
                }, events: { alarm: { title: "Alarm", data: { type: "object", properties: { state: { type: "string" } } } } },
                actions: { reboot: {} } };
            const model = fromThingModel(resolveThingModel(document));
            seed.hubModel = model;
            seed.hubSelectedFormat = "wot";
            seed.preparationFileName = "declared-whole-values.tm.json";
            seed.hubFields = model.fields.map(field => setFieldRole(field, field.column === "alarm" ? "twin" : "telemetry"));
            parityPayloads = { telemetry: { measurement: { nested: { reading: 20 } } },
                twin: { properties: { reported: { alarm: { state: "latest-whole-value" } } } } };
        }
        await wizard(seed);
        if (preview === "error-dtdl") {
            await file("thing-model-upload", JSON.stringify({ notAModel: true }), "unsupported-device-model.json");
        }
        else if (preview === "error-wot") {
            await file("thing-model-upload", JSON.stringify(dtdl), "wrong-format-device-model.json");
        }
        else if (preview === "error-size") {
            await file("thing-model-upload", " ".repeat(2 * 1024 * 1024 + 1), "too-large-device-model.json");
        }
        else if (preview === "error-read") {
            await unreadableFile("thing-model-upload", "unreadable-device-model.json");
        }
        else if (preview === "schema") {
            await click(button("Get Copilot prompt"));
            await until(() => area.querySelector("#model-preparation-prompt"), "Mock cached-schema prompt did not load.");
        }
        else if (["fields", "scope", "review", "parity-fields", "parity-scope", "parity-review"].includes(preview)) {
            await reachReview();
            if (preview.endsWith("fields"))
                await click(step("Add device model"));
            if (preview.endsWith("scope"))
                await click(step("Check model coverage"));
            if (preview === "parity-scope")
                area.querySelector("#coverage-results-heading")!.scrollIntoView({ block: "start" });
            if (preview === "parity-review") {
                area.querySelector(".iot-solution-form")!.scrollIntoView({ block: "start" });
            }
        }
        else if (preview === "raw") {
            await until(() => area.querySelector("#hub-model-telemetry") && tableChoicesReady(), "Mock source metadata did not settle.");
        }
        else if (preview === "dtdl") {
            await file("thing-model-upload", JSON.stringify(dtdl), "device-model.json");
            await click(next());
            await until(() => !!area.querySelector("[data-dtdl-assignment-guidance]") && !next()!.disabled,
                "The current DTDL assignment guidance did not open.");
        }
        document.documentElement.dataset.previewKind = preview;
        document.documentElement.dataset.previewReady = "true";
    })();
}
