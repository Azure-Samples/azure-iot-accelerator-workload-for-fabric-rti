import { stripVTControlCharacters } from "node:util";
import { redactSensitiveText } from "./cloud.js";
import { scenarioGroups, scenarios, type ScenarioId, type ScenarioKey } from "./scenarios.js";

export interface PipelineStage {
  id: ScenarioId;
  name: string;
  status: "not-run" | "passed" | "failed" | "skipped";
  durationMs: number;
  exitCode?: number;
  detail?: string;
}

export function createStage(key: ScenarioKey): PipelineStage {
  return { ...scenarios[key], status: "not-run", durationMs: 0 };
}

export function scenarioResultsTree(stages: readonly PipelineStage[]): string {
  const groups = scenarioGroups
    .map(group => ({ ...group, stages: stages.filter(stage => stage.id.startsWith(`${group.id}.`)) }))
    .filter(group => group.stages.length > 0);
  const labels = { passed: "PASS", failed: "FAIL", skipped: "SKIPPED", "not-run": "NOT RUN" } as const;
  return groups.map((group, groupIndex) => {
    const lastGroup = groupIndex === groups.length - 1;
    const children = [...group.stages]
      .sort((a, b) => a.id.localeCompare(b.id, "en", { numeric: true }))
      .map((stage, index, entries) => {
        const duration = stage.status === "not-run" || stage.status === "skipped"
          ? "" : ` (${(stage.durationMs / 1000).toFixed(1)}s)`;
        const detail = stage.status === "failed" || stage.status === "skipped"
          ? stage.detail?.split("\n")[0] : undefined;
        const suffix = detail ? ` - ${stripVTControlCharacters(redactSensitiveText(detail))}` : "";
        return `${lastGroup ? "    " : "|   "}${index === entries.length - 1 ? "`--" : "|--"} ` +
          `${stage.id} ${stage.name} [${labels[stage.status]}]${duration}${suffix}`;
      });
    return `${lastGroup ? "`--" : "|--"} ${group.id} ${group.name}\n${children.join("\n")}`;
  }).join("\n");
}

function xml(value: string): string {
  return stripVTControlCharacters(redactSensitiveText(value))
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

export function junitReport(runId: string, stages: readonly PipelineStage[], runnerFailure?: string): string {
  const stageFailures = stages.filter(stage => stage.status === "failed").length;
  const extraFailure = runnerFailure !== undefined && stageFailures === 0;
  const failed = stageFailures + Number(extraFailure);
  const tests = stages.length + Number(extraFailure);
  const skipped = stages.filter(stage => stage.status === "not-run" || stage.status === "skipped").length;
  const time = stages.reduce((sum, stage) => sum + stage.durationMs, 0) / 1000;
  return `<?xml version="1.0" encoding="UTF-8"?>\n` +
    `<testsuites tests="${tests}" failures="${failed}" skipped="${skipped}" time="${time.toFixed(3)}">\n` +
    `<testsuite name="Fabric IoT E2E" tests="${tests}" failures="${failed}" errors="0" skipped="${skipped}" time="${time.toFixed(3)}">\n` +
    `<properties><property name="runId" value="${xml(runId)}"/></properties>\n` +
    stages.map(stage => {
      const start = `<testcase classname="fabric.iot.e2e" name="${xml(`${stage.id} ${stage.name}`)}" time="${(stage.durationMs / 1000).toFixed(3)}">`;
      const result = stage.status === "failed"
        ? `<failure message="${xml(stage.detail?.split("\n")[0] ?? "Phase failed")}">${xml(stage.detail ?? "Phase failed")}</failure>`
        : stage.status === "not-run" || stage.status === "skipped"
          ? `<skipped message="${xml(stage.detail ?? "Not run because an earlier phase failed")}"/>` : "";
      return `${start}${result}</testcase>`;
    }).join("\n") +
    (extraFailure ? `\n<testcase classname="fabric.iot.e2e.runner" name="Runner failure" time="0.000"><failure message="${
      xml(runnerFailure.split("\n")[0] ?? "Runner failed")}">${xml(runnerFailure)}</failure></testcase>` : "") +
    "\n</testsuite>\n</testsuites>\n";
}
