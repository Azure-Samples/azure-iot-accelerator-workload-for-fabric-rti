import type { FullConfig, FullResult, Reporter, Suite, TestCase, TestResult } from "@playwright/test/reporter";
import { createStage, scenarioResultsTree, type PipelineStage } from "./report.js";
import { scenarios, type ScenarioKey } from "./scenarios.js";

export default class ScenarioReporter implements Reporter {
  private readonly stages = new Map<string, PipelineStage>();

  onBegin(_config: FullConfig, suite: Suite): void {
    for (const test of suite.allTests()) {
      const key = (Object.keys(scenarios) as ScenarioKey[]).find(key => {
        const scenario = scenarios[key];
        return test.title === `${scenario.id} ${scenario.name}`;
      });
      if (!key) throw new Error(`[reporting] Test has no registered scenario ID: ${test.title}`);
      this.stages.set(test.id, createStage(key));
    }
  }

  onTestEnd(test: TestCase, result: TestResult): void {
    const stage = this.stages.get(test.id);
    if (!stage) throw new Error(`[reporting] Missing scenario result for ${test.title}`);
    stage.durationMs += result.duration;
    stage.status = result.status === "passed" ? "passed"
      : result.status === "skipped" ? test.expectedStatus === "skipped" ? "skipped" : "not-run" : "failed";
    if (result.errors.length) {
      stage.detail = result.errors.map(error => error.message ?? error.value ?? "Unexpected test failure").join("\n");
    }
  }

  onEnd(result: FullResult): void {
    if (process.env.E2E_PIPELINE_RUN === "1" || !this.stages.size) return;
    const stages = [...this.stages.values()];
    const status = result.status === "passed"
      ? stages.every(stage => stage.status === "not-run") ? "NOT RUN" : "PASS"
      : result.status === "failed" ? "FAIL" : result.status.toUpperCase();
    console.log(`\n${status} Fabric IoT E2E scenarios`);
    console.log(scenarioResultsTree(stages));
  }
}
