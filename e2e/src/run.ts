import { randomUUID } from "node:crypto";
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { parseArgs, stripVTControlCharacters } from "node:util";
import { redactSensitiveText } from "./cloud.js";
import { artifacts, requireGuid, root, runConfig } from "./config.js";
import { createPlan, readPlan, writePlan } from "./plan.js";
import { createStage, junitReport, scenarioResultsTree, type PipelineStage } from "./report.js";
import { keepSystemAwake } from "./awake.js";
import { OwnedChildShutdownError, RunInterrupted, runOwnedChild, shutdownController } from "./owned-child.js";

const reportDirectory = join(artifacts, "reports", `run-${randomUUID()}`);
await mkdir(reportDirectory, { recursive: true });
let runId = "";
let failureDetail: string | undefined;
const { values } = parseArgs({ options: {
  "nonlinked-only": { type: "boolean", default: false },
  "ingestion-only": { type: "boolean", default: false },
  "extensions-only": { type: "boolean", default: false },
  "keep-resources": { type: "boolean", default: false },
} });
const environment = { ...process.env, ...(values["keep-resources"] ? { E2E_KEEP_RESOURCES: "1" } : {}) };
let keepResources = values["keep-resources"];
const stages = {
  initialize: createStage("initialize"),
  item: createStage("item"),
  provision: createStage("provision"),
  ingestion: createStage("ingestion"),
  links: createStage("links"),
  combined: createStage("combined"),
  linked: createStage("linked"),
  regularModel: createStage("regularModel"),
  combinedModel: createStage("combinedModel"),
  linkedModel: createStage("linkedModel"),
  regularDashboard: createStage("regularDashboard"),
  combinedDashboard: createStage("combinedDashboard"),
  linkedDashboard: createStage("linkedDashboard"),
  regularAgent: createStage("regularAgent"),
  combinedAgent: createStage("combinedAgent"),
  linkedAgent: createStage("linkedAgent"),
  regularActivator: createStage("regularActivator"),
  combinedActivator: createStage("combinedActivator"),
  linkedActivator: createStage("linkedActivator"),
  cleanup: createStage("cleanup"),
};
const pipeline: Record<keyof typeof stages, PipelineStage> = stages;
const shutdown = shutdownController();
let ownedWorkStopped = true;

for (const key of ["links", "combined", "linked", "combinedModel", "linkedModel", "combinedDashboard", "linkedDashboard", "combinedAgent", "linkedAgent", "combinedActivator", "linkedActivator"] as const) {
  if (values["nonlinked-only"]) {
    stages[key].status = "skipped";
    stages[key].detail = "Excluded by --nonlinked-only.";
  }
}
for (const key of ["regularModel", "combinedModel", "linkedModel", "regularDashboard", "combinedDashboard", "linkedDashboard", "regularAgent", "combinedAgent", "linkedAgent", "regularActivator", "combinedActivator", "linkedActivator"] as const) {
  if (values["ingestion-only"]) {
    stages[key].status = "skipped";
    stages[key].detail = "Excluded by --ingestion-only.";
  }
}

async function step<T>(stage: PipelineStage, action: () => Promise<T>): Promise<T> {
  const start = Date.now();
  console.log(`[${new Date().toISOString()}] START ${stage.id} ${stage.name}`);
  try {
    if (stage !== stages.cleanup) shutdown.signal.throwIfAborted();
    const result = await action();
    stage.status = "passed";
    console.log(`[${new Date().toISOString()}] PASS  ${stage.id} ${stage.name} (${((Date.now() - start) / 1000).toFixed(1)}s)`);
    return result;
  } catch (error) {
    stage.status = "failed";
    stage.detail = redactSensitiveText(error instanceof Error ? error.message : "Unexpected non-Error failure");
    console.error(`[${new Date().toISOString()}] FAIL  ${stage.id} ${stage.name}`);
    console.error(stage.detail.split("\n")[0]);
    throw error;
  } finally {
    stage.durationMs = Date.now() - start;
  }
}

async function child(stage: PipelineStage, args: string[]): Promise<void> {
  let tail = "";
  let code: number;
  try {
    code = await runOwnedChild(args, {
      cwd: root, runId,
      env: { ...environment, E2E_RUN_ID: runId, E2E_PIPELINE_RUN: "1" },
      ...(stage === stages.cleanup ? {} : { signal: shutdown.signal }),
      output(chunk, stream) {
      stream.write(chunk);
      tail = (tail + chunk.toString("utf8")).slice(-20_000);
      },
    });
  } catch (error) {
    if (error instanceof OwnedChildShutdownError) ownedWorkStopped = false;
    throw error;
  }
  stage.exitCode = code;
  if (code !== 0) throw new Error(`[runner] Phase failed (exit ${code}). Run ${runId} is retained.\n${redactSensitiveText(stripVTControlCharacters(tail))}`);
}

let releaseAwake: () => Promise<void> = async () => {};
try {
  releaseAwake = await keepSystemAwake();
  const plan = await step(stages.initialize, async () => {
    const config = runConfig(environment);
    keepResources = config.keepResources;
    const requested = process.env.E2E_RUN_ID;
    if (values["extensions-only"] && !requested) throw new Error("[configuration] --extensions-only requires an explicit E2E_RUN_ID.");
    const value = requested ? await readPlan(config, requireGuid(requested, "E2E_RUN_ID")) : createPlan(config);
    if (!requested) await writePlan(value);
    runId = value.runId;
    return value;
  });
  console.log(`Run ${runId}; ${keepResources ? "KEEP RESOURCES: automatic deletion disabled; retained resources may incur charges" : "run-owned resources will be cleaned up in finally, including on failure"}. Reports: ${reportDirectory}`);
  console.log("Scope: regular/system+UAMI, combined/5 Hubs, linked-only/3 Hubs; one folder per run. Each modeled path has the supplied DTDL or WoT model.");
  const playwright = join(root, "node_modules", "@playwright", "test", "cli.js");
  if (!plan.observed.itemId) {
    await step(stages.item, () => child(stages.item, [playwright, "test", "item.spec.ts"]));
  } else {
    stages.item.status = "skipped";
    stages.item.detail = "Resuming the recorded owned item; no additional item created.";
    console.log(`SKIP  ${stages.item.id} ${stages.item.name}: ${stages.item.detail}`);
  }
  await step(stages.provision, () => child(stages.provision, ["--import", "tsx", join(root, "src", "provision-cli.ts")]));
  if (values["extensions-only"]) {
    if (!plan.verification.rawData?.startsWith("native-kusto")) throw new Error("[prerequisite] Extensions require a recorded successful nonlinked ingestion run.");
    pipeline.ingestion.status = "skipped";
    pipeline.ingestion.detail = "Explicit extensions-only run: previously completed nonlinked outputs retained, not reported as a new pass.";
  } else await step(pipeline.ingestion, () => child(pipeline.ingestion, [playwright, "test", "scenarios/ingestion/nonlinked.spec.ts"]));
  if (!values["nonlinked-only"]) {
    await step(pipeline.links, () => child(pipeline.links, ["--import", "tsx", join(root, "src", "linking-cli.ts")]));
    await step(pipeline.combined, () => child(pipeline.combined, [playwright, "test", "scenarios/ingestion/combined.spec.ts"]));
    await step(pipeline.linked, () => child(pipeline.linked, [playwright, "test", "scenarios/ingestion/linked-only.spec.ts"]));
  }
  if (!values["ingestion-only"]) {
    await step(pipeline.regularModel, () => child(pipeline.regularModel, [playwright, "test", "scenarios/modeling/nonlinked-dtdl.spec.ts"]));
    if (!values["nonlinked-only"]) {
      await step(pipeline.combinedModel, () => child(pipeline.combinedModel, [playwright, "test", "scenarios/modeling/combined-wot.spec.ts"]));
      await step(pipeline.linkedModel, () => child(pipeline.linkedModel, [playwright, "test", "scenarios/modeling/linked-only-wot.spec.ts"]));
    }
    await step(pipeline.regularDashboard, () => child(pipeline.regularDashboard, [playwright, "test", "scenarios/dashboards/nonlinked.spec.ts"]));
    await step(pipeline.regularAgent, () => child(pipeline.regularAgent, [playwright, "test", "scenarios/data-agents/nonlinked.spec.ts"]));
    if (!values["nonlinked-only"]) {
      await step(pipeline.combinedDashboard, () => child(pipeline.combinedDashboard, [playwright, "test", "scenarios/dashboards/combined.spec.ts"]));
      await step(pipeline.combinedAgent, () => child(pipeline.combinedAgent, [playwright, "test", "scenarios/data-agents/combined.spec.ts"]));
      await step(pipeline.linkedDashboard, () => child(pipeline.linkedDashboard, [playwright, "test", "scenarios/dashboards/linked-only.spec.ts"]));
      await step(pipeline.linkedAgent, () => child(pipeline.linkedAgent, [playwright, "test", "scenarios/data-agents/linked-only.spec.ts"]));
    }
    await step(pipeline.regularActivator, () => child(pipeline.regularActivator, [playwright, "test", "scenarios/activator/nonlinked.spec.ts"]));
    if (!values["nonlinked-only"]) {
      await step(pipeline.combinedActivator, () => child(pipeline.combinedActivator, [playwright, "test", "scenarios/activator/combined.spec.ts"]));
      await step(pipeline.linkedActivator, () => child(pipeline.linkedActivator, [playwright, "test", "scenarios/activator/linked-only.spec.ts"]));
    }
  }
} catch (error) {
  process.exitCode = error instanceof RunInterrupted ? error.exitCode : 1;
  failureDetail = redactSensitiveText(error instanceof Error ? error.message : "Unexpected non-Error failure");
  if (!Object.values(pipeline).some(stage => stage.status === "failed")) console.error(failureDetail);
} finally {
  if (!ownedWorkStopped) {
    stages.cleanup.status = "failed";
    stages.cleanup.detail = "[cleanup] Owned work could not be stopped safely; resource deletion was not started.";
  } else if (!runId || keepResources) {
    stages.cleanup.status = "skipped";
    stages.cleanup.detail = runId ? "Resources retained by explicit request; charges may continue." : "No owned run manifest was initialized.";
  } else {
    try {
      await step(stages.cleanup, () => child(stages.cleanup, ["--import", "tsx", join(root, "src", "cleanup-cli.ts")]));
    } catch (error) {
      process.exitCode = 1;
      const cleanupError = redactSensitiveText(error instanceof Error ? error.message : "Unexpected cleanup failure");
      failureDetail = failureDetail ? `${failureDetail}\nCleanup also failed: ${cleanupError}` : cleanupError;
    }
  }
  try {
    await releaseAwake();
  } catch (error) {
    process.exitCode = 1;
    const runtimeError = error instanceof Error ? error.message : "Unexpected execution-request failure";
    failureDetail = failureDetail ? `${failureDetail}\n${runtimeError}` : runtimeError;
    console.error(runtimeError);
  }
  const results = Object.values(pipeline);
  const failed = failureDetail !== undefined || results.some(stage => stage.status === "failed");
  await writeFile(join(reportDirectory, "summary.json"), JSON.stringify({
    runId, status: failed ? "failed" : "passed", suite: "three-iot-data-paths",
    scope: values, keepResources, stages: results, failure: failureDetail,
  }, null, 2));
  await writeFile(join(reportDirectory, "results.junit.xml"), junitReport(runId, results, failureDetail));
  console.log(`\n${failed ? "FAIL" : "PASS"} Fabric IoT E2E scenarios`);
  console.log(scenarioResultsTree(results));
  console.log(`JUnit: ${join(reportDirectory, "results.junit.xml")}`);
  console.log(`JSON:  ${join(reportDirectory, "summary.json")}`);
  shutdown.dispose();
}
