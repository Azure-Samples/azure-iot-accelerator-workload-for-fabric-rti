import test from "node:test";
import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import { readFileSync, rmdirSync, unlinkSync } from "node:fs";
import { dirname, join } from "node:path";
import { createStage, junitReport, scenarioResultsTree, type PipelineStage } from "../src/report.js";
import { artifacts, root } from "../src/config.js";
import { scenarioGroups, scenarios, scenarioTitle } from "../src/scenarios.js";

test("pipeline JUnit shows failed provisioning and unrun ingestion, never a false pass", () => {
  const report = junitReport("example-run", [
    { ...createStage("item"), status: "passed", durationMs: 100 },
    { ...createStage("provision"), status: "failed", durationMs: 2500, detail: "IH500019 <error> & request \"example\"" },
    createStage("ingestion"),
  ]);
  assert.match(report, /tests="3" failures="1" skipped="1"/);
  assert.match(report, /<failure message="IH500019 &lt;error&gt; &amp; request &quot;example&quot;">/);
  assert.match(report, /<skipped message="Not run because an earlier phase failed"\/>/);
  assert.match(report, /time="2\.600"/);
  assert.ok(report.includes(`name="${scenarioTitle("item")}"`));
});

test("JUnit failure output redacts credentials and terminal escape sequences", () => {
  const report = junitReport("example-run", [
    { ...createStage("ingestion"), status: "failed", durationMs: 1, detail: "\u001b[31mSharedAccessKey=fake-secret-value\u001b[0m" },
  ]);
  assert.ok(!report.includes("fake-secret-value"));
  assert.ok(!report.includes("\u001b"));
  assert.match(report, /REDACTED/);
});

test("JUnit includes runner-level failures without changing successful stage outcomes", () => {
  const stages: PipelineStage[] = [{ ...createStage("provision"), status: "passed", durationMs: 1 }];
  const report = junitReport("example-run", stages, "Runner prerequisite failed; SharedAccessKey=fake-runner-secret");
  assert.match(report, /tests="2" failures="1" skipped="0"/);
  assert.match(report, /classname="fabric.iot.e2e.runner" name="Runner failure"/);
  assert.doesNotMatch(report, /fake-runner-secret/);
  assert.equal(stages[0]?.status, "passed");
});

test("JUnit does not duplicate a failure already recorded by a stage", () => {
  const report = junitReport("example-run", [
    { ...createStage("ingestion"), status: "failed", durationMs: 1, detail: "Stage failed" },
  ], "Stage failed");
  assert.match(report, /tests="1" failures="1"/);
  assert.doesNotMatch(report, /name="Runner failure"/);
});

test("scenario IDs are unique, grouped and stable across code and reports", () => {
  const entries = Object.values(scenarios);
  assert.equal(new Set(entries.map(scenario => scenario.id)).size, entries.length);
  for (const scenario of entries) {
    assert.match(scenario.id, /^\d+\.\d+$/);
    assert.ok(scenarioGroups.some(group => scenario.id.startsWith(`${group.id}.`)));
  }
  assert.equal(scenarios.ingestion.id, "2.1");
  assert.equal(scenarios.combinedModel.id, "3.2");
  assert.equal(scenarios.linkedAgent.id, "5.3");
  assert.equal(scenarioTitle("regularModel"), "3.1 DTDL Nonlinked");
  assert.equal(scenarioTitle("combinedModel"), "3.2 WoT Combined");
  assert.equal(scenarioTitle("linkedModel"), "3.3 WoT Linked-only");
});

test("end-of-run tree preserves group order, stable IDs, outcomes and durations", () => {
  const stages: PipelineStage[] = [
    { ...createStage("regularModel"), status: "failed", durationMs: 2500, detail: "Missing native output\nLong stack trace" },
    { ...createStage("combined"), status: "skipped", detail: "Excluded by --nonlinked-only." },
    { ...createStage("ingestion"), status: "passed", durationMs: 1200 },
    createStage("regularDashboard"),
  ];
  assert.equal(scenarioResultsTree(stages), [
    "|-- 2 Ingestion",
    "|   |-- 2.1 Nonlinked: two Hubs, UI ingestion and native DO data [PASS] (1.2s)",
    "|   `-- 2.2 Combined: five Hubs, new shared raw outputs [SKIPPED] - Excluded by --nonlinked-only.",
    "|-- 3 Modeling",
    "|   `-- 3.1 DTDL Nonlinked [FAIL] (2.5s) - Missing native output",
    "`-- 4 Dashboards",
    "    `-- 4.1 Nonlinked: dashboard queries, visuals and folder placement [NOT RUN]",
  ].join("\n"));
});

test("tree handles a selected scenario, no results and sensitive failure messages", () => {
  const tree = scenarioResultsTree([
    { ...createStage("linkedAgent"), status: "failed", durationMs: 1, detail: "\u001b[31mSharedAccessKey=fake-secret-value\u001b[0m" },
  ]);
  assert.match(tree, /^`-- 5 Data Agents\n    `-- 5\.3 .* \[FAIL\] \(0\.0s\) - /);
  assert.ok(!tree.includes("fake-secret-value"));
  assert.ok(!tree.includes("\u001b"));
  assert.match(tree, /REDACTED/);
  assert.equal(scenarioResultsTree([]), "");
});

test("runner prints and persists numbered results on early failure without cloud access", () => {
  const result = spawnSync(process.execPath, ["--import", "tsx", join(root, "src", "run.ts"), "--nonlinked-only", "--ingestion-only", "--keep-resources"], {
    cwd: root,
    env: { ...process.env, E2E_TENANT_ID: "deliberately-invalid-before-cloud-access" },
    encoding: "utf8",
    timeout: 30_000,
  });
  assert.ifError(result.error);
  assert.equal(result.status, 1, result.stderr);
  assert.match(result.stdout, /START 1\.1 Initialize owned run/);
  assert.match(result.stdout, /FAIL Fabric IoT E2E scenarios/);
  assert.match(result.stdout, /1\.1 Initialize owned run \[FAIL\]/);
  assert.match(result.stdout, /2\.1 .+ \[NOT RUN\]/);
  assert.match(result.stdout, /2\.2 .+ \[SKIPPED\] - Excluded by --nonlinked-only\./);
  assert.match(result.stdout, /3\.1 .+ \[SKIPPED\] - Excluded by --ingestion-only\./);
  const summaryPath = /^JSON:  (.+)$/m.exec(result.stdout)?.[1]?.trim();
  assert.ok(summaryPath, "The runner must print the persisted JSON result path.");
  assert.equal(dirname(dirname(summaryPath)), join(artifacts, "reports"));
  try {
    const report: { status: string; stages: PipelineStage[]; keepResources: boolean; scope: { "keep-resources": boolean } } = JSON.parse(readFileSync(summaryPath, "utf8"));
    assert.equal(report.status, "failed");
    assert.equal(report.stages.find(stage => stage.id === "1.1")?.status, "failed");
    assert.equal(report.stages.find(stage => stage.id === "2.1")?.status, "not-run");
    assert.equal(report.keepResources, true);
    assert.equal(report.scope["keep-resources"], true);
  } finally {
    // Deliberate self-test failures must not be published as live E2E failures.
    unlinkSync(summaryPath);
    unlinkSync(join(dirname(summaryPath), "results.junit.xml"));
    rmdirSync(dirname(summaryPath));
  }
});

test("Playwright lists a selected scenario with its original ID, not a passing result", () => {
  const result = spawnSync(process.execPath, [
    join(root, "node_modules", "@playwright", "test", "cli.js"), "test", "--list", "--grep", "\\b3[.]2 ",
    "--reporter", join(root, "src", "scenario-reporter.ts"),
  ], {
    cwd: root,
    env: { ...process.env, E2E_PIPELINE_RUN: "0" },
    encoding: "utf8",
    timeout: 30_000,
  });
  assert.ifError(result.error);
  assert.equal(result.status, 0, result.stderr);
  assert.match(result.stdout, /NOT RUN Fabric IoT E2E scenarios/);
  assert.match(result.stdout, /`-- 3 Modeling\n    `-- 3\.2 .+ \[NOT RUN\]/);
  assert.ok(!result.stdout.includes("PASS Fabric"));
  assert.ok(!result.stdout.includes("3.1 "));
});

test("Playwright reports actual failure and blocked scenarios in the numbered tree", () => {
  const result = spawnSync(process.execPath, [
    join(root, "node_modules", "@playwright", "test", "cli.js"), "test", "scenarios/modeling",
    "--max-failures=1", "--output", join(root, "artifacts", "reporter-check", randomUUID()),
    "--reporter", join(root, "src", "scenario-reporter.ts"),
  ], {
    cwd: root,
    env: { ...process.env, E2E_PIPELINE_RUN: "0", E2E_TENANT_ID: "deliberately-invalid-before-cloud-access" },
    encoding: "utf8",
    timeout: 30_000,
  });
  assert.ifError(result.error);
  assert.equal(result.status, 1, result.stderr);
  assert.match(result.stdout, /FAIL Fabric IoT E2E scenarios/);
  assert.match(result.stdout, /3\.1 .+ \[NOT RUN\]/);
  assert.match(result.stdout, /3\.2 .+ \[FAIL\] \([\d.]+s\)/);
  assert.match(result.stdout, /3\.3 .+ \[NOT RUN\]/);
  assert.ok(!result.stdout.includes("Complete Microsoft Fabric sign-in"));
});
