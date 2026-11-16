import test from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, readFile, rmdir, stat, unlink, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { recordingDirectoryName, workflowRecording, workflowVideoName } from "../src/video-names.js";
import { saveWorkflowVideo } from "../src/recording.js";

test("workflow videos use stable scenario IDs, meaningful names and actual outcomes instead of page hashes", () => {
  assert.equal(workflowVideoName("item", "passed"), "1.3-create-workload-item-passed.webm");
  assert.equal(workflowVideoName("regular", "passed"), "2.1-nonlinked-ingestion-passed.webm");
  assert.equal(workflowVideoName("combined-modeling", "failed"), "3.2-combined-modeling-failed.webm");
  assert.equal(workflowVideoName("linked-dashboard-native", "failed"), "4.3-linked-only-dashboard-native-failed.webm");
  assert.equal(workflowVideoName("combined-data-agent", "passed", 1), "5.2-combined-data-agent-passed-tab-02.webm");
  assert.equal(workflowRecording("regular-modeling").id, "3.1");
  assert.match(workflowRecording("linked-activator").title, /Linked-only.*Activator/);
});

test("video finalization saves the meaningful file before deleting the raw Playwright artifact", async () => {
  const directory = await mkdtemp(join(tmpdir(), "e2e-video-name-"));
  const original = join(directory, "page@synthetic.webm");
  const named = join(directory, workflowVideoName("regular", "passed"));
  try {
    await writeFile(original, "synthetic recording bytes");
    const video = { async path() { return original; } };
    assert.equal(await saveWorkflowVideo(video, directory, "regular", "passed"), "2.1-nonlinked-ingestion-passed.webm");
    assert.equal(await readFile(named, "utf8"), "synthetic recording bytes");
    await assert.rejects(() => stat(original), { code: "ENOENT" });
    await writeFile(original, "a different recording");
    await assert.rejects(() => saveWorkflowVideo(video, directory, "regular", "passed"), { code: "EEXIST" });
    assert.equal(await readFile(named, "utf8"), "synthetic recording bytes");
    assert.equal(await readFile(original, "utf8"), "a different recording", "Never remove the original when saving fails.");
  } finally {
    for (const path of [original, named]) {
      try { await unlink(path); } catch (error) {
        if (!(error && typeof error === "object" && "code" in error && error.code === "ENOENT")) throw error;
      }
    }
    await rmdir(directory);
  }
});

test("recording directories remain unique and sortable without hiding the scenario behind a UUID", () => {
  const first = recordingDirectoryName("regular", "2026-11-04T18:00:00.001Z");
  const second = recordingDirectoryName("regular", "2026-11-04T18:00:00.002Z");
  assert.equal(first, "2.1-nonlinked-ingestion-2026-11-04T18-00-00-001Z");
  assert.ok(first < second);
  assert.ok(!/[<>:"/\\|?*]/.test(first));
  assert.throws(() => workflowRecording("unknown-scenario"), /Unknown/);
  assert.throws(() => recordingDirectoryName("regular", "invalid"), /invalid/);
  assert.throws(() => workflowVideoName("regular", "passed", -1), /invalid/);
});
