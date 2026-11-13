import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { setTimeout as pause } from "node:timers/promises";
import { keepSystemAwake } from "../src/awake.js";

test("the CLI execution request is acquired and released without a persistent power-policy change", async () => {
  const release = await keepSystemAwake();
  assert.equal(typeof release, "function");
  await pause(150);
  await release();
  await release();
});

test("Windows request does not disable screen lock or modify global power configuration", async () => {
  const source = await readFile(new URL("../scripts/keep-awake.ps1", import.meta.url), "utf8");
  assert.match(source, /SetThreadExecutionState\(\[uint32\]2147483649\)/);
  assert.match(source, /SetThreadExecutionState\(\[uint32\]2147483648\)/);
  assert.doesNotMatch(source, /powercfg|Set-ItemProperty|ExecutionPolicy|ES_DISPLAY_REQUIRED|ES_AWAYMODE_REQUIRED/);
});
