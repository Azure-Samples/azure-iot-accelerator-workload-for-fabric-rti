import test from "node:test";
import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { runSimulationWork } from "../src/data.js";
import { captureProcessOwner, processOwnerAlive } from "../src/process-owner.js";

test("a failed device pump cannot tear down transports before the running callback settles", async () => {
  let release!: () => void;
  const gate = new Promise<void>(resolve => { release = resolve; });
  let callbackFinished = false, stopped = false, pumpFailed = false;
  const work = runSimulationWork(async () => {
    await gate;
    callbackFinished = true;
  }, Promise.reject(new Error("synthetic pump fault")), () => {
    assert.equal(callbackFinished, true);
    stopped = true;
  }, () => { pumpFailed = true; });
  await new Promise<void>(resolve => setImmediate(resolve));
  assert.equal(pumpFailed, true);
  assert.equal(stopped, false);
  release();
  await assert.rejects(work, /synthetic pump fault/);
  assert.equal(stopped, true);
});

test("a callback failure aborts and joins the pump before returning", async () => {
  let settle!: () => void;
  const pump = new Promise<void>(resolve => { settle = resolve; });
  let joined = false;
  await assert.rejects(runSimulationWork(async () => { throw new Error("scenario fault"); },
    pump.then(() => { joined = true; }), settle, () => assert.fail("Pump must not fail")), /scenario fault/);
  assert.equal(joined, true);
});

test("process ownership distinguishes the live generation, PID reuse, another run and another host", async () => {
  const runId = randomUUID(), owner = await captureProcessOwner(runId);
  assert.ok(owner);
  assert.equal(await processOwnerAlive(owner, runId), true);
  assert.equal(await processOwnerAlive({ ...owner, stamp: `${owner.stamp}-previous-generation` }, runId), false);
  await assert.rejects(processOwnerAlive(owner, randomUUID()), /another run or host/);
  await assert.rejects(processOwnerAlive({ ...owner, host: "different-host" }, runId), /another run or host/);
});
