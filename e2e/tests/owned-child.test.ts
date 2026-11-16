import test from "node:test";
import assert from "node:assert/strict";
import { spawn, spawnSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import { once } from "node:events";
import { join } from "node:path";
import { pathToFileURL } from "node:url";
import { root } from "../src/config.js";
import { OwnedChildShutdownError, runOwnedChild, stopOwnedChild } from "../src/owned-child.js";
import { captureProcessOwner, processStamp, type ProcessOwner } from "../src/process-owner.js";

test("cancellation stops the owned child and descendant, not a sibling process", async context => {
  const runId = randomUUID();
  const foreign = spawn(process.execPath, ["-e", "setInterval(()=>{},1000)"], {
    detached: true, windowsHide: true, stdio: "ignore",
  });
  assert.ok(foreign.pid);
  const foreignOwner = await captureProcessOwner(runId, foreign.pid);
  assert.ok(foreignOwner);
  context.after(() => stopOwnedChild(foreign, foreignOwner));
  const controller = new AbortController();
  let childPid = 0, descendantPid = 0, output = "";
  const script = `
    const {spawn}=require("node:child_process");
    const child=spawn(process.execPath,["-e","console.log('LEAF_READY');setInterval(()=>{},1000)"],{detached:true,windowsHide:true,stdio:["ignore","pipe","ignore"]});
    child.stdout.once("data",()=>console.log("READY:"+process.pid+":"+child.pid));
    setInterval(()=>{},1000);
  `;
  const work = runOwnedChild(["-e", script], {
    cwd: root, env: process.env, runId, signal: controller.signal,
    output(chunk) {
      output += chunk.toString();
      const ready = /READY:(\d+):(\d+)/.exec(output);
      if (!ready) return;
      childPid = Number(ready[1]);
      descendantPid = Number(ready[2]);
      controller.abort(new Error("synthetic cancellation"));
    },
  });
  await assert.rejects(work, /synthetic cancellation/);
  assert.ok(childPid && descendantPid);
  assert.equal(await processStamp(childPid), undefined);
  assert.equal(await processStamp(descendantPid), undefined);
  assert.equal(await processStamp(foreign.pid), foreignOwner.stamp);
});

test("a reused PID generation is never terminated", async context => {
  const runId = randomUUID();
  const child = spawn(process.execPath, ["-e", "setInterval(()=>{},1000)"], {
    detached: true, windowsHide: true, stdio: "ignore",
  });
  assert.ok(child.pid);
  const owner = await captureProcessOwner(runId, child.pid);
  assert.ok(owner);
  context.after(() => stopOwnedChild(child, owner));
  await assert.rejects(stopOwnedChild(child, { ...owner, stamp: `${owner.stamp}-old` }), /different process generation/);
  assert.equal(await processStamp(child.pid), owner.stamp);
});

test("an exited root cannot certify that its detached descendant stopped", async context => {
  const runId = randomUUID();
  const child = spawn(process.execPath, ["-e", `
    const {spawn}=require("node:child_process");
    const child=spawn(process.execPath,["-e","setInterval(()=>{},1000)"],{detached:true,windowsHide:true,stdio:"ignore"});
    console.log(child.pid);
    process.stdin.once("data",()=>process.exit(0));
  `], { detached: true, windowsHide: true, stdio: ["pipe", "pipe", "pipe"] });
  assert.ok(child.pid);
  const owner = await captureProcessOwner(runId, child.pid);
  assert.ok(owner);
  let descendant: ProcessOwner | undefined;
  context.after(async () => {
    for (const identity of [owner, descendant]) {
      if (identity && await processStamp(identity.pid) === identity.stamp) {
        await stopOwnedChild({ pid: identity.pid, exitCode: null, signalCode: null }, identity);
      }
    }
  });
  const [ready] = await once(child.stdout, "data");
  descendant = await captureProcessOwner(runId, Number(String(ready).trim()));
  assert.ok(descendant);
  const exited = once(child, "close");
  child.stdin.end("exit");
  await exited;
  await assert.rejects(stopOwnedChild(child, owner), error =>
    error instanceof OwnedChildShutdownError && /cleanup is deferred/.test(error.message));
  assert.equal(await processStamp(descendant.pid), descendant.stamp);
});

for (const signal of ["SIGINT", "SIGTERM"] as const) {
  test(`${signal} reaches cleanup once after owned work has stopped`, () => {
    const module = pathToFileURL(join(root, "src", "owned-child.ts")).href;
    const script = `
      import assert from "node:assert/strict";
      import {randomUUID} from "node:crypto";
      import {shutdownController,runOwnedChild,RunInterrupted} from ${JSON.stringify(module)};
      const shutdown=shutdownController();let cleanups=0,stopped=false,sent=false;
      try {
        await runOwnedChild(["-e","console.log('READY');setInterval(()=>{},1000)"],{
          cwd:${JSON.stringify(root)},env:process.env,runId:randomUUID(),signal:shutdown.signal,
          output(chunk){if(!sent&&chunk.toString().includes("READY")){sent=true;process.emit(${JSON.stringify(signal)});process.emit(${JSON.stringify(signal)});}}
        });
        assert.fail("Cancelled work must not pass");
      }catch(error){assert.ok(error instanceof RunInterrupted);stopped=true;}
      finally{assert.equal(stopped,true);cleanups++;shutdown.dispose();}
      assert.equal(cleanups,1);console.log("CLEANUP_ONCE");
    `;
    const result = spawnSync(process.execPath, ["--import", "tsx", "--input-type=module", "-e", script], {
      cwd: root, encoding: "utf8", timeout: 30_000,
    });
    assert.ifError(result.error);
    assert.equal(result.status, 0, result.stderr);
    assert.match(result.stdout, /CLEANUP_ONCE/);
  });
}
