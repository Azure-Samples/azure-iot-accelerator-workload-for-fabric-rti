import { execFile, spawn, type ChildProcess } from "node:child_process";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { setTimeout as pause } from "node:timers/promises";
import { captureProcessOwner, processOwnerAlive, type ProcessOwner } from "./process-owner.js";

const execute = promisify(execFile);

export class RunInterrupted extends Error {
  constructor(readonly signal: "SIGINT" | "SIGTERM") {
    super(`[runner] Interrupted by ${signal}; stopping owned work before cleanup.`);
  }
  get exitCode(): number { return this.signal === "SIGINT" ? 130 : 143; }
}

export class OwnedChildShutdownError extends Error {}

export function shutdownController() {
  const controller = new AbortController();
  const interrupt = (signal: "SIGINT" | "SIGTERM") => {
    if (!controller.signal.aborted) controller.abort(new RunInterrupted(signal));
  };
  const sigint = () => interrupt("SIGINT"), sigterm = () => interrupt("SIGTERM");
  process.on("SIGINT", sigint);
  process.on("SIGTERM", sigterm);
  return {
    signal: controller.signal,
    dispose() { process.off("SIGINT", sigint); process.off("SIGTERM", sigterm); },
  };
}

export async function stopOwnedChild(
  child: Pick<ChildProcess, "pid" | "exitCode" | "signalCode">,
  owner: ProcessOwner,
): Promise<void> {
  if (child.pid !== owner.pid) throw new OwnedChildShutdownError("[ownership] Child process ownership does not match.");
  if (child.exitCode !== null || child.signalCode !== null) {
    throw new OwnedChildShutdownError("[ownership] Owned root exited before its remaining descendants could be verified; cleanup is deferred.");
  }
  if (!await processOwnerAlive(owner, owner.runId)) {
    throw new OwnedChildShutdownError("[ownership] Cannot stop a different process generation.");
  }
  if (process.platform === "win32") {
    try {
      await execute("powershell.exe", ["-NoLogo", "-NoProfile", "-NonInteractive", "-File",
        fileURLToPath(new URL("../scripts/stop-owned-process.ps1", import.meta.url)),
        "-TargetPid", String(owner.pid), "-StartTicks", owner.stamp],
      { windowsHide: true, timeout: 30_000 });
    } catch (error) {
      throw new OwnedChildShutdownError("[runtime] Could not confirm termination of the owned Windows process tree.", { cause: error });
    }
    return;
  }
  const processes = async () => {
    const { stdout } = await execute("ps", ["-axo", "pid=,ppid=,pgid=,stat="], { timeout: 5000 });
    return stdout.trim().split("\n").filter(Boolean).map(line => {
      const [pid, parent, group, state] = line.trim().split(/\s+/);
      return { pid: Number(pid), parent: Number(parent), group: Number(group), state: state ?? "" };
    });
  };
  const snapshot = await processes(), ids = new Set([owner.pid]);
  for (let count = 0; count < 128; count++) {
    const before = ids.size;
    for (const process of snapshot) if (ids.has(process.parent)) ids.add(process.pid);
    if (ids.size === before) break;
    if (ids.size > 128) throw new OwnedChildShutdownError("[ownership] Owned process tree exceeded the bounded shutdown limit.");
  }
  const identities = new Map<number, ProcessOwner>([[owner.pid, owner]]);
  for (const pid of ids) {
    if (pid === owner.pid) continue;
    const identity = await captureProcessOwner(owner.runId, pid);
    if (identity) identities.set(pid, identity);
  }
  const currentTree = await processes(), currentIds = new Set([owner.pid]);
  for (let count = 0; count < 128; count++) {
    const before = currentIds.size;
    for (const process of currentTree) if (currentIds.has(process.parent)) currentIds.add(process.pid);
    if (currentIds.size === before) break;
    if (currentIds.size > 128) throw new OwnedChildShutdownError("[ownership] Owned process tree exceeded the bounded shutdown limit.");
  }
  for (const [pid, identity] of identities) {
    if (!currentIds.has(pid) && await processOwnerAlive(identity, owner.runId)) {
      throw new OwnedChildShutdownError("[ownership] Process ancestry changed before shutdown; refusing an unverified target.");
    }
    if (!await processOwnerAlive(owner, owner.runId)) {
      throw new OwnedChildShutdownError("[ownership] Owned root process generation changed before shutdown.");
    }
  }
  const groups = new Set(currentTree.filter(process => currentIds.has(process.pid) && identities.has(process.group)).map(process => process.group));
  const signalTree = async (signal: "SIGTERM" | "SIGKILL") => {
    for (const group of groups) {
      const identity = identities.get(group)!;
      const current = await captureProcessOwner(owner.runId, group);
      if (current && current.stamp !== identity.stamp) {
        throw new OwnedChildShutdownError("[ownership] An owned process-group ID was reused.");
      }
      try { process.kill(-group, signal); } catch (error) {
        if (!(error instanceof Error && "code" in error && error.code === "ESRCH")) throw error;
      }
    }
    for (const process of snapshot) {
      if (!ids.has(process.pid) || groups.has(process.group)) continue;
      const identity = identities.get(process.pid);
      if (identity && await processOwnerAlive(identity, owner.runId)) {
        try { globalThis.process.kill(process.pid, signal); } catch (error) {
          if (!(error instanceof Error && "code" in error && error.code === "ESRCH")) throw error;
        }
      }
    }
  };
  const alive = async () => (await processes()).some(process =>
    (ids.has(process.pid) || groups.has(process.group)) && !/^[ZX]/.test(process.state));
  await signalTree("SIGTERM");
  const deadline = Date.now() + 30_000;
  while (await alive() && Date.now() < deadline) await pause(50);
  if (await alive()) {
    await signalTree("SIGKILL");
    const forced = Date.now() + 5000;
    while (await alive() && Date.now() < forced) await pause(50);
    if (await alive()) throw new OwnedChildShutdownError("[runtime] Owned process group termination was not confirmed.");
  }
}

export async function runOwnedChild(
  args: string[],
  options: {
    cwd: string;
    env: NodeJS.ProcessEnv;
    runId: string;
    signal?: AbortSignal;
    output: (chunk: Buffer, stream: NodeJS.WriteStream) => void;
  },
): Promise<number> {
  options.signal?.throwIfAborted();
  const child = spawn(process.execPath, args, {
    cwd: options.cwd, env: options.env, detached: true, windowsHide: true, stdio: ["inherit", "pipe", "pipe"],
  });
  child.stdout.on("data", (chunk: Buffer) => options.output(chunk, process.stdout));
  child.stderr.on("data", (chunk: Buffer) => options.output(chunk, process.stderr));
  const completion = new Promise<number>((resolve, reject) => {
    child.once("error", reject);
    child.once("close", (code, signal) => signal
      ? reject(new Error(`[runner] Owned child terminated by ${signal}.`)) : resolve(code ?? 1));
  });
  const owner = (child.pid ? captureProcessOwner(options.runId, child.pid) : Promise.resolve(undefined)).catch(error => {
    throw new OwnedChildShutdownError("[ownership] Could not establish the owned stage process generation.", { cause: error });
  });
  const ownershipFailure = owner.then(() => new Promise<never>(() => {}));
  let cancellation: Promise<never> | undefined;
  let rejectCancellation!: (error: unknown) => void;
  const cancelled = new Promise<never>((_, reject) => { rejectCancellation = reject; });
  const interrupt = () => {
    if (cancellation) return;
    cancellation = (async () => {
      try {
        const identity = await owner;
        if (!identity) throw new OwnedChildShutdownError("[ownership] Child process ownership was not established; cleanup is deferred.");
        await stopOwnedChild(child, identity);
        // A signal exit is expected here; the original interrupt remains the reported failure.
        await completion.catch(() => {});
      } catch (error) {
        if (error instanceof OwnedChildShutdownError) throw error;
        throw new OwnedChildShutdownError("[runtime] Owned child shutdown could not be verified.", { cause: error });
      }
      throw options.signal?.reason ?? new Error("[runner] Owned work was cancelled.");
    })();
    void cancellation.catch(rejectCancellation);
  };
  options.signal?.addEventListener("abort", interrupt, { once: true });
  if (options.signal?.aborted) interrupt();
  try {
    const code = await Promise.race([completion, cancelled, ownershipFailure]);
    await owner;
    if (cancellation) await cancellation;
    return code;
  } catch (error) {
    if (cancellation) await cancellation;
    throw error;
  } finally {
    options.signal?.removeEventListener("abort", interrupt);
  }
}
