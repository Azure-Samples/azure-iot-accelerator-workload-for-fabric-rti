import { execFile } from "node:child_process";
import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import { hostname } from "node:os";
import { promisify } from "node:util";
import { requireGuid } from "./config.js";
import { object, text } from "./cloud.js";
import { updatePlan, type RunPlan } from "./plan.js";

const execute = promisify(execFile);
const host = () => createHash("sha256").update(`${process.platform}\0${hostname()}`).digest("hex");

export interface ProcessOwner {
  runId: string;
  pid: number;
  stamp: string;
  host: string;
}

export async function processStamp(pid: number): Promise<string | undefined> {
  if (!Number.isSafeInteger(pid) || pid <= 0) throw new Error("[ownership] Invalid recorded process ID.");
  try { process.kill(pid, 0); } catch (error) {
    if (error instanceof Error && "code" in error && error.code === "ESRCH") return undefined;
    throw new Error("[ownership] Cannot verify the recorded process.");
  }
  if (process.platform === "win32") {
    try {
      const { stdout } = await execute("powershell.exe", ["-NoLogo", "-NoProfile", "-NonInteractive", "-Command",
        `$p=Get-Process -Id ${pid} -ErrorAction SilentlyContinue; if($null -eq $p){exit 3}; [Console]::Write($p.StartTime.ToUniversalTime().Ticks)`],
      { encoding: "utf8", windowsHide: true, timeout: 5000 });
      if (!/^\d+$/.test(stdout.trim())) throw new Error("Missing native process start time.");
      return stdout.trim();
    } catch (error) {
      if (error instanceof Error && "code" in error && error.code === 3) return undefined;
      throw new Error("[ownership] Cannot read the recorded Windows process generation.");
    }
  }
  if (process.platform === "linux") {
    let stat: string;
    try {
      stat = await readFile(`/proc/${pid}/stat`, "utf8");
    } catch (error) {
      if (error instanceof Error && "code" in error && error.code === "ENOENT") return undefined;
      throw new Error("[ownership] Cannot read the recorded Linux process generation.");
    }
    const fields = stat.slice(stat.lastIndexOf(")") + 2).trim().split(/\s+/);
    if (fields[0] === "Z" || fields[0] === "X") return undefined;
    const started = fields[19];
    if (!started || !/^\d+$/.test(started)) throw new Error("[ownership] Missing native process start time.");
    const boot = (await readFile("/proc/sys/kernel/random/boot_id", "utf8")).trim();
    if (!boot) throw new Error("[ownership] Cannot establish the current boot identity.");
    return `${boot}:${started}`;
  }
  const { stdout } = await execute("ps", ["-p", String(pid), "-o", "lstart="], {
    encoding: "utf8", timeout: 5000, env: { ...process.env, LC_ALL: "C" },
  });
  if (!stdout.trim()) throw new Error("[ownership] Cannot read the recorded process generation.");
  return stdout.trim();
}

export async function captureProcessOwner(runId: string, pid = process.pid): Promise<ProcessOwner | undefined> {
  requireGuid(runId, "process ownership run ID");
  const stamp = await processStamp(pid);
  return stamp === undefined ? undefined : { runId, pid, stamp, host: host() };
}

export async function processOwnerAlive(owner: ProcessOwner, runId: string): Promise<boolean> {
  if (owner.runId !== runId || owner.host !== host() || typeof owner.stamp !== "string" || !owner.stamp) {
    throw new Error("[ownership] Process ownership belongs to another run or host.");
  }
  return await processStamp(owner.pid) === owner.stamp;
}

export async function assertSimulationStopped(plan: RunPlan): Promise<void> {
  if (!["starting", "running-in-process", "stopping"].includes(plan.observed.simulationState ?? "")) return;
  const raw = plan.observed.simulationOwner;
  if (!raw) throw new Error("[cleanup] Stop the run's local devices; this manifest has no process-liveness proof.");
  const value = object(JSON.parse(raw), "simulator process ownership");
  const pid = value.pid;
  if (typeof pid !== "number") throw new Error("[ownership] Missing simulator process ID.");
  const owner: ProcessOwner = { runId: text(value.runId, "simulator run ID"), pid,
    stamp: text(value.stamp, "simulator process generation"), host: text(value.host, "simulator host") };
  if (await processOwnerAlive(owner, plan.runId)) {
    throw new Error("[cleanup] Stop the run's local devices before deleting their fixture; the recorded process is still alive.");
  }
  plan.observed.simulationState = "stopped";
  plan.observed.simulationStoppedAt = new Date().toISOString();
  plan.observed.simulationStopReason = "recorded-owner-process-exited";
  await updatePlan(plan);
  console.log("[simulation] Confirmed the recorded process generation has exited; no live process was stopped.");
}
