import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

export async function keepSystemAwake(): Promise<() => Promise<void>> {
  if (process.platform !== "win32") return async () => {};
  const script = fileURLToPath(new URL("../scripts/keep-awake.ps1", import.meta.url));
  const child = spawn("powershell.exe", ["-NoLogo", "-NoProfile", "-NonInteractive", "-File", script], {
    stdio: ["pipe", "pipe", "pipe"], windowsHide: true,
  });
  let stderr = "", ready = false, released = false, unexpectedExit = false;
  child.stderr.on("data", (chunk: Buffer) => { stderr = (stderr + chunk.toString()).slice(-4000); });
  const exited = new Promise<number | null>(resolve => child.once("exit", code => {
    unexpectedExit = ready && !released;
    if (unexpectedExit) console.error("[runtime] The system execution request ended before the CLI operation.");
    resolve(code);
  }));
  await new Promise<void>((resolve, reject) => {
    const timer = setTimeout(() => {
      child.stdin.end();
      reject(new Error("[runtime] Timed out acquiring the Windows system execution request."));
    }, 15_000);
    const fail = (message: string) => { clearTimeout(timer); reject(new Error(message)); };
    child.once("error", error => fail(`[runtime] Cannot start the Windows execution request: ${error.message}`));
    child.once("exit", code => { if (!ready) fail(`[runtime] Windows execution request failed (${code}): ${stderr}`); });
    let output = "";
    child.stdout.on("data", (chunk: Buffer) => {
      output += chunk.toString();
      if (!ready && output.includes("E2E_AWAKE_READY")) {
        ready = true;
        clearTimeout(timer);
        resolve();
      }
    });
  });
  console.log("[runtime] Holding a process-scoped Windows system execution request; power settings and screen-lock policy are unchanged.");
  return async () => {
    if (released) return;
    released = true;
    child.stdin.end();
    let timer: ReturnType<typeof setTimeout> | undefined;
    try {
      const code = await Promise.race([
        exited,
        new Promise<never>((_, reject) => { timer = setTimeout(() => reject(new Error("[runtime] Windows execution request did not release.")), 5000); }),
      ]);
      if (code !== 0 || unexpectedExit) throw new Error(`[runtime] Windows execution request ended unexpectedly (${code}): ${stderr}`);
    } finally {
      clearTimeout(timer);
    }
  };
}
