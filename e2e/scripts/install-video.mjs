import { spawnSync } from "node:child_process";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("../", import.meta.url));
const result = spawnSync(process.execPath, [join(root, "node_modules", "@playwright", "test", "cli.js"), "install", "ffmpeg"], {
  cwd: root, stdio: "inherit",
  env: { ...process.env, PLAYWRIGHT_BROWSERS_PATH: process.env.PLAYWRIGHT_BROWSERS_PATH ?? join(root, "artifacts", "browsers") },
});
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
