import { readdirSync } from "node:fs";
import { join } from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("../", import.meta.url));
const tests = readdirSync(join(root, "tests"))
  .filter(name => name.endsWith(".test.ts"))
  .sort()
  .map(name => join(root, "tests", name));
if (!tests.length) throw new Error("No unit test files found.");
const result = spawnSync(process.execPath, ["--import", "tsx", "--test", ...tests], {
  cwd: root, stdio: "inherit",
});
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
