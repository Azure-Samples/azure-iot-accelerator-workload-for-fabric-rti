import { ActivatorAttempt } from "./activator.js";
import { Cloud } from "./cloud.js";
import { requireGuid, runConfig } from "./config.js";
import { cleanupRun } from "./cleanup.js";
import { readPlan } from "./plan.js";
import { keepSystemAwake } from "./awake.js";

const config = runConfig();
const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
const cloud = new Cloud(config);
const releaseAwake = await keepSystemAwake();
try {
  await cleanupRun(cloud, plan, () => ActivatorAttempt.cleanupJournals(cloud, plan));
} finally {
  await releaseAwake();
}
