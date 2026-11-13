import { openBrowser, signIn } from "./auth.js";
import { Cloud } from "./cloud.js";
import { requireGuid, runConfig } from "./config.js";
import { requireDeploymentMode } from "./deployment.js";
import { readPlan } from "./plan.js";
import { provision } from "./provision.js";
import { keepSystemAwake } from "./awake.js";

const config = runConfig();
const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
const releaseAwake = await keepSystemAwake();
const context = await openBrowser(config);
try {
  const page = context.pages()[0] ?? await context.newPage();
  await signIn(page, config);
  await requireDeploymentMode(page, config);
  await provision(new Cloud(config), plan);
  console.log("Five owned S1 Hubs passed native checks; linked-Hub, ADR and DPS setup is deferred.");
  console.log("Resources are retained. Fabric ingestion and native data checks have not run.");
} finally {
  try { await context.close(); } finally { await releaseAwake(); }
}
