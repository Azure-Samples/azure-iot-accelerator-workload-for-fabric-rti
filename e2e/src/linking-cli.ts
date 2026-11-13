import { openBrowser, signIn } from "./auth.js";
import { Cloud } from "./cloud.js";
import { requireGuid, runConfig } from "./config.js";
import { requireDeploymentMode } from "./deployment.js";
import { provisionLinks } from "./linking.js";
import { readPlan } from "./plan.js";
import { keepSystemAwake } from "./awake.js";

const config = runConfig();
const plan = await readPlan(config, requireGuid(process.env.E2E_RUN_ID, "E2E_RUN_ID"));
const releaseAwake = await keepSystemAwake();
const context = await openBrowser(config);
try {
  const page = context.pages()[0] ?? await context.newPage();
  await signIn(page, config);
  await requireDeploymentMode(page, config);
  await provisionLinks(new Cloud(config), plan);
  console.log("Native ADR Hub linkage verified: three owned Hubs, namespace A/A/B.");
} finally {
  try { await context.close(); } finally { await releaseAwake(); }
}
