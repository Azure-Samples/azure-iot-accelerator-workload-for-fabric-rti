import { runConfig } from "./config.js";
import { createPlan, writePlan } from "./plan.js";

const plan = createPlan(runConfig());
const path = await writePlan(plan);
console.log(`Plan saved: ${path}`);
console.log("Planned only: 5 S1 Hubs in one new resource group. ADR namespaces, DPS and future linked-Hub associations are deferred.");
console.log("No resources created. UI and native raw-data verification have not run.");
