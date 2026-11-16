// IDs are stable references: append new scenarios rather than renumbering existing ones.
export const scenarioGroups = [
  { id: "1", name: "Setup" },
  { id: "2", name: "Ingestion" },
  { id: "3", name: "Modeling" },
  { id: "4", name: "Dashboards" },
  { id: "5", name: "Data Agents" },
  { id: "6", name: "Browser recording checks" },
  { id: "7", name: "Activator" },
  { id: "8", name: "Cleanup" },
] as const;

export const scenarios = {
  initialize: { id: "1.1", name: "Initialize owned run" },
  deployment: { id: "1.2", name: "Check workload availability (no provisioning)" },
  item: { id: "1.3", name: "Create fresh workload item in the run folder" },
  provision: { id: "1.4", name: "Provision five S1 Hubs (system/UAMI identities)" },
  links: { id: "1.5", name: "Verify native ADR/DPS topology A/A/B" },
  ingestion: { id: "2.1", name: "Nonlinked: two Hubs, UI ingestion and native DO data" },
  combined: { id: "2.2", name: "Combined: five Hubs, new shared raw outputs" },
  linked: { id: "2.3", name: "Linked-only: three positives, two active negative controls" },
  regularModel: { id: "3.1", name: "DTDL Nonlinked" },
  combinedModel: { id: "3.2", name: "WoT Combined" },
  linkedModel: { id: "3.3", name: "WoT Linked-only" },
  regularDashboard: { id: "4.1", name: "Nonlinked: dashboard queries, visuals and folder placement" },
  combinedDashboard: { id: "4.2", name: "Combined: dashboard source bindings and visuals" },
  linkedDashboard: { id: "4.3", name: "Linked-only: dashboard source bindings and visuals" },
  regularAgent: { id: "5.1", name: "Nonlinked: DTDL-modeled Data Agent grounding" },
  combinedAgent: { id: "5.2", name: "Combined: WoT-modeled Data Agent grounding" },
  linkedAgent: { id: "5.3", name: "Linked-only: WoT-modeled Data Agent grounding" },
  videoPass: { id: "6.1", name: "Persist video after a successful workflow" },
  videoFailure: { id: "6.2", name: "Persist video after an intentional workflow failure" },
  pageVisual: { id: "6.3", name: "Offline: visual checks preserve assertion outcomes" },
  pageScopeRegular: { id: "6.4", name: "Offline: nonlinked scoping controls" },
  pageScopeCombined: { id: "6.5", name: "Offline: combined scoping controls" },
  pageScopeLinked: { id: "6.6", name: "Offline: linked-only scoping controls" },
  pageAssignCombined: { id: "6.7", name: "Offline: assign only the unassigned device" },
  pageAssignRegular: { id: "6.8", name: "Offline: nonlinked assignment unavailable" },
  pageAssignLinked: { id: "6.9", name: "Offline: linked devices already assigned" },
  pageDashboard: { id: "6.10", name: "Offline: native dashboard data check" },
  pageAgent: { id: "6.11", name: "Offline: native agent answer check" },
  pageActivator: { id: "6.12", name: "Offline: native Activator arrival check" },
  pageReuse: { id: "6.13", name: "Offline: reuse exact raw tables and Eventstreams" },
  pageComponents: { id: "6.14", name: "Offline: dashboard component field discovery" },
  pageCatalog: { id: "6.15", name: "Offline: exact workload selection with preview badge" },
  pageModelSource: { id: "6.16", name: "Offline: current modeling overview and source flow" },
  pageOwnedTabs: { id: "6.17", name: "Offline: close only manifest-owned Fabric tabs" },
  pageStreamFailure: { id: "6.18", name: "Offline: surface native Eventstream creation errors" },
  pageHubSelector: { id: "6.19", name: "Offline: select a Hub without refreshing its discovery" },
  pageAssignmentOutcome: { id: "6.20", name: "Offline: reject unsuccessful or unknown assignment completion" },
  pageModelFields: { id: "6.21", name: "Offline: upload and select DTDL and WoT fields" },
  pageModelCreation: { id: "6.22", name: "Offline: create modeled data with exact output names" },
  pageModelReady: { id: "6.23", name: "Offline: verify modeled resource summaries" },
  pageIngestionOverview: { id: "6.24", name: "Offline: optional ADR prerequisites and documentation links" },
  regularActivator: { id: "7.1", name: "Nonlinked: Activator setup and source bindings" },
  combinedActivator: { id: "7.2", name: "Combined: Activator setup and source bindings" },
  linkedActivator: { id: "7.3", name: "Linked-only: Activator setup and source bindings" },
  cleanup: { id: "8.1", name: "Delete and verify all run-owned test resources" },
} as const;

export type ScenarioKey = keyof typeof scenarios;
export type ScenarioId = (typeof scenarios)[ScenarioKey]["id"];

export function scenarioTitle(key: ScenarioKey): string {
  const scenario = scenarios[key];
  return `${scenario.id} ${scenario.name}`;
}
