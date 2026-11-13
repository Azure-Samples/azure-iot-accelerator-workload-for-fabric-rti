import { scenarios, type ScenarioKey } from "./scenarios.js";

const recordings = {
  item: { key: "item", label: "create-workload-item" },
  regular: { key: "ingestion", label: "nonlinked-ingestion" },
  combined: { key: "combined", label: "combined-ingestion" },
  "linked-only": { key: "linked", label: "linked-only-ingestion" },
  "regular-modeling": { key: "regularModel", label: "nonlinked-modeling" },
  "combined-modeling": { key: "combinedModel", label: "combined-modeling" },
  "linked-modeling": { key: "linkedModel", label: "linked-only-modeling" },
  "regular-dashboard": { key: "regularDashboard", label: "nonlinked-dashboard" },
  "combined-dashboard": { key: "combinedDashboard", label: "combined-dashboard" },
  "linked-dashboard": { key: "linkedDashboard", label: "linked-only-dashboard" },
  "regular-data-agent": { key: "regularAgent", label: "nonlinked-data-agent" },
  "combined-data-agent": { key: "combinedAgent", label: "combined-data-agent" },
  "linked-data-agent": { key: "linkedAgent", label: "linked-only-data-agent" },
  "regular-activator": { key: "regularActivator", label: "nonlinked-activator" },
  "combined-activator": { key: "combinedActivator", label: "combined-activator" },
  "linked-activator": { key: "linkedActivator", label: "linked-only-activator" },
  "video-pass": { key: "videoPass", label: "browser-recording-pass-check" },
  "video-failure": { key: "videoFailure", label: "browser-recording-failure-check" },
} as const satisfies Record<string, { key: ScenarioKey; label: string }>;

export function workflowRecording(scenario: string) {
  const native = scenario.endsWith("-native");
  const base = native ? scenario.slice(0, -"-native".length) : scenario;
  const recording = Object.entries(recordings).find(([alias]) => alias === base)?.[1];
  if (!recording) throw new Error("[video] Unknown recording scenario; add its stable scenario identity before recording.");
  const entry = scenarios[recording.key];
  return {
    id: entry.id,
    title: entry.name + (native ? " - native app verification" : ""),
    stem: `${entry.id}-${recording.label}${native ? "-native" : ""}`,
  };
}

export function recordingDirectoryName(scenario: string, startedAt: string): string {
  const date = new Date(startedAt);
  if (!Number.isFinite(date.getTime())) throw new Error("[video] Recording start time is invalid.");
  return `${workflowRecording(scenario).stem}-${date.toISOString().replaceAll(":", "-").replace(".", "-")}`;
}

export function workflowVideoName(scenario: string, outcome: "passed" | "failed", pageIndex = 0): string {
  if ((outcome !== "passed" && outcome !== "failed") || !Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new Error("[video] Recording outcome or page index is invalid.");
  }
  return `${workflowRecording(scenario).stem}-${outcome}${pageIndex ? `-tab-${String(pageIndex + 1).padStart(2, "0")}` : ""}.webm`;
}
