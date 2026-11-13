// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { StreamSample } from "./ModelAuthoring.ts";
import type { RoleEvidenceProfile } from "./HubModeling.ts";

export interface SourceSelectionStatus { key: string; ready: boolean; reason: string }
export interface SourceReadiness extends SourceSelectionStatus { revision: number }
export const SOURCE_PENDING = "Select both source tables, then choose Validate source tables.";
export const SOURCE_ANALYSIS_LIMIT = 1000;
export interface SourceAnalysis {
  key: string;
  selectionRevision: number;
  limit: typeof SOURCE_ANALYSIS_LIMIT;
  telemetry: StreamSample;
  twin: StreamSample;
  roleEvidence: RoleEvidenceProfile;
  analyzedAt: string;
}
export function sourceSelectionRevision(context: Record<string, unknown>): number {
  return typeof context.hubSourceSelectionRevision === "number" && Number.isSafeInteger(context.hubSourceSelectionRevision) &&
    context.hubSourceSelectionRevision >= 0 ? context.hubSourceSelectionRevision : 0;
}
export function currentSourceAnalysis(workspaceId: string, context: Record<string, any>): SourceAnalysis | undefined {
  const analysis: SourceAnalysis | undefined = context.hubSourceAnalysis;
  return analysis?.key === modelSourceKey(workspaceId, context) &&
    analysis.selectionRevision === sourceSelectionRevision(context) && analysis.limit === SOURCE_ANALYSIS_LIMIT &&
    !!analysis.telemetry && !!analysis.twin && Array.isArray(analysis.roleEvidence?.bindings) ? analysis : undefined;
}
export function modelSourceKey(workspaceId: string, context: Record<string, unknown>): string {
  return JSON.stringify([workspaceId, context.modelEhId || "", context.modelDbId || "", context.modelDbName || "",
    context.modelQueryServiceUri || "", context.rawTelemetryTable || "", context.rawPropertiesTable || ""]);
}
export function sourceSelectionStatus(input: {
  key: string; loading: boolean; error: string; eventhouseId: string; databaseId: string; queryServiceUri: string;
  eventhouseIds: string[]; databaseIds: string[]; tables: string[]; telemetry: string; properties: string;
}): SourceSelectionStatus {
  const sameTable = !!input.telemetry && input.telemetry === input.properties;
  const ready = !input.loading && !input.error && !!input.queryServiceUri &&
    input.eventhouseIds.includes(input.eventhouseId) && input.databaseIds.includes(input.databaseId) &&
    input.tables.includes(input.telemetry) && input.tables.includes(input.properties) && !sameTable;
  return { key: input.key, ready, reason: input.error
    ? "Resolve the source selection error before inspecting samples or uploading a model."
    : input.loading ? "Wait for the source tables to finish loading."
      : sameTable ? "Choose different tables for telemetry and reported properties."
        : ready ? "Source tables selected." : SOURCE_PENDING };
}
export function sourceActionCurrent(start: SourceReadiness, current: SourceReadiness): boolean {
  return start.ready && current.ready && start.key === current.key && start.revision === current.revision;
}
