// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { Button } from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { callNavigationNavigate } from "../../../controller/NavigationController";
import { DashboardField } from "../DashboardDefinitionGenerator";
import "../IoTSolutionItem.scss";

interface DashboardSummaryStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Summarizes the generated dashboard, explains how to use it, and directs the
 * user to optional modeling when raw-data semantics are no longer sufficient.
 */
export function DashboardSummaryStep({
  wizardContext,
  workloadClient,
  workspaceId,
}: DashboardSummaryStepProps) {
  const dashboardId = wizardContext.dashDashboardId || "";
  const dashboardName = wizardContext.dashDashboardName || "Real-Time Dashboard";
  const telemetryTable = wizardContext.dashTelemetryTable || "Raw telemetry table";
  const propertiesTable = wizardContext.dashPropertiesTable || "Raw properties table";
  const telemetryFields: DashboardField[] = wizardContext.dashTelemetryFields || [];
  const propertyFields: DashboardField[] = wizardContext.dashPropertyFields || [];
  const selectedTelemetry = telemetryFields.filter((field) => field.included);
  const selectedProperties = propertyFields.filter((field) => field.included);
  const timeSeriesCount = selectedTelemetry.filter((field) => field.isNumeric).length;
  const recentValueCount = selectedTelemetry.length - timeSeriesCount;

  const openDashboard = () => {
    if (!dashboardId) return;
    callNavigationNavigate(
      workloadClient,
      "host",
      `/groups/${workspaceId}/kustodashboards/${dashboardId}`
    ).catch(() => {});
  };

  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div className="iot-solution-success-heading">
        <div className="iot-solution-success-title-row">
          <h2 className="iot-solution-step-title">
            Your real-time device dashboard is ready
          </h2>
        </div>
        <p className="iot-solution-step-description">
          <strong>{dashboardName}</strong> provides an immediate view of live and
          recent device behavior using the raw telemetry and property data in your
          Eventhouse. You can use it as generated or extend its queries and visuals
          for your operational scenarios.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">How to use the dashboard</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Select a value for the <strong>Device ID</strong> parameter to display
            data for a specific device.
          </li>
          <li>
            Adjust the dashboard time range to investigate current behavior or
            historical trends.
          </li>
          <li>
            Edit the generated KQL queries, tiles, and layout to add calculations,
            thresholds, comparisons, or other business-specific views.
          </li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Dashboard composition</h3>
        <div className="iot-solution-architecture" aria-label="Generated dashboard composition">
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Raw tables</span>
            <span className="iot-solution-flow-node-detail">
              {telemetryTable} and {propertiesTable}
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Generated KQL queries</span>
            <span className="iot-solution-flow-node-detail">
              Filtered by device and time range
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Dashboard visuals</span>
            <span className="iot-solution-flow-node-detail">
              {timeSeriesCount} charts, {recentValueCount} tables,{" "}
              {selectedProperties.length} value cards
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Operational insight</span>
            <span className="iot-solution-flow-node-detail">
              Trends, state, and device-level investigation
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-prerequisites">
        <h3 className="iot-solution-outcome-heading">A useful next step</h3>
        <p className="iot-solution-step-description">
          This dashboard operates directly on raw device data. If you need consistent
          schemas, semantic context, or business-friendly field names, return to the
          accelerator home page and choose <strong>Model your device data</strong>.
        </p>
      </section>

      <div className="iot-solution-summary-actions">
        <Button appearance="primary" onClick={openDashboard} disabled={!dashboardId}>
          Open dashboard in Fabric
        </Button>
        <span className="iot-solution-finish-hint">
          Select <strong>Finish and Close</strong> when you are done.
        </span>
      </div>
    </div>
  );
}
