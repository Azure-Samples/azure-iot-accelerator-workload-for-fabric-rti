// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useState } from "react";
import { Button, MessageBar, MessageBarBody } from "@fluentui/react-components";
import { ArrowRight20Regular } from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { callNavigationNavigate } from "../../../controller/NavigationController";
import "../IoTSolutionItem.scss";
import type { DocumentedColumn } from "../DataAgentContext.ts";

interface DataAgentSummaryStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Summarizes the configured Data Agent and explains how users can begin asking
 * operational questions in the Data Agent.
 */
export function DataAgentSummaryStep({ wizardContext, workloadClient, workspaceId }: DataAgentSummaryStepProps) {
  const agentId: string = wizardContext.agentId || "";
  const agentName = wizardContext.agentName || "Data Agent";
  const databaseName = wizardContext.agentDbName || "KQL database";
  const tableName = wizardContext.agentTableName || "Modeled data table";
  const columns: DocumentedColumn[] = wizardContext.agentColumns || [];
  const [navigationError, setNavigationError] = useState("");

  const openDataAgent = async () => {
    if (!workspaceId || !agentId) return;
    setNavigationError("");
    try {
      await callNavigationNavigate(workloadClient, "host", `/groups/${workspaceId}/aiskills/${agentId}`);
    } catch (error: unknown) {
      setNavigationError(error instanceof Error ? `Unable to open the Data Agent. ${error.message}` : "Unable to open the Data Agent.");
    }
  };

  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div className="iot-solution-success-heading">
        <div className="iot-solution-success-title-row">
          <h2 className="iot-solution-step-title">
            Your IoT Data Agent is ready
          </h2>
        </div>
        <p className="iot-solution-step-description">
          The agent is grounded in the real-time telemetry and current device state
          stored in the modeled data table. Its generated instructions and examples
          help it interpret the modeled fields and translate natural-language
          questions into queries.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">How to use your Data Agent</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Open your data agent <strong>{agentName}</strong>, and ask a
            question about your devices or operations.
          </li>
          <li>
            Include a device, time range, measurement, property, or operational
            condition when you need a more specific answer.
          </li>
          <li>
            Review the generated answer and query, then refine the question or
            agent instructions if business terminology needs more guidance.
          </li>
        </ul>

        <div className="iot-solution-config-card" style={{ marginTop: "12px" }}>
          <strong>Example questions</strong>
          <ul className="iot-solution-outcome-list">
            <li>Which devices have reported high temperature while in cooling mode?</li>
            <li>Compare average vibration by firmware version over the last 24 hours.</li>
            <li>Which devices have low battery voltage and are currently active?</li>
            <li>Show the latest measurements and reported state for a specific device.</li>
            <li>Which assets show unusual behavior compared with similar devices?</li>
          </ul>
        </div>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Agent configuration</h3>
        <div className="iot-solution-validation">
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Data Agent</span>
            <span className="iot-solution-validation-value">{agentName}</span>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Grounding data</span>
            <span className="iot-solution-validation-value">
              {databaseName} / {tableName}
            </span>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Available fields</span>
            <span className="iot-solution-validation-value">
              {columns.length} modeled columns
            </span>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Reasoning guidance</span>
            <span className="iot-solution-validation-value">
              Generated instructions and example KQL queries
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">How answers are grounded</h3>
        <div className="iot-solution-architecture" aria-label="Data Agent configuration">
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Modeled data</span>
            <span className="iot-solution-flow-node-detail">Telemetry with device state</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">+</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Instructions</span>
            <span className="iot-solution-flow-node-detail">Semantics and query patterns</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Generated KQL</span>
            <span className="iot-solution-flow-node-detail">Executed with user permissions</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Grounded response</span>
            <span className="iot-solution-flow-node-detail">Answers based on Fabric data</span>
          </div>
        </div>
      </section>

      {navigationError && <MessageBar intent="error" role="alert"><MessageBarBody>{navigationError}</MessageBarBody></MessageBar>}
      <div className="iot-solution-summary-actions">
        <Button appearance="primary" onClick={openDataAgent} disabled={!workspaceId || !agentId}>
          Open Data Agent in Fabric
        </Button>
        <span className="iot-solution-finish-hint">
          Select <strong>Finish and Close</strong> when you are done.
        </span>
      </div>
    </div>
  );
}
