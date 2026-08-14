// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { Button } from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import {
  callNavigationNavigate,
  callNavigationOpenInNewBrowserTab,
} from "../../../controller/NavigationController";
import "../IoTSolutionItem.scss";

interface ActivatorSummaryStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Summarizes the resources that feed modeled device data into Activator and
 * guides the user through defining operational rules and actions.
 */
export function ActivatorSummaryStep({
  wizardContext,
  workloadClient,
  workspaceId,
}: ActivatorSummaryStepProps) {
  const activatorId = wizardContext.actActivatorId || "";
  const activatorName = wizardContext.actActivatorName || "Activator";
  const eventstreamName = wizardContext.actEventstreamName || "Eventstream";
  const tableName = wizardContext.actTableName || "Modeled data table";
  const databaseName = wizardContext.actDatabaseName || "KQL database";

  const openActivator = () => {
    if (!activatorId) return;
    callNavigationNavigate(
      workloadClient,
      "host",
      `/groups/${workspaceId}/reflexes/${activatorId}`
    ).catch(() => {});
  };

  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div className="iot-solution-success-heading">
        <div className="iot-solution-success-title-row">
          <h2 className="iot-solution-step-title">
            Your device data is ready for Activator monitoring
          </h2>
        </div>
        <p className="iot-solution-step-description">
          New events in your modeled dataset are now being delivered to Activator.
          Open the Activator resource to define the operational conditions you want
          to monitor and the actions that should run when those conditions are met.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Example monitoring scenarios</h3>
        <ul className="iot-solution-outcome-list">
          <li>Device offline or missing-telemetry detection.</li>
          <li>Temperature, pressure, vibration, or other safety threshold exceeded.</li>
          <li>Production anomaly detection using measurements and current operating state.</li>
          <li>Asset maintenance alerts based on condition, runtime, or device configuration.</li>
          <li>Safety condition monitoring with escalation when a condition persists.</li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Configure rules and responses</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Define thresholds, state transitions, sustained conditions, or the
            absence of expected events.
          </li>
          <li>
            Use modeled telemetry and current device-property fields together to
            create context-aware operational conditions.
          </li>
          <li>
            Send email or Microsoft Teams notifications with contextual device
            information.
          </li>
          <li>
            Trigger Power Automate flows, supported Fabric activities, or webhooks
            to automate investigation and response workflows.
          </li>
        </ul>
        <p className="iot-solution-step-description">
          <a
            href="https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators"
            onClick={(event) => {
              event.preventDefault();
              void callNavigationOpenInNewBrowserTab(
                workloadClient,
                "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators"
              );
            }}
            style={{ color: "var(--colorBrandForegroundLink)" }}
          >
            Learn more about configuring rules, alerts, and actions in Fabric
            Activator
          </a>
          .
        </p>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Resources configured</h3>
        <div className="iot-solution-validation">
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Workspace identity</span>
            <span className="iot-solution-validation-value">
              Ready for secure Fabric resource access
            </span>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Source</span>
            <span className="iot-solution-validation-value">
              {databaseName} / {tableName}
            </span>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">KQL connection</span>
            <span className="iot-solution-validation-value">
              Workspace identity authentication
            </span>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Eventstream</span>
            <span className="iot-solution-validation-value">{eventstreamName}</span>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Activator</span>
            <span className="iot-solution-validation-value">{activatorName}</span>
          </div>
        </div>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Deployed monitoring flow</h3>
        <div className="iot-solution-architecture" aria-label="Activator resource architecture">
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Modeled data table</span>
            <span className="iot-solution-flow-node-detail">
              Telemetry enriched with device state
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Secure KQL connection</span>
            <span className="iot-solution-flow-node-detail">Workspace identity</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Eventstream</span>
            <span className="iot-solution-flow-node-detail">Continuous event delivery</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Activator</span>
            <span className="iot-solution-flow-node-detail">Conditions and automated responses</span>
          </div>
        </div>
      </section>

      <div className="iot-solution-summary-actions">
        <Button appearance="primary" onClick={openActivator} disabled={!activatorId}>
          Open Activator and create rules
        </Button>
        <span className="iot-solution-finish-hint">
          Select <strong>Finish and Close</strong> when you are done.
        </span>
      </div>
    </div>
  );
}
