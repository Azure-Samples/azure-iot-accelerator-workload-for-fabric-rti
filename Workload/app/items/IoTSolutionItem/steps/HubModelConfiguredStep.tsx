// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { Button, Link, MessageBar, MessageBarBody } from "@fluentui/react-components";
import { ArrowRight20Regular } from "@fluentui/react-icons";
import type { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { callNavigationOpenInNewBrowserTab } from "../../../controller/NavigationController";
import { creationFingerprint, creationOutputsConfirmed, type CreationResult } from "../ModeledDataCreation";
import { currentSourceAnalysis, modelSourceKey, type SourceReadiness } from "../ModelSourceReadiness";
import { modelCoverageReady } from "../StreamCoverage";
import { hubInput } from "./HubModelWorkflowSteps";
import "../IoTSolutionItem.scss";

interface ConfiguredProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  sourceState?: SourceReadiness; sourceRevision?: number; workspaceId?: string;
}
function verifiedResult(context: Record<string, any>, source: SourceReadiness | undefined, revision: number, workspaceId: string): CreationResult | undefined {
  const result = context.hubCreation as CreationResult | undefined;
  if (!result || result.phase !== "completed" || !context.hubOutputNames || !context.hubModel) return undefined;
  const input = hubInput(context, source);
  if (!source?.ready || source.key !== modelSourceKey(workspaceId, context) ||
      !currentSourceAnalysis(workspaceId, context) || !modelCoverageReady(input, revision, context.hubStreamChecks)) return undefined;
  const request = { input, coverage: context.hubCoverage, allowWaiting: true, names: context.hubOutputNames };
  return result.fingerprint === creationFingerprint(request) ? result : undefined;
}
export function HubModelConfiguredStep({ wizardContext, workloadClient, onGoToStep, sourceState, sourceRevision = 0, workspaceId = "" }: ConfiguredProps) {
  const input = hubInput(wizardContext, sourceState);
  const result = verifiedResult(wizardContext, sourceState, sourceRevision, workspaceId);
  if (!result || !creationOutputsConfirmed(result, { input, names: wizardContext.hubOutputNames })) {
    return <div className="iot-solution-step iot-solution-outcome-step">
      <h2 className="iot-solution-step-title">Creation needs review</h2>
      <MessageBar intent="warning"><MessageBarBody>
        {!result ? "The current configuration does not have a matching verified creation result. Return to Create modeled data to check it again."
          : "Not all output definitions are confirmed in the current creation result. Return to Create modeled data before finishing."}
      </MessageBarBody></MessageBar>
      <Button onClick={() => onGoToStep("review-plan")}>Return to Create modeled data</Button>
    </div>;
  }
  const telemetryCount = input.fields.filter(field => field.included && field.role === "telemetry").length;
  const twinCount = input.fields.filter(field => field.included && field.role === "twin").length;
  const hasPropertyOutputs = !!result.plannedOutputs?.some(output => output.name === result.names.twinState && output.kind === "view");
  return <div className="iot-solution-step iot-solution-outcome-step iot-solution-model-ready">
    <div className="iot-solution-success-heading"><div className="iot-solution-success-title-row">
      <h2 className="iot-solution-step-title">Your modeled data is ready</h2>
    </div>
    <p className="iot-solution-step-description">
      {result.awaitingData ? telemetryCount === 0
          ? "Your modeled data is ready to add each device's latest property values to new telemetry events when they arrive. It stays empty for device types that send no telemetry."
          : hasPropertyOutputs && twinCount > 0
            ? "Your modeled data is ready to turn new telemetry into typed columns, along with each device's latest property values, when it arrives."
            : "Your modeled data is ready to turn new telemetry into typed columns when it arrives."
        : twinCount === 0 ? "New telemetry now flows into typed columns."
        : telemetryCount > 0
          ? "New telemetry now flows into typed columns, along with each device's latest property values."
          : "Your modeled data is ready to add each device's latest property values to telemetry events. It stays empty for device types that send no telemetry."}{" "}
      Query it by the field names and types from your device model.
    </p>
    </div>
    {result.awaitingData && <MessageBar intent="info" role="status"><MessageBarBody>
      This setup will use matching telemetry when it arrives.
    </MessageBarBody></MessageBar>}
    <section className="iot-solution-outcome-section">
      <h3 className="iot-solution-outcome-heading">What you can do next</h3>
      <ul className="iot-solution-outcome-list">
        <li>
          {hasPropertyOutputs && twinCount > 0
            ? "Set alerts with Activator that account for device state, like different temperature limits for each operating mode"
            : "Set alerts with Activator using the typed fields from your device model"}
          <div><Link href="https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators"
            onClick={(event: React.MouseEvent<HTMLAnchorElement>) => {
              event.preventDefault();
              void callNavigationOpenInNewBrowserTab(workloadClient,
                "https://learn.microsoft.com/en-us/fabric/real-time-intelligence/data-activator/activator-create-activators");
            }}
            target="_blank" rel="noopener noreferrer">
            Learn about Activator
          </Link></div>
        </li>
        <li>
          Build a Fabric data agent that answers questions about your devices in plain language
          <div><Link href="https://learn.microsoft.com/en-us/fabric/data-science/concept-data-agent"
            onClick={(event: React.MouseEvent<HTMLAnchorElement>) => {
              event.preventDefault();
              void callNavigationOpenInNewBrowserTab(workloadClient,
                "https://learn.microsoft.com/en-us/fabric/data-science/concept-data-agent");
            }}
            target="_blank" rel="noopener noreferrer">
            Learn about Fabric data agents
          </Link></div>
        </li>
        <li>Keep using your raw tables for fields that aren't in this model</li>
      </ul>
    </section>
    <section className="iot-solution-outcome-section">
      <h3 id="ready-resources-heading" className="iot-solution-outcome-heading">What was created</h3>
      <p className="iot-solution-step-description">
        These resources were created in the {result.databaseName} KQL database
      </p>
      <div className="iot-solution-validation iot-solution-ready-resources">
        <div className="iot-solution-validation-row"><span className="iot-solution-validation-label">Modeled data table</span>
          <div className="iot-solution-validation-value"><strong>{result.names.modeled}</strong>
            <div>Your typed data: {telemetryCount} telemetry {telemetryCount === 1 ? "column" : "columns"} and{" "}
              {twinCount} property {twinCount === 1 ? "column" : "columns"}, plus device ID and event time, for each event</div>
          </div></div>
        {hasPropertyOutputs && <>
          <div className="iot-solution-validation-row"><span className="iot-solution-validation-label">Normalized properties table</span>
            <div className="iot-solution-validation-value"><strong>{result.names.twinChanges}</strong>
              <div>Stores each property update as its own row: device, property, value, and time</div>
            </div></div>
          <div className="iot-solution-validation-row"><span className="iot-solution-validation-label">Latest property values (materialized view)</span>
            <div className="iot-solution-validation-value"><strong>{result.names.twinState}</strong>
              <div>Keeps each device's latest reported value for every property from the last 30 days</div>
            </div></div>
        </>}
      </div>
    </section>
    <section className="iot-solution-outcome-section">
      <h3 id="ready-architecture-heading" className="iot-solution-outcome-heading">How your modeled data is built</h3>
      <div className="iot-solution-architecture" aria-label="Modeled data architecture">
        <div className="iot-solution-architecture-node"><span className="iot-solution-flow-node-title">Raw IoT tables</span>
          <span className="iot-solution-flow-node-detail">Telemetry and property JSON</span></div>
        <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
        <div className="iot-solution-architecture-node"><span className="iot-solution-flow-node-title">Device model</span>
          <span className="iot-solution-flow-node-detail">Field names, types, and units</span></div>
        {hasPropertyOutputs && twinCount > 0 && <>
          <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
          <div className="iot-solution-architecture-node"><span className="iot-solution-flow-node-title">Latest property values</span>
            <span className="iot-solution-flow-node-detail">Each device's latest state</span></div>
        </>}
        <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
        <div className="iot-solution-architecture-node"><span className="iot-solution-flow-node-title">Modeled data</span>
          <span className="iot-solution-flow-node-detail">{twinCount === 0 ? "Typed telemetry" : telemetryCount > 0
            ? "Typed telemetry with device state" : "Telemetry events with device state"}</span></div>
      </div>
    </section>
    <div className="iot-solution-summary-actions"><span className="iot-solution-finish-hint">
      Select <strong>Finish and Close</strong> when you are done.
    </span></div>
  </div>;
}
