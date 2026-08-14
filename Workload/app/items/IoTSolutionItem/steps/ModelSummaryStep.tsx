// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { DtdlCapability } from "../DtdlModelParser";
import "../IoTSolutionItem.scss";

interface ModelSummaryStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Explains the generated modeled entities, how they work together, and the
 * downstream experiences enabled by the modeled dataset.
 */
export function ModelSummaryStep({ wizardContext }: ModelSummaryStepProps) {
  const databaseName = wizardContext.modelDbName || "KQL database";
  const normalizedTable = wizardContext.propsNormalizedName || "Normalized properties table";
  const lkvView = wizardContext.propsLkvViewName || "Properties LKV view";
  const modeledTable = wizardContext.modeledDataName || "Modeled data table";
  const telemetries: DtdlCapability[] = (wizardContext.modelTelemetries || []).filter(
    (field: DtdlCapability) => field.included
  );
  const properties: DtdlCapability[] = (wizardContext.modelProperties || []).filter(
    (field: DtdlCapability) => field.included
  );

  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div className="iot-solution-success-heading">
        <div className="iot-solution-success-title-row">
          <h2 className="iot-solution-step-title">
            Your modeled device data is ready
          </h2>
        </div>
        <p className="iot-solution-step-description">
          {telemetries.length > 0
            ? "Live telemetry is now transformed into typed, business-friendly columns and enriched with the latest modeled property values for each device."
            : "The modeled table is ready to enrich telemetry events with typed property state. It remains empty for device types that send no telemetry."}{" "}
          This makes device data easier to analyze using the names, types, and
          semantics defined by the device model.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Generated modeled entities</h3>
        <div className="iot-solution-validation">
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Normalized properties table</span>
            <div className="iot-solution-validation-value">
              <strong>{normalizedTable}</strong>
              <div>
                Stores each property update as a device, property, value, and
                timestamp row. This consistent structure makes property history
                easier to query and supplies the last-known-value view.
              </div>
            </div>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Property state view</span>
            <div className="iot-solution-validation-value">
              <strong>{lkvView}</strong>
              <div>
                Maintains the latest property value reported by each device within
                the previous 30 days so current state can be joined to telemetry.
              </div>
            </div>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Modeled data table</span>
            <div className="iot-solution-validation-value">
              <strong>{modeledTable}</strong>
              <div>
                Contains device ID, event time, {telemetries.length} typed telemetry{" "}
                {telemetries.length === 1 ? "column" : "columns"}, and{" "}
                {properties.length} current property{" "}
                {properties.length === 1 ? "column" : "columns"}.
              </div>
            </div>
          </div>
          <div className="iot-solution-validation-row">
            <span className="iot-solution-validation-label">Destination KQL database</span>
            <span className="iot-solution-validation-value">{databaseName}</span>
          </div>
        </div>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">
          How the modeled entities work together
        </h3>
        <div className="iot-solution-architecture" aria-label="Modeled data architecture">
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Raw IoT tables</span>
            <span className="iot-solution-flow-node-detail">Telemetry and property JSON</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Normalized + LKV</span>
            <span className="iot-solution-flow-node-detail">Queryable property history and state</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Device model</span>
            <span className="iot-solution-flow-node-detail">Semantic names, types, units, and components</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Modeled data</span>
            <span className="iot-solution-flow-node-detail">
              {telemetries.length > 0
                ? "Typed telemetry with current state"
                : "Telemetry events with current property state"}
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">What you can do now</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Create Activator rules for context-aware monitoring and alerting, such
            as applying different thresholds based on operating mode or configured
            setpoint.
          </li>
          <li>
            Deploy a Fabric Data Agent that can reason over business-friendly
            device fields, their semantic descriptions, and current device state.
          </li>
          <li>
            Continue using the raw tables for analytics and reporting that require
            fields not included in this model.
          </li>
        </ul>
        <p className="iot-solution-step-description">
          Generated modeled entities create the foundation for context-aware
          alerts, operational monitoring, and AI-powered experiences within
          Microsoft Fabric. This modeled dataset can now be used by Activator
          rules and Fabric Data Agent experiences.
        </p>
      </section>

      <div className="iot-solution-summary-actions">
        <span className="iot-solution-finish-hint">
          Select <strong>Finish and Close</strong> when you are done.
        </span>
      </div>
    </div>
  );
}
