// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import "../IoTSolutionItem.scss";

/**
 * Introduces the dashboard workflow in terms of the operational visibility it
 * creates from raw IoT telemetry and property data.
 */
export function DashboardOverviewStep() {
  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div>
        <h2 className="iot-solution-step-title">
          Turn live device data into operational visibility
        </h2>
        <p className="iot-solution-step-description">
          Generate an extensible Real-Time Dashboard directly from the raw device
          telemetry and property updates already flowing into Fabric. The dashboard
          provides an immediate starting point for monitoring individual devices,
          identifying trends, and understanding current operational state.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">What the dashboard provides</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Time-series charts for numeric telemetry, helping you identify trends,
            changes, and unexpected device behavior.
          </li>
          <li>
            Recent-value tables for non-numeric telemetry, making status and
            categorical changes easy to review.
          </li>
          <li>
            Last-known-value cards for device properties, providing a concise view of
            the latest reported device state.
          </li>
          <li>
            Editable KQL queries and visual layouts that can be extended for your
            operational and business requirements.
          </li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">How your data becomes a dashboard</h3>
        <div className="iot-solution-overview-flow" aria-label="Dashboard generation flow">
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Raw data tables</span>
            <span className="iot-solution-flow-node-detail">
              Telemetry and property updates
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Field discovery</span>
            <span className="iot-solution-flow-node-detail">
              Select the signals that matter
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Real-Time Dashboard</span>
            <span className="iot-solution-flow-node-detail">
              Charts, tables, and value cards
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-prerequisites">
        <h3 className="iot-solution-outcome-heading">What you will need</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Raw data ingestion configured through the{" "}
            <strong>Ingest device data from your IoT Hub</strong> experience. This
            provides the telemetry and properties tables used by the dashboard.
          </li>
          <li>
            Recent device events in those raw data tables so the accelerator can
            discover available telemetry and property fields.
          </li>
          <li>
            Permission to read the Eventhouse data and create a Real-Time Dashboard
            in this Fabric workspace.
          </li>
        </ul>
      </section>
    </div>
  );
}
