// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import "../IoTSolutionItem.scss";

/**
 * Introduces modeling as an optional semantic enrichment layer that combines
 * live telemetry with the device's latest reported state.
 */
export function ModelOverviewStep() {
  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div>
        <h2 className="iot-solution-step-title">
          Enrich your live device data with semantic context
        </h2>
        <p className="iot-solution-step-description">
          Enrich your IoT device data with semantic context to enable more
          powerful operational monitoring and richer AI-powered experiences. Use
          a device model to transform selected telemetry and property fields into
          typed, business-friendly columns and enrich telemetry events with the
          latest device state from twin properties.
        </p>
        <p className="iot-solution-step-description" style={{ marginTop: "8px" }}>
          Modeling is optional: your raw IoT data remains available for analytics
          and reporting without it.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">What modeling unlocks</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Correlate live measurements with current device state, such as
            temperature by operating mode, vibration by firmware version, or
            energy consumption by configured setpoint.
          </li>
          <li>
            Create strongly typed, business-friendly columns instead of repeatedly
            parsing raw JSON in every query.
          </li>
          <li>
            Enables more precise alerts and monitoring by evaluating telemetry in
            the context of device configuration, status, and other reported
            properties.
          </li>
          <li>
            Create a semantic foundation for dashboards, Activator rules,
            operational monitoring, reporting, and AI-powered experiences.
          </li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">How the modeled dataset is created</h3>
        <div className="iot-solution-overview-flow" aria-label="Modeled data creation flow">
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Raw IoT tables</span>
            <span className="iot-solution-flow-node-detail">
              Telemetry and property JSON
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Device model</span>
            <span className="iot-solution-flow-node-detail">
              Names, types, units, and components
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Current device state</span>
            <span className="iot-solution-flow-node-detail">
              Normalized property changes and last-known values
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Modeled data</span>
            <span className="iot-solution-flow-node-detail">
              Typed telemetry enriched with properties
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-prerequisites">
        <h3 className="iot-solution-outcome-heading">What you will need</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Raw data ingestion configured through the{" "}
            <strong>Ingest device data from your IoT Hub</strong> experience. The
            raw telemetry and property tables are the source for the modeled data.
          </li>
          <li>
            A DTDL device model or IoT Central device template export describing
            the telemetry and properties you want as typed columns.
          </li>
          <li>
            Recent device events so the accelerator can compare observed fields
            with the uploaded model.
          </li>
          <li>
            Permission to read the raw data and create tables, update policies,
            and a materialized view in the KQL database.
          </li>
        </ul>
      </section>
    </div>
  );
}
