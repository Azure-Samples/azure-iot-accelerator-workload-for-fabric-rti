// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import "../IoTSolutionItem.scss";

/**
 * Introduces the ingestion workflow in terms of customer outcomes and prerequisites
 * before the user starts configuring individual resources.
 */
export function IngestionOverviewStep() {
  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div>
        <h2 className="iot-solution-step-title">Connect your IoT data to Microsoft Fabric</h2>
        <p className="iot-solution-step-description">
          Configure a continuous data path from Azure IoT Hub into Microsoft Fabric
          Real-Time Intelligence. When setup is complete, live device telemetry and
          property updates will be available in an Eventhouse for monitoring,
          analytics, alerting, and AI-powered exploration.
        </p>
        <p className="iot-solution-step-description iot-solution-outcome-paragraph">
          Azure IoT Hub remains the device connectivity and message routing layer.
          Microsoft Fabric provides the analytics and operational intelligence
          experiences built on your device data.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">What this setup enables</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Continuously ingest device telemetry and twin property updates to enable
            real-time analytics and operational monitoring.
          </li>
          <li>
            Retain raw device data in query-ready Eventhouse tables for historical
            analysis, troubleshooting, and future use cases.
          </li>
          <li>
            Build Real-Time Dashboards that turn live operational data into visible
            trends and actionable insights.
          </li>
          <li>
            Establish a reusable data foundation for optional modeling, alerts, and
            AI-powered experiences.
          </li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">What the wizard configures</h3>
        <div className="iot-solution-overview-flow" aria-label="Ingestion setup flow">
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Azure IoT Hub</span>
            <span className="iot-solution-flow-node-detail">Your device data source</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Eventstreams</span>
            <span className="iot-solution-flow-node-detail">Continuous ingestion</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Eventhouse</span>
            <span className="iot-solution-flow-node-detail">Raw telemetry and properties</span>
          </div>
        </div>
      </section>

      <section className="iot-solution-prerequisites">
        <h3 className="iot-solution-outcome-heading">What you will need</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            An existing Azure IoT Hub with a system-assigned or user-assigned managed
            identity enabled. An identity is needed for IoT Hub to securely route data
            to your Fabric solution.
          </li>
          <li>
            An existing Eventhouse and KQL Database in this Fabric workspace to store
            the ingested device data used for analytics and operational intelligence.
          </li>
          <li>
            Permission to read and configure the IoT Hub, create Fabric items, and
            manage workspace access for the IoT Hub identity. These permissions allow
            the accelerator to securely configure the data path on your behalf.
          </li>
        </ul>
      </section>
    </div>
  );
}
