// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import { ArrowRight20Regular } from "@fluentui/react-icons";
import type { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { callNavigationOpenInNewBrowserTab } from "../../../controller/NavigationController";
import "../IoTSolutionItem.scss";

export function ModelOverviewStep({ workloadClient }: { workloadClient: WorkloadClientAPI }) {
  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div>
        <h2 className="iot-solution-step-title">
          Turn raw IoT data into modeled data
        </h2>
        <p className="iot-solution-step-description">
          Use a device model to turn the telemetry and property fields you choose into typed, business-friendly columns. Each telemetry event also gets the device's latest state from its twin properties.
        </p>
        <p className="iot-solution-step-description" style={{ marginTop: "8px" }}>
          Modeling is optional. Your raw IoT data stays available for analytics and reporting.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">What modeling unlocks</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Compare measurements with device state, such as temperature by operating mode or vibration by firmware version.
          </li>
          <li>
            Query typed columns instead of parsing raw JSON in every query.
          </li>
          <li>
            Set more precise alerts that account for device configuration and status.
          </li>
          <li>
            Build dashboards, Activator rules, reports, and AI experiences on the same modeled data.
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
          <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Device model</span>
            <span className="iot-solution-flow-node-detail">
              Names, types, units, and components
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Current device state</span>
            <span className="iot-solution-flow-node-detail">
              Latest property values
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true"><ArrowRight20Regular /></span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Modeled data</span>
            <span className="iot-solution-flow-node-detail">
              Typed telemetry enriched with properties
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-prerequisites">
        <h3 className="iot-solution-outcome-heading">What you&rsquo;ll need</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Raw telemetry and properties tables from the{" "}
            <strong>Ingest device data from your IoT Hub</strong> experience.
          </li>
          <li>
            Optionally, an Azure Device Registry (ADR) namespace{" "}
            <a
              href="https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal"
              onClick={(event) => {
                event.preventDefault();
                void callNavigationOpenInNewBrowserTab(workloadClient,
                  "https://learn.microsoft.com/en-us/azure/iot/device-registry/get-started-azure-device-registry?tabs=azure-portal");
              }}
              target="_blank"
              rel="noopener noreferrer"
              style={{ color: "var(--colorBrandForegroundLink)" }}
            >
              linked to your IoT Hub
            </a>. Linking an ADR namespace to your IoT Hub allows devices to reference your model.
          </li>
          <li>
            A device model: a Web of Things (WoT) Thing Model (recommended), a Digital Twins Definition Language (DTDL) model, or an IoT Central device template export.
          </li>
          <li>
            At least one device event in your raw tables (more events give better suggestions).
          </li>
          <li>
            Permission to read raw data and to create tables, update policies, and a materialized view in the KQL database.
          </li>
        </ul>
      </section>
    </div>
  );
}
