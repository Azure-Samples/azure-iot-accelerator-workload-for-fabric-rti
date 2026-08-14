// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React from "react";
import "../IoTSolutionItem.scss";

/**
 * Introduces Activator as the condition-evaluation and action layer for
 * state-enriched device data.
 */
export function ActivatorOverviewStep() {
  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div>
        <h2 className="iot-solution-step-title">
          Generate alerts and automated actions based on live IoT data
        </h2>
        <p className="iot-solution-step-description">
          Fabric Activator continuously evaluates your incoming modeled device
          data and business conditions. When a condition is met, it can trigger
          alerts, notifications, workflows, and other automated actions.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">
          Why combine modeled device data and Activator
        </h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Modeled datasets provide business-friendly entities, measurements,
            types, and semantic descriptions, making rules and thresholds easier
            to define and maintain.
          </li>
          <li>
            Modeled events expose the selected telemetry or reported properties as
            typed fields, with the latest property state included when available.
            This supports conditions based on measurements, configuration, or
            operational state.
          </li>
          <li>
            Rules can account for context, such as applying different temperature
            limits by operating mode, detecting vibration anomalies by asset
            configuration, or escalating alerts based on maintenance status.
          </li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">How Activator monitoring works</h3>
        <div className="iot-solution-overview-flow" aria-label="Activator monitoring flow">
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Modeled device data</span>
            <span className="iot-solution-flow-node-detail">
              Typed device events and current state
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Eventstream</span>
            <span className="iot-solution-flow-node-detail">
              Continuously delivers new events
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Activator conditions</span>
            <span className="iot-solution-flow-node-detail">
              Thresholds, state changes, and event absence
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Actions</span>
            <span className="iot-solution-flow-node-detail">
              Email alerts, Microsoft Teams notifications, automated workflows,
              and much more
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-prerequisites">
        <h3 className="iot-solution-outcome-heading">What you will need</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Your modeled dataset containing enriched device data created through
            the <strong>Model your device data</strong> experience. This provides
            typed model fields and semantic context for operational conditions.
          </li>
          <li>
            Live device events flowing into the modeled data table so Activator
            can continuously evaluate new conditions.
          </li>
          <li>
            Permission to create Fabric connections, Eventstreams, and Activator
            items in this workspace.
          </li>
          <li>
            Workspace Identity enabled in this workspace or Admin permission to
            enable it during the setup. Needed for secure communication between
            Fabric resources.
          </li>
        </ul>
      </section>
    </div>
  );
}
