import React from "react";
import "../IoTSolutionItem.scss";

/**
 * Introduces the Data Agent as a natural-language reasoning layer grounded in
 * modeled real-time device data and generated domain instructions.
 */
export function DataAgentOverviewStep() {
  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div>
        <h2 className="iot-solution-step-title">
          Explore device operations using AI
        </h2>
        <p className="iot-solution-step-description">
          Create a Fabric Data Agent that translates natural-language questions
          into queries over your modeled device data, then returns answers and
          operational insights.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Why modeled data improves the agent</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            Business-friendly field names, descriptions, types, units, and
            component context help the agent understand what each measurement
            represents.
          </li>
          <li>
            Each row exposes the available telemetry or reported properties as
            typed model fields, with the latest property state included when
            available.
          </li>
          <li>
            A consistent typed schema reduces ambiguity and helps the agent produce
            more relevant queries, explanations, comparisons, and summaries.
          </li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">How the agent produces answers</h3>
        <div className="iot-solution-overview-flow" aria-label="Data Agent answer flow">
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Modeled device data</span>
            <span className="iot-solution-flow-node-detail">
              Typed device events and current state
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">+</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">AI instructions</span>
            <span className="iot-solution-flow-node-detail">
              Field meaning, query guidance, and examples
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Fabric Data Agent</span>
            <span className="iot-solution-flow-node-detail">
              Interprets questions and generates queries
            </span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-flow-node">
            <span className="iot-solution-flow-node-title">Answers and insights</span>
            <span className="iot-solution-flow-node-detail">
              Trends, comparisons, state, and anomalies
            </span>
          </div>
        </div>
      </section>

      <section className="iot-solution-prerequisites">
        <h3 className="iot-solution-outcome-heading">What you will need</h3>
        <ul className="iot-solution-outcome-list">
          <li>
            A modeled dataset created through the{" "}
            <strong>Model your device data</strong> experience. This supplies the
            typed fields and model metadata used to ground the agent.
          </li>
          <li>
            Device events in the modeled data table so the agent can answer
            questions using current and historical operational data.
          </li>
          <li>
            Permission to read the KQL database and create a Data Agent in this
            Fabric workspace.
          </li>
        </ul>
      </section>
    </div>
  );
}
