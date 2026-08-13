import React from "react";
import { WizardStepProps } from "../../../components/Wizard";
import "../IoTSolutionItem.scss";

interface SummaryRowProps {
  label: string;
  values: string[];
}

function SummaryRow({ label, values }: SummaryRowProps) {
  return (
    <div className="iot-solution-summary-row">
      <dt>{label}</dt>
      <dd>
        <ul className="iot-solution-summary-values">
          {values.map((value) => (
            <li key={value}>{value}</li>
          ))}
        </ul>
      </dd>
    </div>
  );
}

/**
 * Summarizes the completed ingestion architecture and directs users toward
 * outcome-focused experiences they can build on top of the ingested data.
 */
export function IngestionSummaryStep({ wizardContext }: WizardStepProps) {
  const hubName = wizardContext.hubName || "Azure IoT Hub";
  const databaseName = wizardContext.databaseName || "KQL Database";
  const telemetryTableName = wizardContext.telemetryTableName || "Telemetry table";
  const propertiesTableName = wizardContext.propertiesTableName || "Properties table";
  const telemetryStreamName = wizardContext.telemetryStreamName || "Telemetry Eventstream";
  const propertiesStreamName = wizardContext.propertiesStreamName || "Properties Eventstream";

  return (
    <div className="iot-solution-step iot-solution-outcome-step">
      <div className="iot-solution-success-heading">
        <div className="iot-solution-success-title-row">
          <h2 className="iot-solution-step-title">
            Your IoT data is ready for analytics and operational intelligence
          </h2>
        </div>
        <p className="iot-solution-step-description">
          Device telemetry and twin property updates from the{" "}
          <strong>{hubName}</strong> IoT Hub can now flow continuously into the{" "}
          <strong>{databaseName}</strong> KQL database. The raw data is available for
          real-time analysis and for the other accelerator experiences, enabling
          analytics and operational intelligence scenarios.
        </p>
      </div>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">What you can do now</h3>
        <ul className="iot-solution-outcome-list">
          <li>Query recent telemetry and property changes in the Eventhouse.</li>
          <li>Generate a Real-Time Dashboard to monitor device behavior and trends.</li>
          <li>
            Optionally model the raw data to add semantic context, creating the
            foundation for alerts, operational monitoring, and AI-powered experiences.
          </li>
        </ul>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Deployed architecture</h3>
        <div className="iot-solution-architecture" aria-label="Deployed ingestion architecture">
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Devices</span>
            <span className="iot-solution-flow-node-detail">Telemetry and properties</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">{hubName}</span>
            <span className="iot-solution-flow-node-detail">IoT Hub routes</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">Eventstreams</span>
            <span className="iot-solution-flow-node-detail">Telemetry and properties</span>
          </div>
          <span className="iot-solution-flow-arrow" aria-hidden="true">-&gt;</span>
          <div className="iot-solution-architecture-node">
            <span className="iot-solution-flow-node-title">{databaseName}</span>
            <span className="iot-solution-flow-node-detail">Raw data tables</span>
          </div>
        </div>
      </section>

      <section className="iot-solution-outcome-section">
        <h3 className="iot-solution-outcome-heading">Resources configured</h3>
        <dl className="iot-solution-summary">
          <SummaryRow
            label="Eventstreams"
            values={[telemetryStreamName, propertiesStreamName]}
          />
          <SummaryRow
            label="Raw data tables"
            values={[telemetryTableName, propertiesTableName]}
          />
          <SummaryRow
            label="IoT Hub routes"
            values={[
              wizardContext.telemetryRouteName || "Telemetry route",
              wizardContext.propertiesRouteName || "Properties route",
            ]}
          />
          <SummaryRow label="Destination KQL database" values={[databaseName]} />
        </dl>
      </section>

      <p className="iot-solution-finish-hint">
        Select <strong>Finish and Close</strong> to return to the accelerator home page.
      </p>
    </div>
  );
}
