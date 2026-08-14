// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useCallback, useEffect, useRef, useState } from "react";
import {
  Button,
  MessageBar,
  MessageBarBody,
  Spinner,
  Text,
} from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import { DtdlCapability } from "../DtdlModelParser";
import { RAW_EVENT_NORMALIZATION_KQL } from "../RawEventKql";
import "../IoTSolutionItem.scss";

const KUSTO_SCOPE = "https://kusto.kusto.windows.net/.default";
const COVERAGE_EVENT_LIMIT = 1000;

interface ModelCoverageStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
}

/**
 * Compares selected model capabilities with fields observed in the configured
 * raw telemetry and property tables.
 */
export function ModelCoverageStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
}: ModelCoverageStepProps) {
  const [observedTelemetryFields, setObservedTelemetryFields] = useState<string[]>(
    wizardContext.modelObservedTelemetryFields || []
  );
  const [observedPropertyFields, setObservedPropertyFields] = useState<string[]>(
    wizardContext.modelObservedPropertyFields || []
  );
  const [inspecting, setInspecting] = useState(false);
  const [inspectError, setInspectError] = useState("");
  const [verified, setVerified] = useState(!!wizardContext.modelCoverageVerified);
  const [emptySources, setEmptySources] = useState<string[]>(
    wizardContext.modelCoverageEmptySources || []
  );
  const inspectionStarted = useRef(false);

  const selectedTelemetries: DtdlCapability[] = (
    wizardContext.modelTelemetries || []
  ).filter((field: DtdlCapability) => field.included);
  const selectedProperties: DtdlCapability[] = (
    wizardContext.modelProperties || []
  ).filter((field: DtdlCapability) => field.included);

  const inspectCoverage = useCallback(async () => {
    const queryServiceUri = wizardContext.modelQueryServiceUri || "";
    const databaseName = wizardContext.modelDbName || "";
    const telemetryTable = wizardContext.rawTelemetryTable || "";
    const propertiesTable = wizardContext.rawPropertiesTable || "";
    if (!queryServiceUri || !databaseName || !telemetryTable || !propertiesTable) {
      setInspectError("Select the raw data source before verifying model coverage.");
      return;
    }

    setInspecting(true);
    setInspectError("");
    setEmptySources([]);
    setVerified(false);
    resetStepsFrom(stepIndex + 1);
    updateContext("modelCoverageVerified", false);
    updateContext("modelCoverageEmptySources", []);
    updateContext("modeledTablesCreated", false);

    try {
      const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
      const executeFieldQuery = async (query: string): Promise<string[]> => {
        const response = await fetch(`${queryServiceUri}/v1/rest/query`, {
          method: "POST",
          headers: {
            Authorization: "Bearer " + accessToken.token,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ db: databaseName, csl: query }),
        });
        if (!response.ok) {
          throw new Error(`Unable to inspect recent raw data (${response.status}).`);
        }
        const result = await response.json();
        const rows = result?.Tables?.[0]?.Rows || [];
        return rows
          .map((row: unknown[]) => String(row[0] || ""))
          .filter((field: string) => !!field);
      };
      const executeCountQuery = async (tableName: string): Promise<number> => {
        const response = await fetch(`${queryServiceUri}/v1/rest/query`, {
          method: "POST",
          headers: {
            Authorization: "Bearer " + accessToken.token,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ db: databaseName, csl: `${tableName} | count` }),
        });
        if (!response.ok) {
          throw new Error(`Unable to inspect recent raw data (${response.status}).`);
        }
        const result = await response.json();
        return Number(result?.Tables?.[0]?.Rows?.[0]?.[0] || 0);
      };

      const telemetryScope = String(wizardContext.telemetryScopeClause || "").trim();
      const propertiesScope = String(wizardContext.propertiesScopeClause || "").trim();
      const telemetryScopeFilter = telemetryScope ? ` | where ${telemetryScope}` : "";
      const propertiesScopeFilter = propertiesScope ? ` | where ${propertiesScope}` : "";

      const telemetryQuery = `${telemetryTable}
| top ${COVERAGE_EVENT_LIMIT} by ingestion_time() desc
${RAW_EVENT_NORMALIZATION_KQL}${telemetryScopeFilter}
| extend subject = tostring(headers.IoTSubject)
| mv-expand field = bag_keys(data) to typeof(string)
| where field !in ("EventProcessedUtcTime", "PartitionId", "EventEnqueuedUtcTime", "EventHub")
| extend modeledField = iff(subject == "", field, strcat(subject, "_", field))
| distinct modeledField
| order by modeledField asc`;

      const propertiesQuery = `${propertiesTable}
| top ${COVERAGE_EVENT_LIMIT} by ingestion_time() desc
${RAW_EVENT_NORMALIZATION_KQL}${propertiesScopeFilter}
| extend reported = data.properties.reported
| mv-expand property = bag_keys(reported) to typeof(string)
| where property !startswith "$" and property != "iothub-enqueuedtime" and property != "iothub-connection-device-id"
| extend value = reported[property]
| extend isComponent = gettype(value) == "dictionary" and tostring(value["__t"]) == "c"
| mv-expand nestedProperty = iff(isComponent, bag_keys(value), pack_array(property)) to typeof(string)
| where nestedProperty != "__t"
| extend modeledField = iff(isComponent, strcat(property, "_", nestedProperty), property)
| distinct modeledField
| order by modeledField asc`;

      const [telemetryCount, propertyCount] = await Promise.all([
        executeCountQuery(telemetryTable),
        executeCountQuery(propertiesTable),
      ]);
      const unavailableSources = [
        ...(telemetryCount === 0 ? ["telemetry"] : []),
        ...(propertyCount === 0 ? ["properties"] : []),
      ];
      const [telemetryFields, propertyFields] = await Promise.all([
        telemetryCount > 0 ? executeFieldQuery(telemetryQuery) : Promise.resolve([]),
        propertyCount > 0 ? executeFieldQuery(propertiesQuery) : Promise.resolve([]),
      ]);

      setObservedTelemetryFields(telemetryFields);
      setObservedPropertyFields(propertyFields);
      setEmptySources(unavailableSources);
      setVerified(true);
      updateContext("modelObservedTelemetryFields", telemetryFields);
      updateContext("modelObservedPropertyFields", propertyFields);
      updateContext("modelCoverageEmptySources", unavailableSources);
      updateContext("modelCoverageVerified", true);
    } catch (err: unknown) {
      setInspectError(err instanceof Error ? err.message : String(err));
    } finally {
      setInspecting(false);
    }
  }, [stepIndex, wizardContext, workloadClient, resetStepsFrom, updateContext]);

  useEffect(() => {
    if (!inspectionStarted.current && !verified) {
      inspectionStarted.current = true;
      inspectCoverage();
    }
  }, [inspectCoverage, verified]);

  const telemetryNames = new Set(selectedTelemetries.map((field) => field.name));
  const propertyNames = new Set(selectedProperties.map((field) => field.name));
  const uncoveredTelemetry = observedTelemetryFields.filter(
    (field) => !telemetryNames.has(field)
  );
  const uncoveredProperties = observedPropertyFields.filter(
    (field) => !propertyNames.has(field)
  );
  const coveredTelemetry = observedTelemetryFields.length - uncoveredTelemetry.length;
  const coveredProperties = observedPropertyFields.length - uncoveredProperties.length;

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Verify Model Coverage</h2>
      <Text className="iot-solution-step-description">
        Compare fields observed in the selected raw data tables with the
        telemetry and properties selected from your device model. This preview
        shows which fields will become typed columns in the modeled data table.
      </Text>

      {inspecting && (
        <div style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "16px" }}>
          <Spinner size="small" />
          <Text>
            Inspecting up to {COVERAGE_EVENT_LIMIT.toLocaleString()} recent events
            from each raw table...
          </Text>
        </div>
      )}

      {inspectError && (
        <>
          <MessageBar intent="error" style={{ marginTop: "12px" }}>
            <MessageBarBody>{inspectError}</MessageBarBody>
          </MessageBar>
          <Button appearance="secondary" onClick={inspectCoverage} style={{ marginTop: "12px" }}>
            Try again
          </Button>
        </>
      )}

      {verified && (
        <>
          <Text className="iot-solution-step-description" block style={{ marginTop: "16px" }}>
            Coverage is based on the most recent events from raw telemetry and
            properties table after applying the configured scoping clauses.
          </Text>
          {emptySources.length > 0 && (
            <MessageBar intent="warning" style={{ marginTop: "12px" }}>
              <MessageBarBody>
                No raw data is available in the selected{" "}
                {emptySources.length === 2
                  ? "telemetry or properties tables"
                  : `${emptySources[0]} table`}{" "}
                yet, so coverage for {emptySources.length === 2 ? "these fields" : "those fields"}{" "}
                cannot be verified. You can continue using the fields selected
                from your device model, or inspect again after devices begin
                sending data.
              </MessageBarBody>
            </MessageBar>
          )}
          <div className="iot-solution-validation" style={{ marginTop: "12px" }}>
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Telemetry coverage</Text>
              <Text className="iot-solution-validation-value">
                {emptySources.includes("telemetry")
                  ? "Not available — table is empty"
                  : `${coveredTelemetry} of ${observedTelemetryFields.length} observed fields`}
              </Text>
            </div>
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Property coverage</Text>
              <Text className="iot-solution-validation-value">
                {emptySources.includes("properties")
                  ? "Not available — table is empty"
                  : `${coveredProperties} of ${observedPropertyFields.length} observed fields`}
              </Text>
            </div>
          </div>

          {(uncoveredTelemetry.length > 0 || uncoveredProperties.length > 0) && (
            <MessageBar intent="warning" style={{ marginTop: "12px" }}>
              <MessageBarBody>
                <strong>Observed fields not selected in the model</strong>
                <ul style={{ margin: "6px 0", paddingLeft: "20px" }}>
                  {uncoveredTelemetry.length > 0 && (
                    <li>
                      <strong>Telemetry:</strong>{" "}
                      {uncoveredTelemetry.slice(0, 20).join(", ")}
                      {uncoveredTelemetry.length > 20
                        ? `, and ${uncoveredTelemetry.length - 20} more`
                        : ""}
                    </li>
                  )}
                  {uncoveredProperties.length > 0 && (
                    <li>
                      <strong>Properties:</strong>{" "}
                      {uncoveredProperties.slice(0, 20).join(", ")}
                      {uncoveredProperties.length > 20
                        ? `, and ${uncoveredProperties.length - 20} more`
                        : ""}
                    </li>
                  )}
                </ul>
                <div>
                  These fields remain available in the raw JSON tables for
                  analytics and reporting but will not appear in the modeled data
                  table.
                </div>
              </MessageBarBody>
            </MessageBar>
          )}

          {emptySources.length === 0 &&
            observedTelemetryFields.length === 0 &&
            observedPropertyFields.length === 0 && (
            <MessageBar intent="warning" style={{ marginTop: "12px" }}>
              <MessageBarBody>
                No fields were detected in the recent scoped events. You can
                continue, but confirm that devices are sending data and that the
                selected tables and scoping clauses are correct.
              </MessageBarBody>
            </MessageBar>
          )}

          <Button appearance="secondary" onClick={inspectCoverage} style={{ marginTop: "12px" }}>
            Inspect recent data again
          </Button>
        </>
      )}
    </div>
  );
}
