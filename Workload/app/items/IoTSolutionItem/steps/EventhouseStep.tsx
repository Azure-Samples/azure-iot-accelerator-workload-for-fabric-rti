import React, { useState, useEffect, useCallback } from "react";
import {
  Dropdown,
  Label,
  Option,
  Spinner,
  Text,
  Button,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  ArrowClockwise24Regular,
  CheckmarkCircle24Filled,
  Info24Regular,
} from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { FabricPlatformAPIClient } from "../../../clients/FabricPlatformAPIClient";
import { Item } from "../../../clients/FabricPlatformTypes";
import { getEventhouseItem } from "../eventhouseClient";
import "../IoTSolutionItem.scss";

const FABRIC_ITEM_READ_SCOPE = "https://api.fabric.microsoft.com/Item.Read.All";

interface EventhouseStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Eventhouse and database selection.
 *
 * User picks an existing Eventhouse and KQL Database from the current workspace.
 * These will be used to store telemetry data routed from the IoT Hub.
 */
export function EventhouseStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
}: EventhouseStepProps) {
  // Eventhouse state
  const [eventhouses, setEventhouses] = useState<Item[]>([]);
  const [selectedEventhouseId, setSelectedEventhouseId] = useState<string>(
    wizardContext.eventhouseId || ""
  );
  const [selectedEventhouseName, setSelectedEventhouseName] = useState<string>(
    wizardContext.eventhouseName || ""
  );
  const [loadingEventhouses, setLoadingEventhouses] = useState(false);
  const [eventhouseError, setEventhouseError] = useState("");

  // Database state
  const [databases, setDatabases] = useState<Item[]>([]);
  const [selectedDatabaseId, setSelectedDatabaseId] = useState<string>(
    wizardContext.databaseId || ""
  );
  const [selectedDatabaseName, setSelectedDatabaseName] = useState<string>(
    wizardContext.databaseName || ""
  );
  const [loadingDatabases, setLoadingDatabases] = useState(false);
  const [databaseError, setDatabaseError] = useState("");

  // Eventhouse metadata (query/ingestion URIs)
  const [, setQueryServiceUri] = useState<string>(
    wizardContext.queryServiceUri || ""
  );
  const [, setIngestionServiceUri] = useState<string>(
    wizardContext.ingestionServiceUri || ""
  );

  // Permission error state
  const [permissionError, setPermissionError] = useState<string>("");

  /**
   * Pre-flight check: acquire a Fabric API token with consent fallback.
   * Detects missing permissions vs missing admin consent vs user cancellation.
   */
  const ensureFabricApiToken = useCallback(async (): Promise<boolean> => {
    setPermissionError("");
    try {
      // Try silent token acquisition first
      await workloadClient.auth.acquireFrontendAccessToken({
        scopes: [FABRIC_ITEM_READ_SCOPE],
      });
      return true;
    } catch {
      // Silent failed — try interactive consent
      try {
        await workloadClient.auth.acquireAccessToken({
          additionalScopesToConsent: [FABRIC_ITEM_READ_SCOPE],
        });
        // Retry silent after consent
        await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [FABRIC_ITEM_READ_SCOPE],
        });
        return true;
      } catch (interactiveErr: unknown) {
        const errCode = (interactiveErr as { error?: number })?.error;
        const errStr = String(interactiveErr);

        if (errCode === 2) {
          setPermissionError(
            "Required Fabric permissions have not been granted. Contact your Fabric " +
            "administrator or review application permissions before continuing."
          );
        } else if (errCode === 1) {
          setPermissionError(
            "Access to Fabric was not granted. Please complete the sign-in and consent " +
            "process to continue."
          );
        } else if (
          errStr.includes("admin") ||
          errStr.includes("AADSTS65001") ||
          errStr.includes("AADSTS90094")
        ) {
          setPermissionError(
            "Additional administrator approval is required before the accelerator can " +
            "access Fabric resources."
          );
        } else {
          setPermissionError(
            "Unable to authenticate with Microsoft Fabric. Verify application permissions " +
            "and try again."
          );
        }
        return false;
      }
    }
  }, [workloadClient]);

  // Load eventhouses on mount and when workspace changes
  const loadEventhouses = useCallback(async () => {
    if (!workspaceId) return;
    setLoadingEventhouses(true);
    setEventhouseError("");

    // Pre-flight: ensure we have Fabric API permissions
    const hasToken = await ensureFabricApiToken();
    if (!hasToken) {
      setLoadingEventhouses(false);
      return;
    }

    try {
      const client = FabricPlatformAPIClient.create(workloadClient);
      const result = await client.items.listItems(workspaceId, {
        type: "Eventhouse",
      });
      setEventhouses(result.value || []);
      if ((result.value || []).length === 0) {
        setEventhouseError(
          "No Eventhouses were found in the selected workspace. Create an Eventhouse " +
          "and KQL Database before using this accelerator."
        );
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      if (msg.includes("401") || msg.includes("403") || msg.includes("Unauthorized") || msg.includes("Forbidden")) {
        setPermissionError(
          "You do not have permission to view Eventhouses in this workspace."
        );
      } else {
        setEventhouseError(
          `Unable to load available Eventhouses. Verify your Fabric permissions and try again. Error: ${msg}`
        );
      }
      setEventhouses([]);
    } finally {
      setLoadingEventhouses(false);
    }
  }, [workloadClient, workspaceId, ensureFabricApiToken]);

  useEffect(() => {
    loadEventhouses();
  }, [loadEventhouses]);

  // Load databases when an eventhouse is selected
  const loadDatabases = useCallback(
    async (eventhouseId: string) => {
      if (!eventhouseId || !workspaceId) return;
      setLoadingDatabases(true);
      setDatabaseError("");
      setDatabases([]);
      setSelectedDatabaseId("");
      setSelectedDatabaseName("");
      updateContext("databaseId", "");
      updateContext("databaseName", "");
      updateContext("eventhouseValidated", false);

      try {
        // Get eventhouse metadata (includes database IDs and service URIs)
        const ehMeta = await getEventhouseItem(
          workloadClient,
          workspaceId,
          eventhouseId
        );

        if (!ehMeta) {
          throw new Error(
            "Unable to load information for the selected Eventhouse. Verify that you " +
            "have access to this Fabric workspace."
          );
        }

        setQueryServiceUri(ehMeta.properties.queryServiceUri);
        setIngestionServiceUri(ehMeta.properties.ingestionServiceUri);
        updateContext("queryServiceUri", ehMeta.properties.queryServiceUri);
        updateContext(
          "ingestionServiceUri",
          ehMeta.properties.ingestionServiceUri
        );

        const dbIds = ehMeta.properties.databasesItemIds || [];
        if (dbIds.length === 0) {
          setDatabaseError(
            "No KQL Database was found in this Eventhouse. Create or select a KQL " +
            "Database before continuing."
          );
          return;
        }

        // List all KQL databases in the workspace and filter to this eventhouse
        const client = FabricPlatformAPIClient.create(workloadClient);
        const allDbs = await client.items.listItems(workspaceId, {
          type: "KQLDatabase",
        });
        const eventhouseDbs = (allDbs.value || []).filter((db) =>
          dbIds.includes(db.id)
        );

        setDatabases(eventhouseDbs);
        if (eventhouseDbs.length === 0) {
          setDatabaseError(
            "No KQL Database was found in this Eventhouse. Create or select a KQL " +
            "Database before continuing."
          );
        }
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        setDatabaseError(
          msg.startsWith("Unable to load information")
            ? msg
            : `Unable to load KQL Databases for the selected Eventhouse. Error: ${msg}`
        );
      } finally {
        setLoadingDatabases(false);
      }
    },
    [workloadClient, workspaceId, updateContext]
  );

  // Clear downstream context when the Eventhouse or database changes.
  const clearDownstreamFromEventhouse = () => {
    resetStepsFrom(stepIndex + 1);
    const downstreamKeys = [
      "telemetryTableName", "propertiesTableName", "tablesCreated",
      "telemetryStreamName", "telemetryStreamId", "propertiesStreamName", "propertiesStreamId",
      "telemetryEndpointNamespace", "telemetryEndpointEventHubName",
      "telemetryEndpointSourceId",
      "propertiesEndpointNamespace", "propertiesEndpointEventHubName",
      "propertiesEndpointSourceId",
      "eventstreamsCreated",
      "miHasWorkspaceAccess", "miWorkspaceRole", "routingConfigured",
      "telemetryRouteName", "propertiesRouteName",
      "telemetryEndpointName", "propertiesEndpointName",
    ];
    for (const key of downstreamKeys) {
      updateContext(key, key.endsWith("Created") || key.endsWith("Configured") || key.endsWith("Access") ? false : "");
    }
  };

  // When eventhouse selection changes
  const handleEventhouseChange = (eventhouseId: string) => {
    const eh = eventhouses.find((e) => e.id === eventhouseId);
    setSelectedEventhouseId(eventhouseId);
    setSelectedEventhouseName(eh?.displayName || "");
    updateContext("eventhouseId", eventhouseId);
    updateContext("eventhouseName", eh?.displayName || "");
    updateContext("eventhouseValidated", false);
    clearDownstreamFromEventhouse();
    loadDatabases(eventhouseId);
  };

  // When database selection changes
  const handleDatabaseChange = (databaseId: string) => {
    const db = databases.find((d) => d.id === databaseId);
    setSelectedDatabaseId(databaseId);
    setSelectedDatabaseName(db?.displayName || "");
    updateContext("databaseId", databaseId);
    updateContext("databaseName", db?.displayName || "");
    updateContext("eventhouseValidated", !!(databaseId && selectedEventhouseId));
    clearDownstreamFromEventhouse();
  };

  // Reload databases for current eventhouse
  const refreshDatabases = () => {
    if (selectedEventhouseId) {
      loadDatabases(selectedEventhouseId);
    }
  };

  const isComplete = !!(selectedEventhouseId && selectedDatabaseId);

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">
        Select Eventhouse &amp; Database
      </h2>
      <Text className="iot-solution-step-description">
        Choose an existing Eventhouse and KQL Database in this workspace. The Eventhouse will store telemetry,
        properties, and operational data used by dashboards, Activator, and Data Agent experiences.
      </Text>

      {/* Info banner */}
      <div className="iot-solution-info">
        <Info24Regular primaryFill="var(--colorBrandForeground1)" />
        <Text className="iot-solution-info-text">
          Don&apos;t have an Eventhouse and KQL Database yet? Create them from the workspace
          &quot;+ New&quot; menu, then return to this page and select Refresh.{" "}
          <a
            href="https://learn.microsoft.com/en-us/fabric/real-time-intelligence/create-eventhouse"
            target="_blank"
            rel="noopener noreferrer"
            style={{ color: "var(--colorBrandForegroundLink)" }}
          >
            Learn more about Eventhouses
          </a>
        </Text>
      </div>

      {/* Permission error banner */}
      {permissionError && (
        <MessageBar intent="error" className="iot-solution-permission-error">
          <MessageBarBody>
            <strong>Missing API Permission: </strong>
            {permissionError}
          </MessageBarBody>
        </MessageBar>
      )}

      <div className="iot-solution-form">
        {/* Eventhouse Picker */}
        <div className="iot-solution-field">
          <div className="iot-solution-field-header">
            <Label className="iot-solution-field-label" required htmlFor="eventhouse-select">
              Eventhouse
            </Label>
            <Button
              appearance="subtle"
              icon={<ArrowClockwise24Regular />}
              size="small"
              onClick={loadEventhouses}
              disabled={loadingEventhouses}
              title="Refresh Eventhouses"
            />
          </div>
          {loadingEventhouses ? (
            <div className="iot-solution-inline-loading">
              <Spinner size="tiny" />
              <Text className="iot-solution-loading-text">
                Loading Eventhouses...
              </Text>
            </div>
          ) : (
            <Dropdown
              id="eventhouse-select"
              placeholder="Select an Eventhouse"
              value={selectedEventhouseName}
              selectedOptions={selectedEventhouseId ? [selectedEventhouseId] : []}
              onOptionSelect={(_, data) => {
                if (data.optionValue) {
                  handleEventhouseChange(data.optionValue);
                }
              }}
              disabled={eventhouses.length === 0}
            >
              {eventhouses.map((eh) => (
                <Option key={eh.id} value={eh.id}>
                  {eh.displayName}
                </Option>
              ))}
            </Dropdown>
          )}
          {eventhouseError && !loadingEventhouses && (
            <Text className="iot-solution-field-error">{eventhouseError}</Text>
          )}
        </div>

        {/* Database Picker */}
        {selectedEventhouseId && (
          <div className="iot-solution-field">
            <div className="iot-solution-field-header">
              <Label className="iot-solution-field-label" required htmlFor="database-select">
                KQL Database
              </Label>
              <Button
                appearance="subtle"
                icon={<ArrowClockwise24Regular />}
                size="small"
                onClick={refreshDatabases}
                disabled={loadingDatabases}
                title="Refresh Databases"
              />
            </div>
            {loadingDatabases ? (
              <div className="iot-solution-inline-loading">
                <Spinner size="tiny" />
                <Text className="iot-solution-loading-text">
                  Loading databases...
                </Text>
              </div>
            ) : (
              <Dropdown
                id="database-select"
                placeholder="Select a KQL Database"
                value={selectedDatabaseName}
                selectedOptions={selectedDatabaseId ? [selectedDatabaseId] : []}
                onOptionSelect={(_, data) => {
                  if (data.optionValue) {
                    handleDatabaseChange(data.optionValue);
                  }
                }}
                disabled={databases.length === 0}
              >
                {databases.map((db) => (
                  <Option key={db.id} value={db.id}>
                    {db.displayName}
                  </Option>
                ))}
              </Dropdown>
            )}
            {databaseError && !loadingDatabases && (
              <Text className="iot-solution-field-error">{databaseError}</Text>
            )}
          </div>
        )}
      </div>

      {/* Selection Summary */}
      {isComplete && (
        <>
          <div className="iot-solution-success">
            <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
            <Text className="iot-solution-success-text">
              Eventhouse and database selected
            </Text>
          </div>
          <div className="iot-solution-validation">
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Eventhouse</Text>
              <Text className="iot-solution-validation-value">
                {selectedEventhouseName}
              </Text>
            </div>
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Database</Text>
              <Text className="iot-solution-validation-value">
                {selectedDatabaseName}
              </Text>
            </div>
          </div>
        </>
      )}
    </div>
  );
}
