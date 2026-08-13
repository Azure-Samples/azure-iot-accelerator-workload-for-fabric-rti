import React, { useState, useCallback } from "react";
import {
  Input,
  Label,
  Spinner,
  Text,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { CheckmarkCircle24Filled } from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import {
  DashboardField,
  generateDashboardDefinition,
  toBase64,
} from "../DashboardDefinitionGenerator";
import {
  getFabricCreationErrorMessage,
  parseFabricErrorPayload,
} from "../FabricCreationError";
import "../IoTSolutionItem.scss";

const FABRIC_WRITE_SCOPE = "https://api.fabric.microsoft.com/Item.ReadWrite.All";
const FABRIC_API_BASE = "https://api.fabric.microsoft.com/v1";

interface DashboardGenerateStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Generates the Real-Time Dashboard.
 *
 * Reads the data source and selected fields chosen in step 1 (from the shared
 * wizard context) and creates a Fabric Real-Time Dashboard via the REST API.
 * Time-series telemetry is presented as line charts, non-time-series telemetry
 * is presented as tables, and properties are presented as card tiles.
 */
export function DashboardGenerateStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
}: DashboardGenerateStepProps) {
  const queryServiceUri: string = wizardContext.dashQueryServiceUri || "";
  const selectedDatabaseId: string = wizardContext.dashDbId || "";
  const telemetryTableName: string = wizardContext.dashTelemetryTable || "";
  const propertiesTableName: string = wizardContext.dashPropertiesTable || "";
  const telemetryFields: DashboardField[] = wizardContext.dashTelemetryFields || [];
  const propertyFields: DashboardField[] = wizardContext.dashPropertyFields || [];

  const [dashboardName, setDashboardName] = useState<string>(
    wizardContext.dashDashboardName || "IoT Device Dashboard"
  );
  const [isGenerating, setIsGenerating] = useState(false);
  const [generateError, setGenerateError] = useState("");
  const [permissionError, setPermissionError] = useState("");
  const [dashboardCreated, setDashboardCreated] = useState<boolean>(
    wizardContext.dashDashboardCreated || false
  );
  const includedTelemetry = telemetryFields.filter((f) => f.included);
  const includedProperties = propertyFields.filter((f) => f.included);
  const timeSeriesCount = includedTelemetry.filter((f) => f.isNumeric).length;
  const nonTimeSeriesCount = includedTelemetry.filter((f) => !f.isNumeric).length;

  // Generate the Real-Time Dashboard via Fabric REST API
  const generateDashboard = useCallback(async () => {
    if (!queryServiceUri || !selectedDatabaseId || !telemetryTableName || !propertiesTableName) return;
    setIsGenerating(true);
    updateContext("dashDashboardCreating", true);
    setGenerateError("");
    setPermissionError("");

    try {
      // Acquire Fabric write token
      let token: string;
      try {
        const silent = await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [FABRIC_WRITE_SCOPE],
        });
        token = silent.token;
      } catch {
        try {
          await workloadClient.auth.acquireAccessToken({
            additionalScopesToConsent: [FABRIC_WRITE_SCOPE],
          });
          const retry = await workloadClient.auth.acquireFrontendAccessToken({
            scopes: [FABRIC_WRITE_SCOPE],
          });
          token = retry.token;
        } catch (interactiveErr: unknown) {
          const errCode = (interactiveErr as { error?: number })?.error;
          if (errCode === 2) {
            setPermissionError(
              "Unable to access Fabric resources. Verify that the required Fabric permissions have been granted to the application."
            );
          } else {
            setPermissionError(
              "Unable to authenticate with Microsoft Fabric. Verify application permissions and try again."
            );
          }
          setIsGenerating(false);
          return;
        }
      }

      // Generate the dashboard definition JSON
      const dashDef = generateDashboardDefinition({
        title: dashboardName,
        telemetryTable: telemetryTableName,
        propertiesTable: propertiesTableName,
        telemetryFields: includedTelemetry,
        propertyFields: includedProperties,
        queryServiceUri,
        databaseId: selectedDatabaseId,
      });

      const dashJson = JSON.stringify(dashDef);
      const payload = toBase64(dashJson);

      const body = {
        displayName: dashboardName,
        type: "KQLDashboard",
        definition: {
          parts: [
            {
              path: "RealTimeDashboard.json",
              payload,
              payloadType: "InlineBase64",
            },
          ],
        },
      };

      const response = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/items`,
        {
          method: "POST",
          headers: {
            Authorization: "Bearer " + token,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(body),
        }
      );

      if (response.status === 201 || response.status === 200) {
        const result = await response.json();
        setDashboardCreated(true);
        updateContext("dashDashboardId", result.id || "");
        updateContext("dashDashboardName", dashboardName);
        updateContext("dashDashboardCreated", true);
        return;
      }

      // Handle long-running operation (202)
      if (response.status === 202) {
        const location = response.headers.get("Location");
        const retryAfter = parseInt(response.headers.get("Retry-After") || "2", 10);
        if (location) {
          for (let attempt = 0; attempt < 30; attempt++) {
            await new Promise((r) => setTimeout(r, retryAfter * 1000));
            const pollResp = await fetch(location, {
              headers: { Authorization: "Bearer " + token },
            });
            if (pollResp.status === 200) {
              const result = await pollResp.json();
              if (result.status === "Succeeded") {
                const createdId = result.id || "";
                setDashboardCreated(true);
                updateContext("dashDashboardId", createdId);
                updateContext("dashDashboardName", dashboardName);
                updateContext("dashDashboardCreated", true);
                return;
              }
              if (result.status === "Failed" || result.status === "Cancelled") {
                throw new Error(getFabricCreationErrorMessage(
                  pollResp.status,
                  result,
                  "Dashboard",
                  dashboardName
                ));
              }
            } else {
              const errorPayload = parseFabricErrorPayload(await pollResp.text());
              throw new Error(getFabricCreationErrorMessage(
                pollResp.status,
                errorPayload,
                "Dashboard",
                dashboardName
              ));
            }
          }
          throw new Error(
            "Dashboard generation is taking longer than expected. Refresh and verify whether the dashboard was created successfully."
          );
        }
      }

      const errorPayload = parseFabricErrorPayload(await response.text());
      throw new Error(getFabricCreationErrorMessage(
        response.status,
        errorPayload,
        "Dashboard",
        dashboardName
      ));
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setGenerateError(msg);
      setDashboardCreated(false);
      updateContext("dashDashboardCreated", false);
    } finally {
      setIsGenerating(false);
      updateContext("dashDashboardCreating", false);
    }
  }, [
    workloadClient, workspaceId, queryServiceUri, selectedDatabaseId,
    telemetryTableName, propertiesTableName, includedTelemetry, includedProperties,
    dashboardName, updateContext,
  ]);

  const hasSelection = includedTelemetry.length + includedProperties.length > 0;

  if (!hasSelection) {
    return (
      <div className="iot-solution-step">
        <h2 className="iot-solution-step-title">Generate Dashboard</h2>
        <MessageBar intent="warning" style={{ marginTop: "12px" }}>
          <MessageBarBody>
            No fields were selected in the previous step. Go back and inspect your
            tables to select telemetry and property fields to visualize.
          </MessageBarBody>
        </MessageBar>
      </div>
    );
  }

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Generate Dashboard</h2>
      <Text className="iot-solution-step-description">
        Create a Real-Time Dashboard from the fields you selected. Numeric
        time-series telemetry is presented as line charts, non-numeric telemetry is
        presented in recent-value tables, and properties are presented as
        last-known-value cards.
      </Text>

      {permissionError && (
        <MessageBar intent="error" style={{ marginBottom: "12px" }}>
          <MessageBarBody>{permissionError}</MessageBarBody>
        </MessageBar>
      )}

      {/* Selection summary */}
      <div className="iot-solution-config-card" style={{ marginBottom: "16px" }}>
        <Text weight="semibold" size={400} block style={{ marginBottom: "8px" }}>
          Selected Fields
        </Text>
        <Text size={200} block style={{ color: "var(--colorNeutralForeground3)" }}>
          {includedTelemetry.length} telemetry ({timeSeriesCount} time-series, {nonTimeSeriesCount} non-time-series)
          {" · "}
          {includedProperties.length} properties
        </Text>
      </div>

      <Text className="iot-solution-resource-name-note">
        Default resource names are provided below. You can customize them if needed.
      </Text>

      <div className="iot-solution-field" style={{ maxWidth: "400px", marginBottom: "12px" }}>
        <Label className="iot-solution-field-label" htmlFor="dash-name-input">
          Dashboard Name
        </Label>
        <Input
          id="dash-name-input"
          value={dashboardName}
          disabled={isGenerating || dashboardCreated}
          onChange={(_, data) => {
            setDashboardName(data.value);
            setDashboardCreated(false);
            updateContext("dashDashboardName", data.value);
            updateContext("dashDashboardCreated", false);
            updateContext("dashDashboardId", "");
            resetStepsFrom(stepIndex);
          }}
        />
      </div>

      {!isGenerating && !dashboardCreated && (
        <button
          onClick={generateDashboard}
          disabled={!dashboardName.trim() || !hasSelection}
          style={{
            padding: "8px 16px",
            background:
              dashboardName.trim() && hasSelection
                ? "var(--colorBrandBackground)"
                : "var(--colorNeutralBackgroundDisabled)",
            color:
              dashboardName.trim() && hasSelection
                ? "var(--colorNeutralForegroundOnBrand)"
                : "var(--colorNeutralForegroundDisabled)",
            border: "none",
            borderRadius: "var(--borderRadiusMedium)",
            cursor: dashboardName.trim() && hasSelection ? "pointer" : "not-allowed",
            fontFamily: "var(--fontFamilyBase)",
            fontSize: "var(--fontSizeBase300)",
            fontWeight: "var(--fontWeightSemibold)" as any,
          }}
        >
          Generate Dashboard
        </button>
      )}

      {isGenerating && (
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          <Spinner size="small" />
          <Text>Creating dashboard...</Text>
        </div>
      )}

      {generateError && (
        <MessageBar intent="error" style={{ marginTop: "12px" }}>
          <MessageBarBody>{generateError}</MessageBarBody>
        </MessageBar>
      )}

      {dashboardCreated && (
        <div style={{ marginTop: "12px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "6px", marginBottom: "8px" }}>
            <CheckmarkCircle24Filled style={{ color: "var(--colorPaletteGreenForeground1)" }} />
            <Text style={{ color: "var(--colorPaletteGreenForeground1)" }}>
              Dashboard "{dashboardName}" created successfully.
            </Text>
          </div>
        </div>
      )}
    </div>
  );
}
