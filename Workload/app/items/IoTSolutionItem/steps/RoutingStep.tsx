import React, { useState, useCallback } from "react";
import {
  Spinner,
  Text,
  Button,
  Dropdown,
  Option,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  Warning24Filled,
} from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { HubIdentityOption, collectHubIdentities } from "./IoTHubStep";
import {
  buildNativeEventStreamRouting,
  IOT_HUB_ROUTING_API_VERSION,
} from "../IoTHubRouting";
import "../IoTSolutionItem.scss";

const FABRIC_API_BASE = "https://api.fabric.microsoft.com/v1";
const FABRIC_READ_SCOPE = "https://api.fabric.microsoft.com/Workspace.Read.All";
const FABRIC_WRITE_SCOPE =
  "https://api.fabric.microsoft.com/Workspace.ReadWrite.All";
const ARM_SCOPE = "https://management.azure.com/.default";

interface RoutingStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

type CheckStatus = "idle" | "checking" | "has-access" | "no-access" | "error";
type AddStatus = "idle" | "adding" | "success" | "error";
type RoutingStatus = "idle" | "configuring" | "success" | "error";

/**
 * IoT Hub routing configuration.
 *
 * 1. Check if the IoT Hub MI has Contributor/Member access to the Fabric workspace
 * 2. If not, attempt to add it (or instruct user to do it manually)
 * 3. Configure IoT Hub message routes to Eventstream custom endpoints
 */
export function RoutingStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
}: RoutingStepProps) {
  const hubName = wizardContext.hubName || "";
  const subscriptionId = wizardContext.subscriptionId || "";
  const resourceGroup = wizardContext.resourceGroup || "";

  // All managed identities attached to the IoT Hub (system-assigned + user-assigned).
  // The routing endpoint authenticates with exactly one of these, so the user picks
  // which one — and that specific identity is the one we check/grant workspace access for.
  //
  // Recompute from the raw hub details (always stored at validation) so the full set is
  // available even if the pre-computed iotHubIdentities array is stale/missing. Fall back
  // to the stored array, then to a single system-assigned entry from the principal ID.
  const identities: HubIdentityOption[] = React.useMemo(() => {
    const fromDetails = collectHubIdentities(wizardContext.iotHubDetails);
    if (fromDetails.length > 0) {
      return fromDetails;
    }
    const stored = wizardContext.iotHubIdentities as
      | HubIdentityOption[]
      | undefined;
    if (stored && stored.length > 0) {
      return stored;
    }
    const fallback: HubIdentityOption[] = wizardContext.iotHubPrincipalId
      ? [
          {
            kind: "system",
            key: "system",
            principalId: wizardContext.iotHubPrincipalId,
            resourceId: null,
            displayName: "System-assigned",
          },
        ]
      : [];
    return fallback;
  }, [
    wizardContext.iotHubDetails,
    wizardContext.iotHubIdentities,
    wizardContext.iotHubPrincipalId,
  ]);

  const [selectedIdentityKey, setSelectedIdentityKey] = useState<string>(
    wizardContext.routingIdentityKey ||
      identities.find((i) => i.kind === "system")?.key ||
      identities[0]?.key ||
      ""
  );

  const selectedIdentity =
    identities.find((i) => i.key === selectedIdentityKey) || null;
  const principalId = selectedIdentity?.principalId || "";
  const selectedResourceId = selectedIdentity?.resourceId || null;

  // Endpoint details from Eventstream step
  const telemetryEndpointNamespace =
    wizardContext.telemetryEndpointNamespace || "";
  const telemetryEndpointEventHubName =
    wizardContext.telemetryEndpointEventHubName || "";
  const propertiesEndpointNamespace =
    wizardContext.propertiesEndpointNamespace || "";
  const propertiesEndpointEventHubName =
    wizardContext.propertiesEndpointEventHubName || "";
  const telemetryEndpointSourceId =
    wizardContext.telemetryEndpointSourceId || "";
  const propertiesEndpointSourceId =
    wizardContext.propertiesEndpointSourceId || "";
  const telemetryStreamId = wizardContext.telemetryStreamId || "";
  const propertiesStreamId = wizardContext.propertiesStreamId || "";

  const [miCheckStatus, setMiCheckStatus] = useState<CheckStatus>(
    wizardContext.miHasWorkspaceAccess ? "has-access" : "idle"
  );
  const [miRole, setMiRole] = useState<string>(
    wizardContext.miWorkspaceRole || ""
  );
  const [miCheckError, setMiCheckError] = useState("");

  const [addStatus, setAddStatus] = useState<AddStatus>("idle");
  const [addError, setAddError] = useState("");

  const [routingStatus, setRoutingStatus] = useState<RoutingStatus>(
    wizardContext.routingConfigured ? "success" : "idle"
  );
  const [routingError, setRoutingError] = useState("");

  const [permissionError, setPermissionError] = useState("");

  // Switching the routing identity changes which principal we check/grant, so reset
  // the access-check state and persist the selection for the routing configuration.
  const handleIdentityChange = useCallback(
    (key: string) => {
      setSelectedIdentityKey(key);
      setMiCheckStatus("idle");
      setMiRole("");
      setMiCheckError("");
      setAddStatus("idle");
      setAddError("");
      setRoutingStatus("idle");
      setRoutingError("");
      resetStepsFrom(stepIndex);
      const opt = identities.find((i) => i.key === key) || null;
      updateContext("routingIdentityKey", key);
      updateContext("routingIdentityPrincipalId", opt?.principalId || "");
      updateContext("routingIdentityResourceId", opt?.resourceId || null);
      updateContext("miHasWorkspaceAccess", false);
      updateContext("routingConfigured", false);
    },
    [identities, resetStepsFrom, stepIndex, updateContext]
  );

  // Acquire a Fabric token with the given scope, with consent fallback
  const acquireFabricToken = async (scope: string): Promise<string> => {
    try {
      const silent =
        await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [scope],
        });
      return silent.token;
    } catch {
      await workloadClient.auth.acquireAccessToken({
        additionalScopesToConsent: [scope],
      });
      const retry =
        await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [scope],
        });
      return retry.token;
    }
  };

  // Acquire an ARM token with consent fallback
  const acquireArmToken = async (): Promise<string> => {
    try {
      const silent =
        await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [ARM_SCOPE],
        });
      return silent.token;
    } catch {
      await workloadClient.auth.acquireAccessToken({
        additionalScopesToConsent: [ARM_SCOPE],
      });
      const retry =
        await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [ARM_SCOPE],
        });
      return retry.token;
    }
  };

  /**
   * Check if the IoT Hub MI (principalId) is in the workspace role assignments.
   * Uses GET /v1/workspaces/{workspaceId}/roleAssignments
   */
  const checkMiAccess = useCallback(async () => {
    if (!principalId) {
      setMiCheckError(
        "The selected Azure IoT Hub is not configured correctly for Fabric routing."
      );
      setMiCheckStatus("error");
      return;
    }

    setMiCheckStatus("checking");
    setMiCheckError("");
    setPermissionError("");

    try {
      const token = await acquireFabricToken(FABRIC_READ_SCOPE);

      const resp = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/roleAssignments`,
        { headers: { Authorization: `Bearer ${token}` } }
      );

      if (!resp.ok) {
        throw new Error(
          `Failed to list workspace role assignments (${resp.status})`
        );
      }

      const data = await resp.json();
      const assignments: Array<{
        principal: { id: string; type: string };
        role: string;
      }> = data.value || [];

      // Look for the IoT Hub MI principal
      const match = assignments.find(
        (a) =>
          a.principal.id.toLowerCase() === principalId.toLowerCase() &&
          a.principal.type === "ServicePrincipal"
      );

      if (match) {
        setMiRole(match.role);
        setMiCheckStatus("has-access");
        updateContext("miHasWorkspaceAccess", true);
        updateContext("miWorkspaceRole", match.role);
      } else {
        setMiCheckStatus("no-access");
        updateContext("miHasWorkspaceAccess", false);
      }
    } catch {
      setMiCheckError(
        "Unable to access Fabric resources. Verify that the required Fabric permissions have been granted to the application."
      );
      setMiCheckStatus("error");
    }
  }, [principalId, workspaceId, workloadClient, updateContext]);

  /**
   * Add the IoT Hub MI as Contributor to the workspace.
   * Uses POST /v1/workspaces/{workspaceId}/roleAssignments
   * Requires caller to be Member+ and Workspace.ReadWrite.All scope.
   */
  const addMiToWorkspace = useCallback(async () => {
    setAddStatus("adding");
    setAddError("");
    setPermissionError("");

    try {
      const token = await acquireFabricToken(FABRIC_WRITE_SCOPE);

      const resp = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/roleAssignments`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            principal: {
              id: principalId,
              type: "ServicePrincipal",
            },
            role: "Contributor",
          }),
        }
      );

      if (resp.status === 201 || resp.status === 200) {
        setAddStatus("success");
        setMiRole("Contributor");
        setMiCheckStatus("has-access");
        updateContext("miHasWorkspaceAccess", true);
        updateContext("miWorkspaceRole", "Contributor");
      } else {
        const errorText = await resp.text();
        let detail = "";
        try {
          const parsed = JSON.parse(errorText);
          detail = parsed?.message || parsed?.error?.message || errorText;
        } catch {
          detail = errorText;
        }
        throw new Error(
          `Required permissions could not be granted automatically. (${resp.status}): ${detail}`
        );
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setAddError(msg);
      setAddStatus("error");
    }
  }, [principalId, workspaceId, workloadClient, updateContext]);

  /**
   * Configure IoT Hub message routing via ARM API.
   * Creates two native Eventstream endpoints and two routes.
   */
  const configureRouting = useCallback(async () => {
    if (
      !telemetryEndpointNamespace ||
      !telemetryEndpointEventHubName ||
      !propertiesEndpointNamespace ||
      !propertiesEndpointEventHubName ||
      !telemetryEndpointSourceId ||
      !propertiesEndpointSourceId ||
      !telemetryStreamId ||
      !propertiesStreamId
    ) {
      setRoutingError(
        "The selected Eventstream is incomplete or unavailable. Recreate the Eventstream and try again."
      );
      setRoutingStatus("error");
      return;
    }

    setRoutingStatus("configuring");
    setRoutingError("");

    try {
      const token = await acquireArmToken();

      // Get the current IoT Hub configuration
      const hubUrl = `https://management.azure.com/subscriptions/${encodeURIComponent(subscriptionId)}/resourceGroups/${encodeURIComponent(resourceGroup)}/providers/Microsoft.Devices/IotHubs/${encodeURIComponent(hubName)}?api-version=${IOT_HUB_ROUTING_API_VERSION}`;

      const getResp = await fetch(hubUrl, {
        headers: { Authorization: `Bearer ${token}` },
      });
      if (!getResp.ok) {
        throw new Error(
          `Failed to get IoT Hub (${getResp.status}): ${await getResp.text()}`
        );
      }
      const hubData = await getResp.json();

      // Use the same suffix generated during Eventstream creation
      const suffix = wizardContext.sessionSuffix || Math.random().toString(36).substring(2, 8);

      const telemetryEndpointName = `fabric-telemetry-${suffix}`;
      const propertiesEndpointName = `fabric-properties-${suffix}`;

      const telemetryRouteName = `fabric-telemetry-route-${suffix}`;
      const propertiesRouteName = `fabric-properties-route-${suffix}`;

      // Update the IoT Hub with the new routing configuration
      hubData.properties.routing = buildNativeEventStreamRouting(
        hubData.properties.routing,
        {
          telemetryEndpoint: {
            name: telemetryEndpointName,
            endpointUri: `sb://${telemetryEndpointNamespace}`,
            entityPath: telemetryEndpointEventHubName,
            workspaceId,
            eventStreamId: telemetryStreamId,
            sourceId: telemetryEndpointSourceId,
            userAssignedIdentity: selectedResourceId,
          },
          propertiesEndpoint: {
            name: propertiesEndpointName,
            endpointUri: `sb://${propertiesEndpointNamespace}`,
            entityPath: propertiesEndpointEventHubName,
            workspaceId,
            eventStreamId: propertiesStreamId,
            sourceId: propertiesEndpointSourceId,
            userAssignedIdentity: selectedResourceId,
          },
          telemetryRouteName,
          propertiesRouteName,
        },
      );

      const putResp = await fetch(hubUrl, {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(hubData),
      });

      if (!putResp.ok) {
        const errBody = await putResp.text();
        let detail = "";
        try {
          const parsed = JSON.parse(errBody);
          detail = parsed?.error?.message || errBody;
        } catch {
          detail = errBody;
        }
        throw new Error(
          `Unable to configure telemetry routing from Azure IoT Hub to Microsoft Fabric. (${putResp.status}): ${detail}`
        );
      }

      setRoutingStatus("success");
      updateContext("routingConfigured", true);
      updateContext("telemetryRouteName", telemetryRouteName);
      updateContext("propertiesRouteName", propertiesRouteName);
      updateContext("telemetryEndpointName", telemetryEndpointName);
      updateContext("propertiesEndpointName", propertiesEndpointName);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setRoutingError(msg);
      setRoutingStatus("error");
    }
  }, [
    subscriptionId,
    resourceGroup,
    hubName,
    workspaceId,
    telemetryEndpointNamespace,
    telemetryEndpointEventHubName,
    telemetryEndpointSourceId,
    telemetryStreamId,
    propertiesEndpointNamespace,
    propertiesEndpointEventHubName,
    propertiesEndpointSourceId,
    propertiesStreamId,
    selectedResourceId,
    wizardContext.sessionSuffix,
    workloadClient,
    updateContext,
  ]);

  const hasEndpointDetails =
    !!telemetryEndpointNamespace &&
    !!propertiesEndpointNamespace &&
    !!telemetryEndpointSourceId &&
    !!propertiesEndpointSourceId &&
    !!telemetryStreamId &&
    !!propertiesStreamId;

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">
        IoT Hub Routing
      </h2>
      <Text className="iot-solution-step-description">
        Configure new message routes in Azure IoT Hub to continuously send device telemetry and properties
        into Microsoft Fabric for downstream analytics and operational intelligence experiences.
      </Text>

      {permissionError && (
        <MessageBar intent="error" className="iot-solution-permission-error">
          <MessageBarBody>
            <strong>Permission Error: </strong>
            {permissionError}
          </MessageBarBody>
        </MessageBar>
      )}

      {/* Section 1: MI Workspace Access Check */}
      <div className="iot-solution-config-card">
        <Text className="iot-solution-config-card-title">
          <strong>1. Workspace Access for IoT Hub Managed Identity</strong>
        </Text>

        <Text
          className="iot-solution-step-description"
          style={{ display: "block", marginBottom: 8 }}
        >
          Select the managed identity that IoT Hub will use to securely authenticate
          with Fabric. The selected identity will need access to this Fabric workspace.
        </Text>

        <div className="iot-solution-validation" style={{ marginBottom: 8 }}>
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">IoT Hub</Text>
            <Text className="iot-solution-validation-value">{hubName}</Text>
          </div>
          <div
            className="iot-solution-validation-row"
            style={{ alignItems: "center" }}
          >
            <Text className="iot-solution-validation-label">Routing identity</Text>
            {identities.length > 0 ? (
              <Dropdown
                style={{ minWidth: 260 }}
                value={
                  selectedIdentity
                    ? `${selectedIdentity.displayName} (${
                        selectedIdentity.kind === "system"
                          ? "system-assigned"
                          : "user-assigned"
                      })`
                    : ""
                }
                selectedOptions={selectedIdentityKey ? [selectedIdentityKey] : []}
                onOptionSelect={(_, data) => {
                  if (data.optionValue) {
                    handleIdentityChange(data.optionValue);
                  }
                }}
              >
                {identities.map((idn) => (
                  <Option key={idn.key} value={idn.key} text={idn.displayName}>
                    {idn.displayName}{" "}
                    {idn.kind === "system"
                      ? "(system-assigned)"
                      : "(user-assigned)"}
                  </Option>
                ))}
              </Dropdown>
            ) : (
              <Text className="iot-solution-validation-value">
                No managed identity found
              </Text>
            )}
          </div>
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">Principal ID</Text>
            <Text className="iot-solution-validation-value">
              {principalId || "Not available"}
            </Text>
          </div>
        </div>

        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <Button
            appearance="primary"
            onClick={checkMiAccess}
            disabled={miCheckStatus === "checking" || !principalId}
          >
            {miCheckStatus === "checking"
              ? "Checking..."
              : miCheckStatus === "has-access" || miCheckStatus === "no-access"
                ? "Re-check"
                : "Check Access"}
          </Button>
          {miCheckStatus === "checking" && <Spinner size="tiny" />}
        </div>

        {/* Check result */}
        {miCheckStatus === "has-access" && (
          <div className="iot-solution-success" style={{ marginTop: 8 }}>
            <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
            <Text className="iot-solution-success-text">
              IoT Hub MI has <strong>{miRole}</strong> access to this workspace
            </Text>
          </div>
        )}

        {miCheckStatus === "no-access" && (
          <>
            <div
              className="iot-solution-warning"
              style={{
                marginTop: 8,
                display: "flex",
                gap: 8,
                alignItems: "flex-start",
              }}
            >
              <Warning24Filled primaryFill="var(--colorPaletteYellowForeground1)" />
              <Text>
                The Azure IoT Hub managed identity does not have sufficient
                permissions on this Fabric workspace.
              </Text>
            </div>

            <div style={{ marginTop: 8, display: "flex", gap: 8, alignItems: "center" }}>
              <Button
                appearance="primary"
                onClick={addMiToWorkspace}
                disabled={addStatus === "adding"}
              >
                {addStatus === "adding"
                  ? "Adding..."
                  : "Add as Contributor"}
              </Button>
              {addStatus === "adding" && <Spinner size="tiny" />}
            </div>

            {addStatus === "success" && (
              <div className="iot-solution-success" style={{ marginTop: 8 }}>
                <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
                <Text className="iot-solution-success-text">
                  IoT Hub MI added as Contributor
                </Text>
              </div>
            )}

            {addStatus === "error" && (
              <div style={{ marginTop: 8 }}>
                <MessageBar intent="error">
                  <MessageBarBody>
                    <strong>Could not add role assignment: </strong>
                    {addError}
                  </MessageBarBody>
                </MessageBar>
                <div className="iot-solution-info" style={{ marginTop: 8 }}>
                  <Text className="iot-solution-info-text">
                    <strong>Manual steps:</strong> Open the Fabric workspace
                    settings → Manage access → Add people or groups → paste the
                    IoT Hub principal ID (<code>{principalId}</code>) → select{" "}
                    <strong>Contributor</strong> → Add. Then click{" "}
                    <strong>Re-check</strong> above.{" "}
                    <a
                      href="https://learn.microsoft.com/en-us/fabric/real-time-intelligence/event-streams/connect-using-managed-identity#step-2-assign-fabric-workspace-permissions"
                      target="_blank"
                      rel="noopener noreferrer"
                    >
                      See documentation
                    </a>
                  </Text>
                </div>
              </div>
            )}
          </>
        )}

        {miCheckStatus === "error" && (
          <MessageBar intent="error" style={{ marginTop: 8 }}>
            <MessageBarBody>{miCheckError}</MessageBarBody>
          </MessageBar>
        )}
      </div>

      {/* Section 2: Configure IoT Hub Routing */}
      {miCheckStatus === "has-access" && (
        <div className="iot-solution-config-card">
          <Text className="iot-solution-config-card-title">
            <strong>2. Configure Message Routes</strong>
          </Text>

          {!hasEndpointDetails && (
            <MessageBar intent="warning" style={{ marginBottom: 8 }}>
              <MessageBarBody>
                The selected Eventstream is incomplete or unavailable. Recreate
                the Eventstream and try again.
              </MessageBarBody>
            </MessageBar>
          )}

          {hasEndpointDetails && (
            <>
              <Text block style={{ marginBottom: 8 }}>
                This will create two custom routing endpoints and two message routes
                in the IoT Hub using the managed identity selected above.
              </Text>
              <div className="iot-solution-validation" style={{ marginBottom: 8 }}>
                <div className="iot-solution-validation-row">
                  <Text className="iot-solution-validation-label">
                    Telemetry Route
                  </Text>
                  <Text className="iot-solution-validation-value">
                    DeviceMessages → Eventstream custom endpoint
                  </Text>
                </div>
                <div className="iot-solution-validation-row">
                  <Text className="iot-solution-validation-label">
                    Properties Route
                  </Text>
                  <Text className="iot-solution-validation-value">
                    TwinChangeEvents → Eventstream custom endpoint
                  </Text>
                </div>
              </div>

              <Button
                appearance="primary"
                onClick={configureRouting}
                disabled={
                  routingStatus === "configuring" ||
                  routingStatus === "success"
                }
              >
                {routingStatus === "configuring"
                  ? "Configuring..."
                  : routingStatus === "success"
                    ? "Routing Configured"
                    : "Configure Routing"}
              </Button>
              {routingStatus === "configuring" && (
                <Spinner size="tiny" style={{ marginLeft: 8 }} />
              )}
            </>
          )}

          {routingStatus === "success" && (
            <div className="iot-solution-success" style={{ marginTop: 8 }}>
              <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
              <Text className="iot-solution-success-text">
                IoT Hub routing configured successfully
              </Text>
            </div>
          )}

          {routingStatus === "error" && (
            <MessageBar intent="error" style={{ marginTop: 8 }}>
              <MessageBarBody>{routingError}</MessageBarBody>
            </MessageBar>
          )}
        </div>
      )}
    </div>
  );
}
