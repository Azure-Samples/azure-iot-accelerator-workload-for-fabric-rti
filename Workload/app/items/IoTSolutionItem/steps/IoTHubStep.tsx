// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useState, useEffect, useRef } from "react";
import {
  Label,
  Spinner,
  Text,
  Combobox,
  Option,
  Button,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  DismissCircle24Filled,
  Warning24Filled,
} from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import "../IoTSolutionItem.scss";

const ARM_SCOPE = "https://management.azure.com/.default";

const RESOURCE_GRAPH_URL =
  "https://management.azure.com/providers/Microsoft.ResourceGraph/resources?api-version=2021-03-01";

/** Azure Resource Graph query listing every IoT Hub the signed-in user can read, across all accessible subscriptions. */
const IOT_HUBS_QUERY =
  "resources " +
  "| where type =~ 'microsoft.devices/iothubs' " +
  "| project name, resourceGroup, subscriptionId, location, id " +
  "| order by name asc";

/** A single IoT Hub resource as returned by Azure Resource Graph. */
interface IoTHubResource {
  id: string;
  name: string;
  resourceGroup: string;
  subscriptionId: string;
  location: string;
}

interface IoTHubStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
}

/** Build the Authorization header value for a bearer token. */
function authHeader(token: string): string {
  return "Bearer " + token;
}

/** Check whether the IoT Hub has system-assigned or user-assigned managed identity enabled. */
function checkManagedIdentity(data: any): { hasIdentity: boolean; identityType: string } {
  const identity = data?.identity;
  if (!identity || !identity.type || identity.type.toLowerCase() === "none") {
    return { hasIdentity: false, identityType: "None" };
  }
  return { hasIdentity: true, identityType: identity.type };
}

/**
 * A single managed identity associated with an IoT Hub — either the system-assigned
 * identity or one of the user-assigned identities. Any of these can be the one the
 * IoT Hub routing endpoint authenticates with, so we surface all of them.
 */
export interface HubIdentityOption {
  kind: "system" | "user";
  /** Stable key: "system" for the system-assigned identity, otherwise the identity's ARM resource ID. */
  key: string;
  /** Object (principal) ID used for Fabric workspace role assignments. */
  principalId: string;
  /** ARM resource ID for a user-assigned identity; null for the system-assigned identity. */
  resourceId: string | null;
  /** Friendly label (identity resource name for user-assigned; "System-assigned" otherwise). */
  displayName: string;
  clientId?: string;
}

/**
 * Collect every managed identity attached to the IoT Hub: the system-assigned identity
 * (identity.principalId) plus each user-assigned identity (identity.userAssignedIdentities).
 */
export function collectHubIdentities(data: any): HubIdentityOption[] {
  const identity = data?.identity;
  const options: HubIdentityOption[] = [];
  if (!identity) {
    return options;
  }
  const type = (identity.type || "").toLowerCase();

  if (type.includes("systemassigned") && identity.principalId) {
    options.push({
      kind: "system",
      key: "system",
      principalId: identity.principalId,
      resourceId: null,
      displayName: "System-assigned",
    });
  }

  const userAssigned = identity.userAssignedIdentities || {};
  for (const resourceId of Object.keys(userAssigned)) {
    const uai = userAssigned[resourceId] || {};
    if (!uai.principalId) {
      continue;
    }
    const name = resourceId.split("/").pop() || resourceId;
    options.push({
      kind: "user",
      key: resourceId,
      principalId: uai.principalId,
      resourceId,
      displayName: name,
      clientId: uai.clientId,
    });
  }
  return options;
}

/**
 * IoT Hub configuration.
 *
 * The user picks an IoT Hub from a searchable dropdown populated by an Azure Resource
 * Graph query (all hubs the user can access, across subscriptions). A manual-entry
 * fallback (subscription / resource group / hub name) is available if the hub can't be
 * found. Once chosen, the hub is validated via ARM to confirm it exists, is accessible,
 * and has managed identity enabled (required for later Eventstream routing).
 */
export function IoTHubStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
}: IoTHubStepProps) {
  const [subscriptionId, setSubscriptionId] = useState<string>(wizardContext.subscriptionId || "");
  const [resourceGroup, setResourceGroup] = useState<string>(wizardContext.resourceGroup || "");
  const [hubName, setHubName] = useState<string>(wizardContext.hubName || "");
  const [isValidating, setIsValidating] = useState(false);
  const [validationResult, setValidationResult] = useState<any>(wizardContext.iotHubDetails || null);
  const [validationError, setValidationError] = useState<string>(wizardContext.validationError || "");
  const [identityWarning, setIdentityWarning] = useState<string>("");

  // IoT Hub picker (Azure Resource Graph)
  const [hubs, setHubs] = useState<IoTHubResource[]>([]);
  const [loadingHubs, setLoadingHubs] = useState(false);
  const [hubsError, setHubsError] = useState<string>("");
  const [query, setQuery] = useState<string>(wizardContext.hubName || "");
  const [selectedHubId, setSelectedHubId] = useState<string>("");

  const isFirstRender = useRef(true);

  // Sync form fields to wizard context on every change
  useEffect(() => {
    updateContext("subscriptionId", subscriptionId);
  }, [subscriptionId]);

  useEffect(() => {
    updateContext("resourceGroup", resourceGroup);
  }, [resourceGroup]);

  useEffect(() => {
    updateContext("hubName", hubName);
  }, [hubName]);

  // Clear validation when the selected hub changes — but not on the initial mount,
  // so a previously validated hub keeps its checkmark when the wizard is reopened.
  useEffect(() => {
    if (isFirstRender.current) {
      isFirstRender.current = false;
      return;
    }
    if (validationResult || validationError) {
      setValidationResult(null);
      setValidationError("");
      setIdentityWarning("");
      updateContext("iotHubDetails", null);
      updateContext("validationError", "");
      updateContext("iotHubValidated", false);
    }
  }, [subscriptionId, resourceGroup, hubName]);

  /** Load the list of IoT Hubs the user can access via Azure Resource Graph. */
  const loadHubs = async () => {
    setLoadingHubs(true);
    setHubsError("");
    try {
      const token = await acquireTokenWithConsent(workloadClient, ARM_SCOPE);
      const collected: IoTHubResource[] = [];
      let skipToken: string | undefined;

      do {
        const body: any = {
          query: IOT_HUBS_QUERY,
          options: { resultFormat: "objectArray", "$top": 1000 },
        };
        if (skipToken) {
          body.options["$skipToken"] = skipToken;
        }

        const response = await fetch(RESOURCE_GRAPH_URL, {
          method: "POST",
          headers: {
            Authorization: authHeader(token.token),
            "Content-Type": "application/json",
          },
          body: JSON.stringify(body),
        });

        if (!response.ok) {
          const errorBody = await response.text();
          let message = `HTTP ${response.status}`;
          try {
            const parsed = JSON.parse(errorBody);
            message = parsed?.error?.message || parsed?.Message || message;
          } catch {
            // use status text
          }
          throw new Error(message);
        }

        const json = await response.json();
        const rows: any[] = Array.isArray(json?.data) ? json.data : [];
        for (const r of rows) {
          collected.push({
            id: r.id,
            name: r.name,
            resourceGroup: r.resourceGroup,
            subscriptionId: r.subscriptionId,
            location: r.location,
          });
        }
        skipToken = json?.["$skipToken"] || undefined;
      } while (skipToken);

      setHubs(collected);

      // Restore prior selection (if any) so reopening the wizard keeps the picker in sync.
      if (wizardContext.hubName) {
        const prior = collected.find(
          (h) =>
            h.name === wizardContext.hubName &&
            (!wizardContext.resourceGroup || h.resourceGroup === wizardContext.resourceGroup) &&
            (!wizardContext.subscriptionId || h.subscriptionId === wizardContext.subscriptionId)
        );
        if (prior) {
          setSelectedHubId(prior.id);
          setQuery(prior.name);
        }
      }
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setHubsError(msg);
    } finally {
      setLoadingHubs(false);
    }
  };

  // Load hubs once on mount.
  useEffect(() => {
    void loadHubs();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const validateIoTHub = async (sub: string, rg: string, hub: string) => {
    setIsValidating(true);
    setValidationResult(null);
    setValidationError("");
    setIdentityWarning("");

    // Clear all downstream context and completion state.
    resetStepsFrom(stepIndex + 1);
    const downstreamKeys = [
      "eventhouseId", "eventhouseName", "databaseId", "databaseName",
      "queryServiceUri", "ingestionServiceUri", "eventhouseValidated",
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
      updateContext(key, key.endsWith("Validated") || key.endsWith("Created") || key.endsWith("Configured") || key.endsWith("Access") ? false : "");
    }

    try {
      const token = await acquireTokenWithConsent(workloadClient, ARM_SCOPE);

      const url = `https://management.azure.com/subscriptions/${encodeURIComponent(sub)}/resourceGroups/${encodeURIComponent(rg)}/providers/Microsoft.Devices/IotHubs/${encodeURIComponent(hub)}?api-version=2023-06-30`;

      const response = await fetch(url, {
        headers: { Authorization: authHeader(token.token) },
      });

      if (!response.ok) {
        const errorBody = await response.text();
        let message = `HTTP ${response.status}`;
        try {
          const parsed = JSON.parse(errorBody);
          message = parsed?.error?.message || parsed?.Message || message;
        } catch {
          // use status text
        }
        throw new Error(message);
      }

      const data = await response.json();

      // Check managed identity
      const { hasIdentity, identityType } = checkManagedIdentity(data);
      if (!hasIdentity) {
        setIdentityWarning(
          "Managed Identity is required for routing Azure IoT Hub telemetry into Microsoft Fabric. " +
          "Enable a system-assigned or user-assigned identity before continuing."
        );
        updateContext("iotHubValidated", false);
        // Still show the result so user sees what was found, but don't mark as validated
        setValidationResult(data);
        updateContext("iotHubDetails", data);
        return;
      }

      setValidationResult(data);
      updateContext("iotHubDetails", data);
      updateContext("iotHubIdentityType", identityType);
      updateContext("iotHubPrincipalId", data.identity?.principalId || "");
      updateContext("iotHubIdentities", collectHubIdentities(data));
      updateContext("iotHubValidated", true);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setValidationError(msg);
      updateContext("validationError", msg);
      updateContext("iotHubValidated", false);
    } finally {
      setIsValidating(false);
    }
  };

  /** Handle picking a hub from the dropdown: populate the fields and validate it. */
  const onHubSelect = (hubId: string | undefined) => {
    const hub = hubs.find((h) => h.id === hubId);
    if (!hub) return;
    setSelectedHubId(hub.id);
    setQuery(hub.name);
    setSubscriptionId(hub.subscriptionId);
    setResourceGroup(hub.resourceGroup);
    setHubName(hub.name);
    void validateIoTHub(hub.subscriptionId, hub.resourceGroup, hub.name);
  };

  // When the input shows the selected hub's name (i.e. the user isn't actively searching),
  // don't filter — so the full list is available as soon as the dropdown is opened.
  const selectedHub = hubs.find((h) => h.id === selectedHubId);
  const activeQuery = selectedHub && query === selectedHub.name ? "" : query.trim();
  const filteredHubs = activeQuery
    ? hubs.filter(
        (h) =>
          h.name.toLowerCase().includes(activeQuery.toLowerCase()) ||
          h.resourceGroup.toLowerCase().includes(activeQuery.toLowerCase())
      )
    : hubs;

  const isFormComplete = !!(subscriptionId.trim() && resourceGroup.trim() && hubName.trim());

  const primaryBtnStyle = (enabled: boolean): React.CSSProperties => ({
    padding: "8px 16px",
    background: enabled ? "var(--colorBrandBackground)" : "var(--colorNeutralBackgroundDisabled)",
    color: enabled ? "var(--colorNeutralForegroundOnBrand)" : "var(--colorNeutralForegroundDisabled)",
    border: "none",
    borderRadius: "var(--borderRadiusMedium)",
    cursor: enabled ? "pointer" : "not-allowed",
    fontFamily: "var(--fontFamilyBase)",
    fontSize: "var(--fontSizeBase300)",
    fontWeight: "var(--fontWeightSemibold)" as any,
  });

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">
        Select your IoT Hub
      </h2>
      <Text className="iot-solution-step-description">
        Select the Azure IoT Hub you want to connect to your Fabric solution.
        The hub is validated to ensure it exists, you have access to it, and managed identity is enabled
        (required for message routing to Fabric).
      </Text>

      <div className="iot-solution-form">
        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="hub-combobox">
            IoT Hub
          </Label>
          {loadingHubs ? (
            <Spinner size="tiny" label="Loading IoT Hubs..." />
          ) : hubsError ? (
            <>
              <MessageBar intent="error">
                <MessageBarBody>
                  Unable to load available Azure IoT Hubs. Error: {hubsError}
                </MessageBarBody>
              </MessageBar>
              <div style={{ marginTop: 8 }}>
                <Button size="small" onClick={() => { void loadHubs(); }}>Retry</Button>
              </div>
            </>
          ) : (
            <>
              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <Combobox
                  id="hub-combobox"
                  placeholder="Search for an IoT Hub..."
                  value={query}
                  selectedOptions={selectedHubId ? [selectedHubId] : []}
                  onChange={(ev) => setQuery((ev.target as HTMLInputElement).value)}
                  onOptionSelect={(_, data) => onHubSelect(data.optionValue)}
                  style={{ minWidth: "360px" }}
                >
                  {filteredHubs.length === 0 ? (
                    <Option key="__none__" value="__none__" text="" disabled>
                      No matching IoT Hubs
                    </Option>
                  ) : (
                    filteredHubs.map((h) => (
                      <Option key={h.id} value={h.id} text={h.name}>
                        <div style={{ display: "flex", flexDirection: "column" }}>
                          <span>{h.name}</span>
                          <span style={{ fontSize: "var(--fontSizeBase200)", color: "var(--colorNeutralForeground3)" }}>
                            {h.resourceGroup} · {h.location}
                          </span>
                        </div>
                      </Option>
                    ))
                  )}
                </Combobox>
                <Button size="small" appearance="secondary" onClick={() => { void loadHubs(); }} disabled={loadingHubs}>
                  Refresh
                </Button>
              </div>
              <Text style={{ display: "block", marginTop: 4, fontSize: "var(--fontSizeBase200)", color: "var(--colorNeutralForeground3)" }}>
                {hubs.length} IoT Hub{hubs.length === 1 ? "" : "s"} found.
              </Text>
            </>
          )}
        </div>
      </div>

      {/* Re-validate button (e.g. after enabling managed identity in the portal) */}
      {!isValidating && selectedHubId && (
        <div>
          <button
            onClick={() => void validateIoTHub(subscriptionId, resourceGroup, hubName)}
            disabled={!isFormComplete}
            style={primaryBtnStyle(isFormComplete)}
          >
            Re-validate IoT Hub
          </button>
        </div>
      )}

      {/* Loading State */}
      {isValidating && (
        <div className="iot-solution-loading">
          <Spinner size="medium" />
          <Text className="iot-solution-loading-text">
            Validating IoT Hub...
          </Text>
        </div>
      )}

      {/* Validation Error */}
      {validationError && (
        <div className="iot-solution-error">
          <DismissCircle24Filled primaryFill="var(--colorPaletteRedForeground1)" />
          <Text className="iot-solution-error-text">{validationError}</Text>
        </div>
      )}

      {/* Managed Identity Warning */}
      {identityWarning && (
        <div className="iot-solution-warning">
          <Warning24Filled primaryFill="var(--colorPaletteDarkOrangeForeground1)" />
          <Text className="iot-solution-warning-text">{identityWarning}</Text>
        </div>
      )}

      {/* Validation Success */}
      {validationResult && !identityWarning && (
        <div className="iot-solution-success">
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">IoT Hub validated successfully</Text>
        </div>
      )}

      {/* Hub Details (shown on both success and identity warning) */}
      {validationResult && (
        <div className="iot-solution-validation">
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">Name</Text>
            <Text className="iot-solution-validation-value">{validationResult.name}</Text>
          </div>
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">Location</Text>
            <Text className="iot-solution-validation-value">{validationResult.location}</Text>
          </div>
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">Identity</Text>
            <Text className="iot-solution-validation-value">
              {validationResult.identity?.type || "None"}
            </Text>
          </div>
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">State</Text>
            <Text className="iot-solution-validation-value">
              {validationResult.properties?.state}
            </Text>
          </div>
        </div>
      )}
    </div>
  );
}
