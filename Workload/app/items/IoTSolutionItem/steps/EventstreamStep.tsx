import React, { useState, useCallback, useMemo } from "react";
import {
  Input,
  Label,
  Spinner,
  Text,
  Button,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  DismissCircle24Filled,
} from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import {
  getFabricCreationErrorMessage,
  parseFabricErrorPayload,
} from "../FabricCreationError";
import "../IoTSolutionItem.scss";

const FABRIC_WRITE_SCOPE =
  "https://api.fabric.microsoft.com/Item.ReadWrite.All";
const FABRIC_API_BASE = "https://api.fabric.microsoft.com/v1";

function generateSuffix(): string {
  return Math.random().toString(36).substring(2, 6);
}

/**
 * Build the Eventstream topology JSON for:
 *   Custom Endpoint source → Default Stream → SQL operator → Eventhouse (processed ingestion)
 *
 * The SQL operator extracts the raw JSON body and EventHub metadata headers,
 * outputting two columns: data (string) and headers (dynamic).
 * The destination uses ProcessedIngestion mode with the target table pre-configured.
 */
function buildEventstreamTopology(
  streamDisplayName: string,
  workspaceId: string,
  kqlDatabaseId: string,
  databaseName: string,
  tableName: string,
): object {
  const sourceName = "CustomEndpoint-Source";
  const defaultStreamName = `${streamDisplayName}-stream`;
  const destName = "Eventhouse";

  const sqlQuery =
    ` SELECT\n` +
    `     JSON_STRINGIFY(stream) AS data,\n` +
    `     GETMETADATAPROPERTYVALUE(stream, '[EventHub]') AS headers\n` +
    ` INTO [${destName}]\n` +
    ` FROM [${defaultStreamName}] AS stream`;

  return {
    sources: [
      {
        name: sourceName,
        type: "CustomEndpoint",
        properties: {},
      },
    ],
    destinations: [
      {
        name: destName,
        type: "Eventhouse",
        properties: {
          dataIngestionMode: "ProcessedIngestion",
          workspaceId: workspaceId,
          itemId: kqlDatabaseId,
          databaseName: databaseName,
          tableName: tableName,
          inputSerialization: {
            type: "Json",
            properties: { encoding: "UTF8" },
          },
        },
        inputNodes: [{ name: "SqlCode" }],
        inputSchemas: [
          { name: "SqlCode", schema: { columns: [] } },
        ],
      },
    ],
    streams: [
      {
        name: defaultStreamName,
        type: "DefaultStream",
        properties: {},
        inputNodes: [{ name: sourceName }],
      },
    ],
    operators: [
      {
        name: "SqlCode",
        type: "SQL",
        inputNodes: [{ name: defaultStreamName }],
        properties: {
          query: sqlQuery,
          advancedSettings: null,
        },
        inputSchemas: [
          { name: defaultStreamName, schema: { columns: [] } },
        ],
      },
    ],
    compatibilityLevel: "1.1",
  };
}

interface CreationStatus {
  step: string;
  status: "pending" | "running" | "success" | "error";
  error?: string;
}

interface EndpointDetails {
  eventstreamId: string;
  eventstreamName: string;
  sourceId?: string;
  fullyQualifiedNamespace?: string;
  eventHubName?: string;
}

interface EventstreamStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Eventstream creation.
 *
 * Creates two Eventstream items in the workspace:
 * - Telemetry stream: routes device telemetry from IoT Hub custom endpoint → KQL table
 * - Properties stream: routes device twin property changes → KQL table
 *
 * Each Eventstream will have a Custom Endpoint source (Event Hub-compatible)
 * and a KQL Database destination using the mappings from the table setup step.
 */
export function EventstreamStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
}: EventstreamStepProps) {
  // Include sanitized IoT Hub name so stream names are scoped per hub
  const sanitizedHubName = (wizardContext.hubName || "hub")
    .replace(/[^a-zA-Z0-9\-_]/g, "_");
  const defaults = useMemo(
    () => {
      const sharedSuffix = wizardContext.sessionSuffix || generateSuffix();
      return {
        telemetry: `TelemetryStream_${sanitizedHubName}_${sharedSuffix}`,
        properties: `PropertiesStream_${sanitizedHubName}_${sharedSuffix}`,
        suffix: sharedSuffix,
      };
    },
    [sanitizedHubName, wizardContext.sessionSuffix]
  );

  const [telemetryStreamName, setTelemetryStreamName] = useState<string>(
    wizardContext.telemetryStreamName || defaults.telemetry
  );
  const [propertiesStreamName, setPropertiesStreamName] = useState<string>(
    wizardContext.propertiesStreamName || defaults.properties
  );

  const [creating, setCreating] = useState(false);
  const [permissionError, setPermissionError] = useState("");
  const [steps, setSteps] = useState<CreationStatus[]>([]);
  const [streamsCreated, setStreamsCreated] = useState(
    !!wizardContext.eventstreamsCreated
  );

  // Context from previous steps
  const databaseId = wizardContext.databaseId || "";
  const databaseName = wizardContext.databaseName || "";
  const telemetryTableName = wizardContext.telemetryTableName || "";
  const propertiesTableName = wizardContext.propertiesTableName || "";

  // Endpoint details are fetched after creation and saved to wizard context
  // for the IoT Hub routing step — no UI display needed.

  /**
   * Fetch Custom Endpoint connection details for an Eventstream.
   * Two-step process:
   *   1. GET topology → find the CustomEndpoint source → get its id
   *   2. GET /sources/{sourceId}/connection → get Event Hub namespace, name, keys
   * The Eventstream may take a moment to provision, so we retry.
   */
  const fetchEndpointDetails = async (
    token: string,
    eventstreamId: string,
    eventstreamName: string,
    retries = 8,
    delayMs = 3000
  ): Promise<EndpointDetails> => {
    const headers = { Authorization: `Bearer ${token}` };

    // Step 1: Get topology to find Custom Endpoint source ID
    let sourceId = "";
    for (let attempt = 0; attempt < retries; attempt++) {
      if (attempt > 0) {
        await new Promise((r) => setTimeout(r, delayMs));
      }
      try {
        const topoResp = await fetch(
          `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams/${eventstreamId}/topology`,
          { headers }
        );
        if (!topoResp.ok) continue;
        const topology = await topoResp.json();
        const customSource = (topology.sources || []).find(
          (s: { type: string }) => s.type === "CustomEndpoint"
        );
        if (customSource?.id) {
          sourceId = customSource.id;
          break;
        }
      } catch {
        // Retry
      }
    }

    if (!sourceId) {
      return { eventstreamId, eventstreamName };
    }

    // Step 2: Get connection details from the source
    for (let attempt = 0; attempt < retries; attempt++) {
      if (attempt > 0) {
        await new Promise((r) => setTimeout(r, delayMs));
      }
      try {
        const connResp = await fetch(
          `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams/${eventstreamId}/sources/${sourceId}/connection`,
          { headers }
        );
        if (!connResp.ok) continue;
        const conn = await connResp.json();
        return {
          eventstreamId,
          eventstreamName,
          sourceId,
          fullyQualifiedNamespace: conn.fullyQualifiedNamespace,
          eventHubName: conn.eventHubName,
        };
      } catch {
        // Retry
      }
    }

    return { eventstreamId, eventstreamName, sourceId };
  };

  const invalidateRouting = () => {
    resetStepsFrom(stepIndex + 1);
    const routingKeys = [
      "telemetryStreamId",
      "propertiesStreamId",
      "telemetryEndpointNamespace",
      "telemetryEndpointEventHubName",
      "telemetryEndpointSourceId",
      "propertiesEndpointNamespace",
      "propertiesEndpointEventHubName",
      "propertiesEndpointSourceId",
      "miHasWorkspaceAccess",
      "miWorkspaceRole",
      "routingConfigured",
      "telemetryRouteName",
      "propertiesRouteName",
      "telemetryEndpointName",
      "propertiesEndpointName",
    ];
    for (const key of routingKeys) {
      updateContext(
        key,
        key.endsWith("Configured") || key.endsWith("Access") ? false : ""
      );
    }
  };

  const handleTelemetryNameChange = (value: string) => {
    setTelemetryStreamName(value);
    setStreamsCreated(false);
    updateContext("eventstreamsCreated", false);
    invalidateRouting();
    setSteps([]);
  };

  const handlePropertiesNameChange = (value: string) => {
    setPropertiesStreamName(value);
    setStreamsCreated(false);
    updateContext("eventstreamsCreated", false);
    invalidateRouting();
    setSteps([]);
  };

  // Create a single Eventstream with topology via Fabric REST API
  // Topology: Custom Endpoint source → stream → SQL operator → Eventhouse destination
  const createEventstreamItem = async (
    token: string,
    displayName: string,
    description: string,
    tableName: string,
  ): Promise<{ id: string; displayName: string }> => {
    const topology = buildEventstreamTopology(
      displayName,
      workspaceId,
      databaseId,
      databaseName,
      tableName,
    );

    const topologyBase64 = btoa(
      unescape(encodeURIComponent(JSON.stringify(topology)))
    );

    const body: Record<string, unknown> = {
      displayName,
      description,
      definition: {
        format: "eventstream",
        parts: [
          {
            path: "eventstream.json",
            payload: topologyBase64,
            payloadType: "InlineBase64",
          },
        ],
      },
    };

    const response = await fetch(
      `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams`,
      {
        method: "POST",
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      }
    );

    if (response.status === 201 || response.status === 200) {
      return await response.json();
    }

    // Handle long-running operation (202)
    if (response.status === 202) {
      const location = response.headers.get("Location");
      const retryAfter = parseInt(
        response.headers.get("Retry-After") || "2",
        10
      );

      if (location) {
        // Poll for completion
        for (let attempt = 0; attempt < 30; attempt++) {
          await new Promise((r) =>
            setTimeout(r, retryAfter * 1000)
          );
          const pollResp = await fetch(location, {
            headers: { Authorization: `Bearer ${token}` },
          });
          if (pollResp.status === 200) {
            const result = await pollResp.json();
            if (result.status === "Succeeded" && result.id) {
              return { id: result.id, displayName };
            }
            if (result.status === "Succeeded") {
              // Result might be in a different format — try to get the item
              break;
            }
            if (
              result.status === "Failed" ||
              result.status === "Cancelled"
            ) {
              throw new Error(getFabricCreationErrorMessage(
                pollResp.status,
                result,
                "Eventstream",
                displayName
              ));
            }
          }
        }
      }

      // If we got a 202 but couldn't poll, try to find the item by name
      const listResp = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams`,
        { headers: { Authorization: `Bearer ${token}` } }
      );
      if (listResp.ok) {
        const items = await listResp.json();
        const found = (items.value || []).find(
          (i: { displayName: string }) => i.displayName === displayName
        );
        if (found) return { id: found.id, displayName: found.displayName };
      }

      throw new Error(
        "Eventstream creation is still in progress. Refresh and verify that the Eventstream was successfully created."
      );
    }

    // Error response
    const errorPayload = parseFabricErrorPayload(await response.text());
    throw new Error(getFabricCreationErrorMessage(
      response.status,
      errorPayload,
      "Eventstream",
      displayName
    ));
  };

  const createEventstreams = useCallback(async () => {
    if (!workspaceId) return;

    setCreating(true);
    updateContext("eventstreamsCreating", true);
    setPermissionError("");
    setSteps([]);

    // Acquire Fabric API token with consent fallback
    let token: string;
    try {
      // Try silent first
      const silent =
        await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [FABRIC_WRITE_SCOPE],
        });
      token = silent.token;
    } catch {
      try {
        await workloadClient.auth.acquireAccessToken({
          additionalScopesToConsent: [FABRIC_WRITE_SCOPE],
        });
        const retry =
          await workloadClient.auth.acquireFrontendAccessToken({
            scopes: [FABRIC_WRITE_SCOPE],
          });
        token = retry.token;
      } catch (interactiveErr: unknown) {
        const errCode = (interactiveErr as { error?: number })?.error;
        if (errCode === 2) {
          setPermissionError(
            "Unable to access Fabric resources. Verify that the required Fabric " +
            "permissions have been granted to the application."
          );
        } else if (errCode === 1) {
          setPermissionError(
            "Access to Fabric was not granted. Please complete the sign-in and " +
            "consent process to continue."
          );
        } else {
          setPermissionError(
            "Unable to authenticate with Microsoft Fabric. Verify application " +
            "permissions and try again."
          );
        }
        setCreating(false);
        updateContext("eventstreamsCreating", false);
        return;
      }
    }

    const statusArr: CreationStatus[] = [
      { step: `Create ${telemetryStreamName}`, status: "pending" },
      { step: `Create ${propertiesStreamName}`, status: "pending" },
    ];
    setSteps([...statusArr]);

    let telemetryId = wizardContext.telemetryStreamId || "";
    let propertiesId = wizardContext.propertiesStreamId || "";
    let allSucceeded = true;

    // Create telemetry stream
    statusArr[0].status = "running";
    setSteps([...statusArr]);
    try {
      if (!telemetryId) {
        const result = await createEventstreamItem(
          token,
          telemetryStreamName,
          `Telemetry stream: Custom Endpoint → ${databaseName} (processed ingestion)`,
          telemetryTableName,
        );
        telemetryId = result.id;
        updateContext("telemetryStreamId", telemetryId);
        updateContext("telemetryStreamName", telemetryStreamName);
      }
      statusArr[0].status = "success";
    } catch (err: unknown) {
      statusArr[0].status = "error";
      statusArr[0].error =
        err instanceof Error ? err.message : String(err);
      allSucceeded = false;
    }
    setSteps([...statusArr]);

    // Create properties stream
    statusArr[1].status = "running";
    setSteps([...statusArr]);
    try {
      if (!propertiesId) {
        const result = await createEventstreamItem(
          token,
          propertiesStreamName,
          `Properties stream: Custom Endpoint → ${databaseName} (processed ingestion)`,
          propertiesTableName,
        );
        propertiesId = result.id;
        updateContext("propertiesStreamId", propertiesId);
        updateContext("propertiesStreamName", propertiesStreamName);
      }
      statusArr[1].status = "success";
    } catch (err: unknown) {
      statusArr[1].status = "error";
      statusArr[1].error =
        err instanceof Error ? err.message : String(err);
      allSucceeded = false;
    }
    setSteps([...statusArr]);

    if (allSucceeded) {
      // Fetch endpoint details from topology API
      statusArr.push({
        step: "Fetching Custom Endpoint connection details...",
        status: "running",
      });
      setSteps([...statusArr]);

      try {
        const [telemetryDetails, propertiesDetails] = await Promise.all([
          fetchEndpointDetails(token, telemetryId, telemetryStreamName),
          fetchEndpointDetails(token, propertiesId, propertiesStreamName),
        ]);
        if (
          !telemetryDetails.fullyQualifiedNamespace ||
          !telemetryDetails.eventHubName ||
          !telemetryDetails.sourceId ||
          !propertiesDetails.fullyQualifiedNamespace ||
          !propertiesDetails.eventHubName ||
          !propertiesDetails.sourceId
        ) {
          throw new Error(
            "Eventstream creation is still in progress. Refresh and verify that the Eventstream was successfully created."
          );
        }
        statusArr[statusArr.length - 1].status = "success";

        // Save endpoint details to wizard context for IoT Hub routing step
        updateContext("telemetryEndpointNamespace", telemetryDetails.fullyQualifiedNamespace);
        updateContext("telemetryEndpointEventHubName", telemetryDetails.eventHubName);
        updateContext("telemetryEndpointSourceId", telemetryDetails.sourceId);
        updateContext("propertiesEndpointNamespace", propertiesDetails.fullyQualifiedNamespace);
        updateContext("propertiesEndpointEventHubName", propertiesDetails.eventHubName);
        updateContext("propertiesEndpointSourceId", propertiesDetails.sourceId);
      } catch (err: unknown) {
        allSucceeded = false;
        statusArr[statusArr.length - 1].status = "error";
        statusArr[statusArr.length - 1].error =
          err instanceof Error
            ? err.message
            : "Eventstream creation is still in progress. Refresh and verify that the Eventstream was successfully created.";
      }
      setSteps([...statusArr]);

      if (allSucceeded) {
        setStreamsCreated(true);
        updateContext("telemetryStreamName", telemetryStreamName);
        updateContext("telemetryStreamId", telemetryId);
        updateContext("propertiesStreamName", propertiesStreamName);
        updateContext("propertiesStreamId", propertiesId);
        updateContext("sessionSuffix", defaults.suffix);
        updateContext("eventstreamsCreated", true);
      }
    }

    setCreating(false);
    updateContext("eventstreamsCreating", false);
  }, [
    workspaceId,
    databaseId,
    databaseName,
    telemetryStreamName,
    telemetryTableName,
    propertiesStreamName,
    propertiesTableName,
    workloadClient,
    updateContext,
    wizardContext.telemetryStreamId,
    wizardContext.propertiesStreamId,
  ]);

  const canCreate = !!(
    telemetryStreamName.trim() &&
    propertiesStreamName.trim() &&
    workspaceId &&
    !creating &&
    !streamsCreated
  );

  // When tables are being reused, the existing Eventstreams are used as-is — there is
  // nothing to create in this step (eventstreamsCreated was already set upstream).
  if (wizardContext.reuseExistingTables) {
    return (
      <div className="iot-solution-step">
        <h2 className="iot-solution-step-title">Eventstreams</h2>
        <Text className="iot-solution-step-description">
          You chose to reuse existing raw data tables, so no new Eventstreams need to
          be created in this step. The existing Eventstreams will continue to route
          data from your IoT Hub to the raw data tables as before.
        </Text>
        <div className="iot-solution-success">
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">Using existing Eventstreams</Text>
        </div>
        <div className="iot-solution-validation">
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">Telemetry Eventstream</Text>
            <Text className="iot-solution-validation-value">
              {wizardContext.telemetryStreamName || "—"}
            </Text>
          </div>
          <div className="iot-solution-validation-row">
            <Text className="iot-solution-validation-label">Properties Eventstream</Text>
            <Text className="iot-solution-validation-value">
              {wizardContext.propertiesStreamName || "—"}
            </Text>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">
        Create Eventstreams
      </h2>
      <Text className="iot-solution-step-description">
        Create Eventstreams to continuously ingest device telemetry and property
        updates. The Eventstreams prepare incoming IoT Hub messages for
        query-ready storage and deliver them to the raw KQL tables, where they
        become available for real-time analytics and the other accelerator
        experiences.
      </Text>

      {/* Permission error */}
      {permissionError && (
        <MessageBar intent="error" className="iot-solution-permission-error">
          <MessageBarBody>
            <strong>Missing API Permission: </strong>
            {permissionError}
          </MessageBarBody>
        </MessageBar>
      )}

      <div className="iot-solution-form">
        <Text className="iot-solution-resource-name-note">
          Default resource names are provided below. You can customize them if needed.
        </Text>

        {/* Telemetry stream name */}
        <div className="iot-solution-field">
          <Label
            className="iot-solution-field-label"
            required
            htmlFor="telemetry-stream"
          >
            Telemetry Stream Name
          </Label>
          <Input
            id="telemetry-stream"
            value={telemetryStreamName}
            onChange={(_, data) => handleTelemetryNameChange(data.value)}
            disabled={creating || streamsCreated}
            placeholder="e.g., TelemetryStream"
          />
          <Text className="iot-solution-field-hint">
            Routes device telemetry → {telemetryTableName || "telemetry table"}
          </Text>
        </div>

        {/* Properties stream name */}
        <div className="iot-solution-field">
          <Label
            className="iot-solution-field-label"
            required
            htmlFor="properties-stream"
          >
            Properties Stream Name
          </Label>
          <Input
            id="properties-stream"
            value={propertiesStreamName}
            onChange={(_, data) => handlePropertiesNameChange(data.value)}
            disabled={creating || streamsCreated}
            placeholder="e.g., PropertiesStream"
          />
          <Text className="iot-solution-field-hint">
            Routes device properties → {propertiesTableName || "properties table"}
          </Text>
        </div>

        {/* Create button */}
        <Button
          appearance="primary"
          onClick={createEventstreams}
          disabled={!canCreate}
        >
          {creating ? "Creating Eventstreams..." : "Create Eventstreams"}
        </Button>
      </div>

      {/* Progress steps */}
      {steps.length > 0 && (
        <div className="iot-solution-creation-steps">
          {steps.map((s, i) => (
            <div key={i} className="iot-solution-creation-step">
              {s.status === "running" && <Spinner size="tiny" />}
              {s.status === "success" && (
                <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
              )}
              {s.status === "error" && (
                <DismissCircle24Filled primaryFill="var(--colorPaletteRedForeground1)" />
              )}
              {s.status === "pending" && (
                <span className="iot-solution-step-dot" />
              )}
              <div className="iot-solution-creation-step-content">
                <Text
                  className={`iot-solution-creation-step-label ${s.status}`}
                >
                  {s.step}
                </Text>
                {s.error && (
                  <Text className="iot-solution-field-error">{s.error}</Text>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Success */}
      {streamsCreated && (
        <div className="iot-solution-success">
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">
            Eventstreams created successfully
          </Text>
        </div>
      )}
    </div>
  );
}
