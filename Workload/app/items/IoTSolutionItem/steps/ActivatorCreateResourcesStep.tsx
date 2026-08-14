// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useState, useCallback } from "react";
import {
  Button,
  Input,
  Label,
  Spinner,
  Text,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  CheckmarkCircle24Filled,
  DismissCircle24Filled,
} from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import {
  getFabricCreationErrorMessage,
  hasFabricErrorCode,
  parseFabricErrorPayload,
} from "../FabricCreationError";
import "../IoTSolutionItem.scss";

const FABRIC_WRITE_SCOPE = "https://api.fabric.microsoft.com/Item.ReadWrite.All";
const FABRIC_CONNECTION_SCOPE = "https://api.fabric.microsoft.com/Connection.ReadWrite.All";
const FABRIC_WORKSPACE_SCOPE = "https://api.fabric.microsoft.com/Workspace.ReadWrite.All";
const KUSTO_SCOPE = "https://kusto.kusto.windows.net/.default";
const FABRIC_API_BASE = "https://api.fabric.microsoft.com/v1";

interface CreationStep {
  label: string;
  status: "pending" | "running" | "done" | "error";
  error?: string;
}

interface ActivatorCreateResourcesStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

function kustoTypeToEventstreamType(kustoType: string): string {
  switch (kustoType) {
    case "string": return "Nvarchar(max)";
    case "real": return "Float";
    case "int": return "BigInt";
    case "long": return "BigInt";
    case "datetime": return "DateTime";
    case "bool": return "Bit";
    case "dynamic": return "Record";
    case "guid": return "Nvarchar(max)";
    case "timespan": return "Nvarchar(max)";
    case "decimal": return "Float";
    default: return "Nvarchar(max)";
  }
}

/**
 * Activator Wizard Step 3: Create the resources.
 *
 * Using the table selected in step 1 and the workspace identity verified in
 * step 2, this creates the connection to the KQL database, the Activator
 * (Reflex), and an Eventstream that streams the table into the Activator.
 */
export function ActivatorCreateResourcesStep({
  wizardContext,
  updateContext,
  workloadClient,
  workspaceId,
}: ActivatorCreateResourcesStepProps) {
  const queryServiceUri: string = wizardContext.actQueryServiceUri || "";
  const selectedDatabaseName: string = wizardContext.actDatabaseName || "";
  const selectedTable: string = wizardContext.actTableName || "";
  const nameSuffix: string = wizardContext.actNameSuffix || "iot";

  const [activatorName, setActivatorName] = useState<string>(
    wizardContext.actActivatorName || ""
  );
  const [eventstreamName, setEventstreamName] = useState<string>(
    wizardContext.actEventstreamName || ""
  );

  const [isCreating, setIsCreating] = useState(false);
  const [permissionError, setPermissionError] = useState("");
  const [creationSteps, setCreationSteps] = useState<CreationStep[]>([]);
  const [created, setCreated] = useState(!!wizardContext.activatorCreated);

  const stopCreating = useCallback(() => {
    setIsCreating(false);
    updateContext("actResourcesCreating", false);
  }, [updateContext]);

  const acquireFabricWriteToken = async (): Promise<string> => {
    const allScopes = [FABRIC_WRITE_SCOPE, FABRIC_CONNECTION_SCOPE, FABRIC_WORKSPACE_SCOPE];
    try {
      const result = await workloadClient.auth.acquireFrontendAccessToken({
        scopes: allScopes,
      });
      return result.token;
    } catch {
      await workloadClient.auth.acquireAccessToken({
        additionalScopesToConsent: allScopes,
      });
      const retry = await workloadClient.auth.acquireFrontendAccessToken({
        scopes: allScopes,
      });
      return retry.token;
    }
  };

  const pollLongRunningOperation = async (
    locationUrl: string,
    retryAfterSec: number,
    token: string,
    resourceType: string,
    displayName: string,
    timeoutMessage: string
  ): Promise<any> => {
    let url = locationUrl;
    let delay = retryAfterSec * 1000 || 2000;
    for (let attempt = 0; attempt < 60; attempt++) {
      await new Promise((r) => setTimeout(r, delay));
      const pollResp = await fetch(url, {
        headers: { Authorization: "Bearer " + token },
      });
      if (pollResp.status === 200 || pollResp.status === 201) {
        const body = parseFabricErrorPayload(await pollResp.text()) as Record<string, unknown>;
        if (body.status === "Failed" || body.status === "Cancelled") {
          throw new Error(getFabricCreationErrorMessage(
            pollResp.status,
            body,
            resourceType,
            displayName
          ));
        }
        if (body.status === "Running" || body.status === "NotStarted") {
          continue;
        }
        return body;
      }
      if (pollResp.status === 202) {
        const retryHeader = pollResp.headers.get("Retry-After");
        if (retryHeader) delay = parseInt(retryHeader, 10) * 1000 || delay;
        const newLocation = pollResp.headers.get("Location");
        if (newLocation) url = newLocation;
        continue;
      }
      const body = parseFabricErrorPayload(await pollResp.text());
      throw new Error(getFabricCreationErrorMessage(
        pollResp.status,
        body,
        resourceType,
        displayName
      ));
    }
    throw new Error(timeoutMessage);
  };

  const findConnectionByName = async (
    token: string,
    displayName: string
  ): Promise<string> => {
    let url = `${FABRIC_API_BASE}/connections`;
    for (let page = 0; page < 20; page++) {
      const response = await fetch(url, {
        headers: { Authorization: "Bearer " + token },
      });
      if (!response.ok) return "";

      const result = await response.json();
      const connection = (result.value || []).find(
        (item: { displayName?: string }) => item.displayName === displayName
      );
      if (connection?.id) return connection.id;

      if (result.continuationUri) {
        url = result.continuationUri;
      } else if (result.continuationToken) {
        url = `${FABRIC_API_BASE}/connections?continuationToken=${encodeURIComponent(result.continuationToken)}`;
      } else {
        break;
      }
    }
    return "";
  };

  const findWorkspaceResourceByName = async (
    token: string,
    collection: "items" | "eventstreams",
    displayName: string
  ): Promise<string> => {
    let url = `${FABRIC_API_BASE}/workspaces/${workspaceId}/${collection}`;
    for (let page = 0; page < 20; page++) {
      const response = await fetch(url, {
        headers: { Authorization: "Bearer " + token },
      });
      if (!response.ok) return "";

      const result = await response.json();
      const resource = (result.value || []).find(
        (item: { displayName?: string }) => item.displayName === displayName
      );
      if (resource?.id) return resource.id;

      if (result.continuationUri) {
        url = result.continuationUri;
      } else if (result.continuationToken) {
        url = `${FABRIC_API_BASE}/workspaces/${workspaceId}/${collection}?continuationToken=${encodeURIComponent(result.continuationToken)}`;
      } else {
        break;
      }
    }
    return "";
  };

  const createResources = useCallback(async () => {
    if (!queryServiceUri || !selectedDatabaseName || !selectedTable) return;

    setIsCreating(true);
    updateContext("actResourcesCreating", true);
    setPermissionError("");

    const steps: CreationStep[] = [
      { label: "Create connection to KQL database", status: "pending" },
      { label: "Create Activator", status: "pending" },
      { label: "Query table schema", status: "pending" },
      { label: "Create Eventstream", status: "pending" },
    ];
    setCreationSteps([...steps]);

    let fabricToken: string;
    try {
      fabricToken = await acquireFabricWriteToken();
    } catch {
      setPermissionError(
        "Unable to authenticate with Microsoft Fabric. Verify application permissions and try again."
      );
      stopCreating();
      return;
    }

    let connectionId: string = wizardContext.actConnectionId || "";
    let createdActivatorId: string = wizardContext.actActivatorId || "";
    let eventstreamId: string = wizardContext.actEventstreamId || "";
    let schemaColumns: { Name: string; CslType: string }[] = [];

    // --- Step 1: Create connection ---
    steps[0].status = "running";
    setCreationSteps([...steps]);
    try {
      if (!connectionId) {
        const connectionName = `KQL-${selectedDatabaseName}-${nameSuffix}`;
        const connBody = {
          connectivityType: "ShareableCloud",
          displayName: connectionName,
          connectionDetails: {
            type: "AzureDataExplorer",
            creationMethod: "AzureDataExplorer.Contents",
            parameters: [
              {
                dataType: "Text",
                name: "cluster",
                value: queryServiceUri,
              },
            ],
          },
          credentialDetails: {
            singleSignOnType: "None",
            connectionEncryption: "NotEncrypted",
            skipTestConnection: false,
            credentials: {
              credentialType: "WorkspaceIdentity",
            },
          },
          privacyLevel: "Organizational",
        };
        const connResp = await fetch(`${FABRIC_API_BASE}/connections`, {
          method: "POST",
          headers: {
            Authorization: "Bearer " + fabricToken,
            "Content-Type": "application/json",
          },
          body: JSON.stringify(connBody),
        });
        if (!connResp.ok) {
          const errorPayload = parseFabricErrorPayload(await connResp.text());
          if (hasFabricErrorCode(errorPayload, "DuplicateConnectionName")) {
            connectionId = await findConnectionByName(fabricToken, connectionName);
          }
          if (!connectionId) {
            throw new Error(getFabricCreationErrorMessage(
              connResp.status,
              errorPayload,
              "Connection",
              connectionName
            ));
          }
        } else {
          const connResult = await connResp.json();
          connectionId = connResult.id;
        }
        updateContext("actConnectionId", connectionId);
      }
      steps[0].status = "done";
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      steps[0].status = "error";
      steps[0].error = msg;
      setCreationSteps([...steps]);
      stopCreating();
      return;
    }
    setCreationSteps([...steps]);

    // --- Step 2: Create Activator ---
    steps[1].status = "running";
    setCreationSteps([...steps]);
    try {
      if (!createdActivatorId) {
        const actBody = {
          displayName: activatorName,
          type: "Reflex",
        };
        const actResp = await fetch(
          `${FABRIC_API_BASE}/workspaces/${workspaceId}/items`,
          {
            method: "POST",
            headers: {
              Authorization: "Bearer " + fabricToken,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(actBody),
          }
        );

        if (actResp.status === 201 || actResp.status === 200) {
          const actResult = await actResp.json();
          createdActivatorId = actResult.id;
        } else if (actResp.status === 202) {
          const locationUrl = actResp.headers.get("Location") || "";
          const retryAfter = parseInt(actResp.headers.get("Retry-After") || "2", 10);
          if (!locationUrl) {
            throw new Error(
              "Activator creation could not be validated. Verify whether the Activator resource exists in the workspace."
            );
          }
          const pollResult = await pollLongRunningOperation(
            locationUrl,
            retryAfter,
            fabricToken,
            "Activator",
            activatorName,
            "Activator creation could not be validated. Verify whether the Activator resource exists in the workspace."
          );
          createdActivatorId = pollResult?.id || "";
        } else {
          const errorPayload = parseFabricErrorPayload(await actResp.text());
          if (hasFabricErrorCode(errorPayload, "ItemDisplayNameAlreadyInUse")) {
            createdActivatorId = await findWorkspaceResourceByName(
              fabricToken,
              "items",
              activatorName
            );
          }
          if (!createdActivatorId) {
            throw new Error(getFabricCreationErrorMessage(
              actResp.status,
              errorPayload,
              "Activator",
              activatorName
            ));
          }
        }
        if (!createdActivatorId) {
          throw new Error(
            "Activator creation could not be validated. Verify whether the Activator resource exists in the workspace."
          );
        }
        updateContext("actActivatorId", createdActivatorId);
        updateContext("actActivatorName", activatorName);
      }
      steps[1].status = "done";
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      steps[1].status = "error";
      steps[1].error = msg;
      setCreationSteps([...steps]);
      stopCreating();
      return;
    }
    setCreationSteps([...steps]);

    // --- Step 3: Query table schema ---
    steps[2].status = "running";
    setCreationSteps([...steps]);
    try {
      const kustoToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
      const schemaResp = await fetch(`${queryServiceUri}/v1/rest/mgmt`, {
        method: "POST",
        headers: {
          Authorization: "Bearer " + kustoToken.token,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          db: selectedDatabaseName,
          csl: `.show table ${selectedTable} schema as json`,
        }),
      });
      if (!schemaResp.ok) {
        throw new Error(
          "The selected table does not contain a usable schema for Activator configuration."
        );
      }
      const schemaResult = await schemaResp.json();
      const rows = schemaResult?.Tables?.[0]?.Rows || schemaResult?.[0]?.Rows || [];
      const schemaJsonStr = Array.isArray(rows[0]) ? rows[0][1] : rows[0]?.Schema;
      if (!schemaJsonStr) {
        throw new Error(
          "The selected table does not contain a usable schema for Activator configuration."
        );
      }
      const schemaObj = JSON.parse(schemaJsonStr);
      schemaColumns = (schemaObj.OrderedColumns || []).map((col: any) => ({
        Name: col.Name,
        CslType: col.CslType,
      }));
      if (schemaColumns.length === 0) {
        throw new Error(
          "The selected table contains no data columns and cannot be used for Activator configuration."
        );
      }
      steps[2].status = "done";
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      steps[2].status = "error";
      steps[2].error = msg;
      setCreationSteps([...steps]);
      stopCreating();
      return;
    }
    setCreationSteps([...steps]);

    // --- Step 4: Create Eventstream ---
    steps[3].status = "running";
    setCreationSteps([...steps]);
    try {
      const streamName = `${eventstreamName}-stream`;

      const topology = {
        sources: [
          {
            id: crypto.randomUUID(),
            name: "kql-source",
            type: "AzureDataExplorer",
            properties: {
              dataConnectionId: connectionId,
              databaseName: selectedDatabaseName,
              tableNames: selectedTable,
            },
          },
        ],
        destinations: [
          {
            id: crypto.randomUUID(),
            name: "Activator",
            type: "Activator",
            properties: {
              workspaceId,
              itemId: createdActivatorId,
              inputSerialization: {
                type: "Json",
                properties: { encoding: "UTF8" },
              },
            },
            inputNodes: [{ name: streamName }],
            inputSchemas: [
              {
                name: streamName,
                schema: {
                  columns: schemaColumns.map((col) => ({
                    name: col.Name,
                    type: kustoTypeToEventstreamType(col.CslType),
                    fields: col.CslType === "dynamic" ? ([] as string[]) : null,
                    items: null as string | null,
                  })),
                },
              },
            ],
          },
        ],
        streams: [
          {
            id: crypto.randomUUID(),
            name: streamName,
            type: "DefaultStream",
            properties: {},
            inputNodes: [{ name: "kql-source" }],
          },
        ],
        operators: [] as unknown[],
        compatibilityLevel: "1.1",
      };

      const eventstreamProperties = {
        retentionTimeInDays: 1,
        eventThroughputLevel: "Low",
        schemaMode: "None",
      };

      const toBase64 = (obj: any) =>
        btoa(unescape(encodeURIComponent(JSON.stringify(obj))));

      const esBody = {
        displayName: eventstreamName,
        definition: {
          format: "eventstream",
          parts: [
            {
              path: "eventstream.json",
              payload: toBase64(topology),
              payloadType: "InlineBase64",
            },
            {
              path: "eventstreamProperties.json",
              payload: toBase64(eventstreamProperties),
              payloadType: "InlineBase64",
            },
          ],
        },
      };

      if (!eventstreamId) {
        const esResp = await fetch(
          `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams`,
          {
            method: "POST",
            headers: {
              Authorization: "Bearer " + fabricToken,
              "Content-Type": "application/json",
            },
            body: JSON.stringify(esBody),
          }
        );

        if (esResp.status === 201 || esResp.status === 200) {
          const esResult = await esResp.json();
          eventstreamId = esResult.id || "";
        } else if (esResp.status === 202) {
          const locationUrl = esResp.headers.get("Location") || "";
          const retryAfter = parseInt(esResp.headers.get("Retry-After") || "2", 10);
          if (!locationUrl) {
            throw new Error(
              "Eventstream creation failed. Review configuration and try again."
            );
          }
          const pollResult = await pollLongRunningOperation(
            locationUrl,
            retryAfter,
            fabricToken,
            "Eventstream",
            eventstreamName,
            "Eventstream creation failed. Review configuration and try again."
          );
          eventstreamId = pollResult?.id || "";
        } else {
          const errorPayload = parseFabricErrorPayload(await esResp.text());
          if (hasFabricErrorCode(errorPayload, "ItemDisplayNameAlreadyInUse")) {
            eventstreamId = await findWorkspaceResourceByName(
              fabricToken,
              "eventstreams",
              eventstreamName
            );
          }
          if (!eventstreamId) {
            throw new Error(getFabricCreationErrorMessage(
              esResp.status,
              errorPayload,
              "Eventstream",
              eventstreamName
            ));
          }
        }
        updateContext("actEventstreamId", eventstreamId);
        updateContext("actEventstreamName", eventstreamName);
      }

      steps[3].status = "done";
      setCreationSteps([...steps]);

      // All succeeded
      setCreated(true);
      updateContext("activatorCreated", true);
      updateContext("actActivatorId", createdActivatorId);
      updateContext("actActivatorName", activatorName);
      updateContext("actEventstreamId", eventstreamId);
      updateContext("actEventstreamName", eventstreamName);
      updateContext("actConnectionId", connectionId);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      steps[3].status = "error";
      steps[3].error = msg;
      setCreationSteps([...steps]);
      stopCreating();
      return;
    }

    stopCreating();
  }, [
    queryServiceUri,
    selectedDatabaseName,
    selectedTable,
    nameSuffix,
    activatorName,
    eventstreamName,
    workspaceId,
    workloadClient,
    updateContext,
    wizardContext.actConnectionId,
    wizardContext.actActivatorId,
    wizardContext.actEventstreamId,
    stopCreating,
  ]);

  const canCreate = !!(
    selectedTable &&
    queryServiceUri &&
    selectedDatabaseName &&
    eventstreamName.trim() &&
    activatorName.trim() &&
    !created &&
    !isCreating
  );

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Create Activator Resources</h2>
      <Text className="iot-solution-step-description">
        Create the secure KQL connection, Eventstream, and Activator that deliver
        new events from
        {selectedTable ? ` "${selectedTable}"` : " your modeled data table"} for
        continuous condition evaluation. You will define the monitoring rules and
        actions in Activator after this setup is complete.
      </Text>

      {permissionError && (
        <MessageBar intent="error" className="iot-solution-permission-error">
          <MessageBarBody>
            <strong>Error: </strong>
            {permissionError}
          </MessageBarBody>
        </MessageBar>
      )}

      <div className="iot-solution-form">
        <Text weight="semibold" size={400} block style={{ marginTop: "4px" }}>
          Output Resources
        </Text>
        <Text className="iot-solution-resource-name-note">
          Default resource names are provided below. You can customize them if needed.
        </Text>

        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="act-activator-name">
            Activator Name
          </Label>
          <Input
            id="act-activator-name"
            value={activatorName}
            onChange={(_, data) => {
              setActivatorName(data.value);
              updateContext("actActivatorName", data.value);
            }}
            disabled={isCreating || created || !!wizardContext.actActivatorId}
          />
        </div>

        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="act-eventstream-name">
            Eventstream Name
          </Label>
          <Input
            id="act-eventstream-name"
            value={eventstreamName}
            onChange={(_, data) => {
              setEventstreamName(data.value);
              updateContext("actEventstreamName", data.value);
            }}
            disabled={isCreating || created || !!wizardContext.actEventstreamId}
          />
        </div>

        {/* Create button */}
        <Button
          appearance="primary"
          onClick={createResources}
          disabled={!canCreate}
          style={{ marginTop: "8px" }}
        >
          {isCreating ? (
            <>
              <Spinner size="tiny" style={{ marginRight: "6px" }} />
              Setting Up Activator...
            </>
          ) : (
            "Set Up Activator"
          )}
        </Button>
      </div>

      {/* Progress steps */}
      {creationSteps.length > 0 && (
        <div className="iot-solution-creation-steps">
          {creationSteps.map((s, i) => (
            <div key={i} className="iot-solution-creation-step">
              {s.status === "running" && <Spinner size="tiny" />}
              {s.status === "done" && (
                <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
              )}
              {s.status === "error" && (
                <DismissCircle24Filled primaryFill="var(--colorPaletteRedForeground1)" />
              )}
              {s.status === "pending" && (
                <span className="iot-solution-step-dot" />
              )}
              <div className="iot-solution-creation-step-content">
                <Text className={`iot-solution-creation-step-label ${s.status}`}>
                  {s.label}
                </Text>
                {s.error && (
                  <Text className="iot-solution-field-error">{s.error}</Text>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {created && (
        <div className="iot-solution-success">
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">
            Activator data feed created. Select <strong>Next</strong> to review
            the configured resources and monitoring examples.
          </Text>
        </div>
      )}
    </div>
  );
}
