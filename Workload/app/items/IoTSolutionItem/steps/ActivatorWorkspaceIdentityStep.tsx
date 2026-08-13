import React, { useState, useCallback, useEffect } from "react";
import {
  Button,
  Spinner,
  Text,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { CheckmarkCircle24Filled, ShieldKeyhole24Regular } from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import "../IoTSolutionItem.scss";

const FABRIC_WRITE_SCOPE = "https://api.fabric.microsoft.com/Item.ReadWrite.All";
const FABRIC_CONNECTION_SCOPE = "https://api.fabric.microsoft.com/Connection.ReadWrite.All";
const FABRIC_WORKSPACE_SCOPE = "https://api.fabric.microsoft.com/Workspace.ReadWrite.All";
const FABRIC_API_BASE = "https://api.fabric.microsoft.com/v1";

interface ActivatorWorkspaceIdentityStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Activator Wizard Step 2: Verify (and optionally enable) the workspace identity.
 *
 * Fabric Activator monitors the modeled-data table through an Eventstream. That
 * Eventstream reads from the Eventhouse (Azure Data Explorer) using the
 * workspace's managed identity as the connection credential — this avoids
 * storing any secrets. This step checks whether the workspace identity exists
 * and, if not, lets the user provision it before any resources are created.
 */
export function ActivatorWorkspaceIdentityStep({
  updateContext,
  workloadClient,
  workspaceId,
}: ActivatorWorkspaceIdentityStepProps) {
  const [checking, setChecking] = useState(false);
  const [provisioning, setProvisioning] = useState(false);
  const [identityPresent, setIdentityPresent] = useState(false);
  const [error, setError] = useState("");
  const [checkedOnce, setCheckedOnce] = useState(false);

  const acquireFabricWriteToken = useCallback(async (): Promise<string> => {
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
  }, [workloadClient]);

  const pollLongRunningOperation = useCallback(async (
    locationUrl: string,
    retryAfterSec: number,
    token: string
  ): Promise<any> => {
    let url = locationUrl;
    let delay = retryAfterSec * 1000 || 2000;
    for (let attempt = 0; attempt < 60; attempt++) {
      await new Promise((r) => setTimeout(r, delay));
      const pollResp = await fetch(url, {
        headers: { Authorization: "Bearer " + token },
      });
      if (pollResp.status === 200 || pollResp.status === 201) {
        return pollResp.json().catch(() => ({}));
      }
      if (pollResp.status === 202) {
        const retryHeader = pollResp.headers.get("Retry-After");
        if (retryHeader) delay = parseInt(retryHeader, 10) * 1000 || delay;
        const newLocation = pollResp.headers.get("Location");
        if (newLocation) url = newLocation;
        continue;
      }
      const body = await pollResp.json().catch(() => ({}));
      const status = body?.status;
      if (status === "Succeeded") return body;
      if (status === "Failed") {
        throw new Error(
          "The operation did not complete successfully. Retry the operation or review workspace configuration."
        );
      }
      if (status === "Running" || status === "NotStarted") {
        const retryHeader = pollResp.headers.get("Retry-After");
        if (retryHeader) delay = parseInt(retryHeader, 10) * 1000 || delay;
        continue;
      }
      return body;
    }
    throw new Error(
      "The operation did not complete successfully. Retry the operation or review workspace configuration."
    );
  }, []);

  const markReady = useCallback((ready: boolean) => {
    setIdentityPresent(ready);
    updateContext("actWorkspaceIdentityReady", ready);
  }, [updateContext]);

  // Check whether the workspace already has a managed identity.
  const checkIdentity = useCallback(async () => {
    if (!workspaceId) return;
    setChecking(true);
    setError("");
    try {
      const token = await acquireFabricWriteToken();
      const wsResp = await fetch(`${FABRIC_API_BASE}/workspaces/${workspaceId}`, {
        headers: { Authorization: "Bearer " + token },
      });
      if (!wsResp.ok) throw new Error(`Failed to get workspace (${wsResp.status})`);
      const wsData = await wsResp.json();
      markReady(!!wsData.workspaceIdentity);
    } catch {
      setError("Unable to verify workspace identity configuration.");
      markReady(false);
    } finally {
      setChecking(false);
      setCheckedOnce(true);
    }
  }, [workspaceId, acquireFabricWriteToken, markReady]);

  // Provision (enable) the workspace identity.
  const enableIdentity = useCallback(async () => {
    if (!workspaceId) return;
    setProvisioning(true);
    setError("");
    try {
      const token = await acquireFabricWriteToken();
      const provResp = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/provisionIdentity`,
        {
          method: "POST",
          headers: { Authorization: "Bearer " + token },
        }
      );
      if (provResp.status === 200) {
        // Already provisioned
      } else if (provResp.status === 202) {
        const location = provResp.headers.get("Location") || "";
        const retryAfter = parseInt(provResp.headers.get("Retry-After") || "5", 10);
        if (location) {
          await pollLongRunningOperation(location, retryAfter, token);
        }
      } else {
        const errText = await provResp.text();
        throw new Error(`Failed to enable workspace identity (${provResp.status}): ${errText}`);
      }
      markReady(true);
    } catch {
      setError(
        "Workspace identity could not be enabled. Verify that you have Fabric Workspace Admin permissions."
      );
      markReady(false);
    } finally {
      setProvisioning(false);
    }
  }, [workspaceId, acquireFabricWriteToken, pollLongRunningOperation, markReady]);

  // Auto-check on mount
  useEffect(() => {
    checkIdentity();
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Workspace Identity</h2>
      <Text className="iot-solution-step-description">
        Workspace identity enables Microsoft Fabric services to securely access
        and manage the resources required for the Activator data connection and
        automated experiences. The Eventstream uses this identity to read modeled
        data from your Eventhouse without storing credentials.
      </Text>

      {error && (
        <MessageBar intent="error" style={{ marginBottom: "12px" }}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {checking && (
        <div style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "12px" }}>
          <Spinner size="small" />
          <Text>Checking workspace identity...</Text>
        </div>
      )}

      {!checking && identityPresent && (
        <div className="iot-solution-success" style={{ marginTop: "12px" }}>
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">
            Workspace identity is ready. Fabric can securely access the resources
            required by this Activator experience.
          </Text>
        </div>
      )}

      {!checking && checkedOnce && !identityPresent && (
        <div className="iot-solution-config-card" style={{ marginTop: "12px" }}>
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "8px" }}>
            <ShieldKeyhole24Regular />
            <Text weight="semibold" size={400}>
              Workspace identity is not enabled
            </Text>
          </div>
          <Text size={200} block style={{ marginBottom: "12px", color: "var(--colorNeutralForeground3)" }}>
          Enable workspace identity so Fabric can securely create and use the
          Eventhouse connection required by Activator. You need the Admin role
          on this workspace to enable it.
          </Text>
          <Button
            appearance="primary"
            onClick={enableIdentity}
            disabled={provisioning}
          >
            {provisioning ? (
              <>
                <Spinner size="tiny" style={{ marginRight: "6px" }} />
                Enabling...
              </>
            ) : (
              "Enable workspace identity"
            )}
          </Button>
        </div>
      )}
    </div>
  );
}
