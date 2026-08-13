import React, { useState, useCallback, useMemo, useEffect } from "react";
import {
  Input,
  Label,
  Spinner,
  Text,
  Button,
  Dropdown,
  Option,
  Radio,
  RadioGroup,
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
import "../IoTSolutionItem.scss";

const KUSTO_SCOPE = "https://kusto.kusto.windows.net/.default";
const FABRIC_API_BASE = "https://api.fabric.microsoft.com/v1";
const FABRIC_RW_SCOPE = "https://api.fabric.microsoft.com/Item.ReadWrite.All";

const TELEMETRY_ROLE = "iothub-telemetry-raw";
const PROPERTIES_ROLE = "iothub-properties-raw";

/** Build the Authorization header value for a bearer token. */
function authHeader(token: string): string {
  return "Bearer " + token;
}

function generateSuffix(): string {
  return Math.random().toString(36).substring(2, 6);
}

/** Extract @iot.meta key=value pairs from a KQL table docstring. */
function parseIotMeta(docstring: string): Record<string, string> {
  const meta: Record<string, string> = {};
  const regex = /@iot\.meta\s+(\S+?)=(\S+)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(docstring)) !== null) {
    meta[match[1]] = match[2];
  }
  return meta;
}

/** Escape a string for use inside a double-quoted KQL string literal. */
function toKqlString(value: string): string {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .replace(/\r/g, "")
    .replace(/\n/g, "\\n");
}

/**
 * Append an IoT Hub source to a table docstring's `@iot.meta source=` line. Multiple
 * sources share a single line separated by "; " (matching the modeled-data docstring
 * format) rather than one `@iot.meta source=` line per source. If the source is already
 * listed the docstring is returned unchanged; if there is no source line yet, one is added.
 */
function appendIotMetaSource(docstring: string, newSource: string): string {
  const sourceLineRegex = /^(@iot\.meta[ \t]+source=)(.*)$/m;
  const match = sourceLineRegex.exec(docstring);
  if (match) {
    const existing = match[2]
      .split(";")
      .map((s) => s.trim())
      .filter(Boolean);
    if (existing.includes(newSource)) return docstring;
    const updatedLine = `${match[1]}${[...existing, newSource].join("; ")}`;
    return docstring.replace(sourceLineRegex, () => updatedLine);
  }
  return `${docstring}\n@iot.meta source=${newSource}`;
}

interface TableCreationStatus {
  step: string;
  status: "pending" | "running" | "success" | "error";
  error?: string;
}

/** An existing raw table (with the Eventstream that feeds it) that can be reused. */
interface ReusableTable {
  tableName: string;
  eventstreamId: string;
  eventstreamName: string;
  docstring: string;
}

interface TableSetupStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

/**
 * Raw data table configuration.
 *
 * Two modes:
 * - Create: creates raw KQL tables (data: string, headers: dynamic) in the selected
 *   Eventhouse database, tagged with @iot.meta docstrings. The Eventstream step then
 *   wires Custom Endpoint → SQL transform → processed ingestion into these tables.
 * - Reuse: pick existing raw telemetry/properties tables already fed by an Eventstream
 *   custom endpoint in this workspace. Reusing skips Eventstream creation (the existing
 *   streams are used) and appends this IoT Hub as an additional source on the selected
 *   tables' docstrings (@iot.meta source=<hub>).
 */
export function TableSetupStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
}: TableSetupStepProps) {
  const sanitizedHubName = (wizardContext.hubName || "hub").replace(/[^a-zA-Z0-9_]/g, "_");
  const defaults = useMemo(() => {
    const sharedSuffix = wizardContext.sessionSuffix || generateSuffix();
    return {
      telemetry: `TelemetryRawIoTHub_${sanitizedHubName}_${sharedSuffix}`,
      properties: `PropertiesRawIoTHub_${sanitizedHubName}_${sharedSuffix}`,
      suffix: sharedSuffix,
    };
  }, [sanitizedHubName, wizardContext.sessionSuffix]);

  const [mode, setMode] = useState<"create" | "reuse">(
    wizardContext.reuseExistingTables ? "reuse" : "create"
  );

  // ---- Create mode state ----
  const [telemetryTableName, setTelemetryTableName] = useState<string>(
    wizardContext.telemetryTableName || defaults.telemetry
  );
  const [propertiesTableName, setPropertiesTableName] = useState<string>(
    wizardContext.propertiesTableName || defaults.properties
  );
  const [creating, setCreating] = useState(false);
  const [permissionError, setPermissionError] = useState("");
  const [creationSteps, setCreationSteps] = useState<TableCreationStatus[]>([]);
  const [tablesCreated, setTablesCreated] = useState(!!wizardContext.tablesCreated);

  // ---- Reuse mode state ----
  const [loadingReuse, setLoadingReuse] = useState(false);
  const [reuseError, setReuseError] = useState("");
  const [telemetryOptions, setTelemetryOptions] = useState<ReusableTable[]>([]);
  const [propertiesOptions, setPropertiesOptions] = useState<ReusableTable[]>([]);
  const [selectedTelemetry, setSelectedTelemetry] = useState<string>(
    wizardContext.reuseExistingTables ? wizardContext.telemetryTableName || "" : ""
  );
  const [selectedProperties, setSelectedProperties] = useState<string>(
    wizardContext.reuseExistingTables ? wizardContext.propertiesTableName || "" : ""
  );
  const [applying, setApplying] = useState(false);
  const [reuseSteps, setReuseSteps] = useState<TableCreationStatus[]>([]);

  const queryServiceUri = wizardContext.queryServiceUri || "";
  const databaseName = wizardContext.databaseName || "";
  const databaseId = wizardContext.databaseId || "";
  const iotHubFqdn = `${wizardContext.hubName || "unknown"}.azure-devices.net`;

  // ==========================================================================
  // Shared helpers
  // ==========================================================================

  // Execute a KQL management command; returns the parsed response (or null if empty).
  const executeKqlMgmt = async (token: string, db: string, csl: string): Promise<any> => {
    const response = await fetch(`${queryServiceUri}/v1/rest/mgmt`, {
      method: "POST",
      headers: {
        Authorization: authHeader(token),
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ db, csl }),
    });
    if (!response.ok) {
      const errorText = await response.text();
      let detail = "";
      try {
        const parsed = JSON.parse(errorText);
        detail = parsed?.error?.message || parsed?.error?.["@message"] || errorText;
      } catch {
        detail = errorText;
      }
      throw new Error(
        `Database operation failed. Verify permissions and database configuration. (${response.status}): ${detail}`
      );
    }
    const text = await response.text();
    if (!text) return null;
    try {
      return JSON.parse(text);
    } catch {
      return null;
    }
  };

  /** Acquire a Fabric API token (Item.ReadWrite.All) with interactive consent fallback. */
  const acquireFabricToken = async (): Promise<string> => {
    try {
      const silent = await workloadClient.auth.acquireFrontendAccessToken({
        scopes: [FABRIC_RW_SCOPE],
      });
      return silent.token;
    } catch {
      await workloadClient.auth.acquireAccessToken({
        additionalScopesToConsent: [FABRIC_RW_SCOPE],
      });
      const retry = await workloadClient.auth.acquireFrontendAccessToken({
        scopes: [FABRIC_RW_SCOPE],
      });
      return retry.token;
    }
  };

  /**
   * Fetch the Custom Endpoint connection (namespace + event hub name) for an Eventstream.
   * Existing streams are already provisioned, so a single attempt is enough.
   */
  const fetchEndpointDetails = async (
    token: string,
    eventstreamId: string
  ): Promise<{ namespace?: string; eventHubName?: string; sourceId?: string }> => {
    const headers = { Authorization: authHeader(token) };
    const topoResp = await fetch(
      `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams/${eventstreamId}/topology`,
      { headers }
    );
    if (!topoResp.ok) return {};
    const topology = await topoResp.json();
    const customSource = (topology.sources || []).find(
      (s: { type: string; id?: string }) => s.type === "CustomEndpoint"
    );
    if (!customSource?.id) return {};
    const connResp = await fetch(
      `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams/${eventstreamId}/sources/${customSource.id}/connection`,
      { headers }
    );
    if (!connResp.ok) return {};
    const conn = await connResp.json();
    return {
      namespace: conn.fullyQualifiedNamespace,
      eventHubName: conn.eventHubName,
      sourceId: customSource.id,
    };
  };

  // ==========================================================================
  // Create mode
  // ==========================================================================

  const invalidateDownstream = () => {
    resetStepsFrom(stepIndex + 1);
    const downstreamKeys = [
      "telemetryStreamName",
      "telemetryStreamId",
      "propertiesStreamName",
      "propertiesStreamId",
      "telemetryEndpointNamespace",
      "telemetryEndpointEventHubName",
      "telemetryEndpointSourceId",
      "propertiesEndpointNamespace",
      "propertiesEndpointEventHubName",
      "propertiesEndpointSourceId",
      "eventstreamsCreated",
      "miHasWorkspaceAccess",
      "miWorkspaceRole",
      "routingConfigured",
      "telemetryRouteName",
      "propertiesRouteName",
      "telemetryEndpointName",
      "propertiesEndpointName",
    ];
    for (const key of downstreamKeys) {
      updateContext(
        key,
        key.endsWith("Created") ||
          key.endsWith("Configured") ||
          key.endsWith("Access")
          ? false
          : ""
      );
    }
    updateContext("sessionSuffix", generateSuffix());
  };

  const handleTelemetryNameChange = (value: string) => {
    setTelemetryTableName(value);
    setTablesCreated(false);
    updateContext("tablesCreated", false);
    invalidateDownstream();
    setCreationSteps([]);
  };

  const handlePropertiesNameChange = (value: string) => {
    setPropertiesTableName(value);
    setTablesCreated(false);
    updateContext("tablesCreated", false);
    invalidateDownstream();
    setCreationSteps([]);
  };

  const createTables = useCallback(async () => {
    if (!queryServiceUri || !databaseName) return;

    setCreating(true);
    updateContext("tableSetupCreating", true);
    setPermissionError("");
    setCreationSteps([]);

    let token: string;
    try {
      const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
      token = accessToken.token;
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      if (msg.includes("not configured for the scope")) {
        setPermissionError(
          "Required database permissions are missing. Verify access to the selected KQL Database."
        );
      } else {
        setPermissionError(
          `Unable to authenticate with the selected KQL Database. ${msg}`
        );
      }
      setCreating(false);
      updateContext("tableSetupCreating", false);
      return;
    }

    const commands: { label: string; csl: string }[] = [
      {
        label: `Create table ${telemetryTableName}`,
        csl: `.create table ${telemetryTableName} (data: string, headers: dynamic)`,
      },
      {
        label: `Set docstring on ${telemetryTableName}`,
        csl: `.alter table ${telemetryTableName} docstring "Raw device telemetry from IoT Hub\\n@iot.meta tableRole=${TELEMETRY_ROLE}\\n@iot.meta source=${iotHubFqdn}"`,
      },
      {
        label: `Disable streaming ingestion on ${telemetryTableName}`,
        csl: `.alter table ${telemetryTableName} policy streamingingestion '{"IsEnabled": false}'`,
      },
      {
        label: `Create table ${propertiesTableName}`,
        csl: `.create table ${propertiesTableName} (data: string, headers: dynamic)`,
      },
      {
        label: `Set docstring on ${propertiesTableName}`,
        csl: `.alter table ${propertiesTableName} docstring "Raw device property changes from IoT Hub\\n@iot.meta tableRole=${PROPERTIES_ROLE}\\n@iot.meta source=${iotHubFqdn}"`,
      },
    ];

    const statusArr: TableCreationStatus[] = commands.map((c) => ({
      step: c.label,
      status: "pending" as const,
    }));
    setCreationSteps([...statusArr]);

    let allSucceeded = true;
    for (let i = 0; i < commands.length; i++) {
      statusArr[i].status = "running";
      setCreationSteps([...statusArr]);
      try {
        await executeKqlMgmt(token, databaseName, commands[i].csl);
        statusArr[i].status = "success";
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        statusArr[i].status = "error";
        statusArr[i].error = msg;
        allSucceeded = false;
      }
      setCreationSteps([...statusArr]);
    }

    if (allSucceeded) {
      setTablesCreated(true);
      updateContext("telemetryTableName", telemetryTableName);
      updateContext("propertiesTableName", propertiesTableName);
      updateContext("sessionSuffix", defaults.suffix);
      updateContext("tablesCreated", true);
      // Creating fresh tables → the Eventstream step must run to create the streams.
      updateContext("reuseExistingTables", false);
    }

    setCreating(false);
    updateContext("tableSetupCreating", false);
  }, [
    queryServiceUri,
    databaseName,
    telemetryTableName,
    propertiesTableName,
    workloadClient,
    updateContext,
    iotHubFqdn,
    defaults.suffix,
  ]);

  const canCreate = !!(
    telemetryTableName.trim() &&
    propertiesTableName.trim() &&
    queryServiceUri &&
    databaseName &&
    !creating &&
    !tablesCreated
  );

  // ==========================================================================
  // Reuse mode
  // ==========================================================================

  /**
   * Discover reusable raw tables: Eventstreams in this workspace that have a Custom
   * Endpoint source and ingest into a Kusto table in the selected database. The target
   * tables are classified by their @iot.meta tableRole docstring.
   */
  const loadReusableTables = useCallback(async () => {
    if (!queryServiceUri || !databaseName) return;
    setLoadingReuse(true);
    setReuseError("");
    setTelemetryOptions([]);
    setPropertiesOptions([]);

    try {
      const fabricToken = await acquireFabricToken();

      // 1. List eventstreams in the workspace.
      const listResp = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams`,
        { headers: { Authorization: authHeader(fabricToken) } }
      );
      if (!listResp.ok) {
        throw new Error(
          `Unable to load available Eventstreams. (${listResp.status})`
        );
      }
      const listJson = await listResp.json();
      const eventstreams: { id: string; displayName: string }[] = listJson.value || [];

      // 2. Inspect each topology for a CustomEndpoint source + Eventhouse destination
      //    targeting a table in the selected database.
      const tableToStream = new Map<string, { eventstreamId: string; eventstreamName: string }>();
      await Promise.all(
        eventstreams.map(async (es) => {
          try {
            const topoResp = await fetch(
              `${FABRIC_API_BASE}/workspaces/${workspaceId}/eventstreams/${es.id}/topology`,
              { headers: { Authorization: authHeader(fabricToken) } }
            );
            if (!topoResp.ok) return;
            const topology = await topoResp.json();
            const hasCustomEndpoint = (topology.sources || []).some(
              (s: { type: string }) => s.type === "CustomEndpoint"
            );
            if (!hasCustomEndpoint) return;
            for (const dest of topology.destinations || []) {
              if (dest.type !== "Eventhouse") continue;
              const p = dest.properties || {};
              const inThisDb =
                (databaseId && p.itemId === databaseId) ||
                (databaseName && p.databaseName === databaseName);
              if (!inThisDb || !p.tableName) continue;
              if (!tableToStream.has(p.tableName)) {
                tableToStream.set(p.tableName, {
                  eventstreamId: es.id,
                  eventstreamName: es.displayName,
                });
              }
            }
          } catch {
            // ignore individual eventstream failures
          }
        })
      );

      if (tableToStream.size === 0) {
        setLoadingReuse(false);
        return;
      }

      // 3. Read table docstrings from the selected database to classify roles.
      const kustoToken = (await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE)).token;
      const result = await executeKqlMgmt(
        kustoToken,
        databaseName,
        ".show tables details | project TableName, DocString"
      );
      const rows = result?.Tables?.[0]?.Rows || result?.[0]?.Rows || [];
      const docMap: Record<string, string> = {};
      for (const r of rows) {
        const name = Array.isArray(r) ? r[0] : r.TableName;
        const doc = Array.isArray(r) ? r[1] || "" : r.DocString || "";
        if (name) docMap[name] = doc;
      }

      const telem: ReusableTable[] = [];
      const props: ReusableTable[] = [];
      for (const [tableName, stream] of tableToStream.entries()) {
        const doc = docMap[tableName];
        if (!doc) continue;
        const role = parseIotMeta(doc).tableRole;
        const entry: ReusableTable = {
          tableName,
          eventstreamId: stream.eventstreamId,
          eventstreamName: stream.eventstreamName,
          docstring: doc,
        };
        if (role === TELEMETRY_ROLE) telem.push(entry);
        else if (role === PROPERTIES_ROLE) props.push(entry);
      }
      telem.sort((a, b) => a.tableName.localeCompare(b.tableName));
      props.sort((a, b) => a.tableName.localeCompare(b.tableName));
      setTelemetryOptions(telem);
      setPropertiesOptions(props);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setReuseError(msg);
    } finally {
      setLoadingReuse(false);
    }
  }, [queryServiceUri, databaseName, databaseId, workspaceId, workloadClient]);

  // Load reusable tables when entering reuse mode (once).
  useEffect(() => {
    if (
      mode === "reuse" &&
      telemetryOptions.length === 0 &&
      propertiesOptions.length === 0 &&
      !loadingReuse &&
      !reuseError
    ) {
      void loadReusableTables();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode]);

  const handleModeChange = (value: "create" | "reuse") => {
    setMode(value);
    // Switching modes invalidates any prior completion for this step.
    setTablesCreated(false);
    setCreationSteps([]);
    setReuseSteps([]);
    updateContext("tablesCreated", false);
    invalidateDownstream();
    if (value === "create") {
      // Creating fresh tables again → drop reuse wiring so the Eventstream step runs.
      updateContext("reuseExistingTables", false);
    }
  };

  const applyReuse = useCallback(async () => {
    const telem = telemetryOptions.find((t) => t.tableName === selectedTelemetry);
    const props = propertiesOptions.find((t) => t.tableName === selectedProperties);
    if (!telem || !props) return;

    setApplying(true);
    updateContext("tableSetupCreating", true);
    setPermissionError("");

    const statusArr: TableCreationStatus[] = [
      { step: "Fetch Eventstream endpoint details", status: "pending" },
      { step: `Tag ${telem.tableName} with this IoT Hub source`, status: "pending" },
      { step: `Tag ${props.tableName} with this IoT Hub source`, status: "pending" },
    ];
    setReuseSteps([...statusArr]);

    try {
      // 1. Endpoint details for both existing eventstreams (for the routing step).
      statusArr[0].status = "running";
      setReuseSteps([...statusArr]);
      const fabricToken = await acquireFabricToken();
      const [telemEp, propsEp] = await Promise.all([
        fetchEndpointDetails(fabricToken, telem.eventstreamId),
        fetchEndpointDetails(fabricToken, props.eventstreamId),
      ]);
      if (
        !telemEp.namespace ||
        !telemEp.eventHubName ||
        !telemEp.sourceId ||
        !propsEp.namespace ||
        !propsEp.eventHubName ||
        !propsEp.sourceId
      ) {
        throw new Error(
          "The selected Eventstream is incomplete or unavailable. Recreate the Eventstream and try again."
        );
      }
      statusArr[0].status = "success";
      setReuseSteps([...statusArr]);

      // 2. Append this IoT Hub as a source on both tables' docstrings.
      const kustoToken = (await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE)).token;
      const targets = [telem, props];
      for (let idx = 0; idx < targets.length; idx++) {
        const tbl = targets[idx];
        const stepIdx = idx + 1;
        statusArr[stepIdx].status = "running";
        setReuseSteps([...statusArr]);
        try {
          const newDoc = appendIotMetaSource(tbl.docstring, iotHubFqdn);
          if (newDoc !== tbl.docstring) {
            await executeKqlMgmt(
              kustoToken,
              databaseName,
              `.alter table ${tbl.tableName} docstring "${toKqlString(newDoc)}"`
            );
          }
          statusArr[stepIdx].status = "success";
        } catch (err: unknown) {
          statusArr[stepIdx].status = "error";
          statusArr[stepIdx].error = err instanceof Error ? err.message : String(err);
          setReuseSteps([...statusArr]);
          throw err;
        }
        setReuseSteps([...statusArr]);
      }

      // 3. Wire wizard context: tables + eventstreams + endpoints (skip Eventstream step).
      updateContext("telemetryTableName", telem.tableName);
      updateContext("propertiesTableName", props.tableName);
      updateContext("tablesCreated", true);
      updateContext("reuseExistingTables", true);

      updateContext("telemetryStreamId", telem.eventstreamId);
      updateContext("telemetryStreamName", telem.eventstreamName);
      updateContext("propertiesStreamId", props.eventstreamId);
      updateContext("propertiesStreamName", props.eventstreamName);

      updateContext("telemetryEndpointNamespace", telemEp.namespace);
      updateContext("telemetryEndpointEventHubName", telemEp.eventHubName || "");
      updateContext("telemetryEndpointSourceId", telemEp.sourceId);
      updateContext("propertiesEndpointNamespace", propsEp.namespace);
      updateContext("propertiesEndpointEventHubName", propsEp.eventHubName || "");
      updateContext("propertiesEndpointSourceId", propsEp.sourceId);

      updateContext("eventstreamsCreated", true);

      setTablesCreated(true);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setPermissionError(msg);
    } finally {
      setApplying(false);
      updateContext("tableSetupCreating", false);
    }
  }, [
    telemetryOptions,
    propertiesOptions,
    selectedTelemetry,
    selectedProperties,
    databaseName,
    iotHubFqdn,
    workloadClient,
    updateContext,
    workspaceId,
  ]);

  const canApplyReuse = !!(
    selectedTelemetry &&
    selectedProperties &&
    !applying &&
    !tablesCreated
  );

  const handleReuseSelectionChange = (which: "telemetry" | "properties", value: string) => {
    if (which === "telemetry") setSelectedTelemetry(value);
    else setSelectedProperties(value);
    setTablesCreated(false);
    updateContext("tablesCreated", false);
    invalidateDownstream();
    setReuseSteps([]);
  };

  // ==========================================================================
  // Render
  // ==========================================================================

  const renderProgress = (list: TableCreationStatus[]) =>
    list.length > 0 ? (
      <div className="iot-solution-creation-steps">
        {list.map((s, i) => (
          <div key={i} className="iot-solution-creation-step">
            {s.status === "running" && <Spinner size="tiny" />}
            {s.status === "success" && (
              <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
            )}
            {s.status === "error" && (
              <DismissCircle24Filled primaryFill="var(--colorPaletteRedForeground1)" />
            )}
            {s.status === "pending" && <span className="iot-solution-step-dot" />}
            <div className="iot-solution-creation-step-content">
              <Text className={`iot-solution-creation-step-label ${s.status}`}>{s.step}</Text>
              {s.error && <Text className="iot-solution-field-error">{s.error}</Text>}
            </div>
          </div>
        ))}
      </div>
    ) : null;

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Create or Reuse Raw Data Tables</h2>
      <Text className="iot-solution-step-description">
        Raw data tables in the <strong>{databaseName}</strong> database will store raw device telemetry, properties,
        and operational data that will be used by dashboards, Activator rules, AI experiences, and other Fabric
        Real-Time Intelligence workloads.
        <br />
        <br />
        You may choose to create new raw data tables for this IoT Hub or combine data from multiple IoT Hubs
        by reusing existing tables.
      </Text>

      {/* Mode selector */}
      <div className="iot-solution-form">
        <RadioGroup
          value={mode}
          onChange={(_, data) => handleModeChange(data.value as "create" | "reuse")}
          disabled={creating || applying}
        >
          <Radio value="create" label="Create new tables" />
          <Radio value="reuse" label="Reuse existing tables" />
        </RadioGroup>
      </div>

      {/* Permission / apply error */}
      {permissionError && (
        <MessageBar intent="error" className="iot-solution-permission-error">
          <MessageBarBody>
            <strong>
              {mode === "reuse" ? "Couldn't reuse tables: " : "Missing API Permission: "}
            </strong>
            {permissionError}
          </MessageBarBody>
        </MessageBar>
      )}

      {mode === "create" ? (
        <div className="iot-solution-form">
          <Text className="iot-solution-resource-name-note">
            Default resource names are provided below. You can customize them if needed.
          </Text>

          {/* Telemetry table name */}
          <div className="iot-solution-field">
            <Label className="iot-solution-field-label" required htmlFor="telemetry-table">
              Telemetry Table Name
            </Label>
            <Input
              id="telemetry-table"
              value={telemetryTableName}
              onChange={(_, data) => handleTelemetryNameChange(data.value)}
              disabled={creating || tablesCreated}
              placeholder="e.g., TelemetryRawIoTHub"
            />
            <Text className="iot-solution-field-hint">
              Stores raw device telemetry messages
            </Text>
          </div>

          {/* Properties table name */}
          <div className="iot-solution-field">
            <Label className="iot-solution-field-label" required htmlFor="properties-table">
              Properties Table Name
            </Label>
            <Input
              id="properties-table"
              value={propertiesTableName}
              onChange={(_, data) => handlePropertiesNameChange(data.value)}
              disabled={creating || tablesCreated}
              placeholder="e.g., PropertiesRawIoTHub"
            />
            <Text className="iot-solution-field-hint">
              Stores device twin property updates
            </Text>
          </div>

          <Button appearance="primary" onClick={createTables} disabled={!canCreate}>
            {creating ? "Creating Tables..." : "Create Tables"}
          </Button>
        </div>
      ) : (
        <div className="iot-solution-form">
          {loadingReuse ? (
            <Spinner size="tiny" label="Finding reusable tables..." />
          ) : reuseError ? (
            <>
              <MessageBar intent="error">
                <MessageBarBody>
                  Unable to retrieve existing telemetry tables. {reuseError}
                </MessageBarBody>
              </MessageBar>
              <div style={{ marginTop: 8 }}>
                <Button size="small" onClick={() => { void loadReusableTables(); }}>
                  Retry
                </Button>
              </div>
            </>
          ) : telemetryOptions.length === 0 && propertiesOptions.length === 0 ? (
            <MessageBar intent="warning">
              <MessageBarBody>
                No existing raw tables fed by an Eventstream custom endpoint were found in{" "}
                <strong>{databaseName}</strong>. Switch to "Create new tables" to set them up.
              </MessageBarBody>
            </MessageBar>
          ) : (
            <>
              <div className="iot-solution-field">
                <Label className="iot-solution-field-label" required htmlFor="reuse-telemetry">
                  Telemetry Table
                </Label>
                <Dropdown
                  id="reuse-telemetry"
                  placeholder="Select a telemetry table"
                  value={selectedTelemetry}
                  selectedOptions={selectedTelemetry ? [selectedTelemetry] : []}
                  onOptionSelect={(_, data) =>
                    handleReuseSelectionChange("telemetry", data.optionValue || "")
                  }
                  disabled={applying || tablesCreated}
                >
                  {telemetryOptions.map((t) => (
                    <Option key={t.tableName} value={t.tableName} text={t.tableName}>
                      <div style={{ display: "flex", flexDirection: "column" }}>
                        <span>{t.tableName}</span>
                        <span
                          style={{
                            fontSize: "var(--fontSizeBase200)",
                            color: "var(--colorNeutralForeground3)",
                          }}
                        >
                          via {t.eventstreamName}
                        </span>
                      </div>
                    </Option>
                  ))}
                </Dropdown>
              </div>

              <div className="iot-solution-field">
                <Label className="iot-solution-field-label" required htmlFor="reuse-properties">
                  Properties Table
                </Label>
                <Dropdown
                  id="reuse-properties"
                  placeholder="Select a properties table"
                  value={selectedProperties}
                  selectedOptions={selectedProperties ? [selectedProperties] : []}
                  onOptionSelect={(_, data) =>
                    handleReuseSelectionChange("properties", data.optionValue || "")
                  }
                  disabled={applying || tablesCreated}
                >
                  {propertiesOptions.map((t) => (
                    <Option key={t.tableName} value={t.tableName} text={t.tableName}>
                      <div style={{ display: "flex", flexDirection: "column" }}>
                        <span>{t.tableName}</span>
                        <span
                          style={{
                            fontSize: "var(--fontSizeBase200)",
                            color: "var(--colorNeutralForeground3)",
                          }}
                        >
                          via {t.eventstreamName}
                        </span>
                      </div>
                    </Option>
                  ))}
                </Dropdown>
              </div>

              <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <Button appearance="primary" onClick={applyReuse} disabled={!canApplyReuse}>
                  {applying ? "Applying..." : "Use selected tables"}
                </Button>
                <Button
                  size="small"
                  appearance="secondary"
                  onClick={() => { void loadReusableTables(); }}
                  disabled={applying || loadingReuse}
                >
                  Refresh
                </Button>
              </div>
            </>
          )}
        </div>
      )}

      {/* Progress */}
      {mode === "create" ? renderProgress(creationSteps) : renderProgress(reuseSteps)}

      {/* Success summary */}
      {tablesCreated && (
        <>
          <div className="iot-solution-success">
            <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
            <Text className="iot-solution-success-text">
              {mode === "reuse" ? "Existing tables selected" : "Tables created successfully"}
            </Text>
          </div>
          <div className="iot-solution-validation">
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Telemetry Table</Text>
              <Text className="iot-solution-validation-value">
                {mode === "reuse" ? selectedTelemetry : telemetryTableName}
              </Text>
            </div>
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Properties Table</Text>
              <Text className="iot-solution-validation-value">
                {mode === "reuse" ? selectedProperties : propertiesTableName}
              </Text>
            </div>
            {mode === "reuse" ? (
              <div className="iot-solution-validation-row">
                <Text className="iot-solution-validation-label">Eventstreams</Text>
                <Text className="iot-solution-validation-value">
                  Reusing existing (creation step is skipped)
                </Text>
              </div>
            ) : (
              <div className="iot-solution-validation-row">
                <Text className="iot-solution-validation-label">Schema</Text>
                <Text className="iot-solution-validation-value">
                  data: string, headers: dynamic
                </Text>
              </div>
            )}
          </div>
        </>
      )}
    </div>
  );
}
