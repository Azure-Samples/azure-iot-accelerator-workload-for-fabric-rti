// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import {
  Dropdown,
  Input,
  Label,
  Option,
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
import { FabricPlatformAPIClient } from "../../../clients/FabricPlatformAPIClient";
import { Item } from "../../../clients/FabricPlatformTypes";
import { getEventhouseItem } from "../eventhouseClient";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import { buildDataAgentContext, buildDataAgentDataSource, dataAgentExamplesMarkdown, defaultDataAgentName, documentedColumns, requireEventTimeColumn, normalizeDocumentedSource, type DataAgentContext } from "../DataAgentContext.ts";
import { parseIotMeta, parseKqlTableResult, quoteKqlIdentifier } from "../ModelAuthoring.ts";
import { DataAgentCreationUnsuccessfulError, resolveCreatedDataAgentId } from "../DataAgentCreation";
import "../IoTSolutionItem.scss";

const FABRIC_WRITE_SCOPE = "https://api.fabric.microsoft.com/Item.ReadWrite.All";
const FABRIC_API_BASE = "https://api.fabric.microsoft.com/v1";
const KUSTO_SCOPE = "https://kusto.kusto.windows.net/.default";

interface CreationStatus {
  step: string;
  status: "pending" | "running" | "success" | "error";
  error?: string;
}

interface TableColumn {
  name: string;
  type: string;
  docstring?: string;
}

/** Filter tables whose docstring contains @iot.meta tableRole=<role>. */
function filterTablesByRole(
  allTables: string[],
  docstrings: Record<string, string>,
  role: string
): string[] {
  return allTables.filter((t) => {
    const doc = docstrings[t];
    if (!doc) return false;
    const meta = parseIotMeta(doc);
    return meta.tableRole === role;
  });
}

interface DataAgentSetupStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
  phase: "source" | "configure";
}

/**
 * Data Agent Setup Step:
 * Creates a Fabric Data Agent backed by a KQL database, pointed at the
 * modeled data table. Generates contextual AI instructions and few-shot
 * examples based on the table schema.
 */
export function DataAgentSetupStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
  phase,
}: DataAgentSetupStepProps) {
  // ---------------------------------------------------------------------------
  // Eventhouse / Database selection
  // ---------------------------------------------------------------------------
  const [eventhouses, setEventhouses] = useState<Item[]>([]);
  const [selectedEventhouseId, setSelectedEventhouseId] = useState<string>(
    wizardContext.agentEhId || ""
  );
  const [selectedEventhouseName, setSelectedEventhouseName] = useState<string>(
    wizardContext.agentEhName || ""
  );
  const [loadingEventhouses, setLoadingEventhouses] = useState(false);
  const [eventhouseError, setEventhouseError] = useState("");

  const [databases, setDatabases] = useState<Item[]>([]);
  const [selectedDatabaseId, setSelectedDatabaseId] = useState<string>(
    wizardContext.agentDbId || ""
  );
  const [selectedDatabaseName, setSelectedDatabaseName] = useState<string>(
    wizardContext.agentDbName || ""
  );
  const [loadingDatabases, setLoadingDatabases] = useState(phase === "source" && !!wizardContext.agentEhId);
  const [databaseError, setDatabaseError] = useState("");

  const [queryServiceUri, setQueryServiceUri] = useState<string>(
    wizardContext.agentQueryServiceUri || ""
  );

  // Table selection
  const [tables, setTables] = useState<string[]>([]);
  const [tableDocstrings, setTableDocstrings] = useState<Record<string, string>>(
    wizardContext.agentTableName && wizardContext.agentTableDocstring
      ? { [wizardContext.agentTableName]: wizardContext.agentTableDocstring }
      : {}
  );
  const [loadingTables, setLoadingTables] = useState(
    phase === "source" && !!wizardContext.agentDbId && !!wizardContext.agentDbName && !!wizardContext.agentQueryServiceUri
  );
  const [tablesLoaded, setTablesLoaded] = useState(false);
  const [selectedTable, setSelectedTable] = useState<string>(
    wizardContext.agentTableName || ""
  );
  const [tableDropdownOpen, setTableDropdownOpen] = useState(false);

  // Schema inspection
  const [columns, setColumns] = useState<TableColumn[]>(
    wizardContext.agentColumns || []
  );
  const [inspectingSchema, setInspectingSchema] = useState(false);
  const [schemaError, setSchemaError] = useState("");

  // Data Agent name
  const suffix = useRef(
    wizardContext.agentSuffix || Math.random().toString(36).slice(2, 6)
  );
  const [agentName, setAgentName] = useState<string>(
    wizardContext.agentName || ""
  );

  // Creation state
  const [creating, setCreating] = useState(false);
  const [permissionError, setPermissionError] = useState("");
  const [creationSteps, setCreationSteps] = useState<CreationStatus[]>([]);
  const [agentCreated, setAgentCreated] = useState(
    !!wizardContext.agentCreated
  );
  const [creationUnconfirmed, setCreationUnconfirmed] = useState(false);
  const creationRequest = useRef<AbortController>();
  useEffect(() => () => creationRequest.current?.abort(),
    [workspaceId, selectedEventhouseId, selectedDatabaseId, selectedDatabaseName, selectedTable, queryServiceUri, agentName, columns, tableDocstrings]);

  const invalidateAgent = useCallback(() => {
    creationRequest.current?.abort();
    creationRequest.current = undefined;
    setCreating(false);
    setCreationUnconfirmed(false);
    setCreationSteps([]);
    updateContext("agentCreating", false);
    resetStepsFrom(stepIndex + 1);
    setAgentCreated(false);
    updateContext("agentCreated", false);
    updateContext("agentId", "");
    updateContext("agentName", "");
    updateContext("agentSourceSelected", false);
    updateContext("agentColumns", []);
    updateContext("agentTableDocstring", "");
  }, [resetStepsFrom, stepIndex, updateContext]);

  // ---------------------------------------------------------------------------
  // Load Eventhouses
  // ---------------------------------------------------------------------------
  const loadEventhouses = useCallback(async () => {
    if (!workspaceId) return;
    setLoadingEventhouses(true);
    setEventhouseError("");
    try {
      const client = FabricPlatformAPIClient.create(workloadClient);
      const result = await client.items.listItems(workspaceId, { type: "Eventhouse" });
      setEventhouses(result.value || []);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setEventhouseError(
        `Unable to load available Eventhouses. Verify your Fabric permissions and try again. Error: ${msg}`
      );
    } finally {
      setLoadingEventhouses(false);
    }
  }, [workloadClient, workspaceId]);

  // ---------------------------------------------------------------------------
  // Load Databases when Eventhouse is selected
  // ---------------------------------------------------------------------------
  const loadDatabases = useCallback(
    async (eventhouseId: string, signal: AbortSignal) => {
      if (!eventhouseId || !workspaceId) return;
      setLoadingDatabases(true);
      setDatabaseError("");
      setDatabases([]);
      try {
        const ehMeta = await getEventhouseItem(workloadClient, workspaceId, eventhouseId);
        if (signal.aborted) return;
        if (!ehMeta) {
          throw new Error(
            "Unable to load information for the selected Eventhouse. Verify that you have access to this Fabric workspace."
          );
        }
        setQueryServiceUri(ehMeta.properties.queryServiceUri);
        updateContext("agentQueryServiceUri", ehMeta.properties.queryServiceUri);

        const dbIds = ehMeta.properties.databasesItemIds || [];
        if (dbIds.length === 0) {
          setDatabaseError(
            "No KQL Database was found in this Eventhouse. Create or select a KQL Database before continuing."
          );
          return;
        }
        const client = FabricPlatformAPIClient.create(workloadClient);
        const allDbs = await client.items.listItems(workspaceId, { type: "KQLDatabase" });
        if (signal.aborted) return;
        const eventhouseDbs = (allDbs.value || []).filter((db) => dbIds.includes(db.id));
        setDatabases(eventhouseDbs);
      } catch (err: unknown) {
        if (signal.aborted) return;
        const msg = err instanceof Error ? err.message : String(err);
        setDatabaseError(
          msg.startsWith("Unable to load information")
            ? msg
            : `Unable to load KQL Databases for the selected Eventhouse. Error: ${msg}`
        );
      } finally {
        if (!signal.aborted) setLoadingDatabases(false);
      }
    },
    [workloadClient, workspaceId, updateContext]
  );

  // ---------------------------------------------------------------------------
  // Load Tables when Database is selected
  // ---------------------------------------------------------------------------
  const loadTables = useCallback(
    async (dbName: string, qsUri: string, signal: AbortSignal) => {
      if (!dbName || !qsUri) return;
      setLoadingTables(true);
      setTablesLoaded(false);
      setTables([]);
      setTableDocstrings({});
      setPermissionError("");
      setSchemaError("");
      try {
        const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
        if (signal.aborted) return;
        const response = await fetch(`${qsUri}/v1/rest/mgmt`, {
          method: "POST",
          signal,
          headers: {
            Authorization: `Bearer ${accessToken.token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({ db: dbName, csl: ".show tables details | project TableName, DocString" }),
        });
        if (!response.ok) {
          throw new Error(
            `Unable to load available tables from the selected database. Error: (${response.status})`
          );
        }
        const result = await response.json();
        if (signal.aborted) return;
        const rows = result?.Tables?.[0]?.Rows || result?.[0]?.Rows || [];
        const tableNames: string[] = [];
        const docMap: Record<string, string> = {};
        for (const r of rows) {
          const name = Array.isArray(r) ? r[0] : r.TableName;
          const doc = Array.isArray(r) ? r[1] || "" : r.DocString || "";
          if (name) {
            tableNames.push(name);
            if (doc) { parseIotMeta(doc); docMap[name] = doc; }
          }
        }
        const modeledTables = filterTablesByRole(tableNames, docMap, "modeled-data");
        setTables(tableNames);
        setTableDocstrings(docMap);
        setTablesLoaded(true);
        if (selectedTable && !modeledTables.includes(selectedTable)) {
          setSelectedTable("");
          setColumns([]);
          updateContext("agentTableName", "");
          updateContext("agentTableDocstring", "");
          updateContext("agentSourceSelected", false);
          updateContext("agentColumns", []);
        }
      } catch (err: unknown) {
        if (signal.aborted) return;
        setTables([]);
        const msg = err instanceof Error ? err.message : String(err);
        const errorMessage = msg.startsWith("Unable to load available tables")
            ? msg
            : `Unable to load available tables from the selected database. Error: ${msg}`;
        setPermissionError(errorMessage);
        setSchemaError(errorMessage);
      } finally {
        if (!signal.aborted) setLoadingTables(false);
      }
    },
    [workloadClient, selectedTable, updateContext]
  );

  // Auto-load eventhouses on mount
  useEffect(() => {
    if (phase === "source" && eventhouses.length === 0 && !loadingEventhouses) {
      loadEventhouses();
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-load databases when eventhouse changes
  useEffect(() => {
    if (phase !== "source" || !selectedEventhouseId) return undefined;
    const abort = new AbortController();
    void loadDatabases(selectedEventhouseId, abort.signal);
    return () => abort.abort();
  }, [selectedEventhouseId]); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-load tables when database changes
  useEffect(() => {
    if (phase !== "source" || !selectedDatabaseId || !selectedDatabaseName || !queryServiceUri) return undefined;
    const abort = new AbortController();
    void loadTables(selectedDatabaseName, queryServiceUri, abort.signal);
    return () => abort.abort();
  }, [selectedDatabaseId, selectedDatabaseName, queryServiceUri]); // eslint-disable-line react-hooks/exhaustive-deps

  // Generate default agent name when table is selected
  useEffect(() => {
    if (phase === "configure" && selectedTable && !agentCreated && !agentName) {
      const doc = tableDocstrings[selectedTable] || "";
      const name = defaultDataAgentName(doc, selectedTable, suffix.current);
      setAgentName(name);
      updateContext("agentName", name);
      updateContext("agentSuffix", suffix.current);
    }
  }, [selectedTable]); // eslint-disable-line react-hooks/exhaustive-deps

  // ---------------------------------------------------------------------------
  // Inspect table schema when table is selected
  // ---------------------------------------------------------------------------
  useEffect(() => {
    if (
      phase !== "source" ||
      !queryServiceUri ||
      !selectedDatabaseName ||
      !selectedTable
    ) {
      return undefined;
    }

    const abort = new AbortController();
    const inspect = async () => {
      setInspectingSchema(true);
      setSchemaError("");
      updateContext("agentSourceSelected", false);
      try {
        const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
        if (abort.signal.aborted) return;
        const resp = await fetch(
          `${queryServiceUri}/v1/rest/mgmt`,
          {
            method: "POST",
            signal: abort.signal,
            headers: {
              Authorization: `Bearer ${accessToken.token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              db: selectedDatabaseName,
              csl: `.show table ${quoteKqlIdentifier(selectedTable)} schema as json`,
            }),
          }
        );
        if (!resp.ok) {
          throw new Error(`Unable to inspect the modeled data table (${resp.status}).`);
        }
        const result = parseKqlTableResult(await resp.json());
        const schemaIndex = result.Columns.findIndex(column => column.ColumnName === "Schema");
        const schemaJson = result.Rows[0]?.[schemaIndex];
        if (typeof schemaJson !== "string") {
          throw new Error("The selected table schema could not be retrieved.");
        }
        const schema = JSON.parse(schemaJson);
        const cols = documentedColumns(schema);
        requireEventTimeColumn(cols);
        if (abort.signal.aborted) return;
        setColumns(cols);
        updateContext("agentColumns", cols);
        updateContext("agentTableDocstring", tableDocstrings[selectedTable] || "");
        updateContext("agentSourceSelected", true);
      } catch (err: unknown) {
        if (abort.signal.aborted) return;
        setColumns([]);
        setSchemaError(err instanceof Error ? err.message : String(err));
        updateContext("agentColumns", []);
        updateContext("agentSourceSelected", false);
      } finally {
        if (!abort.signal.aborted) setInspectingSchema(false);
      }
    };

    void inspect();
    return () => abort.abort();
  }, [
    phase,
    queryServiceUri,
    selectedDatabaseName,
    selectedTable,
    tableDocstrings,
    workloadClient,
    updateContext,
  ]);

  // ---------------------------------------------------------------------------
  // Build AI instructions from schema
  // ---------------------------------------------------------------------------
  const documentedSource = useMemo(() => {
    if (!selectedTable || columns.length === 0) return undefined;
    return normalizeDocumentedSource({
      database: selectedDatabaseName, table: selectedTable, docstring: tableDocstrings[selectedTable] || "",
      columns,
    });
  }, [selectedTable, selectedDatabaseName, columns, tableDocstrings]);
  const contextAnalysis = useMemo<{ value?: DataAgentContext; error?: string }>(() => {
    try { return { value: documentedSource ? buildDataAgentContext(documentedSource) : undefined, error: undefined }; }
    catch (error: unknown) { return { value: undefined, error: error instanceof Error ? error.message : String(error) }; }
  }, [documentedSource]);
  useEffect(() => {
    if (contextAnalysis.error) {
      setSchemaError(contextAnalysis.error);
      if (phase === "configure") setPermissionError(contextAnalysis.error);
      updateContext("agentSourceSelected", false);
    }
  }, [contextAnalysis.error, phase, updateContext]);
  const sourceContext = contextAnalysis.value;
  const aiInstructions = sourceContext?.sourceInstructions || "";

  // ---------------------------------------------------------------------------
  // Build few-shot examples from schema
  // ---------------------------------------------------------------------------
  const fewShots = useMemo(() => {
    return (sourceContext?.examples || []).map(example => ({ id: crypto.randomUUID(), ...example }));
  }, [sourceContext]);

  // ---------------------------------------------------------------------------
  // Create Data Agent
  // ---------------------------------------------------------------------------
  const createAgent = useCallback(async () => {
    if (!workspaceId || !selectedDatabaseId || !selectedTable || !agentName.trim() || !sourceContext || !documentedSource ||
        creating || agentCreated || creationUnconfirmed || creationRequest.current) return;
    const controller = new AbortController();
    creationRequest.current = controller;
    const isCurrent = () => creationRequest.current === controller && !controller.signal.aborted;
    setCreating(true);
    setAgentCreated(false);
    updateContext("agentCreated", false);
    updateContext("agentId", "");
    updateContext("agentCreating", true);
    setPermissionError("");
    setCreationSteps([]);

    // Acquire Fabric write token
    let token: string;
    try {
      const silent = await workloadClient.auth.acquireFrontendAccessToken({
        scopes: [FABRIC_WRITE_SCOPE],
      });
      token = silent.token;
    } catch {
      if (!isCurrent()) return;
      try {
        await workloadClient.auth.acquireAccessToken({
          additionalScopesToConsent: [FABRIC_WRITE_SCOPE],
        });
        const retry = await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [FABRIC_WRITE_SCOPE],
        });
        token = retry.token;
      } catch {
        if (isCurrent()) {
          setPermissionError(
            "Unable to authenticate with Microsoft Fabric. Verify application permissions and try again."
          );
          setCreating(false);
          updateContext("agentCreating", false);
          creationRequest.current = undefined;
        }
        return;
      }
    }
    if (!isCurrent()) return;

    const steps: CreationStatus[] = [
      { step: "Build Data Agent definition", status: "pending" },
      { step: `Create Data Agent "${agentName}"`, status: "pending" },
    ];
    setCreationSteps([...steps]);

    // Step 1: Build definition
    steps[0].status = "running";
    setCreationSteps([...steps]);

    const dsName = `kusto-${selectedDatabaseName}`;

    const datasourceConfig = buildDataAgentDataSource(documentedSource, workspaceId, selectedDatabaseId, crypto.randomUUID());

    // Build fewshots.json
    const fewShotsConfig = {
      $schema:
        "https://developer.microsoft.com/json-schemas/fabric/item/dataAgent/definition/fewShots/1.0.0/schema.json",
      fewShots: fewShots,
    };

    // Build stage_config.json
    const stageConfig = {
      $schema:
        "https://developer.microsoft.com/json-schemas/fabric/item/dataAgent/definition/stageConfiguration/1.0.0/schema.json",
      aiInstructions: sourceContext.agentInstructions,
    };

    // Build data_agent.json
    const dataAgentConfig = {
      $schema:
        "https://developer.microsoft.com/json-schemas/fabric/item/dataAgent/definition/dataAgent/2.1.0/schema.json",
    };

    // Build publish_info.json
    const publishInfoConfig = {
      $schema:
        "https://developer.microsoft.com/json-schemas/fabric/item/dataAgent/definition/publishInfo/1.0.0/schema.json",
      description: "",
    };

    // Base64-encode helper
    const toB64 = (obj: object) =>
      btoa(unescape(encodeURIComponent(JSON.stringify(obj, null, 2))));

    const definition = {
      parts: [
        {
          path: "Files/Config/data_agent.json",
          payload: toB64(dataAgentConfig),
          payloadType: "InlineBase64",
        },
        {
          path: "Files/Config/draft/stage_config.json",
          payload: toB64(stageConfig),
          payloadType: "InlineBase64",
        },
        {
          path: `Files/Config/draft/${dsName}/datasource.json`,
          payload: toB64(datasourceConfig),
          payloadType: "InlineBase64",
        },
        {
          path: `Files/Config/draft/${dsName}/fewshots.json`,
          payload: toB64(fewShotsConfig),
          payloadType: "InlineBase64",
        },
        {
          path: "Files/Config/publish_info.json",
          payload: toB64(publishInfoConfig),
          payloadType: "InlineBase64",
        },
        {
          path: "Files/Config/published/stage_config.json",
          payload: toB64(stageConfig),
          payloadType: "InlineBase64",
        },
        {
          path: `Files/Config/published/${dsName}/datasource.json`,
          payload: toB64(datasourceConfig),
          payloadType: "InlineBase64",
        },
        {
          path: `Files/Config/published/${dsName}/fewshots.json`,
          payload: toB64(fewShotsConfig),
          payloadType: "InlineBase64",
        },
      ],
    };

    steps[0].status = "success";
    setCreationSteps([...steps]);

    // Step 2: Create Data Agent via REST API
    steps[1].status = "running";
    setCreationSteps([...steps]);

    let outcomeUnconfirmed = true;
    setCreationUnconfirmed(true);
    try {
      const resp = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/dataAgents`,
        {
          method: "POST",
          signal: controller.signal,
          redirect: "manual",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            displayName: agentName,
            description: sourceContext.agentDescription,
            definition,
          }),
        }
      );

      if (!isCurrent()) return;
      outcomeUnconfirmed = resp.status === 0 || resp.status === 200 || resp.status === 201 || resp.status === 202 ||
        [301, 302, 303, 307, 308].includes(resp.status);
      setCreationUnconfirmed(outcomeUnconfirmed);
      const agentId = await resolveCreatedDataAgentId(resp, workspaceId, agentName, token, controller.signal);
      if (!isCurrent()) return;
      setCreationUnconfirmed(false);
      steps[1].status = "success";
      setCreationSteps([...steps]);
      setAgentCreated(true);
      updateContext("agentId", agentId);
      updateContext("agentCreated", true);
      updateContext("agentName", agentName);
      updateContext("agentEhId", selectedEventhouseId);
      updateContext("agentEhName", selectedEventhouseName);
      updateContext("agentDbId", selectedDatabaseId);
      updateContext("agentDbName", selectedDatabaseName);
      updateContext("agentTableName", selectedTable);
      updateContext("agentQueryServiceUri", queryServiceUri);
    } catch (err: unknown) {
      if (!isCurrent()) return;
      if (err instanceof DataAgentCreationUnsuccessfulError) {
        outcomeUnconfirmed = false;
        setCreationUnconfirmed(false);
      }
      steps[1].status = "error";
      const message = err instanceof Error ? err.message : String(err);
      steps[1].error = outcomeUnconfirmed && !message.includes("before creating another")
        ? `${message} Verify the existing operation or agent in Fabric before creating another.` : message;
      setCreationSteps([...steps]);
    } finally {
      if (isCurrent()) {
        setCreating(false);
        updateContext("agentCreating", false);
        creationRequest.current = undefined;
      }
    }
  }, [
    workspaceId,
    selectedEventhouseId,
    selectedEventhouseName,
    selectedDatabaseId,
    selectedDatabaseName,
    selectedTable,
    queryServiceUri,
    agentName,
    columns,
    tableDocstrings,
    aiInstructions,
    sourceContext,
    documentedSource,
    fewShots,
    workloadClient,
    updateContext,
    creating,
    agentCreated,
    creationUnconfirmed,
  ]);

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------
  const canCreate = !!(
    workspaceId &&
    selectedEventhouseId &&
    selectedDatabaseId &&
    selectedTable &&
    agentName.trim() &&
    columns.length > 0 &&
    !!sourceContext &&
    !creating &&
    !creationUnconfirmed &&
    !agentCreated
  );

  if (phase === "source") {
    return (
      <div className="iot-solution-step">
        <h2 className="iot-solution-step-title">Select modeled data source</h2>
        <Text className="iot-solution-step-description">
          Select the modeled data table that will ground the agent's answers. The
          table contains typed real-time device telemetry enriched with current
          device state from modeled properties.
        </Text>
        <MessageBar intent="info" style={{ marginTop: "12px" }}>
          <MessageBarBody>
            Modeled data gives the agent consistent field names, data types,
            descriptions, units, and device-property context, helping it generate
            more relevant queries and better operational answers.
          </MessageBarBody>
        </MessageBar>

        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="agent-eh">
            Eventhouse
          </Label>
          {loadingEventhouses ? (
            <Spinner size="tiny" label="Loading Eventhouses..." />
          ) : eventhouseError ? (
            <Text className="iot-solution-error-text">{eventhouseError}</Text>
          ) : (
            <Dropdown
              id="agent-eh"
              placeholder="Select an Eventhouse"
              value={selectedEventhouseName}
              selectedOptions={selectedEventhouseId ? [selectedEventhouseId] : []}
              onOptionSelect={(_e, data) => {
                const eh = eventhouses.find((item) => item.id === data.optionValue);
                if (!eh) return;
                if (eh.id === selectedEventhouseId) return;
                invalidateAgent();
                setLoadingDatabases(true);
                setDatabaseError("");
                setLoadingTables(false);
                setTablesLoaded(false);
                setPermissionError("");
                setSchemaError("");
                setSelectedEventhouseId(eh.id);
                setSelectedEventhouseName(eh.displayName);
                setSelectedDatabaseId("");
                setSelectedDatabaseName("");
                setSelectedTable("");
                setColumns([]);
                setDatabases([]);
                setTables([]);
                setQueryServiceUri("");
                updateContext("agentEhId", eh.id);
                updateContext("agentEhName", eh.displayName);
                updateContext("agentDbId", "");
                updateContext("agentDbName", "");
                updateContext("agentTableName", "");
                updateContext("agentQueryServiceUri", "");
              }}
            >
              {eventhouses.map((eh) => (
                <Option key={eh.id} value={eh.id} text={eh.displayName}>
                  {eh.displayName}
                </Option>
              ))}
            </Dropdown>
          )}
        </div>

        {selectedEventhouseId && (
          <div className="iot-solution-field">
            <Label className="iot-solution-field-label" required htmlFor="agent-db">
              KQL database
            </Label>
            {loadingDatabases ? (
              <Spinner size="tiny" label="Loading databases..." />
            ) : databaseError ? (
              <Text className="iot-solution-error-text">{databaseError}</Text>
            ) : (
              <Dropdown
                id="agent-db"
                placeholder="Select a database"
                value={selectedDatabaseName}
                selectedOptions={selectedDatabaseId ? [selectedDatabaseId] : []}
                onOptionSelect={(_e, data) => {
                  const db = databases.find((item) => item.id === data.optionValue);
                  if (!db) return;
                  if (db.id === selectedDatabaseId) return;
                  invalidateAgent();
                  setLoadingTables(true);
                  setTablesLoaded(false);
                  setPermissionError("");
                  setSchemaError("");
                  setSelectedDatabaseId(db.id);
                  setSelectedDatabaseName(db.displayName);
                  setSelectedTable("");
                  setColumns([]);
                  setTables([]);
                  updateContext("agentDbId", db.id);
                  updateContext("agentDbName", db.displayName);
                  updateContext("agentTableName", "");
                }}
              >
                {databases.map((db) => (
                  <Option key={db.id} value={db.id} text={db.displayName}>
                    {db.displayName}
                  </Option>
                ))}
              </Dropdown>
            )}
          </div>
        )}

        {selectedDatabaseId && (
          <div className="iot-solution-field">
            <Label className="iot-solution-field-label" required htmlFor="agent-table">
              Modeled data table
            </Label>
            {loadingTables ? (
              <Spinner size="tiny" label="Loading tables..." />
            ) : tablesLoaded && filterTablesByRole(tables, tableDocstrings, "modeled-data").length > 0 ? (
              <Dropdown
                id="agent-table"
                placeholder="Select a modeled data table"
                value={selectedTable}
                selectedOptions={selectedTable ? [selectedTable] : []}
                open={tableDropdownOpen}
                onOpenChange={(_e, data) => {
                  setTableDropdownOpen(data.open);
                }}
                onOptionSelect={(_e, data) => {
                  const table = data.optionValue as string;
                  invalidateAgent();
                  setSelectedTable(table);
                  setTableDropdownOpen(false);
                  setColumns([]);
                  updateContext("agentTableName", table);
                  updateContext("agentTableDocstring", tableDocstrings[table] || "");
                }}
              >
                {filterTablesByRole(tables, tableDocstrings, "modeled-data").map((table) => (
                  <Option key={table} value={table}>{table}</Option>
                ))}
              </Dropdown>
            ) : tablesLoaded ? (
              <MessageBar intent="warning">
                <MessageBarBody>
                  No modeled data tables were found in this database. Create modeled
                  data first, then return to set up the Data Agent.
                </MessageBarBody>
              </MessageBar>
            ) : null}
          </div>
        )}

        {selectedTable && inspectingSchema && (
          <Spinner size="small" label="Inspecting modeled fields..." />
        )}
        {schemaError && (
          <MessageBar intent="error">
            <MessageBarBody>{schemaError}</MessageBarBody>
          </MessageBar>
        )}
        {columns.length > 0 && (
          <div className="iot-solution-validation">
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Selected source</Text>
              <Text className="iot-solution-validation-value">
                {selectedDatabaseName} / {selectedTable}
              </Text>
            </div>
            <div className="iot-solution-validation-row">
              <Text className="iot-solution-validation-label">Modeled fields</Text>
              <Text className="iot-solution-validation-value">
                {columns.length} typed columns available to the agent
              </Text>
            </div>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Configure Data Agent</h2>
      <Text className="iot-solution-step-description">
        Name the agent and review the generated guidance that combines your
        modeled schema, semantic metadata, query rules, and example questions.
      </Text>

      {permissionError && (
        <MessageBar intent="error" className="iot-solution-permission-error">
          <MessageBarBody>{permissionError}</MessageBarBody>
        </MessageBar>
      )}

      <div className="iot-solution-validation">
        <div className="iot-solution-validation-row">
          <Text className="iot-solution-validation-label">Grounding data</Text>
          <Text className="iot-solution-validation-value">
            {selectedDatabaseName} / {selectedTable}
          </Text>
        </div>
        <div className="iot-solution-validation-row">
          <Text className="iot-solution-validation-label">Available fields</Text>
          <Text className="iot-solution-validation-value">
            {columns.length} modeled columns
          </Text>
        </div>
      </div>

      {agentCreated && (
        <Text className="iot-solution-resource-name-note">
          Default resource names are provided below. You can customize them if needed.
        </Text>
      )}

      <div className="iot-solution-field">
        <Label className="iot-solution-field-label" required htmlFor="agent-name">
          Data Agent name
        </Label>
        <Input
          id="agent-name"
          value={agentName}
          onChange={(_, data) => {
            setAgentName(data.value);
            updateContext("agentName", data.value);
          }}
          disabled={creating || agentCreated}
          placeholder="e.g., DataAgent_MyModel_abc1"
        />
      </div>

      {aiInstructions && !agentCreated && (
        <div className="iot-solution-field">
          <Label className="iot-solution-field-label">
            AI instructions (auto-generated)
          </Label>
          <Text className="iot-solution-field-hint">
            These instructions become part of the Data Agent and teach it how to
            interpret the modeled fields, distinguish telemetry from device state,
            apply time and device filters, and generate appropriate KQL. The
            included example questions and queries provide additional patterns for
            producing relevant answers.
          </Text>
          <pre
            style={{
              fontSize: "11px",
              background: "var(--colorNeutralBackground3)",
              padding: "12px",
              borderRadius: "6px",
              maxHeight: "300px",
              overflow: "auto",
              whiteSpace: "pre-wrap",
              wordBreak: "break-word",
            }}
          >
            {(sourceContext?.agentInstructions || "") + "\n\n" + aiInstructions}
            {fewShots.length > 0 && (
              `\n## Examples (${fewShots.length})\n\n` +
              dataAgentExamplesMarkdown(fewShots)
            )}
          </pre>
        </div>
      )}

      {!agentCreated && (
        <Button
          appearance="primary"
          onClick={createAgent}
          disabled={!canCreate}
          className="iot-solution-create-button"
        >
          {creating ? <Spinner size="tiny" label="Creating..." /> : "Create Data Agent"}
        </Button>
      )}

      {creationSteps.length > 0 && (
        <div className="iot-solution-creation-steps">
          {creationSteps.map((status, index) => (
            <div key={index} className="iot-solution-creation-step">
              {status.status === "running" && <Spinner size="tiny" />}
              {status.status === "success" && (
                <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
              )}
              {status.status === "error" && (
                <DismissCircle24Filled primaryFill="var(--colorPaletteRedForeground1)" />
              )}
              {status.status === "pending" && <span className="iot-solution-step-dot" />}
              <div className="iot-solution-creation-step-content">
                <Text className={`iot-solution-creation-step-label ${status.status}`}>
                  {status.step}
                </Text>
                {status.error && (
                  <Text className="iot-solution-error-text">{status.error}</Text>
                )}
              </div>
            </div>
          ))}
        </div>
      )}

      {agentCreated && (
        <div className="iot-solution-success">
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">
            Data Agent created. Select <strong>Next</strong> for usage guidance
            and example operational questions.
          </Text>
        </div>
      )}
    </div>
  );
}
