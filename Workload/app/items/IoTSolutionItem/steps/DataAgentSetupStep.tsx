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
import {
  buildDataAgentQueryGuidelines,
  buildLatestValueQuery,
} from "../dataAgentQueryGuidance";
import {
  getFabricCreationErrorMessage,
  parseFabricErrorPayload,
} from "../FabricCreationError";
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

/** Per-column metadata captured from the device model, embedded in the modeled-data docstring. */
interface ColumnDoc {
  kind?: "telemetry" | "property";
  displayName?: string;
  unit?: string;
  component?: string;
  description?: string;
}

/**
 * Parse the "Columns (extracted from the device model)" section of a modeled-data table
 * docstring into per-column metadata. The model-setup step writes one line per column as
 * `- <name> — "<displayName>" (<kind>, <type>, unit: <unit>, component: <c>): <description>`;
 * every part except the name is optional.
 */
function parseColumnDocs(docstring: string): Record<string, ColumnDoc> {
  const result: Record<string, ColumnDoc> = {};
  const lineRegex = /^- (\S+?)(?: — "([^"]*)")? \(([^)]*)\)(?::\s*(.*))?$/;
  for (const raw of docstring.split(/\r?\n/)) {
    const m = lineRegex.exec(raw.trim());
    if (!m) continue;
    const [, name, displayName, attrs, description] = m;
    const doc: ColumnDoc = {};
    if (displayName) doc.displayName = displayName;
    if (description) doc.description = description;
    for (const attr of attrs.split(",").map((a) => a.trim())) {
      if (attr === "telemetry" || attr === "property") doc.kind = attr;
      else if (attr.startsWith("unit:")) doc.unit = attr.slice("unit:".length).trim();
      else if (attr.startsWith("component:")) doc.component = attr.slice("component:".length).trim();
    }
    result[name] = doc;
  }
  return result;
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
  const [loadingDatabases, setLoadingDatabases] = useState(false);
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
  const [loadingTables, setLoadingTables] = useState(false);
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

  const invalidateAgent = useCallback(() => {
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
    async (eventhouseId: string) => {
      if (!eventhouseId || !workspaceId) return;
      setLoadingDatabases(true);
      setDatabaseError("");
      setDatabases([]);
      try {
        const ehMeta = await getEventhouseItem(workloadClient, workspaceId, eventhouseId);
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
        const eventhouseDbs = (allDbs.value || []).filter((db) => dbIds.includes(db.id));
        setDatabases(eventhouseDbs);
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

  // ---------------------------------------------------------------------------
  // Load Tables when Database is selected
  // ---------------------------------------------------------------------------
  const loadTables = useCallback(
    async (dbName: string, qsUri: string) => {
      if (!dbName || !qsUri) return;
      setLoadingTables(true);
      setTables([]);
      setTableDocstrings({});
      setPermissionError("");
      try {
        const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
        const response = await fetch(`${qsUri}/v1/rest/mgmt`, {
          method: "POST",
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
        const rows = result?.Tables?.[0]?.Rows || result?.[0]?.Rows || [];
        const tableNames: string[] = [];
        const docMap: Record<string, string> = {};
        for (const r of rows) {
          const name = Array.isArray(r) ? r[0] : r.TableName;
          const doc = Array.isArray(r) ? r[1] || "" : r.DocString || "";
          if (name) {
            tableNames.push(name);
            if (doc) docMap[name] = doc;
          }
        }
        setTables(tableNames);
        setTableDocstrings(docMap);
        const modeledTables = filterTablesByRole(tableNames, docMap, "modeled-data");
        if (selectedTable && !modeledTables.includes(selectedTable)) {
          setSelectedTable("");
          setColumns([]);
          updateContext("agentTableName", "");
          updateContext("agentTableDocstring", "");
          updateContext("agentSourceSelected", false);
          updateContext("agentColumns", []);
        }
      } catch (err: unknown) {
        setTables([]);
        const msg = err instanceof Error ? err.message : String(err);
        setPermissionError(
          msg.startsWith("Unable to load available tables")
            ? msg
            : `Unable to load available tables from the selected database. Error: ${msg}`
        );
      } finally {
        setLoadingTables(false);
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
    if (phase === "source" && selectedEventhouseId) loadDatabases(selectedEventhouseId);
  }, [selectedEventhouseId]); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-load tables when database changes
  useEffect(() => {
    if (phase === "source" && selectedDatabaseName && queryServiceUri) {
      loadTables(selectedDatabaseName, queryServiceUri);
    }
  }, [selectedDatabaseName, queryServiceUri]); // eslint-disable-line react-hooks/exhaustive-deps

  // Generate default agent name when table is selected
  useEffect(() => {
    if (phase === "configure" && selectedTable && !agentCreated && !agentName) {
      const doc = tableDocstrings[selectedTable] || "";
      const meta = parseIotMeta(doc);
      const baseName = meta.deviceModel
        ? meta.deviceModel.replace(/[^a-zA-Z0-9_]/g, "_")
        : selectedTable.replace(/[^a-zA-Z0-9_]/g, "_");
      const name = `DataAgent_${baseName}_${suffix.current}`;
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
      return;
    }

    const inspect = async () => {
      setInspectingSchema(true);
      setSchemaError("");
      updateContext("agentSourceSelected", false);
      try {
        const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
        const resp = await fetch(
          `${queryServiceUri}/v1/rest/query`,
          {
            method: "POST",
            headers: {
              Authorization: `Bearer ${accessToken.token}`,
              "Content-Type": "application/json",
            },
            body: JSON.stringify({
              db: selectedDatabaseName,
              csl: `.show table ${selectedTable} schema as json`,
            }),
          }
        );
        if (!resp.ok) {
          throw new Error(`Unable to inspect the modeled data table (${resp.status}).`);
        }
        const result = await resp.json();
        const schemaJson = result?.Tables?.[0]?.Rows?.[0]?.[1];
        if (!schemaJson) {
          throw new Error("The selected table schema could not be retrieved.");
        }
        const schema = JSON.parse(schemaJson);
        const cols: TableColumn[] = (
          schema?.OrderedColumns || []
        ).map((c: any) => ({
          name: c.Name,
          type: c.CslType || c.Type,
        }));
        if (cols.length === 0) {
          throw new Error("The selected table does not contain any fields.");
        }
        setColumns(cols);
        updateContext("agentColumns", cols);
        updateContext("agentTableDocstring", tableDocstrings[selectedTable] || "");
        updateContext("agentSourceSelected", true);
      } catch (err: unknown) {
        setColumns([]);
        setSchemaError(err instanceof Error ? err.message : String(err));
        updateContext("agentColumns", []);
        updateContext("agentSourceSelected", false);
      } finally {
        setInspectingSchema(false);
      }
    };

    inspect();
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
  const aiInstructions = useMemo(() => {
    if (!selectedTable || columns.length === 0) return "";

    const columnDocs = parseColumnDocs(tableDocstrings[selectedTable] || "");
    const tableMeta = parseIotMeta(tableDocstrings[selectedTable] || "");
    const deviceModel = tableMeta.deviceModel;
    const rowKind = tableMeta.rowKind;

    // Build a rich, human-readable description for a column from the device-model metadata,
    // falling back to a generic label when the model did not declare any.
    const describe = (c: TableColumn, fallback: string): string => {
      const meta = columnDocs[c.name];
      if (!meta) return fallback;
      const main = [meta.displayName, meta.description].filter(Boolean).join(" — ");
      const qualifiers: string[] = [];
      if (meta.unit) qualifiers.push(`unit: ${meta.unit}`);
      if (meta.component) qualifiers.push(`component: ${meta.component}`);
      const suffix = qualifiers.length ? `${main ? " " : ""}(${qualifiers.join(", ")})` : "";
      return `${main}${suffix}` || fallback;
    };

    const telemetryCols = columns.filter(
      (c) =>
        c.name !== "deviceId" &&
        c.name !== "enqueuedTime" &&
        (columnDocs[c.name]?.kind === "telemetry" ||
          (!columnDocs[c.name]?.kind &&
            ["real", "int", "long", "decimal", "double", "bool"].includes(c.type)))
    );
    const propertyCols = columns.filter(
      (c) =>
        c.name !== "deviceId" &&
        c.name !== "enqueuedTime" &&
        (columnDocs[c.name]?.kind === "property" ||
          (!columnDocs[c.name]?.kind && !telemetryCols.find((t) => t.name === c.name)))
    );
    const rowDescription =
      rowKind === "property-update"
        ? "Each row represents a reported property update from an IoT device. Only properties included in that update are populated; other property columns can be null."
        : "Each row represents a telemetry event from an IoT device, enriched with the device's latest selected reported properties.";
    let instructions =
      `You are an IoT data assistant. You help users explore and analyze IoT device data stored in a KQL (Kusto Query Language) database.\n\n` +
      `## Data Model\n\n` +
      (deviceModel ? `This data comes from \`${deviceModel}\` devices. ` : ``) +
      `The main table is \`${selectedTable}\` in the \`${selectedDatabaseName}\` database. ` +
      `${rowDescription}\n\n` +
      `### Key Columns\n` +
      `- \`deviceId\` (string): The unique identifier of the IoT device that produced this reading.\n` +
      `- \`enqueuedTime\` (datetime): The timestamp when the message was received by IoT Hub. Use this as the primary time axis for all time-based queries.\n\n`;

    if (telemetryCols.length > 0) {
      instructions += `### Telemetry Columns (real-time measurements)\n`;
      for (const c of telemetryCols) {
        instructions += `- \`${c.name}\` (${c.type}): ${describe(c, "Device telemetry value.")}\n`;
      }
      instructions += `\n`;
    }

    if (propertyCols.length > 0) {
      instructions += rowKind === "property-update"
        ? `### Property Columns (values included in each reported-property update)\n`
        : `### Property Columns (device-reported metadata, enriched via last-known-value lookup)\n`;
      for (const c of propertyCols) {
        instructions += `- \`${c.name}\` (${c.type}): ${describe(c, "Device-reported property.")}\n`;
      }
      instructions += `\n`;
    }

    instructions += buildDataAgentQueryGuidelines(rowKind);

    return instructions;
  }, [selectedTable, selectedDatabaseName, columns, tableDocstrings]);

  // ---------------------------------------------------------------------------
  // Build few-shot examples from schema
  // ---------------------------------------------------------------------------
  const fewShots = useMemo(() => {
    if (!selectedTable || columns.length === 0) return [];

    const numericCol = columns.find((c) =>
      ["real", "int", "long", "decimal", "double"].includes(c.type)
    );
    const stringCol = columns.find(
      (c) =>
        c.type === "string" &&
        c.name !== "deviceId"
    );

    const examples: { id: string; question: string; query: string }[] = [
      {
        id: crypto.randomUUID(),
        question: "What devices are in the system?",
        query: `${selectedTable} | distinct deviceId`,
      },
      {
        id: crypto.randomUUID(),
        question: "Show me the latest modeled records for a device",
        query: `${selectedTable} | where deviceId == "device-01" | top 10 by enqueuedTime desc`,
      },
    ];

    if (numericCol) {
      examples.push({
        id: crypto.randomUUID(),
        question: `What is the average ${numericCol.name} over the last hour?`,
        query: `${selectedTable} | where enqueuedTime >= ago(1h) | summarize avg(${numericCol.name}) by bin(enqueuedTime, 5m) | render timechart`,
      });
      examples.push({
        id: crypto.randomUUID(),
        question: `Show the latest ${numericCol.name} for each device`,
        query: buildLatestValueQuery(selectedTable, numericCol.name),
      });
    }

    if (stringCol) {
      examples.push({
        id: crypto.randomUUID(),
        question: `Show all devices and their ${stringCol.name}`,
        query: buildLatestValueQuery(selectedTable, stringCol.name),
      });
    }

    examples.push({
      id: crypto.randomUUID(),
      question: "How many readings per device in the last 24 hours?",
      query: `${selectedTable} | where enqueuedTime >= ago(24h) | summarize Count=count() by deviceId | order by Count desc`,
    });

    return examples;
  }, [selectedTable, columns]);

  // ---------------------------------------------------------------------------
  // Create Data Agent
  // ---------------------------------------------------------------------------
  const createAgent = useCallback(async () => {
    if (!workspaceId || !selectedDatabaseId || !selectedTable || !agentName.trim()) return;
    setCreating(true);
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
      try {
        await workloadClient.auth.acquireAccessToken({
          additionalScopesToConsent: [FABRIC_WRITE_SCOPE],
        });
        const retry = await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [FABRIC_WRITE_SCOPE],
        });
        token = retry.token;
      } catch {
        setPermissionError(
          "Unable to authenticate with Microsoft Fabric. Verify application permissions and try again."
        );
        setCreating(false);
        updateContext("agentCreating", false);
        return;
      }
    }

    const steps: CreationStatus[] = [
      { step: "Build Data Agent definition", status: "pending" },
      { step: `Create Data Agent "${agentName}"`, status: "pending" },
    ];
    setCreationSteps([...steps]);

    // Step 1: Build definition
    steps[0].status = "running";
    setCreationSteps([...steps]);

    const dsName = `kusto-${selectedDatabaseName}`;

    // Map KQL/CslType to .NET type names used by the Data Agent API
    const kqlToSystemType = (cslType: string): string => {
      const map: Record<string, string> = {
        string: "System.String",
        bool: "System.SByte",
        datetime: "System.DateTime",
        int: "System.Int32",
        long: "System.Int64",
        real: "System.Double",
        double: "System.Double",
        decimal: "System.Decimal",
        dynamic: "System.Object",
        guid: "System.Guid",
        timespan: "System.TimeSpan",
      };
      return map[cslType] || "System.Object";
    };

    // Build the datasource.json
    const selectedTableMeta = parseIotMeta(tableDocstrings[selectedTable] || "");
    const modeledTableDescription =
      selectedTableMeta.rowKind === "property-update"
        ? "Modeled IoT device property-update events."
        : "Modeled IoT device events with typed model fields and available device state.";
    const tableId = crypto.randomUUID();
    const tableGroupId = crypto.randomUUID();
    const datasourceConfig = {
      $schema:
        "https://developer.microsoft.com/json-schemas/fabric/item/dataAgent/definition/dataSource/1.0.0/schema.json",
      artifactId: selectedDatabaseId,
      workspaceId: workspaceId,
      dataSourceInstructions: null as string | null,
      displayName: selectedDatabaseName,
      type: "kusto",
      userDescription: null as string | null,
      metadata: {},
      elements: [
        {
          id: tableGroupId,
          is_selected: false,
          display_name: "Tables",
          type: "table_grouping",
          description: null as string | null,
          children: [
            {
              id: tableId,
              is_selected: true,
              display_name: selectedTable,
              type: "kusto.table",
              description: `${modeledTableDescription}\n@iot.meta tableRole=modeled-data`,
              children: columns.map((col) => ({
                id: col.name,
                is_selected: true,
                display_name: col.name,
                type: "kusto.column",
                data_type: kqlToSystemType(col.type),
                description:
                  col.name === "deviceId"
                    ? "Unique IoT device identifier"
                    : col.name === "enqueuedTime"
                      ? "Timestamp when the message was received by IoT Hub"
                      : (null as string | null),
                children: [] as never[],
              })),
            },
          ],
        },
      ],
    };

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
      aiInstructions: aiInstructions,
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
      ],
    };

    steps[0].status = "success";
    setCreationSteps([...steps]);

    // Step 2: Create Data Agent via REST API
    steps[1].status = "running";
    setCreationSteps([...steps]);

    try {
      const resp = await fetch(
        `${FABRIC_API_BASE}/workspaces/${workspaceId}/dataAgents`,
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            displayName: agentName,
            description: `AI-powered Data Agent for querying IoT device data from ${selectedTable}`,
            definition,
          }),
        }
      );

      if (resp.status === 201 || resp.status === 200) {
        const result = await resp.json();
        steps[1].status = "success";
        setCreationSteps([...steps]);
        setAgentCreated(true);
        updateContext("agentCreated", true);
        updateContext("agentId", result.id);
        updateContext("agentName", agentName);
        updateContext("agentEhId", selectedEventhouseId);
        updateContext("agentEhName", selectedEventhouseName);
        updateContext("agentDbId", selectedDatabaseId);
        updateContext("agentDbName", selectedDatabaseName);
        updateContext("agentTableName", selectedTable);
        updateContext("agentQueryServiceUri", queryServiceUri);
      } else if (resp.status === 202) {
        // LRO — poll for completion
        const location = resp.headers.get("Location");
        const retryAfter = parseInt(resp.headers.get("Retry-After") || "2", 10);
        let succeeded = false;

        if (location) {
          for (let attempt = 0; attempt < 30; attempt++) {
            await new Promise((r) => setTimeout(r, retryAfter * 1000));
            const pollResp = await fetch(location, {
              headers: { Authorization: `Bearer ${token}` },
            });
            if (pollResp.status === 200) {
              const pollResult = await pollResp.json();
              if (pollResult.status === "Succeeded") {
                succeeded = true;
                const agentId = pollResult.id || pollResult.resourceId || "";
                steps[1].status = "success";
                setCreationSteps([...steps]);
                setAgentCreated(true);
                updateContext("agentCreated", true);
                updateContext("agentId", agentId);
                updateContext("agentName", agentName);
                updateContext("agentEhId", selectedEventhouseId);
                updateContext("agentEhName", selectedEventhouseName);
                updateContext("agentDbId", selectedDatabaseId);
                updateContext("agentDbName", selectedDatabaseName);
                updateContext("agentTableName", selectedTable);
                updateContext("agentQueryServiceUri", queryServiceUri);
                break;
              }
              if (pollResult.status === "Failed") {
                throw new Error(getFabricCreationErrorMessage(
                  pollResp.status,
                  pollResult,
                  "Data Agent",
                  agentName
                ));
              }
            } else {
              const errorPayload = parseFabricErrorPayload(await pollResp.text());
              throw new Error(getFabricCreationErrorMessage(
                pollResp.status,
                errorPayload,
                "Data Agent",
                agentName
              ));
            }
          }
          if (!succeeded) {
            throw new Error(
              "Data Agent creation is taking longer than expected. Refresh and verify whether the agent was created successfully."
            );
          }
        } else {
          throw new Error(
            "Data Agent creation is taking longer than expected. Refresh and verify whether the agent was created successfully."
          );
        }
      } else {
        const errorPayload = parseFabricErrorPayload(await resp.text());
        throw new Error(getFabricCreationErrorMessage(
          resp.status,
          errorPayload,
          "Data Agent",
          agentName
        ));
      }
    } catch (err: unknown) {
      steps[1].status = "error";
      steps[1].error = err instanceof Error ? err.message : String(err);
      setCreationSteps([...steps]);
    }

    setCreating(false);
    updateContext("agentCreating", false);
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
    fewShots,
    workloadClient,
    updateContext,
  ]);

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------
  const canCreate = !!(
    selectedEventhouseId &&
    selectedDatabaseId &&
    selectedTable &&
    agentName.trim() &&
    columns.length > 0 &&
    !creating &&
    !agentCreated
  );

  if (phase === "source") {
    return (
      <div className="iot-solution-step">
        <h2 className="iot-solution-step-title">Select Modeled Data Source</h2>
        <Text className="iot-solution-step-description">
          Select the modeled data table that will ground the agent's answers. The
          table contains typed device events and model metadata, with current
          property state when available.
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
                invalidateAgent();
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
              KQL Database
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
                  invalidateAgent();
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
              Modeled Data Table
            </Label>
            {loadingTables ? (
              <Spinner size="tiny" label="Loading tables..." />
            ) : filterTablesByRole(tables, tableDocstrings, "modeled-data").length > 0 ? (
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
            ) : (
              <MessageBar intent="warning">
                <MessageBarBody>
                  No modeled data tables were found in this database. Create modeled
                  data first, then return to set up the Data Agent.
                </MessageBarBody>
              </MessageBar>
            )}
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

      <Text className="iot-solution-resource-name-note">
        Default resource names are provided below. You can customize them if needed.
      </Text>

      <div className="iot-solution-field">
        <Label className="iot-solution-field-label" required htmlFor="agent-name">
          Data Agent Name
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
            AI Instructions (auto-generated)
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
            {aiInstructions}
            {fewShots.length > 0 && (
              `\n## Examples (${fewShots.length})\n\n` +
              fewShots.map((example) => `Q: ${example.question}\n${example.query}`).join("\n\n")
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
