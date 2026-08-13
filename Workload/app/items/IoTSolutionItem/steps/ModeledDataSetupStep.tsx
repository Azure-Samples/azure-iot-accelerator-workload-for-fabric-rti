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
import { DtdlCapability } from "../DtdlModelParser";
import { RAW_EVENT_NORMALIZATION_KQL } from "../RawEventKql";
import "../IoTSolutionItem.scss";

const FABRIC_ITEM_READ_SCOPE = "https://api.fabric.microsoft.com/Item.Read.All";
const KUSTO_SCOPE = "https://kusto.kusto.windows.net/.default";

function generateSuffix(): string {
  return Math.random().toString(36).substring(2, 6);
}

function sanitize(name: string): string {
  return name.replace(/[^a-zA-Z0-9_]/g, "_");
}

interface CreationStatus {
  step: string;
  status: "pending" | "running" | "success" | "error";
  error?: string;
}

interface ModeledDataSetupStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
  phase: "source" | "create";
}

/**
 * Modeled Data Setup Step:
 *
 * User selects Eventhouse → Database → raw telemetry/properties tables via dropdowns,
 * then creates the KQL entities that transform raw data into modeled form:
 *
 * 1. PropertiesNormalized table + update policy (entity-value rows)
 * 2. PropertiesLKV materialized view using arg_max (last 30 days)
 * 3. ModeledData table with typed columns + update policy (telemetry + LKV join)
 */

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

/** Escape a string for use inside a double-quoted KQL string literal. */
function toKqlString(value: string): string {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/"/g, '\\"')
    .replace(/\r/g, "")
    .replace(/\n/g, "\\n");
}

/** One human/agent readable line describing a modeled-data column and its device-model metadata. */
function formatColumnDoc(c: DtdlCapability): string {
  // Collapse any internal whitespace/newlines so each column stays on a single docstring line.
  const oneLine = (s: string) => s.replace(/\s+/g, " ").trim();
  const attrs = [c.kind.toLowerCase(), c.kustoType];
  if (c.unit) attrs.push(`unit: ${oneLine(c.unit)}`);
  if (c.component) attrs.push(`component: ${oneLine(c.component)}`);
  const displayName = c.displayName ? oneLine(c.displayName) : "";
  const label = displayName && displayName !== c.name ? ` — "${displayName}"` : "";
  const desc = c.description ? `: ${oneLine(c.description)}` : "";
  return `- ${c.name}${label} (${attrs.join(", ")})${desc}`;
}

/**
 * Build a column-reference section for the modeled-data table docstring from the parsed
 * device model. Captures per-column display names, units, component and descriptions when
 * the DTDL model declares them, so downstream tools (e.g. the Data Agent) can describe the
 * columns accurately. Returns "" when there are no columns.
 */
function buildColumnDocSection(
  telemetries: DtdlCapability[],
  properties: DtdlCapability[]
): string {
  if (telemetries.length === 0 && properties.length === 0) return "";
  const lines: string[] = [
    "Columns (extracted from the device model):",
    "- deviceId (string): device identifier from the IoT Hub connection.",
    "- enqueuedTime (datetime): time the message was enqueued by IoT Hub.",
  ];
  for (const t of telemetries) lines.push(formatColumnDoc(t));
  for (const p of properties) lines.push(formatColumnDoc(p));
  return lines.join("\n");
}

export function ModeledDataSetupStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
  phase,
}: ModeledDataSetupStepProps) {
  const modelName = sanitize(wizardContext.modelName || "Model");
  const suffix = useMemo(() => generateSuffix(), []);

  const telemetries: DtdlCapability[] = (wizardContext.modelTelemetries || []).filter(
    (t: DtdlCapability) => t.included
  );
  const properties: DtdlCapability[] = (wizardContext.modelProperties || []).filter(
    (p: DtdlCapability) => p.included
  );

  // ---------------------------------------------------------------------------
  // Eventhouse / Database selection
  // ---------------------------------------------------------------------------
  const [eventhouses, setEventhouses] = useState<Item[]>([]);
  const [selectedEventhouseId, setSelectedEventhouseId] = useState<string>(
    wizardContext.modelEhId || ""
  );
  const [selectedEventhouseName, setSelectedEventhouseName] = useState<string>(
    wizardContext.modelEhName || ""
  );
  const [loadingEventhouses, setLoadingEventhouses] = useState(false);
  const [eventhouseError, setEventhouseError] = useState("");

  const [databases, setDatabases] = useState<Item[]>([]);
  const [selectedDatabaseId, setSelectedDatabaseId] = useState<string>(
    wizardContext.modelDbId || ""
  );
  const [selectedDatabaseName, setSelectedDatabaseName] = useState<string>(
    wizardContext.modelDbName || ""
  );
  const [loadingDatabases, setLoadingDatabases] = useState(false);
  const [databaseError, setDatabaseError] = useState("");

  const [queryServiceUri, setQueryServiceUri] = useState<string>(
    wizardContext.modelQueryServiceUri || ""
  );

  // Table list + selection
  const [tables, setTables] = useState<string[]>([]);
  const [tableDocstrings, setTableDocstrings] = useState<Record<string, string>>({});
  const [loadingTables, setLoadingTables] = useState(false);
  const [showAllTelemetryTables, setShowAllTelemetryTables] = useState(false);
  const [showAllPropertiesTables, setShowAllPropertiesTables] = useState(false);
  const [telemetryDropdownOpen, setTelemetryDropdownOpen] = useState(false);
  const [propertiesDropdownOpen, setPropertiesDropdownOpen] = useState(false);
  const skipTelemetryClose = useRef(false);
  const skipPropertiesClose = useRef(false);

  const [rawTelemetryTable, setRawTelemetryTable] = useState<string>(
    wizardContext.rawTelemetryTable || ""
  );
  const [rawPropertiesTable, setRawPropertiesTable] = useState<string>(
    wizardContext.rawPropertiesTable || ""
  );

  // Scoping clauses (optional filtering for raw tables)
  const [telemetryScopeClause, setTelemetryScopeClause] = useState<string>(
    wizardContext.telemetryScopeClause || ""
  );
  const [propertiesScopeClause, setPropertiesScopeClause] = useState<string>(
    wizardContext.propertiesScopeClause || ""
  );
  const [telemetryScopeValid, setTelemetryScopeValid] = useState<boolean | null>(
    wizardContext.telemetryScopeClause?.trim() && wizardContext.modelRawDataSelected
      ? true
      : null
  ); // null=not validated, true=ok, false=error
  const [propertiesScopeValid, setPropertiesScopeValid] = useState<boolean | null>(
    wizardContext.propertiesScopeClause?.trim() && wizardContext.modelRawDataSelected
      ? true
      : null
  );
  const [telemetryScopeError, setTelemetryScopeError] = useState("");
  const [propertiesScopeError, setPropertiesScopeError] = useState("");
  const [validatingTelemetryScope, setValidatingTelemetryScope] = useState(false);
  const [validatingPropertiesScope, setValidatingPropertiesScope] = useState(false);
  const [telemetryScopePreview, setTelemetryScopePreview] = useState<Record<string, unknown>[]>([]);
  const [propertiesScopePreview, setPropertiesScopePreview] = useState<Record<string, unknown>[]>([]);

  // Output names
  const [propsNormalizedName, setPropsNormalizedName] = useState<string>(
    wizardContext.propsNormalizedName || `PropertiesNormalized_${modelName}_${suffix}`
  );
  const [propsLkvViewName, setPropsLkvViewName] = useState<string>(
    wizardContext.propsLkvViewName || `PropertiesLKV_${modelName}_${suffix}`
  );
  const [modeledDataName, setModeledDataName] = useState<string>(
    wizardContext.modeledDataName || `ModeledData_${modelName}_${suffix}`
  );

  // Creation state
  const [creating, setCreating] = useState(false);
  const [permissionError, setPermissionError] = useState("");
  const [creationSteps, setCreationSteps] = useState<CreationStatus[]>([]);
  const [tablesCreated, setTablesCreated] = useState(
    !!wizardContext.modeledTablesCreated
  );

  // ---------------------------------------------------------------------------
  // Fabric token helper
  // ---------------------------------------------------------------------------
  const ensureFabricApiToken = useCallback(async (): Promise<boolean> => {
    setPermissionError("");
    try {
      await workloadClient.auth.acquireFrontendAccessToken({
        scopes: [FABRIC_ITEM_READ_SCOPE],
      });
      return true;
    } catch {
      try {
        await workloadClient.auth.acquireAccessToken({
          additionalScopesToConsent: [FABRIC_ITEM_READ_SCOPE],
        });
        await workloadClient.auth.acquireFrontendAccessToken({
          scopes: [FABRIC_ITEM_READ_SCOPE],
        });
        return true;
      } catch (interactiveErr: unknown) {
        const errCode = (interactiveErr as { error?: number })?.error;
        if (errCode === 2) {
          setPermissionError(
            "Unable to access Fabric resources. Verify that the required Fabric permissions have been granted to the application."
          );
        } else if (errCode === 1) {
          setPermissionError(
            "Access to Fabric was not granted. Please complete the sign-in and consent process to continue."
          );
        }
        return false;
      }
    }
  }, [workloadClient]);

  // ---------------------------------------------------------------------------
  // Load eventhouses
  // ---------------------------------------------------------------------------
  const loadEventhouses = useCallback(async () => {
    if (!workspaceId) return;
    setLoadingEventhouses(true);
    setEventhouseError("");
    const hasToken = await ensureFabricApiToken();
    if (!hasToken) { setLoadingEventhouses(false); return; }
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
  }, [workloadClient, workspaceId, ensureFabricApiToken]);

  // ---------------------------------------------------------------------------
  // Load databases for selected eventhouse
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
        updateContext("modelQueryServiceUri", ehMeta.properties.queryServiceUri);

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
  // Load table names from a KQL database
  // ---------------------------------------------------------------------------
  const loadTables = useCallback(
    async (dbName: string, qsUri: string) => {
      if (!dbName || !qsUri) return;
      setLoadingTables(true);
      setTables([]);
      setTableDocstrings({});
      setShowAllTelemetryTables(false);
      setShowAllPropertiesTables(false);
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
            docMap[name] = doc;
          }
        }
        setTables(tableNames);
        setTableDocstrings(docMap);
      } catch (err: unknown) {
        console.error("Failed to list tables:", err);
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
    [workloadClient]
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

  // ---------------------------------------------------------------------------
  // Selection handlers
  // ---------------------------------------------------------------------------
  const handleEventhouseSelect = (_: any, data: any) => {
    const eh = eventhouses.find((e) => e.id === data.optionValue);
    if (eh) {
      setSelectedEventhouseId(eh.id);
      setSelectedEventhouseName(eh.displayName);
      setSelectedDatabaseId("");
      setSelectedDatabaseName("");
      setQueryServiceUri("");
      setTables([]);
      setRawTelemetryTable("");
      setRawPropertiesTable("");
      resetCreation();
      updateContext("modelEhId", eh.id);
      updateContext("modelEhName", eh.displayName);
      updateContext("modelDbId", "");
      updateContext("modelDbName", "");
      updateContext("modelQueryServiceUri", "");
      updateContext("rawTelemetryTable", "");
      updateContext("rawPropertiesTable", "");
    }
  };

  const handleDatabaseSelect = (_: any, data: any) => {
    const db = databases.find((d) => d.id === data.optionValue);
    if (db) {
      setSelectedDatabaseId(db.id);
      setSelectedDatabaseName(db.displayName);
      setRawTelemetryTable("");
      setRawPropertiesTable("");
      resetCreation();
      updateContext("modelDbId", db.id);
      updateContext("modelDbName", db.displayName);
      updateContext("rawTelemetryTable", "");
      updateContext("rawPropertiesTable", "");
    }
  };

  const handleTelemetryTableSelect = (_: any, data: any) => {
    setRawTelemetryTable(data.optionText || "");
    resetCreation();
    updateContext("rawTelemetryTable", data.optionText || "");
  };

  const handlePropertiesTableSelect = (_: any, data: any) => {
    setRawPropertiesTable(data.optionText || "");
    resetCreation();
    updateContext("rawPropertiesTable", data.optionText || "");
  };

  const resetCreation = () => {
    resetStepsFrom(phase === "source" ? stepIndex + 1 : stepIndex);
    setTablesCreated(false);
    updateContext("modeledTablesCreated", false);
    if (phase === "source") {
      updateContext("modelCoverageVerified", false);
      updateContext("modelObservedTelemetryFields", []);
      updateContext("modelObservedPropertyFields", []);
    }
    setCreationSteps([]);
  };

  // ---------------------------------------------------------------------------
  // Execute a KQL management command
  // ---------------------------------------------------------------------------
  const executeKqlCommand = async (
    token: string,
    db: string,
    csl: string
  ): Promise<void> => {
    const response = await fetch(`${queryServiceUri}/v1/rest/mgmt`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ db, csl }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      let detail = "";
      try {
        const parsed = JSON.parse(errorText);
        detail =
          parsed?.error?.["@message"] ||
          parsed?.error?.message ||
          errorText;
      } catch {
        detail = errorText;
      }
      throw new Error(
        `Failed to create modeled data assets. Verify database permissions and configuration. (${response.status}): ${detail}`
      );
    }
  };

  // ---------------------------------------------------------------------------
  // Validate a scoping clause by running a preview query against the raw table
  // ---------------------------------------------------------------------------
  const validateScopeClause = async (
    tableName: string,
    clause: string,
    kind: "telemetry" | "properties"
  ) => {
    const setValidating = kind === "telemetry" ? setValidatingTelemetryScope : setValidatingPropertiesScope;
    const setValid = kind === "telemetry" ? setTelemetryScopeValid : setPropertiesScopeValid;
    const setError = kind === "telemetry" ? setTelemetryScopeError : setPropertiesScopeError;
    const setPreview = kind === "telemetry" ? setTelemetryScopePreview : setPropertiesScopePreview;

    setValidating(true);
    setError("");
    setPreview([]);

    try {
      const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
      const kql = `${tableName}\n${RAW_EVENT_NORMALIZATION_KQL}\n| where ${clause} | take 10`;
      const response = await fetch(`${queryServiceUri}/v1/rest/query`, {
        method: "POST",
        headers: {
          Authorization: `Bearer ${accessToken.token}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ db: selectedDatabaseName, csl: kql }),
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
        setValid(false);
        setError(detail);
        return;
      }

      const json = await response.json();
      const primaryTable = json?.Tables?.[0];
      if (!primaryTable) {
        setValid(true);
        return;
      }
      const cols: string[] = primaryTable.Columns?.map((c: { ColumnName: string }) => c.ColumnName) || [];
      const rows: Record<string, unknown>[] = (primaryTable.Rows || []).map((row: unknown[]) => {
        const obj: Record<string, unknown> = {};
        cols.forEach((col, i) => { obj[col] = row[i]; });
        return obj;
      });
      setPreview(rows);
      setValid(true);
    } catch (err: unknown) {
      setValid(false);
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setValidating(false);
    }
  };

  // ---------------------------------------------------------------------------
  // Build KQL commands
  // ---------------------------------------------------------------------------
  const buildCommands = useCallback((): { label: string; csl: string }[] => {
    const commands: { label: string; csl: string }[] = [];

    // Whether the model has any component capabilities (column names like `<component>_<leaf>`).
    // Component data arrives differently: telemetry as separate messages tagged with IoTSubject,
    // and reported properties nested under the component key with a `__t: "c"` marker.
    const hasComponents = [...telemetries, ...properties].some((c) => c.component);

    // --- 1. PropertiesNormalized table + inline update policy ---
    commands.push({
      label: `Create table ${propsNormalizedName}`,
      csl: `.create table ${propsNormalizedName} (deviceId: string, propertyName: string, propertyValue: dynamic, enqueuedTime: datetime)`,
    });

    const propsScopeFilter = propertiesScopeClause.trim()
      ? ` | where ${propertiesScopeClause.trim()}`
      : "";
    const propsBase = `${rawPropertiesTable}\n${RAW_EVENT_NORMALIZATION_KQL}${propsScopeFilter} | extend rp = data.properties.reported | extend keys = bag_keys(rp) | mv-expand key = keys to typeof(string) | where key !startswith "$" and key != "iothub-enqueuedtime" and key != "iothub-connection-device-id"`;
    // When the model has components, flatten each component object (marked with __t="c") into
    // `<component>_<subProperty>` rows so they match the modeled component columns. Non-component
    // object-valued properties (e.g. Object-schema properties, which have no __t) are kept whole.
    const propsQuery = hasComponents
      ? `${propsBase} | extend val = rp[key] | extend isComponent = (gettype(val) == "dictionary" and tostring(val["__t"]) == "c") | mv-expand subkey = iff(isComponent, bag_keys(val), pack_array(key)) to typeof(string) | where subkey != "__t" | project deviceId = tostring(headers.IoTConnectionDeviceId), propertyName = iff(isComponent, strcat(key, "_", subkey), key), propertyValue = iff(isComponent, val[subkey], val), enqueuedTime = todatetime(headers.IoTEnqueueTime)`
      : `${propsBase} | project deviceId = tostring(headers.IoTConnectionDeviceId), propertyName = key, propertyValue = rp[key], enqueuedTime = todatetime(headers.IoTEnqueueTime)`;

    commands.push({
      label: `Set update policy on ${propsNormalizedName}`,
      csl: buildUpdatePolicyCmd(propsNormalizedName, rawPropertiesTable, propsQuery),
    });

    // --- 2. PropertiesLKV materialized view ---
    commands.push({
      label: `Create materialized view ${propsLkvViewName}`,
      csl: `.create materialized-view ${propsLkvViewName} on table ${propsNormalizedName} {\n` +
        `  ${propsNormalizedName}\n` +
        `  | where ingestion_time() > ago(30d)\n` +
        `  | summarize arg_max(enqueuedTime, propertyValue) by deviceId, propertyName\n` +
        `}`,
    });

    // --- 3. ModeledData table with typed columns + inline update policy ---
    const columns: string[] = [
      "deviceId: string",
      "enqueuedTime: datetime",
    ];
    for (const t of telemetries) {
      columns.push(`${t.name}: ${t.kustoType}`);
    }
    for (const p of properties) {
      columns.push(`${p.name}: ${p.kustoType}`);
    }

    commands.push({
      label: `Create table ${modeledDataName}`,
      csl: `.create table ${modeledDataName} (${columns.join(", ")})`,
    });

    // Build the inline modeled data query.
    // Group telemetries by component (undefined = root interface). Component telemetry arrives as
    // separate messages tagged with IoTSubject=<component>; we pick each source's payload via a
    // subject-guarded bag before extracting typed columns. This prevents cross-component key
    // collisions (e.g. two components both reporting a "current" field).
    const telGroups = new Map<string, DtdlCapability[]>();
    for (const t of telemetries) {
      const key = t.component || "";
      const arr = telGroups.get(key);
      if (arr) arr.push(t);
      else telGroups.set(key, [t]);
    }
    const telBagVar = (comp: string): string =>
      comp ? `__telC_${sanitize(comp)}` : "__telRoot";

    const sourceBagExprs: string[] = [];
    const componentTelProjExprs: string[] = [];
    for (const [comp, caps] of telGroups.entries()) {
      const bagVar = telBagVar(comp);
      const cond = comp ? `subject == '${comp}'` : `subject == ''`;
      sourceBagExprs.push(`${bagVar} = iff(${cond}, telemetry, dynamic(null))`);
      for (const t of caps) {
        const castFn = kustoTypeToCast(t.kustoType);
        componentTelProjExprs.push(`${t.name} = ${castFn}(${bagVar}['${t.leafName}'])`);
      }
    }
    // Legacy (no components): extract each column directly from the telemetry payload.
    const telemetryProjections = telemetries.map((t) => {
      const castFn = kustoTypeToCast(t.kustoType);
      return `${t.name} = ${castFn}(telemetry['${t.leafName}'])`;
    }).join(", ");

    const propertyProjections = properties.map((p) => {
      const castFn = kustoTypeToCast(p.kustoType);
      return `${p.name} = ${castFn}(bag["${p.name}"])`;
    });
    const modeledColumnNames = [
      "deviceId",
      "enqueuedTime",
      ...telemetries.map((t) => t.name),
      ...properties.map((p) => p.name),
    ];

    const telScopeFilter = telemetryScopeClause.trim()
      ? ` | where ${telemetryScopeClause.trim()}`
      : "";
    let modelQuery: string;
    let modeledSourceTable: string;
    if (telemetries.length > 0) {
      modeledSourceTable = rawTelemetryTable;
      modelQuery =
        `${properties.length > 0
          ? `let lkv_bag = ${propsLkvViewName} | summarize bag = make_bag(bag_pack(propertyName, propertyValue)) by deviceId; `
          : ""}` +
        `${rawTelemetryTable}\n${RAW_EVENT_NORMALIZATION_KQL}${telScopeFilter} | extend deviceId = tostring(headers.IoTConnectionDeviceId), enqueuedTime = todatetime(headers.IoTEnqueueTime), telemetry = data`;

      if (hasComponents) {
        modelQuery += `, subject = tostring(headers.IoTSubject)`;
        modelQuery += ` | extend ${sourceBagExprs.join(", ")}`;
        modelQuery += ` | extend ${componentTelProjExprs.join(", ")}`;
      } else {
        modelQuery += ` | extend ${telemetryProjections}`;
      }

      if (properties.length > 0) {
        modelQuery +=
          ` | lookup kind=leftouter lkv_bag on deviceId` +
          ` | extend ${propertyProjections.join(", ")}`;
      }
      modelQuery += ` | project ${modeledColumnNames.join(", ")}`;
    } else {
      // ModeledData always represents telemetry events. A property-only model still creates the
      // same schema and enrichment policy, but the table remains empty when no telemetry arrives.
      modeledSourceTable = rawTelemetryTable;
      modelQuery =
        `let lkv_bag = ${propsLkvViewName} | summarize bag = make_bag(bag_pack(propertyName, propertyValue)) by deviceId; ` +
        `${rawTelemetryTable}\n${RAW_EVENT_NORMALIZATION_KQL}${telScopeFilter}` +
        ` | extend deviceId = tostring(headers.IoTConnectionDeviceId), enqueuedTime = todatetime(headers.IoTEnqueueTime)` +
        ` | lookup kind=leftouter lkv_bag on deviceId` +
        ` | extend ${propertyProjections.join(", ")}` +
        ` | project ${modeledColumnNames.join(", ")}`;
    }

    commands.push({
      label: `Set update policy on ${modeledDataName}`,
      csl: buildUpdatePolicyCmd(modeledDataName, modeledSourceTable, modelQuery),
    });

    // --- 4. Set docstrings with @iot.meta tags ---
    const clusterName = queryServiceUri.replace(/^https?:\/\//, "").replace(/\/$/, "");
    const dbName = selectedDatabaseName;

    commands.push({
      label: `Set docstring on ${propsNormalizedName}`,
      csl: `.alter table ${propsNormalizedName} docstring "Normalized device property changes in Entity-Attribute-Value (EAV) format\\n@iot.meta tableRole=properties-eav\\n@iot.meta source=${clusterName}/${dbName}/${rawPropertiesTable}\\n@iot.meta deviceModel=${modelName}"`,
    });

    commands.push({
      label: `Set docstring on ${propsLkvViewName}`,
      csl: `.alter materialized-view ${propsLkvViewName} docstring "Property last known values\\n@iot.meta viewRole=properties-lkv\\n@iot.meta source=${clusterName}/${dbName}/${propsNormalizedName}\\n@iot.meta deviceModel=${modelName}"`,
    });

    commands.push({
      label: `Set docstring on ${modeledDataName}`,
      csl: `.alter table ${modeledDataName} docstring "${toKqlString(
        [
          "Modeled device telemetry and properties",
          "@iot.meta tableRole=modeled-data",
          "@iot.meta rowKind=telemetry-event",
          `@iot.meta source=${clusterName}/${dbName}/${rawTelemetryTable}; ${clusterName}/${dbName}/${propsLkvViewName}`,
          `@iot.meta deviceModel=${modelName}`,
          ...(buildColumnDocSection(telemetries, properties)
            ? ["", buildColumnDocSection(telemetries, properties)]
            : []),
        ].join("\n")
      )}"`,
    });

    return commands;
  }, [
    rawTelemetryTable,
    rawPropertiesTable,
    telemetryScopeClause,
    propertiesScopeClause,
    propsNormalizedName,
    propsLkvViewName,
    modeledDataName,
    telemetries,
    properties,
    queryServiceUri,
    selectedDatabaseName,
    modelName,
  ]);

  // ---------------------------------------------------------------------------
  // Execute creation
  // ---------------------------------------------------------------------------
  const createAll = useCallback(async () => {
    if (!queryServiceUri || !selectedDatabaseName) return;

    setCreating(true);
    updateContext("modeledDataCreating", true);
    setPermissionError("");
    setCreationSteps([]);

    let token: string;
    try {
      const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
      token = accessToken.token;
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setPermissionError(`Unable to access the selected KQL Database. ${msg}`);
      setCreating(false);
      updateContext("modeledDataCreating", false);
      return;
    }

    const commands = buildCommands();
    const statusArr: CreationStatus[] = commands.map((c) => ({
      step: c.label,
      status: "pending" as const,
    }));
    setCreationSteps([...statusArr]);

    let allSucceeded = true;

    for (let i = 0; i < commands.length; i++) {
      statusArr[i].status = "running";
      setCreationSteps([...statusArr]);

      try {
        await executeKqlCommand(token, selectedDatabaseName, commands[i].csl);
        statusArr[i].status = "success";
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        statusArr[i].status = "error";
        statusArr[i].error = msg;
        allSucceeded = false;
        setCreationSteps([...statusArr]);
        break;
      }

      setCreationSteps([...statusArr]);
    }

    if (allSucceeded) {
      setTablesCreated(true);
      updateContext("modeledTablesCreated", true);
      updateContext("propsNormalizedName", propsNormalizedName);
      updateContext("propsLkvViewName", propsLkvViewName);
      updateContext("modeledDataName", modeledDataName);
      updateContext("rawTelemetryTable", rawTelemetryTable);
      updateContext("rawPropertiesTable", rawPropertiesTable);
      updateContext("telemetryScopeClause", telemetryScopeClause);
      updateContext("propertiesScopeClause", propertiesScopeClause);
      updateContext("modelQueryServiceUri", queryServiceUri);
      updateContext("modelDbId", selectedDatabaseId);
      updateContext("modelDbName", selectedDatabaseName);
      updateContext("modelEhId", selectedEventhouseId);
      updateContext("modelEhName", selectedEventhouseName);
    }

    setCreating(false);
    updateContext("modeledDataCreating", false);
  }, [
    queryServiceUri,
    selectedDatabaseId,
    selectedDatabaseName,
    selectedEventhouseId,
    selectedEventhouseName,
    workloadClient,
    buildCommands,
    updateContext,
    propsNormalizedName,
    propsLkvViewName,
    modeledDataName,
    rawTelemetryTable,
    rawPropertiesTable,
  ]);

  const telScopeOk = !telemetryScopeClause.trim() || telemetryScopeValid === true;
  const propsScopeOk = !propertiesScopeClause.trim() || propertiesScopeValid === true;
  const sourceReady = !!(
    selectedEventhouseId &&
    selectedDatabaseId &&
    selectedDatabaseName &&
    queryServiceUri &&
    rawTelemetryTable &&
    rawPropertiesTable &&
    telScopeOk &&
    propsScopeOk
  );

  useEffect(() => {
    if (phase === "source") {
      updateContext("modelRawDataSelected", sourceReady);
    }
  }, [phase, sourceReady, updateContext]);

  const canCreate = !!(
    rawTelemetryTable.trim() &&
    rawPropertiesTable.trim() &&
    propsNormalizedName.trim() &&
    propsLkvViewName.trim() &&
    modeledDataName.trim() &&
    queryServiceUri &&
    selectedDatabaseName &&
    telScopeOk &&
    propsScopeOk &&
    !creating &&
    !tablesCreated
  );

  // ---------------------------------------------------------------------------
  // Render
  // ---------------------------------------------------------------------------
  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">
        {phase === "source" ? "Select Raw Data Source" : "Create Modeled Data"}
      </h2>
      <Text className="iot-solution-step-description">
        {phase === "source"
          ? "Select the raw telemetry and property tables that contain data for this device model. Optional scoping clauses can limit the modeled data to matching devices or events."
          : telemetries.length > 0
            ? "Create the tables, materialized view, and update policies that convert raw JSON IoT device data into typed columns and enrich telemetry events with the latest device state from twin properties."
            : "Create the normalized property assets and a modeled table that enriches telemetry events with typed property state. If this device type sends no telemetry, the modeled table remains empty."}
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
        {phase === "source" && (
          <>
        {/* Eventhouse selection */}
        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="model-eh-dropdown">
            Eventhouse
          </Label>
          {loadingEventhouses ? (
            <Spinner size="tiny" label="Loading Eventhouses..." />
          ) : eventhouseError ? (
            <MessageBar intent="error"><MessageBarBody>{eventhouseError}</MessageBarBody></MessageBar>
          ) : (
            <Dropdown
              id="model-eh-dropdown"
              placeholder="Select an Eventhouse"
              value={selectedEventhouseName}
              selectedOptions={selectedEventhouseId ? [selectedEventhouseId] : []}
              onOptionSelect={handleEventhouseSelect}
              disabled={creating}
            >
              {eventhouses.map((eh) => (
                <Option key={eh.id} value={eh.id}>{eh.displayName}</Option>
              ))}
            </Dropdown>
          )}
        </div>

        {/* Database selection */}
        {selectedEventhouseId && (
          <div className="iot-solution-field">
            <Label className="iot-solution-field-label" required htmlFor="model-db-dropdown">
              Database
            </Label>
            {loadingDatabases ? (
              <Spinner size="tiny" label="Loading databases..." />
            ) : databaseError ? (
              <MessageBar intent="error"><MessageBarBody>{databaseError}</MessageBarBody></MessageBar>
            ) : (
              <Dropdown
                id="model-db-dropdown"
                placeholder="Select a database"
                value={selectedDatabaseName}
                selectedOptions={selectedDatabaseId ? [selectedDatabaseId] : []}
                onOptionSelect={handleDatabaseSelect}
                disabled={creating}
              >
                {databases.map((db) => (
                  <Option key={db.id} value={db.id}>{db.displayName}</Option>
                ))}
              </Dropdown>
            )}
          </div>
        )}

        {/* Raw table selection */}
        {selectedDatabaseName && (
          <>
            <Text weight="semibold" size={400} block style={{ marginTop: "12px" }}>
              Source Tables
            </Text>
            <Text size={200} style={{ color: "var(--colorNeutralForeground3)", marginBottom: "8px" }} block>
              Select the raw telemetry and property tables created in the ingestion
              experience. These source tables remain unchanged and continue to
              contain the complete raw JSON data.
            </Text>

            <div className="iot-solution-field">
              <Label className="iot-solution-field-label" required htmlFor="raw-tel-dropdown">
                Raw Telemetry Table
              </Label>
              {loadingTables ? (
                <Spinner size="tiny" label="Loading tables..." />
              ) : (
                <Dropdown
                  id="raw-tel-dropdown"
                  placeholder="Select telemetry table"
                  value={rawTelemetryTable}
                  selectedOptions={rawTelemetryTable ? [rawTelemetryTable] : []}
                  open={telemetryDropdownOpen}
                  onOpenChange={(_e, data) => {
                    if (!data.open && skipTelemetryClose.current) {
                      skipTelemetryClose.current = false;
                      return;
                    }
                    setTelemetryDropdownOpen(data.open);
                  }}
                  onOptionSelect={(_e, data) => {
                    if (data.optionValue === "__show_all_telemetry__") {
                      skipTelemetryClose.current = true;
                      setShowAllTelemetryTables(true);
                      return;
                    }
                    handleTelemetryTableSelect(_e, data);
                    setTelemetryDropdownOpen(false);
                  }}
                  disabled={creating}
                >
                  {(showAllTelemetryTables
                    ? tables
                    : filterTablesByRole(tables, tableDocstrings, "iothub-telemetry-raw").length > 0
                      ? filterTablesByRole(tables, tableDocstrings, "iothub-telemetry-raw")
                      : tables
                  ).map((t) => (
                    <Option key={t} value={t}>{t}</Option>
                  ))}
                  {!showAllTelemetryTables && filterTablesByRole(tables, tableDocstrings, "iothub-telemetry-raw").length > 0 && filterTablesByRole(tables, tableDocstrings, "iothub-telemetry-raw").length < tables.length && (
                    <Option key="__show_all_telemetry__" value="__show_all_telemetry__" text="">
                      <span style={{ color: "var(--colorBrandForeground1)", fontSize: "12px" }}>Show all tables...</span>
                    </Option>
                  )}
                </Dropdown>
              )}
            </div>

            {/* Telemetry scoping clause */}
            {rawTelemetryTable && (
              <div className="iot-solution-field" style={{ marginLeft: "8px", borderLeft: "2px solid var(--colorNeutralStroke2)", paddingLeft: "12px" }}>
                <Label className="iot-solution-field-label" htmlFor="tel-scope">
                  Telemetry Scoping Clause (optional)
                </Label>
                <Text size={200} style={{ color: "var(--colorNeutralForeground3)", marginBottom: "4px" }} block>
                  Filter raw telemetry rows to only include devices matching this model.
                  Reference <code>headers</code> and <code>data</code> (parsed as JSON) fields.
                </Text>
                <Text size={200} style={{ color: "var(--colorNeutralForeground4)", fontStyle: "italic", marginBottom: "8px" }} block>
                  Examples: <code>headers.IoTConnectionDeviceId startswith "thermostat"</code> · <code>data.deviceType == "hvac"</code>
                </Text>
                <Input
                  id="tel-scope"
                  value={telemetryScopeClause}
                  onChange={(_, d) => {
                    setTelemetryScopeClause(d.value);
                    setTelemetryScopeValid(null);
                    setTelemetryScopeError("");
                    setTelemetryScopePreview([]);
                    updateContext("telemetryScopeClause", d.value);
                    resetCreation();
                  }}
                  placeholder={'e.g., headers.IoTConnectionDeviceId startswith "thermostat"'}
                  disabled={creating}
                  style={{ fontFamily: "monospace", fontSize: "12px" }}
                />
                <div style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "6px" }}>
                  <Button
                    size="small"
                    appearance={telemetryScopeValid === true ? "subtle" : "primary"}
                    onClick={() => validateScopeClause(rawTelemetryTable, telemetryScopeClause, "telemetry")}
                    disabled={!telemetryScopeClause.trim() || validatingTelemetryScope || creating}
                  >
                    {validatingTelemetryScope ? <Spinner size="tiny" label="Validating..." /> : telemetryScopeValid === true ? "✓ Re-validate" : "Validate"}
                  </Button>
                  {telemetryScopeValid === true && (
                    <Text size={200} style={{ color: "var(--colorPaletteGreenForeground1)" }}>
                      Valid — {telemetryScopePreview.length} row{telemetryScopePreview.length !== 1 ? "s" : ""} returned
                    </Text>
                  )}
                  {telemetryScopeValid === false && (
                    <Text size={200} style={{ color: "var(--colorPaletteRedForeground1)" }}>
                      Failed: {telemetryScopeError}
                    </Text>
                  )}
                </div>
              </div>
            )}

            <div className="iot-solution-field">
              <Label className="iot-solution-field-label" required htmlFor="raw-prop-dropdown">
                Raw Properties Table
              </Label>
              {loadingTables ? (
                <Spinner size="tiny" label="Loading tables..." />
              ) : (
                <Dropdown
                  id="raw-prop-dropdown"
                  placeholder="Select properties table"
                  value={rawPropertiesTable}
                  selectedOptions={rawPropertiesTable ? [rawPropertiesTable] : []}
                  open={propertiesDropdownOpen}
                  onOpenChange={(_e, data) => {
                    if (!data.open && skipPropertiesClose.current) {
                      skipPropertiesClose.current = false;
                      return;
                    }
                    setPropertiesDropdownOpen(data.open);
                  }}
                  onOptionSelect={(_e, data) => {
                    if (data.optionValue === "__show_all_properties__") {
                      skipPropertiesClose.current = true;
                      setShowAllPropertiesTables(true);
                      return;
                    }
                    handlePropertiesTableSelect(_e, data);
                    setPropertiesDropdownOpen(false);
                  }}
                  disabled={creating}
                >
                  {(showAllPropertiesTables
                    ? tables
                    : filterTablesByRole(tables, tableDocstrings, "iothub-properties-raw").length > 0
                      ? filterTablesByRole(tables, tableDocstrings, "iothub-properties-raw")
                      : tables
                  ).map((t) => (
                    <Option key={t} value={t}>{t}</Option>
                  ))}
                  {!showAllPropertiesTables && filterTablesByRole(tables, tableDocstrings, "iothub-properties-raw").length > 0 && filterTablesByRole(tables, tableDocstrings, "iothub-properties-raw").length < tables.length && (
                    <Option key="__show_all_properties__" value="__show_all_properties__" text="">
                      <span style={{ color: "var(--colorBrandForeground1)", fontSize: "12px" }}>Show all tables...</span>
                    </Option>
                  )}
                </Dropdown>
              )}
            </div>

            {/* Properties scoping clause */}
            {rawPropertiesTable && (
              <div className="iot-solution-field" style={{ marginLeft: "8px", borderLeft: "2px solid var(--colorNeutralStroke2)", paddingLeft: "12px" }}>
                <Label className="iot-solution-field-label" htmlFor="prop-scope">
                  Properties Scoping Clause (optional)
                </Label>
                <Text size={200} style={{ color: "var(--colorNeutralForeground3)", marginBottom: "4px" }} block>
                  Filter raw property rows to only include devices matching this model.
                  Reference <code>headers</code> and <code>data</code> (parsed as JSON) fields.
                </Text>
                <Text size={200} style={{ color: "var(--colorNeutralForeground4)", fontStyle: "italic", marginBottom: "8px" }} block>
                  Examples: <code>headers.IoTConnectionDeviceId startswith "thermostat"</code> · <code>data.properties.reported.deviceModel == "ThermoPro X200"</code>
                </Text>
                <Input
                  id="prop-scope"
                  value={propertiesScopeClause}
                  onChange={(_, d) => {
                    setPropertiesScopeClause(d.value);
                    setPropertiesScopeValid(null);
                    setPropertiesScopeError("");
                    setPropertiesScopePreview([]);
                    updateContext("propertiesScopeClause", d.value);
                    resetCreation();
                  }}
                  placeholder={'e.g., headers.IoTConnectionDeviceId startswith "thermostat"'}
                  disabled={creating}
                  style={{ fontFamily: "monospace", fontSize: "12px" }}
                />
                <div style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "6px" }}>
                  <Button
                    size="small"
                    appearance={propertiesScopeValid === true ? "subtle" : "primary"}
                    onClick={() => validateScopeClause(rawPropertiesTable, propertiesScopeClause, "properties")}
                    disabled={!propertiesScopeClause.trim() || validatingPropertiesScope || creating}
                  >
                    {validatingPropertiesScope ? <Spinner size="tiny" label="Validating..." /> : propertiesScopeValid === true ? "✓ Re-validate" : "Validate"}
                  </Button>
                  {propertiesScopeValid === true && (
                    <Text size={200} style={{ color: "var(--colorPaletteGreenForeground1)" }}>
                      Valid — {propertiesScopePreview.length} row{propertiesScopePreview.length !== 1 ? "s" : ""} returned
                    </Text>
                  )}
                  {propertiesScopeValid === false && (
                    <Text size={200} style={{ color: "var(--colorPaletteRedForeground1)" }}>
                      Failed: {propertiesScopeError}
                    </Text>
                  )}
                </div>
              </div>
            )}
          </>
        )}
          </>
        )}

        {/* Output entity names */}
        {phase === "create" && rawTelemetryTable && rawPropertiesTable && (
          <>
            <div className="iot-solution-validation">
              <div className="iot-solution-validation-row">
                <Text className="iot-solution-validation-label">Raw telemetry source</Text>
                <Text className="iot-solution-validation-value">{rawTelemetryTable}</Text>
              </div>
              <div className="iot-solution-validation-row">
                <Text className="iot-solution-validation-label">Raw property source</Text>
                <Text className="iot-solution-validation-value">{rawPropertiesTable}</Text>
              </div>
              <div className="iot-solution-validation-row">
                <Text className="iot-solution-validation-label">Selected model fields</Text>
                <Text className="iot-solution-validation-value">
                  {telemetries.length} telemetry and {properties.length} properties
                </Text>
              </div>
            </div>
            <Text weight="semibold" size={400} block style={{ marginTop: "12px" }}>
              Modeled Data Outputs
            </Text>
            <MessageBar intent="info">
              <MessageBarBody>
                The normalized table makes property changes easy to query. The
                materialized view maintains each device's latest property state.
                The modeled data table then combines that state with live telemetry
                and exposes selected model fields as properly typed columns.
              </MessageBarBody>
            </MessageBar>
            <Text className="iot-solution-resource-name-note">
              Default resource names are provided below. You can customize them if needed.
            </Text>

            <div className="iot-solution-field">
              <Label className="iot-solution-field-label" required htmlFor="props-normalized">
                Normalized Properties Table
              </Label>
              <Input
                id="props-normalized"
                value={propsNormalizedName}
                onChange={(_, data) => { setPropsNormalizedName(data.value); resetCreation(); }}
                disabled={creating || tablesCreated}
              />
              <Text className="iot-solution-field-hint">
                Converts nested property JSON into one device, property, value, and
                timestamp row per update. This supports property history analysis
                and supplies the last-known-value view.
              </Text>
            </div>

            <div className="iot-solution-field">
              <Label className="iot-solution-field-label" required htmlFor="props-lkv">
                Properties LKV Materialized View
              </Label>
              <Input
                id="props-lkv"
                value={propsLkvViewName}
                onChange={(_, data) => { setPropsLkvViewName(data.value); resetCreation(); }}
                disabled={creating || tablesCreated}
              />
              <Text className="iot-solution-field-hint">
                Maintains the latest reported value for each device property within
                the previous 30 days, representing current device state.
              </Text>
            </div>

            <div className="iot-solution-field">
              <Label className="iot-solution-field-label" required htmlFor="modeled-data">
                Modeled Data Table
              </Label>
              <Input
                id="modeled-data"
                value={modeledDataName}
                onChange={(_, data) => { setModeledDataName(data.value); resetCreation(); }}
                disabled={creating || tablesCreated}
              />
              <Text className="iot-solution-field-hint">
                {telemetries.length > 0
                  ? `Creates ${telemetries.length} typed telemetry columns and enriches each event with ${properties.length} current property columns, plus device ID and event time.`
                  : `Creates a telemetry-event table with ${properties.length} current property columns, plus device ID and event time. The table remains empty if no telemetry events arrive.`}
              </Text>
            </div>

            {/* Create button */}
            <Button
              appearance="primary"
              onClick={createAll}
              disabled={!canCreate}
              style={{ marginTop: "8px" }}
            >
              {creating ? "Creating Modeled Data..." : "Create Modeled Data"}
            </Button>
          </>
        )}
      </div>

      {/* Progress steps */}
      {phase === "create" && creationSteps.length > 0 && (
        <div className="iot-solution-creation-steps">
          {creationSteps.map((s, i) => (
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
                <Text className={`iot-solution-creation-step-label ${s.status}`}>
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

      {phase === "create" && tablesCreated && (
        <div className="iot-solution-success">
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">
            Modeled data created. Select <strong>Next</strong> to review how
            the generated entities work together.
          </Text>
        </div>
      )}

      {phase === "source" && sourceReady && (
        <div className="iot-solution-success">
          <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
          <Text className="iot-solution-success-text">
            IoT data source is ready. Select <strong>Next</strong> to compare
            recent fields with the device model.
          </Text>
        </div>
      )}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function kustoTypeToCast(kustoType: string): string {
  switch (kustoType) {
    case "real": return "todouble";
    case "int": return "toint";
    case "long": return "tolong";
    case "bool": return "tobool";
    case "datetime": return "todatetime";
    case "dynamic": return "todynamic";
    case "string": return "tostring";
    case "date": return "todatetime";
    default: return "tostring";
  }
}

/** Escape a string for embedding inside a JSON string value */
function buildUpdatePolicyCmd(tableName: string, sourceTable: string, query: string): string {
  // Build the policy JSON, then use @"..." KQL verbatim string (escape " as "")
  const policyJson = JSON.stringify([{
    IsEnabled: true,
    Source: sourceTable,
    Query: query,
    IsTransactional: false,
    PropagateIngestionProperties: false,
  }]);
  const escaped = policyJson.replace(/"/g, '""');
  return `.alter table ${tableName} policy update @"${escaped}"`;
}
