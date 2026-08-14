// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  Button,
  Dropdown,
  Label,
  Option,
  Spinner,
  Text,
  Checkbox,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  ArrowClockwise24Regular,
} from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { FabricPlatformAPIClient } from "../../../clients/FabricPlatformAPIClient";
import { Item } from "../../../clients/FabricPlatformTypes";
import { getEventhouseItem } from "../eventhouseClient";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import { RAW_EVENT_NORMALIZATION_KQL } from "../RawEventKql";
import "../IoTSolutionItem.scss";

const FABRIC_ITEM_READ_SCOPE = "https://api.fabric.microsoft.com/Item.Read.All";
const KUSTO_SCOPE = "https://kusto.kusto.windows.net/.default";

interface DashboardDataStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
}

interface TelemetryField {
  key: string;
  isNumeric: boolean;
  included: boolean;
  component?: string;
  leaf?: string;
}

interface PropertyField {
  key: string;
  included: boolean;
  component?: string;
  leaf?: string;
}

/**
 * Dashboard data source selection and field inspection.
 *
 * User selects an Eventhouse, database, and the raw telemetry/properties tables.
 * Then clicks "Inspect Tables" to extract distinct telemetry names (with numeric flag)
 * and property names via KQL queries. Results are shown as editable lists.
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

export function DashboardDataStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
}: DashboardDataStepProps) {
  // Eventhouse state
  const [eventhouses, setEventhouses] = useState<Item[]>([]);
  const [selectedEventhouseId, setSelectedEventhouseId] = useState<string>(
    wizardContext.dashEhId || ""
  );
  const [selectedEventhouseName, setSelectedEventhouseName] = useState<string>(
    wizardContext.dashEhName || ""
  );
  const [loadingEventhouses, setLoadingEventhouses] = useState(false);
  const [eventhouseError, setEventhouseError] = useState("");

  // Database state
  const [databases, setDatabases] = useState<Item[]>([]);
  const [selectedDatabaseId, setSelectedDatabaseId] = useState<string>(
    wizardContext.dashDbId || ""
  );
  const [selectedDatabaseName, setSelectedDatabaseName] = useState<string>(
    wizardContext.dashDbName || ""
  );
  const [loadingDatabases, setLoadingDatabases] = useState(false);
  const [databaseError, setDatabaseError] = useState("");

  // Query service URI for KQL queries
  const [queryServiceUri, setQueryServiceUri] = useState<string>(
    wizardContext.dashQueryServiceUri || ""
  );

  // Table names (typed by user or pre-filled from ingestion wizard)
  const [telemetryTableName, setTelemetryTableName] = useState<string>(
    wizardContext.dashTelemetryTable || ""
  );
  const [propertiesTableName, setPropertiesTableName] = useState<string>(
    wizardContext.dashPropertiesTable || ""
  );

  // Tables list from database (with docstring metadata)
  const [tables, setTables] = useState<string[]>([]);
  const [tableDocstrings, setTableDocstrings] = useState<Record<string, string>>({});
  const [loadingTables, setLoadingTables] = useState(false);
  const [showAllTelemetryTables, setShowAllTelemetryTables] = useState(false);
  const [showAllPropertiesTables, setShowAllPropertiesTables] = useState(false);
  const [telemetryDropdownOpen, setTelemetryDropdownOpen] = useState(false);
  const [propertiesDropdownOpen, setPropertiesDropdownOpen] = useState(false);
  const skipTelemetryClose = useRef(false);
  const skipPropertiesClose = useRef(false);

  // Inspection state
  const [isInspecting, setIsInspecting] = useState(false);
  const [inspectError, setInspectError] = useState("");
  const [telemetryFields, setTelemetryFields] = useState<TelemetryField[]>(
    wizardContext.dashTelemetryFields || []
  );
  const [propertyFields, setPropertyFields] = useState<PropertyField[]>(
    wizardContext.dashPropertyFields || []
  );
  const [inspected, setInspected] = useState<boolean>(
    wizardContext.dashInspected || false
  );


  // Permission error
  const [permissionError, setPermissionError] = useState("");

  // Fabric API token
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

  // Load eventhouses
  const loadEventhouses = useCallback(async () => {
    if (!workspaceId) return;
    setLoadingEventhouses(true);
    setEventhouseError("");
    const hasToken = await ensureFabricApiToken();
    if (!hasToken) {
      setLoadingEventhouses(false);
      return;
    }
    try {
      const client = FabricPlatformAPIClient.create(workloadClient);
      const result = await client.items.listItems(workspaceId, {
        type: "Eventhouse",
      });
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

  // Load databases for selected eventhouse
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
        updateContext("dashQueryServiceUri", ehMeta.properties.queryServiceUri);

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

  // Load table names from a KQL database
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
        const docstrings: Record<string, string> = {};
        for (const r of rows) {
          const name = Array.isArray(r) ? r[0] : r.TableName;
          const doc = Array.isArray(r) ? r[1] : r.DocString;
          if (name) {
            tableNames.push(name);
            if (doc) docstrings[name] = doc;
          }
        }
        setTables(tableNames);
        setTableDocstrings(docstrings);
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

  // Execute a KQL query and return parsed rows
  const executeKqlQuery = async (
    qsUri: string,
    dbName: string,
    query: string
  ): Promise<any[]> => {
    const accessToken = await acquireTokenWithConsent(workloadClient, KUSTO_SCOPE);
    const response = await fetch(`${qsUri}/v1/rest/query`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken.token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ db: dbName, csl: query }),
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
        `Unable to retrieve data required for dashboard generation. (${response.status}): ${detail}`
      );
    }
    const result = await response.json();
    // Kusto v1 response: Tables[0].Columns + Tables[0].Rows
    const columns = result?.Tables?.[0]?.Columns || [];
    const rows = result?.Tables?.[0]?.Rows || [];
    // Map rows to objects using column names
    return rows.map((row: any[]) => {
      const obj: Record<string, any> = {};
      columns.forEach((col: any, idx: number) => {
        obj[col.ColumnName] = row[idx];
      });
      return obj;
    });
  };

  // Inspect tables — runs KQL queries to extract telemetry and property fields
  const inspectTables = useCallback(async () => {
    if (!queryServiceUri || !selectedDatabaseName || !telemetryTableName || !propertiesTableName) return;
    resetStepsFrom(stepIndex);
    setInspected(false);
    updateContext("dashInspected", false);
    updateContext("dashDashboardCreated", false);
    updateContext("dashDashboardId", "");
    setIsInspecting(true);
    setInspectError("");

    try {
      // Query 1: Get distinct telemetry field names and whether they are numeric.
      // Component telemetry arrives as separate messages tagged with IoTSubject=<component>;
      // prefix those fields as `<component>_<field>` so they match the modeled-data columns.
      // Root telemetry has an empty IoTSubject and stays unprefixed.
      const telemetryQuery = `${telemetryTableName}
| where ingestion_time() >= ago(30d)
${RAW_EVENT_NORMALIZATION_KQL}
| extend enqueuedTime = todatetime(headers.IoTEnqueueTime)
| where isnotempty(enqueuedTime)
| extend subject = tostring(headers.IoTSubject)
| extend j = data
| extend keys = bag_keys(j)
| mv-expand k = keys to typeof(string)
| where k !in ("EventProcessedUtcTime", "PartitionId", "EventEnqueuedUtcTime", "EventHub")
| extend v = j[k]
| extend valueType = gettype(v)
| extend Key = iff(subject == "", k, strcat(subject, "_", k))
| summarize arg_max(enqueuedTime, valueType) by Key, Component = subject, Leaf = k
| extend IsNumericType = valueType in ("int","long","real","decimal","double")
| project Key, IsNumericType, Component, Leaf
| order by Key asc`;

      // Query 2: Get distinct property names. Component reported properties arrive nested under the
      // component key with a `__t: "c"` marker; flatten those into `<component>_<field>` to match the
      // modeled columns. Non-component object properties (no __t) are kept whole.
      const propertiesQuery = `${propertiesTableName}
| where ingestion_time() >= ago(30d)
${RAW_EVENT_NORMALIZATION_KQL}
| extend enqueuedTime = todatetime(headers.IoTEnqueueTime)
| where isnotempty(enqueuedTime)
| extend rp = data.properties.reported
| extend keys = bag_keys(rp)
| mv-expand key = keys to typeof(string)
| where key !startswith "$" and key != "iothub-enqueuedtime" and key != "iothub-connection-device-id"
| extend val = rp[key]
| extend isComponent = (gettype(val) == "dictionary" and tostring(val["__t"]) == "c")
| mv-expand subkey = iff(isComponent, bag_keys(val), pack_array(key)) to typeof(string)
| where subkey != "__t"
| extend Component = iff(isComponent, key, ""), Leaf = subkey
| extend Key = iff(isComponent, strcat(key, "_", subkey), key)
| distinct Key, Component, Leaf
| order by Key asc`;

      const [telemetryRows, propertyRows] = await Promise.all([
        executeKqlQuery(queryServiceUri, selectedDatabaseName, telemetryQuery),
        executeKqlQuery(queryServiceUri, selectedDatabaseName, propertiesQuery),
      ]);

      const tFields: TelemetryField[] = telemetryRows.map((row) => ({
        key: row.Key,
        isNumeric: row.IsNumericType === true || row.IsNumericType === 1,
        included: true,
        component: row.Component ? String(row.Component) : undefined,
        leaf: row.Leaf != null ? String(row.Leaf) : undefined,
      }));

      const pFields: PropertyField[] = propertyRows.map((row) => ({
        key: row.Key,
        included: true,
        component: row.Component ? String(row.Component) : undefined,
        leaf: row.Leaf != null ? String(row.Leaf) : undefined,
      }));

      setTelemetryFields(tFields);
      setPropertyFields(pFields);
      setInspected(true);

      updateContext("dashTelemetryFields", tFields);
      updateContext("dashPropertyFields", pFields);
      updateContext("dashInspected", true);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      setInspectError(msg);
      setInspected(false);
      updateContext("dashInspected", false);
    } finally {
      setIsInspecting(false);
    }
  }, [
    queryServiceUri,
    selectedDatabaseName,
    telemetryTableName,
    propertiesTableName,
    resetStepsFrom,
    stepIndex,
    updateContext,
  ]);

  const invalidateGeneratedDashboard = () => {
    resetStepsFrom(stepIndex + 1);
    updateContext("dashDashboardCreated", false);
    updateContext("dashDashboardId", "");
  };

  // Toggle a telemetry field inclusion
  const toggleTelemetry = (index: number) => {
    const updated = [...telemetryFields];
    updated[index].included = !updated[index].included;
    setTelemetryFields(updated);
    updateContext("dashTelemetryFields", updated);
    invalidateGeneratedDashboard();
  };

  // Toggle a property field inclusion
  const toggleProperty = (index: number) => {
    const updated = [...propertyFields];
    updated[index].included = !updated[index].included;
    setPropertyFields(updated);
    updateContext("dashPropertyFields", updated);
    invalidateGeneratedDashboard();
  };

  const setAllTelemetries = (included: boolean) => {
    const updated = telemetryFields.map((field) => ({ ...field, included }));
    setTelemetryFields(updated);
    updateContext("dashTelemetryFields", updated);
    invalidateGeneratedDashboard();
  };

  const setAllProperties = (included: boolean) => {
    const updated = propertyFields.map((field) => ({ ...field, included }));
    setPropertyFields(updated);
    updateContext("dashPropertyFields", updated);
    invalidateGeneratedDashboard();
  };

  // Eventhouse selection handler
  const handleEventhouseChange = (eventhouseId: string) => {
    const eh = eventhouses.find((e) => e.id === eventhouseId);
    setSelectedEventhouseId(eventhouseId);
    setSelectedEventhouseName(eh?.displayName || "");
    updateContext("dashEhId", eventhouseId);
    updateContext("dashEhName", eh?.displayName || "");
    // Clear downstream
    setSelectedDatabaseId("");
    setSelectedDatabaseName("");
    setTables([]);
    setTelemetryTableName("");
    setPropertiesTableName("");
    setTelemetryFields([]);
    setPropertyFields([]);
    setInspected(false);
    updateContext("dashDbId", "");
    updateContext("dashDbName", "");
    updateContext("dashTelemetryTable", "");
    updateContext("dashPropertiesTable", "");
    updateContext("dashInspected", false);
    invalidateGeneratedDashboard();
    loadDatabases(eventhouseId);
  };

  // Database selection handler
  const handleDatabaseChange = (databaseId: string) => {
    const db = databases.find((d) => d.id === databaseId);
    setSelectedDatabaseId(databaseId);
    setSelectedDatabaseName(db?.displayName || "");
    updateContext("dashDbId", databaseId);
    updateContext("dashDbName", db?.displayName || "");
    // Clear downstream
    setTelemetryTableName("");
    setPropertiesTableName("");
    setTelemetryFields([]);
    setPropertyFields([]);
    setInspected(false);
    updateContext("dashTelemetryTable", "");
    updateContext("dashPropertiesTable", "");
    updateContext("dashInspected", false);
    invalidateGeneratedDashboard();
    if (db?.displayName && queryServiceUri) {
      loadTables(db.displayName, queryServiceUri);
    }
  };

  // Table selection handlers
  const handleTelemetryTableChange = (table: string) => {
    setTelemetryTableName(table);
    updateContext("dashTelemetryTable", table);
    setTelemetryFields([]);
    setPropertyFields([]);
    setInspected(false);
    updateContext("dashInspected", false);
    invalidateGeneratedDashboard();
  };

  const handlePropertiesTableChange = (table: string) => {
    setPropertiesTableName(table);
    updateContext("dashPropertiesTable", table);
    setTelemetryFields([]);
    setPropertyFields([]);
    setInspected(false);
    updateContext("dashInspected", false);
    invalidateGeneratedDashboard();
  };

  // Load eventhouses on mount
  useEffect(() => {
    loadEventhouses();
  }, [loadEventhouses]);

  // Load databases if eventhouse is pre-selected
  useEffect(() => {
    if (selectedEventhouseId && databases.length === 0 && !loadingDatabases) {
      loadDatabases(selectedEventhouseId);
    }
  }, [selectedEventhouseId]);

  // Load tables if database is pre-selected
  useEffect(() => {
    if (selectedDatabaseName && queryServiceUri && tables.length === 0 && !loadingTables) {
      loadTables(selectedDatabaseName, queryServiceUri);
    }
  }, [selectedDatabaseName, queryServiceUri]);

  const canInspect =
    !!selectedEventhouseId &&
    !!selectedDatabaseId &&
    !!telemetryTableName &&
    !!propertiesTableName &&
    !isInspecting;

  const includedTelemetryCount = telemetryFields.filter((f) => f.included).length;
  const includedPropertyCount = propertyFields.filter((f) => f.included).length;
  const timeSeriesCount = telemetryFields.filter((f) => f.included && f.isNumeric).length;
  const nonTimeSeriesCount = telemetryFields.filter((f) => f.included && !f.isNumeric).length;

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Select Data Source and Fields</h2>
      <Text className="iot-solution-step-description">
        Select the raw data tables containing your IoT device data. Then inspect the
        tables to discover available telemetry and property fields.
      </Text>

      {permissionError && (
        <MessageBar intent="error" style={{ marginBottom: "12px" }}>
          <MessageBarBody>{permissionError}</MessageBarBody>
        </MessageBar>
      )}

      {/* Eventhouse & Database Selection */}
      <div className="iot-solution-form">
        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="dash-eh-select">
            Eventhouse
          </Label>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <Dropdown
              id="dash-eh-select"
              placeholder={loadingEventhouses ? "Loading..." : "Select an Eventhouse"}
              value={selectedEventhouseName}
              selectedOptions={selectedEventhouseId ? [selectedEventhouseId] : []}
              onOptionSelect={(_, data) => {
                if (data.optionValue) handleEventhouseChange(data.optionValue);
              }}
              disabled={loadingEventhouses || eventhouses.length === 0}
              style={{ minWidth: "280px" }}
            >
              {eventhouses.map((eh) => (
                <Option key={eh.id} value={eh.id}>{eh.displayName}</Option>
              ))}
            </Dropdown>
            <button
              onClick={() => loadEventhouses()}
              style={{ background: "none", border: "none", cursor: "pointer", padding: "4px" }}
              title="Refresh"
            >
              <ArrowClockwise24Regular />
            </button>
          </div>
          {eventhouseError && (
            <Text style={{ color: "var(--colorPaletteRedForeground1)", fontSize: "var(--fontSizeBase200)" }}>
              {eventhouseError}
            </Text>
          )}
        </div>

        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="dash-db-select">
            KQL Database
          </Label>
          <Dropdown
            id="dash-db-select"
            placeholder={loadingDatabases ? "Loading..." : "Select a database"}
            value={selectedDatabaseName}
            selectedOptions={selectedDatabaseId ? [selectedDatabaseId] : []}
            onOptionSelect={(_, data) => {
              if (data.optionValue) handleDatabaseChange(data.optionValue);
            }}
            disabled={!selectedEventhouseId || loadingDatabases || databases.length === 0}
            style={{ minWidth: "280px" }}
          >
            {databases.map((db) => (
              <Option key={db.id} value={db.id}>{db.displayName}</Option>
            ))}
          </Dropdown>
          {databaseError && (
            <Text style={{ color: "var(--colorPaletteRedForeground1)", fontSize: "var(--fontSizeBase200)" }}>
              {databaseError}
            </Text>
          )}
        </div>

        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="dash-telemetry-table">
            Raw Telemetry Table
          </Label>
          <Dropdown
            id="dash-telemetry-table"
            placeholder={loadingTables ? "Loading..." : "Select raw telemetry table"}
            value={telemetryTableName}
            selectedOptions={telemetryTableName ? [telemetryTableName] : []}
            open={telemetryDropdownOpen}
            onOpenChange={(_, data) => {
              if (!data.open && skipTelemetryClose.current) {
                skipTelemetryClose.current = false;
                return;
              }
              setTelemetryDropdownOpen(data.open);
            }}
            onOptionSelect={(_, data) => {
              if (data.optionValue === "__show_all_telemetry__") {
                setShowAllTelemetryTables(true);
                skipTelemetryClose.current = true;
                return;
              }
              if (data.optionValue) handleTelemetryTableChange(data.optionValue);
              setTelemetryDropdownOpen(false);
            }}
            disabled={!selectedDatabaseId || loadingTables || tables.length === 0}
            style={{ minWidth: "280px" }}
          >
            {(showAllTelemetryTables
              ? tables
              : filterTablesByRole(tables, tableDocstrings, "iothub-telemetry-raw")
            ).map((t) => (
              <Option key={t} value={t}>{t}</Option>
            ))}
            {!showAllTelemetryTables && tables.length > 0 && !loadingTables && (
              <Option
                key="__show_all_telemetry__"
                value="__show_all_telemetry__"
                style={{ fontStyle: "italic", color: "var(--colorBrandForeground1)" }}
              >
                Show all tables...
              </Option>
            )}
          </Dropdown>
        </div>

        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="dash-properties-table">
            Raw Properties Table
          </Label>
          <Dropdown
            id="dash-properties-table"
            placeholder={loadingTables ? "Loading..." : "Select raw properties table"}
            value={propertiesTableName}
            selectedOptions={propertiesTableName ? [propertiesTableName] : []}
            open={propertiesDropdownOpen}
            onOpenChange={(_, data) => {
              if (!data.open && skipPropertiesClose.current) {
                skipPropertiesClose.current = false;
                return;
              }
              setPropertiesDropdownOpen(data.open);
            }}
            onOptionSelect={(_, data) => {
              if (data.optionValue === "__show_all_properties__") {
                setShowAllPropertiesTables(true);
                skipPropertiesClose.current = true;
                return;
              }
              if (data.optionValue) handlePropertiesTableChange(data.optionValue);
              setPropertiesDropdownOpen(false);
            }}
            disabled={!selectedDatabaseId || loadingTables || tables.length === 0}
            style={{ minWidth: "280px" }}
          >
            {(showAllPropertiesTables
              ? tables
              : filterTablesByRole(tables, tableDocstrings, "iothub-properties-raw")
            ).map((t) => (
              <Option key={t} value={t}>{t}</Option>
            ))}
            {!showAllPropertiesTables && tables.length > 0 && !loadingTables && (
              <Option
                key="__show_all_properties__"
                value="__show_all_properties__"
                style={{ fontStyle: "italic", color: "var(--colorBrandForeground1)" }}
              >
                Show all tables...
              </Option>
            )}
          </Dropdown>
        </div>
      </div>

      {/* Inspect Button */}
      {!isInspecting && (
        <div style={{ marginTop: "16px" }}>
          <button
            onClick={inspectTables}
            disabled={!canInspect}
            style={{
              padding: "8px 16px",
              background: canInspect ? "var(--colorBrandBackground)" : "var(--colorNeutralBackgroundDisabled)",
              color: canInspect ? "var(--colorNeutralForegroundOnBrand)" : "var(--colorNeutralForegroundDisabled)",
              border: "none",
              borderRadius: "var(--borderRadiusMedium)",
              cursor: canInspect ? "pointer" : "not-allowed",
              fontFamily: "var(--fontFamilyBase)",
              fontSize: "var(--fontSizeBase300)",
              fontWeight: "var(--fontWeightSemibold)" as any,
            }}
          >
            {inspected ? "Re-inspect Tables" : "Inspect Tables"}
          </button>
        </div>
      )}

      {isInspecting && (
        <div style={{ marginTop: "16px", display: "flex", alignItems: "center", gap: "8px" }}>
          <Spinner size="small" />
          <Text>Inspecting tables...</Text>
        </div>
      )}

      {inspectError && (
        <MessageBar intent="error" style={{ marginTop: "12px" }}>
          <MessageBarBody>{inspectError}</MessageBarBody>
        </MessageBar>
      )}

      {/* Results */}
      {inspected && telemetryFields.length > 0 && (
        <div style={{ marginTop: "20px" }}>
          <Text className="iot-solution-step-description" block style={{ marginBottom: "12px" }}>
            The fields below were discovered from recent events in the selected raw
            data tables, covering up to the last 30 days. Select the fields to include
            in the dashboard. Numeric telemetry is used for time-series charts,
            non-numeric telemetry is presented in recent-value tables, and properties
            are presented as last-known-value cards.
          </Text>
          <div className="iot-solution-config-card">
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "12px", marginBottom: "8px" }}>
              <div>
                <Text weight="semibold" size={400} block>
                  Telemetry Fields
                </Text>
                <Text size={200} block style={{ color: "var(--colorNeutralForeground3)" }}>
                  {includedTelemetryCount} selected — {timeSeriesCount} time-series (numeric), {nonTimeSeriesCount} non-time-series
                </Text>
              </div>
              <Button
                appearance="subtle"
                size="small"
                onClick={() => setAllTelemetries(!telemetryFields.every((field) => field.included))}
              >
                {telemetryFields.every((field) => field.included) ? "Deselect all" : "Select all"}
              </Button>
            </div>
            <div className="iot-solution-field-list">
              {telemetryFields.map((field, idx) => (
                <div key={field.key} className="iot-solution-field-list-item">
                  <Checkbox
                    checked={field.included}
                    onChange={() => toggleTelemetry(idx)}
                    label={field.key}
                  />
                  <span
                    className={`iot-solution-field-badge ${
                      field.isNumeric ? "iot-solution-field-badge--numeric" : "iot-solution-field-badge--text"
                    }`}
                  >
                    {field.isNumeric ? "Time-series" : "Non-time-series"}
                  </span>
                </div>
              ))}
            </div>
          </div>

          <div className="iot-solution-config-card" style={{ marginTop: "12px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: "12px", marginBottom: "8px" }}>
              <div>
                <Text weight="semibold" size={400} block>
                  Property Fields
                </Text>
                <Text size={200} block style={{ color: "var(--colorNeutralForeground3)" }}>
                  {includedPropertyCount} selected
                </Text>
              </div>
              {propertyFields.length > 0 && (
                <Button
                  appearance="subtle"
                  size="small"
                  onClick={() => setAllProperties(!propertyFields.every((field) => field.included))}
                >
                  {propertyFields.every((field) => field.included) ? "Deselect all" : "Select all"}
                </Button>
              )}
            </div>
            <div className="iot-solution-field-list">
              {propertyFields.map((field, idx) => (
                <div key={field.key} className="iot-solution-field-list-item">
                  <Checkbox
                    checked={field.included}
                    onChange={() => toggleProperty(idx)}
                    label={field.key}
                  />
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {inspected && telemetryFields.length === 0 && propertyFields.length === 0 && (
        <MessageBar intent="warning" style={{ marginTop: "12px" }}>
          <MessageBarBody>
            No telemetry or property fields were found in the last 30 days. Make sure your
            devices are sending data and the correct tables are selected.
          </MessageBarBody>
        </MessageBar>
      )}
    </div>
  );
}
