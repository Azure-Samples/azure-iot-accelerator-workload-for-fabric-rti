// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useState, useCallback, useEffect, useRef } from "react";
import {
  Dropdown,
  Label,
  Option,
  Spinner,
  Text,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { FabricPlatformAPIClient } from "../../../clients/FabricPlatformAPIClient";
import { Item } from "../../../clients/FabricPlatformTypes";
import { getEventhouseItem } from "../eventhouseClient";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import "../IoTSolutionItem.scss";

const FABRIC_ITEM_READ_SCOPE = "https://api.fabric.microsoft.com/Item.Read.All";
const KUSTO_SCOPE = "https://kusto.kusto.windows.net/.default";

interface ActivatorSelectTableStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
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

/**
 * Activator Wizard Step 1: Select the modeled-data table.
 *
 * User picks an Eventhouse, database, and the modeled-data KQL table that
 * Activator should monitor. The selection (plus derived default resource names)
 * is stored in the shared wizard context for the following steps.
 */
export function ActivatorSelectTableStep({
  stepIndex,
  wizardContext,
  updateContext,
  resetStepsFrom,
  workloadClient,
  workspaceId,
}: ActivatorSelectTableStepProps) {
  const suffix = useRef(wizardContext.actNameSuffix || Math.random().toString(36).substring(2, 6));

  const [eventhouses, setEventhouses] = useState<Item[]>([]);
  const [selectedEventhouseId, setSelectedEventhouseId] = useState<string>(
    wizardContext.actEventhouseId || ""
  );
  const [selectedEventhouseName, setSelectedEventhouseName] = useState<string>(
    wizardContext.actEventhouseName || ""
  );
  const [loadingEventhouses, setLoadingEventhouses] = useState(false);
  const [eventhouseError, setEventhouseError] = useState("");

  const [databases, setDatabases] = useState<Item[]>([]);
  const [selectedDatabaseId, setSelectedDatabaseId] = useState<string>(
    wizardContext.actDatabaseId || ""
  );
  const [selectedDatabaseName, setSelectedDatabaseName] = useState<string>(
    wizardContext.actDatabaseName || ""
  );
  const [loadingDatabases, setLoadingDatabases] = useState(false);
  const [databaseError, setDatabaseError] = useState("");

  const [queryServiceUri, setQueryServiceUri] = useState<string>(
    wizardContext.actQueryServiceUri || ""
  );

  const [tables, setTables] = useState<string[]>([]);
  const [tableDocstrings, setTableDocstrings] = useState<Record<string, string>>({});
  const [loadingTables, setLoadingTables] = useState(false);
  const [tableDropdownOpen, setTableDropdownOpen] = useState(false);
  const [selectedTable, setSelectedTable] = useState<string>(
    wizardContext.actTableName || ""
  );

  const [permissionError, setPermissionError] = useState("");

  const invalidateResources = useCallback(() => {
    resetStepsFrom(stepIndex + 1);
    updateContext("activatorCreated", false);
    updateContext("actActivatorId", "");
    updateContext("actEventstreamId", "");
    updateContext("actConnectionId", "");
  }, [resetStepsFrom, stepIndex, updateContext]);

  // Persist the generated suffix so resource names stay stable across steps.
  useEffect(() => {
    if (!wizardContext.actNameSuffix) {
      updateContext("actNameSuffix", suffix.current);
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

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
        updateContext("actQueryServiceUri", ehMeta.properties.queryServiceUri);

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
            Authorization: "Bearer " + accessToken.token,
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
        const modeledTables = filterTablesByRole(tableNames, docMap, "modeled-data");
        if (selectedTable && !modeledTables.includes(selectedTable)) {
          setSelectedTable("");
          updateContext("actTableName", "");
          updateContext("actActivatorName", "");
          updateContext("actEventstreamName", "");
          invalidateResources();
        }
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
    [workloadClient, selectedTable, updateContext, invalidateResources]
  );

  // Auto-load eventhouses on mount
  useEffect(() => {
    if (eventhouses.length === 0 && !loadingEventhouses) {
      loadEventhouses();
    }
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-load databases when eventhouse is (pre)selected
  useEffect(() => {
    if (selectedEventhouseId) loadDatabases(selectedEventhouseId);
  }, [selectedEventhouseId]); // eslint-disable-line react-hooks/exhaustive-deps

  // Auto-load tables when database is (pre)selected
  useEffect(() => {
    if (selectedDatabaseName && queryServiceUri) {
      loadTables(selectedDatabaseName, queryServiceUri);
    }
  }, [selectedDatabaseName, queryServiceUri]); // eslint-disable-line react-hooks/exhaustive-deps

  // Generate default resource names when a table is selected
  const applyDefaultNames = (table: string, docMap: Record<string, string>) => {
    const doc = docMap[table] || "";
    const meta = parseIotMeta(doc);
    const baseName = meta.deviceModel
      ? meta.deviceModel.replace(/[^a-zA-Z0-9_]/g, "_")
      : table.replace(/[^a-zA-Z0-9_]/g, "_");
    updateContext("actActivatorName", `Activator_${baseName}_${suffix.current}`);
    updateContext("actEventstreamName", `Activator_Stream_${baseName}_${suffix.current}`);
  };

  useEffect(() => {
    if (
      selectedTable &&
      Object.keys(tableDocstrings).length > 0 &&
      (!wizardContext.actActivatorName || !wizardContext.actEventstreamName)
    ) {
      applyDefaultNames(selectedTable, tableDocstrings);
    }
  }, [selectedTable, tableDocstrings]); // eslint-disable-line react-hooks/exhaustive-deps

  return (
    <div className="iot-solution-step">
      <h2 className="iot-solution-step-title">Select Data for Monitoring</h2>
      <Text className="iot-solution-step-description">
        Select the modeled dataset that Activator will evaluate as new device
        events arrive. Modeled data is recommended because it provides
        business-friendly measurements and enriches telemetry with current device
        properties, enabling rules that correlate live behavior with device state.
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
        {/* Eventhouse selection */}
        <div className="iot-solution-field">
          <Label className="iot-solution-field-label" required htmlFor="act-eh-dropdown">
            Eventhouse
          </Label>
          {loadingEventhouses ? (
            <Spinner size="tiny" label="Loading Eventhouses..." />
          ) : eventhouseError ? (
            <MessageBar intent="error"><MessageBarBody>{eventhouseError}</MessageBarBody></MessageBar>
          ) : (
            <Dropdown
              id="act-eh-dropdown"
              placeholder="Select an Eventhouse"
              value={selectedEventhouseName}
              selectedOptions={selectedEventhouseId ? [selectedEventhouseId] : []}
              onOptionSelect={(_, data) => {
                const eh = eventhouses.find((e) => e.id === data.optionValue);
                if (eh) {
                 invalidateResources();
                  setSelectedEventhouseId(eh.id);
                  setSelectedEventhouseName(eh.displayName);
                  setSelectedDatabaseId("");
                  setSelectedDatabaseName("");
                  setSelectedTable("");
                  setDatabases([]);
                  setTables([]);
                  updateContext("actEventhouseId", eh.id);
                  updateContext("actEventhouseName", eh.displayName);
                  updateContext("actDatabaseId", "");
                  updateContext("actDatabaseName", "");
                  updateContext("actTableName", "");
                  loadDatabases(eh.id);
                }
              }}
              style={{ minWidth: "320px" }}
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
            <Label className="iot-solution-field-label" required htmlFor="act-db-dropdown">
              Database
            </Label>
            {loadingDatabases ? (
              <Spinner size="tiny" label="Loading databases..." />
            ) : databaseError ? (
              <MessageBar intent="error"><MessageBarBody>{databaseError}</MessageBarBody></MessageBar>
            ) : (
              <Dropdown
                id="act-db-dropdown"
                placeholder="Select a database"
                value={selectedDatabaseName}
                selectedOptions={selectedDatabaseId ? [selectedDatabaseId] : []}
                onOptionSelect={(_, data) => {
                  const db = databases.find((d) => d.id === data.optionValue);
                  if (db) {
                    invalidateResources();
                    setSelectedDatabaseId(db.id);
                    setSelectedDatabaseName(db.displayName);
                    setSelectedTable("");
                    setTables([]);
                    updateContext("actDatabaseId", db.id);
                    updateContext("actDatabaseName", db.displayName);
                    updateContext("actTableName", "");
                  }
                }}
                style={{ minWidth: "320px" }}
              >
                {databases.map((db) => (
                  <Option key={db.id} value={db.id}>{db.displayName}</Option>
                ))}
              </Dropdown>
            )}
          </div>
        )}

        {/* Table selection */}
        {selectedDatabaseName && (
          <div className="iot-solution-field">
            <Label className="iot-solution-field-label" required htmlFor="act-table-dropdown">
              Modeled Data Table
            </Label>
            {loadingTables ? (
              <Spinner size="tiny" label="Loading tables..." />
            ) : filterTablesByRole(tables, tableDocstrings, "modeled-data").length > 0 ? (
              <Dropdown
                id="act-table-dropdown"
                placeholder="Select a table"
                value={selectedTable}
                selectedOptions={selectedTable ? [selectedTable] : []}
                open={tableDropdownOpen}
                onOpenChange={(_e, data) => {
                  setTableDropdownOpen(data.open);
                }}
                onOptionSelect={(_, data) => {
                  const table = data.optionValue as string;
                  invalidateResources();
                  setSelectedTable(table);
                  updateContext("actTableName", table);
                  applyDefaultNames(table, tableDocstrings);
                  setTableDropdownOpen(false);
                }}
                style={{ minWidth: "320px" }}
              >
                {filterTablesByRole(tables, tableDocstrings, "modeled-data").map((t) => (
                  <Option key={t} value={t}>{t}</Option>
                ))}
              </Dropdown>
            ) : (
              <MessageBar intent="warning">
                <MessageBarBody>
                  No modeled data tables were found in this database. Create modeled
                  data first, then return to configure Activator.
                </MessageBarBody>
              </MessageBar>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
