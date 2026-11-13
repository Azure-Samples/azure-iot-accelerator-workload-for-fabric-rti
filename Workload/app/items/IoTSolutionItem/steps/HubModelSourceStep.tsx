// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { useEffect, useLayoutEffect, useRef, useState } from "react";
import { Button, Dropdown, Label, MessageBar, MessageBarBody, Option, Spinner, Text } from "@fluentui/react-components";
import { CheckmarkCircle24Filled } from "@fluentui/react-icons";
import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { WizardStepProps } from "../../../components/Wizard";
import { FabricPlatformAPIClient } from "../../../clients/FabricPlatformAPIClient";
import { FabricPlatformError } from "../../../clients/FabricPlatformClient";
import { Item } from "../../../clients/FabricPlatformTypes";
import { acquireTokenWithConsent } from "../../../controller/AuthenticationController";
import { getEventhouseItem } from "../eventhouseClient";
import { canonicalJson, executeAuthoringRead, parseIotMeta, readPayloadSamples } from "../ModelAuthoring";
import type { AssignmentSource } from "../HubAssignment.ts";
import { clearAutomaticHints, prepareRoleEvidence } from "../HubModeling.ts";
import { rememberCreationAttempt } from "../ModeledDataCreation.ts";
import { currentSourceAnalysis, modelSourceKey, sourceSelectionRevision, sourceSelectionStatus, SOURCE_ANALYSIS_LIMIT,
  type SourceAnalysis, type SourceSelectionStatus } from "../ModelSourceReadiness.ts";
import "../IoTSolutionItem.scss";

interface HubModelSourceStepProps extends WizardStepProps {
  workloadClient: WorkloadClientAPI;
  workspaceId: string;
  onSourceStateChange?: (state: SourceSelectionStatus) => void;
}

export function HubModelSourceStep({
  workloadClient, workspaceId, wizardContext, updateContext, resetStepsFrom, stepIndex, onSourceStateChange,
}: HubModelSourceStepProps) {
  const [items, setItems] = useState<Item[]>([]);
  const [databaseIds, setDatabaseIds] = useState<string[]>([]);
  const [tables, setTables] = useState<string[]>([]);
  const [itemsOwner, setItemsOwner] = useState("");
  const [databasesOwner, setDatabasesOwner] = useState("");
  const [tablesOwner, setTablesOwner] = useState("");
  const [docstrings, setDocstrings] = useState<Record<string, string>>({});
  const [tableRoles, setTableRoles] = useState<Record<string, string>>({});
  const [metadataWarnings, setMetadataWarnings] = useState<{ table: string; message: string }[]>([]);
  const [showAllTelemetry, setShowAllTelemetry] = useState(false);
  const [showAllProperties, setShowAllProperties] = useState(false);
  const [telemetryOpen, setTelemetryOpen] = useState(false);
  const [propertiesOpen, setPropertiesOpen] = useState(false);
  const telemetryExpansionEvent = useRef<React.SyntheticEvent | Event>();
  const propertiesExpansionEvent = useRef<React.SyntheticEvent | Event>();
  const [loadingItems, setLoadingItems] = useState(true);
  const [loadingDatabases, setLoadingDatabases] = useState(false);
  const [loadingTables, setLoadingTables] = useState(false);
  const [itemError, setItemError] = useState("");
  const [databaseError, setDatabaseError] = useState("");
  const [tableError, setTableError] = useState("");
  const [resourceRetry, setResourceRetry] = useState(0);
  const [databaseRetry, setDatabaseRetry] = useState(0);
  const [tableRetry, setTableRetry] = useState(0);
  const [itemRetryable, setItemRetryable] = useState(true);
  const [databaseRetryable, setDatabaseRetryable] = useState(true);
  const [tableRetryable, setTableRetryable] = useState(true);
  const [analyzing, setAnalyzing] = useState(false);
  const [analysisError, setAnalysisError] = useState("");
  const analysisRequest = useRef<AbortController>();
  const analysisRunning = useRef(false);
  const mounted = useRef(true);
  const heading = useRef<HTMLHeadingElement>(null);
  const ehId: string = wizardContext.modelEhId || "";
  const dbId: string = wizardContext.modelDbId || "";
  const dbName: string = wizardContext.modelDbName || "";
  const queryServiceUri: string = wizardContext.modelQueryServiceUri || "";
  const telemetryTable: string = wizardContext.rawTelemetryTable || "";
  const twinTable: string = wizardContext.rawPropertiesTable || "";

  useEffect(() => { heading.current?.focus(); }, []);

  useEffect(() => {
    let current = true;
    setLoadingItems(true);
    setItemError("");
    setItems([]);
    setItemRetryable(true);
    const load = async () => {
      try {
        if (!workspaceId) throw new Error("This item's workspace context could not be loaded. Data source selection is unavailable.");
        await acquireTokenWithConsent(workloadClient, "https://api.fabric.microsoft.com/Item.Read.All");
        if (!current) return;
        const client = FabricPlatformAPIClient.create(workloadClient);
        const allItems = await client.items.getAllItems(workspaceId);
        if (!Array.isArray(allItems) || !allItems.every(item =>
          typeof item?.id === "string" && typeof item.type === "string" && typeof item.displayName === "string")) {
          throw new Error("The workspace returned an unexpected resource list. Retry loading resources.");
        }
        if (current) { setItems(allItems); setItemsOwner(workspaceId); }
      } catch (err: unknown) {
        if (current) {
          setItemError(err instanceof Error ? err.message : "The item list request failed.");
          setItemRetryable(!(typeof err === "object" && err !== null && "isRetriable" in err && err.isRetriable === false));
        }
      } finally {
        if (current) setLoadingItems(false);
      }
    };
    void load();
    return () => { current = false; };
  }, [workloadClient, workspaceId, resourceRetry]);

  useEffect(() => {
    if (!ehId) {
      setDatabaseIds([]);
      setDatabasesOwner("");
      setDatabaseError("");
      setDatabaseRetryable(true);
      setLoadingDatabases(false);
      return undefined;
    }
    let current = true;
    setLoadingDatabases(true);
    setDatabaseError("");
    setDatabaseRetryable(true);
    setDatabaseIds([]);
    setDatabasesOwner("");
    const load = async () => {
      try {
        const eventhouse = await getEventhouseItem(workloadClient, workspaceId, ehId);
        if (!eventhouse?.properties?.queryServiceUri || !Array.isArray(eventhouse.properties.databasesItemIds) ||
            !eventhouse.properties.databasesItemIds.every(id => typeof id === "string")) {
          throw new Error("The Eventhouse returned incomplete database information.");
        }
        if (!current) return;
        setDatabaseIds(eventhouse.properties.databasesItemIds);
        setDatabasesOwner(JSON.stringify([workspaceId, ehId]));
        updateContext("modelQueryServiceUri", eventhouse.properties.queryServiceUri);
      } catch (err: unknown) {
        if (current) {
          setDatabaseError(err instanceof Error ? err.message
            : typeof err === "object" && err !== null && "message" in err && typeof err.message === "string"
              ? err.message : "The database list request failed.");
          setDatabaseRetryable(err instanceof FabricPlatformError
            ? err.statusCode === 408 || err.statusCode === 429 || err.statusCode >= 500
            : !(typeof err === "object" && err !== null && "isRetriable" in err && err.isRetriable === false));
        }
      } finally {
        if (current) setLoadingDatabases(false);
      }
    };
    void load();
    return () => { current = false; };
  }, [ehId, workloadClient, workspaceId, databaseRetry, updateContext]);

  useEffect(() => {
    if (!dbName || !queryServiceUri || !databaseIds.includes(dbId) ||
        databasesOwner !== JSON.stringify([workspaceId, ehId])) {
      setTables([]);
      setTablesOwner("");
      setDocstrings({});
      setTableRoles({});
      setMetadataWarnings([]);
      setTableError("");
      setLoadingTables(false);
      return undefined;
    }
    const controller = new AbortController();
    setLoadingTables(true);
    setTableError("");
    setTableRetryable(true);
    setTables([]);
    setTablesOwner("");
    setDocstrings({});
    setTableRoles({});
    setMetadataWarnings([]);
    setShowAllTelemetry(false);
    setShowAllProperties(false);
    setTelemetryOpen(false);
    setPropertiesOpen(false);
    telemetryExpansionEvent.current = undefined;
    propertiesExpansionEvent.current = undefined;
    const load = async () => {
      try {
        const token = await acquireTokenWithConsent(workloadClient, "https://kusto.kusto.windows.net/.default");
        if (controller.signal.aborted) return;
        const result = await executeAuthoringRead(
          { queryServiceUri, databaseName: dbName }, token.token,
          ".show tables details | project TableName, DocString", "mgmt", controller.signal
        );
        if (controller.signal.aborted) return;
        const nameIndex = result.Columns.findIndex(column => column.ColumnName === "TableName");
        const docIndex = result.Columns.findIndex(column => column.ColumnName === "DocString");
        if (result.Columns.length !== 2 || nameIndex < 0 || docIndex < 0 ||
            result.Columns.some(column => column.ColumnType?.toLowerCase() !== "string") ||
            !result.Rows.every(row => typeof row[nameIndex] === "string" &&
              (typeof row[docIndex] === "string" || row[docIndex] === null)) ||
            new Set(result.Rows.map(row => row[nameIndex])).size !== result.Rows.length) {
          throw new Error("The database returned malformed table metadata.");
        }
        const roles: Record<string, string> = Object.create(null);
        const warnings: { table: string; message: string }[] = [];
        const documents = Object.fromEntries(result.Rows.map(row => {
          const table = String(row[nameIndex]);
          const document = typeof row[docIndex] === "string" ? row[docIndex] : "";
          try {
            roles[table] = parseIotMeta(document).tableRole || "";
            return [table, document];
          } catch (error: unknown) {
            if (!(error instanceof Error)) throw error;
            warnings.push({ table, message: error.message });
            return [table, ""];
          }
        }));
        setTables(result.Rows.map(row => String(row[nameIndex])).sort((a, b) => a.localeCompare(b)));
        setTablesOwner(JSON.stringify([workspaceId, ehId, dbId, dbName, queryServiceUri]));
        setDocstrings(documents);
        setTableRoles(roles);
        setMetadataWarnings(warnings);
      } catch (err: unknown) {
        if (!controller.signal.aborted) {
          setTableError(err instanceof Error ? err.message : "The table list request failed.");
          setTableRetryable(!(typeof err === "object" && err !== null && "isRetriable" in err && err.isRetriable === false));
        }
      } finally {
        if (!controller.signal.aborted) setLoadingTables(false);
      }
    };
    void load();
    return () => { controller.abort(); };
  }, [workspaceId, ehId, dbName, dbId, queryServiceUri, databaseIds, databasesOwner, workloadClient, tableRetry]);

  const eventhouses = items.filter(item => item.type === "Eventhouse");
  const databases = items.filter(item => item.type === "KQLDatabase" && databaseIds.includes(item.id));
  const itemsCurrent = itemsOwner === workspaceId;
  const databasesCurrent = databasesOwner === JSON.stringify([workspaceId, ehId]);
  const tablesCurrent = tablesOwner === JSON.stringify([workspaceId, ehId, dbId, dbName, queryServiceUri]);
  const loading = loadingItems || loadingDatabases || loadingTables;
  const databaseFieldLoading = !!ehId && (loadingItems || loadingDatabases);
  const tableFieldLoading = !!dbId && loading;
  const error = itemError || databaseError || tableError;
  const sameTable = !!telemetryTable && telemetryTable === twinTable;
  const sourceKey = modelSourceKey(workspaceId, wizardContext);
  const selectionStatus = sourceSelectionStatus({ key: sourceKey, loading, error, eventhouseId: ehId, databaseId: dbId, queryServiceUri,
    eventhouseIds: itemsCurrent ? eventhouses.map(item => item.id) : [],
    databaseIds: databasesCurrent ? databases.map(item => item.id) : [], tables: tablesCurrent ? tables : [],
    telemetry: telemetryTable, properties: twinTable });
  const selectionRevision = sourceSelectionRevision(wizardContext);
  const revisionRef = useRef(selectionRevision); revisionRef.current = selectionRevision;
  const analysis = currentSourceAnalysis(workspaceId, wizardContext);
  const liveSelection = useRef({ key: sourceKey, revision: selectionRevision, ready: selectionStatus.ready });
  liveSelection.current = { key: sourceKey, revision: selectionRevision, ready: selectionStatus.ready };
  const ready = selectionStatus.ready && !!analysis;
  const status: SourceSelectionStatus = { key: sourceKey, ready, reason: !selectionStatus.ready ? selectionStatus.reason
    : analyzing ? "Validating source tables..." : analysisError ? "Validation failed. Choose Validate again."
      : ready ? "Source tables validated" : "Validate source tables before continuing." };
  useEffect(() => {
    mounted.current = true;
    return () => { mounted.current = false; analysisRequest.current?.abort(); };
  }, []);
  useLayoutEffect(() => {
    analysisRequest.current?.abort();
    analysisRequest.current = undefined; analysisRunning.current = false;
    setAnalyzing(false); setAnalysisError("");
  }, [sourceKey, selectionRevision, selectionStatus.ready]);
  useLayoutEffect(() => { onSourceStateChange?.(status); }, [sourceKey, ready, status.reason, onSourceStateChange]);
  const assignmentSource: AssignmentSource | undefined = ready ? { key: sourceKey,
    telemetry: { table: telemetryTable, docstring: docstrings[telemetryTable] || "" },
    twin: { table: twinTable, docstring: docstrings[twinTable] || "" } } : undefined;
  const assignmentSourceKey = canonicalJson(assignmentSource || null);
  useLayoutEffect(() => {
    if (canonicalJson(wizardContext.hubAssignmentSource || null) !== assignmentSourceKey) {
      updateContext("hubAssignmentSource", assignmentSource);
    }
  }, [assignmentSourceKey, updateContext]);

  useEffect(() => {
    updateContext("modelRawDataSelected", ready);
  }, [ready, ehId, dbId, queryServiceUri, telemetryTable, twinTable, updateContext]);

  const invalidateProof = () => {
    onSourceStateChange?.({ key: sourceKey, ready: false, reason: "Wait for the source tables to finish loading." });
    resetStepsFrom(stepIndex + 1);
    updateContext("hubAssignmentIntent", undefined);
    updateContext("hubAssignmentHubSummary", undefined);
    updateContext("hubCreationHistory", rememberCreationAttempt(wizardContext.hubCreationHistory || [], wizardContext.hubCreation));
    updateContext("hubCreation", undefined);
    updateContext("modelRawDataSelected", false);
    updateContext("modelCoverageVerified", false);
    updateContext("hubCoverage", undefined);
    updateContext("hubStreamChecks", undefined);
    updateContext("hubPreview", undefined);
    updateContext("hubAllowWaiting", false);
  };
  const invalidate = () => {
    analysisRequest.current?.abort(); analysisRequest.current = undefined; analysisRunning.current = false;
    setAnalyzing(false); setAnalysisError("");
    revisionRef.current++; updateContext("hubSourceSelectionRevision", revisionRef.current);
    updateContext("hubSourceAnalysis", undefined);
    invalidateProof();
    updateContext("hubFields", clearAutomaticHints(wizardContext.hubFields || []));
    updateContext("hubRoleEvidence", undefined);
    updateContext("hubRoleEvidenceKey", undefined);
    updateContext("hubHintCounts", undefined);
    updateContext("hubHintBindings", undefined);
  };
  const retryItems = () => { invalidate(); setLoadingItems(true); setItems([]); setResourceRetry(value => value + 1); };
  const retryDatabases = () => {
    invalidate(); setLoadingItems(true); setLoadingDatabases(true); setItems([]); setDatabaseIds([]);
    setResourceRetry(value => value + 1); setDatabaseRetry(value => value + 1);
  };
  const retryTables = () => { invalidate(); setLoadingTables(true); setTables([]); setTableRetry(value => value + 1); };
  const failure = (label: string, detail: string, retryable: boolean, retry: () => void) => detail && <MessageBar intent="error" role="alert"><MessageBarBody>
    We couldn't load {label}.{retryable ? " Try again." : " The service marked this request as not retryable."}
    <details className="iot-solution-model-details"><summary>Details</summary>{detail}</details>
    {retryable && <Button appearance="subtle" disabled={loading} onClick={retry}>Retry</Button>}
  </MessageBarBody></MessageBar>;
  const telemetryMatches = tables.filter(table => tableRoles[table] === "iothub-telemetry-raw");
  const propertyMatches = tables.filter(table => tableRoles[table] === "iothub-properties-raw");
  const analyze = async () => {
    if (!selectionStatus.ready || analysisRunning.current) return;
    analysisRequest.current?.abort();
    invalidateProof(); updateContext("hubSourceAnalysis", undefined);
    const controller = new AbortController(); analysisRequest.current = controller;
    analysisRunning.current = true; setAnalyzing(true); setAnalysisError("");
    const owner = { key: sourceKey, revision: selectionRevision };
    const current = () => mounted.current && analysisRequest.current === controller &&
      liveSelection.current.ready && liveSelection.current.key === owner.key && liveSelection.current.revision === owner.revision;
    try {
      const token = await acquireTokenWithConsent(workloadClient, "https://kusto.kusto.windows.net/.default");
      if (controller.signal.aborted || !current()) return;
      const samples = await readPayloadSamples({ queryServiceUri, databaseName: dbName, telemetryTable, twinTable },
        token.token, SOURCE_ANALYSIS_LIMIT, controller.signal);
      if (controller.signal.aborted || !current()) return;
      const completed: SourceAnalysis = { key: owner.key, selectionRevision: owner.revision,
        limit: SOURCE_ANALYSIS_LIMIT, ...samples, roleEvidence: prepareRoleEvidence(samples.telemetry, samples.twin),
        analyzedAt: new Date().toISOString() };
      updateContext("hubSourceAnalysis", completed);
    } catch (failure: unknown) {
      if (!controller.signal.aborted && current()) {
        setAnalysisError(failure instanceof Error ? failure.message : "Couldn't analyze the selected device data. Try again.");
        controller.abort();
      }
    } finally {
      if (current()) { analysisRunning.current = false; setAnalyzing(false); analysisRequest.current = undefined; }
    }
  };

  return (
    <section className="iot-solution-step iot-solution-model-source-section" aria-labelledby="hub-source-heading">
      <h2 ref={heading} id="hub-source-heading" tabIndex={-1} className="iot-solution-step-title">Select your data source</h2>
      <Text className="iot-solution-step-description" block>Choose the eventhouse, KQL database, and tables that hold your raw IoT data. Modeling doesn't change these tables. They keep the complete raw JSON.</Text>
      <div className="iot-solution-form">
      <div className="iot-solution-field">
        <div className="iot-solution-field-header">
          <Label htmlFor="hub-model-eventhouse" className="iot-solution-field-label" required>Eventhouse</Label>
        </div>
          {failure("Eventhouses", itemError, itemRetryable, retryItems)}
          <Dropdown id="hub-model-eventhouse" placeholder={loadingItems ? "Loading..." : "Select an Eventhouse"}
            aria-busy={loadingItems} value={loadingItems ? "" : wizardContext.modelEhName || ""} selectedOptions={ehId ? [ehId] : []}
            disabled={loadingItems || !!itemError || eventhouses.length === 0}
            onOptionSelect={(_, data) => {
              if (loadingItems || itemError || !itemsCurrent) return;
              const item = eventhouses.find(candidate => candidate.id === data.optionValue);
              if (!item) return;
              invalidate();
              setLoadingDatabases(true);
              setDatabaseRetry(value => value + 1);
              updateContext("modelEhId", item.id);
              updateContext("modelEhName", item.displayName);
              for (const key of ["modelDbId", "modelDbName", "modelQueryServiceUri", "rawTelemetryTable", "rawPropertiesTable"]) {
                updateContext(key, "");
              }
              setTables([]);
            }}>
            {eventhouses.map(item => <Option key={item.id} value={item.id}>{item.displayName}</Option>)}
          </Dropdown>
          {!loadingItems && !itemError && eventhouses.length === 0 && <Text size={200}>No Eventhouses found in this workspace.</Text>}
      </div>
      <div className="iot-solution-field">
          <div className="iot-solution-field-header">
            <Label htmlFor="hub-model-database" className="iot-solution-field-label" required>KQL database</Label>
          </div>
            {failure("databases", databaseError, databaseRetryable, retryDatabases)}
            <Dropdown id="hub-model-database" placeholder={databaseFieldLoading ? "Loading..." : "Select a database"}
              aria-busy={databaseFieldLoading} value={databaseFieldLoading ? "" : dbName} selectedOptions={dbId ? [dbId] : []}
              disabled={!ehId || loadingItems || loadingDatabases || !!itemError || !!databaseError || !databasesCurrent || !databases.length}
              onOptionSelect={(_, data) => {
                if (loadingItems || loadingDatabases || itemError || databaseError || !databasesCurrent) return;
                const item = databases.find(candidate => candidate.id === data.optionValue);
                if (!item) return;
                invalidate();
                setLoadingTables(true);
                setTableRetry(value => value + 1);
                updateContext("modelDbId", item.id);
                updateContext("modelDbName", item.displayName);
                updateContext("rawTelemetryTable", "");
                updateContext("rawPropertiesTable", "");
                setTables([]);
              }}>
              {databases.map(item => <Option key={item.id} value={item.id}>{item.displayName}</Option>)}
            </Dropdown>
          {ehId && databasesCurrent && !loading && !error && databases.length === 0 && (
            <Text size={200}>No databases found in {wizardContext.modelEhName}. Choose another Eventhouse.</Text>
          )}
        </div>
      <>
            {failure("tables", tableError, tableRetryable, retryTables)}
            {tablesCurrent && metadataWarnings.length > 0 && <MessageBar intent="warning" role="status"><MessageBarBody>
              Some table metadata could not be read. These tables remain available for manual selection, without metadata hints.
              <details className="iot-solution-model-details"><summary>Table metadata warnings</summary>
                <ul>{metadataWarnings.map(warning => <li key={warning.table}>{warning.table}: {warning.message}</li>)}</ul>
              </details>
            </MessageBarBody></MessageBar>}
            <div className="iot-solution-model-source-grid iot-solution-model-source-grid--tables">
              {([
                ["rawTelemetryTable", "hub-model-telemetry", "Raw telemetry table", telemetryTable, true],
                ["rawPropertiesTable", "hub-model-twin", "Raw properties table", twinTable, false],
              ] as const).map(([key, id, label, value, telemetry]) => {
                const matches = telemetry ? telemetryMatches : propertyMatches;
                const showAll = telemetry ? showAllTelemetry : showAllProperties;
                const expansionEvent = telemetry ? telemetryExpansionEvent : propertiesExpansionEvent;
                const setOpen = telemetry ? setTelemetryOpen : setPropertiesOpen;
                return (
                <div className="iot-solution-field" key={key}>
                  <Label htmlFor={id} className="iot-solution-field-label" required>{label}</Label>
                  <Dropdown id={id} placeholder={tableFieldLoading ? "Loading..." : "Select a table"}
                    aria-busy={tableFieldLoading} value={tableFieldLoading ? "" : value}
                    selectedOptions={value ? [value] : []} disabled={!dbId || loading || !!error || !tablesCurrent || !tables.length}
                    open={telemetry ? telemetryOpen : propertiesOpen}
                    onOpenChange={(event, data) => {
                      const selectingExpansion = expansionEvent.current === event;
                      expansionEvent.current = undefined;
                      if (!data.open && selectingExpansion) return;
                      setOpen(data.open);
                    }}
                    onOptionSelect={(event, data) => {
                      if (loading || error || !dbId || !tablesCurrent) return;
                      if (data.optionValue === (telemetry ? "__show_all_telemetry__" : "__show_all_properties__")) {
                        expansionEvent.current = event;
                        if (telemetry) setShowAllTelemetry(true); else setShowAllProperties(true);
                        // Fluent closes before keyboard selection, but after mouse selection.
                        setOpen(true);
                        return;
                      }
                      if (!data.optionValue || !tables.includes(data.optionValue)) return;
                      if (data.optionValue === value) { setOpen(false); return; }
                      invalidate();
                      updateContext(key, data.optionValue);
                      setOpen(false);
                    }}>
                    {(showAll || !matches.length ? tables : matches).map(table => <Option key={table} value={table}>{table}</Option>)}
                    {!showAll && matches.length > 0 && matches.length < tables.length && <Option
                      value={telemetry ? "__show_all_telemetry__" : "__show_all_properties__"} text="">
                      <span className="iot-solution-model-show-all">Show All Tables...</span>
                    </Option>}
                  </Dropdown>
                </div>
              ); })}
            </div>
          {sameTable && <MessageBar intent="error" role="alert"><MessageBarBody>
            Choose different tables for telemetry and reported properties.
          </MessageBarBody></MessageBar>}
          {dbId && tablesCurrent && !loading && !error && tables.length === 0 && (
            <Text>No tables found in {dbName}. Choose another database.</Text>
          )}
      </>
      {!analyzing && <div>
        <Button appearance="primary" disabled={!selectionStatus.ready} onClick={analyze}>
          {analysis || analysisError ? "Validate again" : "Validate source tables"}
        </Button>
      </div>}
      {analyzing && <div className="iot-solution-loading" role="status">
        <Spinner size="medium" aria-labelledby="source-tables-validation-loading" />
        <Text id="source-tables-validation-loading" className="iot-solution-loading-text">Validating source tables...</Text>
      </div>}
      {analysisError && <MessageBar intent="error" role="alert"><MessageBarBody>{analysisError}</MessageBarBody></MessageBar>}
      {ready && !analyzing && <div className="iot-solution-success" role="status">
        <CheckmarkCircle24Filled primaryFill="var(--colorPaletteGreenForeground1)" />
        <Text className="iot-solution-success-text">Source tables validated</Text>
      </div>}
      </div>
    </section>
  );
}
