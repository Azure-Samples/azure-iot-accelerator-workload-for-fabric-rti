// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

// Synthetic in-memory Kusto contract fixture. Never performs fetch or creates real resources.
import { decodeDoSamples, type KqlTableResult } from "../app/items/IoTSolutionItem/ModelAuthoring.ts";
import { fromThingModel, inspectFieldEvidence, modelingFingerprint, setFieldRole, type CoverageSnapshot } from "../app/items/IoTSolutionItem/HubModeling.ts";
import { resolveThingModel } from "../app/items/IoTSolutionItem/ThingModelParser.ts";
import { buildModeledDataPlan, type CreationCommand, type OutputColumn } from "../app/items/IoTSolutionItem/ModeledDataPlan.ts";
import type { CreationRequest, CreationServices } from "../app/items/IoTSolutionItem/ModeledDataCreation.ts";

export function mockResult(columns: string[], rows: unknown[][]): KqlTableResult {
  return { Columns: columns.map(ColumnName => ({ ColumnName })), Rows: rows };
}
export function creationRequest(waiting = false, enrichTwins = true): CreationRequest {
  const document = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel", id: "urn:mock:device", title: "Device",
    properties: { state: { type: "string", description: "Reported state." }, temperature: { type: "number", unit: "degC" } } };
  const model = fromThingModel(resolveThingModel(document));
  const input = {
    model, fields: model.fields.map(field => ({ ...setFieldRole(field, field.path[0] === "state" ? "twin" : "telemetry"),
      included: field.path[0] !== "state" || enrichTwins })),
    source: { queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databaseName: "Database", telemetryTable: "Telemetry", twinTable: "Twins" },
    scope: { mode: waiting ? "reference" as const : "data" as const, additional: "specified" as const, operator: "and" as const,
      telemetry: "", twin: "", reference: { id: document.id } }, limit: 100,
  };
  const sample = (stream: "telemetry" | "twin") => decodeDoSamples([{
    data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
      event_kind: stream === "telemetry" ? "measurement" : "lifecycle", event_time: "2026-10-01T12:00:00Z",
      payload: { value: stream === "telemetry" ? { temperature: 22 } : { properties: { reported: { state: null } } } } }),
    headers: { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "", IoTSubject: "" },
    user_headers: {},
  }], stream, 100);
  const sourceProbe = { telemetry: sample("telemetry"), twin: sample("twin") };
  const telemetry = waiting ? decodeDoSamples([], "telemetry", 100) : sourceProbe.telemetry;
  const twin = waiting ? decodeDoSamples([], "twin", 100) : sourceProbe.twin;
  const coverage: CoverageSnapshot = { fingerprint: modelingFingerprint(input), kind: "processing",
    telemetry, twin, sourceProbe, nativeIssues: [], fields: inspectFieldEvidence(input.fields, telemetry, twin), checkedAt: "2026-10-01" };
  return { input, coverage, allowWaiting: waiting, names: { modeled: "Modeled_Output", twinChanges: "Changes_Output", twinState: "State_Output" } };
}
export interface MockTable {
  id: string; folder: string; columns: OutputColumn[]; doc: string; columnDocs: Record<string, string>;
  policy: unknown[] | null; streaming: boolean | null;
}
export function createCreationMock(request: CreationRequest) {
  const plan = buildModeledDataPlan(request.input, request.coverage, request.allowWaiting, request.names);
  const tables = new Map<string, MockTable>();
  for (const name of [request.input.source.telemetryTable, request.input.source.twinTable]) {
    tables.set(name, { id: `id-${name}`, folder: "", columns: [{ name: "data", type: "string" }, { name: "headers", type: "dynamic" },
      { name: "user_headers", type: "dynamic" }], doc: "", columnDocs: {}, policy: null, streaming: false });
  }
  const views = new Map<string, { folder: string; doc: string; command: Extract<CreationCommand, { kind: "view" }> }>();
  const functions: string[] = [];
  const calls: { csl: string; kind: string }[] = [], writes: string[] = [];
  const options = {
    role: "Database Database Admin",
    fail: undefined as ((csl: string, kind: string) => Error | undefined) | undefined,
    before: undefined as ((csl: string) => void) | undefined,
    schemaMismatch: undefined as string | undefined,
    policyReadbackMismatch: false, documentationMismatch: false,
    inspect: undefined as (() => Promise<CoverageSnapshot>) | undefined,
  };
  const schema = (columns: OutputColumn[]) => mockResult(["ColumnName", "ColumnOrdinal", "DataType", "ColumnType"],
    columns.map((column, index) => [column.name, index, "", column.type]));
  const service: CreationServices = {
    inspect: async () => options.inspect ? options.inspect() : request.coverage,
    request: async (csl, kind) => {
      calls.push({ csl, kind }); options.before?.(csl);
      const failure = options.fail?.(csl, kind); if (failure) throw failure;
      if (kind === "query") {
        if (!csl.endsWith("| getschema")) throw new Error("Mock only permits generated schema queries.");
        const view = [...views.values()].find(value => csl.startsWith(`['${value.command.target}']\n`));
        const policy = plan.commands.find((command): command is Extract<CreationCommand, { kind: "policy" }> =>
          command.kind === "policy" && (csl === `${command.query}\n| take 0\n| getschema` ||
            csl === `${command.query.replace(`['${plan.names.twinState}']`, "(datatable(deviceId:string, propertyName:string, event_time:datetime, propertyValue:dynamic)[])")}\n| take 0\n| getschema`));
        const output = view?.command || plan.commands.find((command): command is Extract<CreationCommand, { kind: "table" }> =>
          command.kind === "table" && command.target === policy?.target);
        if (!output) throw new Error("Unexpected generated schema query.");
        const columns = output.columns.map(column => ({ ...column }));
        if (options.schemaMismatch === output.target) columns.reverse();
        return schema(columns);
      }
      if (csl.includes(" principal roles")) return mockResult(["Role"], [[options.role]]);
      if (csl === ".show tables | project TableName") return mockResult(["TableName"], [...tables.keys()].map(name => [name]));
      if (csl === ".show materialized-views | project Name") return mockResult(["Name"], [...views.keys()].map(name => [name]));
      if (csl === ".show functions | project Name") return mockResult(["Name"], functions.map(name => [name]));
      const name = /\['([^']+)'\]/.exec(csl)?.[1];
      if (csl.startsWith(".show table ") && name) {
        const table = tables.get(name); if (!table) throw new Error(`Unknown mock table ${name}.`);
        if (csl.endsWith(" details")) return mockResult(["TableName", "TableId", "Folder", "StreamingIngestionPolicy"],
          [[name, table.id, table.folder, table.streaming === null ? null : JSON.stringify({ IsEnabled: table.streaming })]]);
        if (csl.endsWith(" schema as json")) return mockResult(["TableName", "Schema", "DatabaseName", "Folder", "DocString"],
          [[name, JSON.stringify({ OrderedColumns: table.columns.map(column => ({
            Name: column.name, CslType: column.type, DocString: options.documentationMismatch ? "mismatch" : table.columnDocs[column.name] || "",
          })) }), request.input.source.databaseName, table.folder, table.doc]]);
        if (csl.endsWith(" policy update")) {
          const value = options.policyReadbackMismatch && table.policy?.length ? [] : table.policy;
          return mockResult(["Policy"], [[JSON.stringify(value)]]);
        }
        if (csl.endsWith(" policy ingestiontime")) return mockResult(["Policy"], [[JSON.stringify({ IsEnabled: true })]]);
      }
      if (csl.startsWith(".show materialized-view ") && name) {
        const view = views.get(name); if (!view) throw new Error("Unknown mock view.");
        return mockResult(["Name", "SourceTable", "Query", "IsEnabled", "Folder", "DocString"],
          [[name, view.command.source, view.command.query, true, view.folder, view.doc]]);
      }
      if (csl.startsWith(".set ") && name) {
        writes.push(csl);
        const command = plan.commands.find((item): item is Extract<CreationCommand, { kind: "table" }> => item.kind === "table" && item.target === name);
        if (!command || command.csl !== csl || !csl.endsWith(")[]")) throw new Error("Unexpected empty-table creation.");
        if (tables.has(name) || views.has(name) || functions.includes(name)) throw new Error(`Table ${name} already exists.`);
        tables.set(name, { id: `new-${name}`, folder: "", columns: command.columns,
          doc: "", columnDocs: {}, policy: null, streaming: false });
        return mockResult(["ExtentId", "OriginalSize", "ExtentSize", "CompressedSize", "IndexSize", "RowCount"], []);
      }
      if (csl.startsWith(".create materialized-view ")) {
        writes.push(csl);
        const command = plan.commands.find((item): item is Extract<CreationCommand, { kind: "view" }> =>
          item.kind === "view" && csl.includes(` ${"['" + item.target + "']"} on table `));
        if (!command || command.csl !== csl || views.has(command.target) || tables.has(command.target) ||
            functions.includes(command.target)) throw new Error("Unexpected/existing materialized view.");
        views.set(command.target, { command, folder: "", doc: "" });
        return mockResult(["Name"], [[command.target]]);
      }
      const command = plan.commands.find(item => item.csl === csl);
      if (command?.kind === "viewDoc" && name) {
        const view = views.get(name); if (!view) throw new Error("Unknown view documentation target.");
        writes.push(csl); view.doc = command.docstring;
        return mockResult(["Name"], [[name]]);
      }
      if (command && name && ["policy", "tableDoc", "columnDocs"].includes(command.kind)) {
        writes.push(csl);
        const table = tables.get(name); if (!table) throw new Error("Unknown mutation target.");
        if (command.kind === "policy") table.policy = [{ IsEnabled: true, Source: command.source,
          Query: command.query, IsTransactional: false, PropagateIngestionProperties: false }];
        if (command.kind === "tableDoc") table.doc = command.docstring;
        if (command.kind === "columnDocs") table.columnDocs = plan.documentation.columns;
        return mockResult(["TableName"], [[name]]);
      }
      throw new Error(`Unexpected mock request: ${csl}`);
    },
  };
  return { service, tables, views, functions, calls, writes, options, plan };
}
