// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { canonicalJson, isJsonObject, quoteKqlIdentifier, type KqlTableResult } from "./ModelAuthoring.ts";
import { coverageReady, modelValidation, modelingFingerprint, nativeBindingIssues, processingBindingSamples, verifiedEmptyCoverage, type CoverageSnapshot, type ModelingInput } from "./HubModeling.ts";
import { buildModeledDataPlan, SYSTEM_EVENT_TIME_DOCSTRING, type CreationCommand, type ModeledDataPlan, type OutputColumn } from "./ModeledDataPlan.ts";

export interface CreationRequest {
  input: ModelingInput;
  coverage: CoverageSnapshot;
  // Policy permission for verified empty matches, not a consent flag or a raw-source proof bypass.
  allowWaiting: boolean;
  names: ModeledDataPlan["names"];
}
export interface CreationServices {
  request: (csl: string, kind: "query" | "mgmt") => Promise<KqlTableResult>;
  inspect: (input: ModelingInput) => Promise<CoverageSnapshot>;
}
export interface CreationStep {
  label: string;
  status: "running" | "passed" | "failed";
  detail?: string;
}
export interface CreatedObject { name: string; kind: "table" | "view"; tableId?: string }
export function creationOutputsConfirmed(result: CreationResult, request: Pick<CreationRequest, "input" | "names">): boolean {
  const expected = result.plannedOutputs || [];
  return expected.some(output => output.name === request.names.modeled && output.kind === "table") &&
    result.databaseName === request.input.source.databaseName &&
    (["modeled", "twinChanges", "twinState"] as const).every(key => result.names[key] === request.names[key]) &&
    expected.every(output => result.objects.some(object => object.name === output.name && object.kind === output.kind));
}
export interface CreationResult {
  operationId: string;
  names: ModeledDataPlan["names"];
  databaseName: string;
  phase: "checking" | "blocked" | "ready" | "creating" | "failed" | "completed";
  fingerprint: string;
  steps: CreationStep[];
  objects: CreatedObject[];
  plannedOutputs?: Pick<CreatedObject, "name" | "kind">[];
  awaitingData: boolean;
  error?: string;
  uncertainCommand?: string;
  coverage?: CoverageSnapshot;
}
export interface CreationAttempt {
  operationId: string;
  phase: "failed" | "completed";
  names: ModeledDataPlan["names"];
  databaseName: string;
  objects: CreatedObject[];
  error?: string;
  uncertainCommand?: string;
}
export function rememberCreationAttempt(history: CreationAttempt[], result?: CreationResult): CreationAttempt[] {
  if (!result || (result.phase !== "failed" && result.phase !== "completed")) return history;
  const attempt: CreationAttempt = {
    operationId: result.operationId, phase: result.phase,
    names: { ...result.names }, databaseName: result.databaseName, objects: result.objects.map(object => ({ ...object })),
    error: result.error, uncertainCommand: result.uncertainCommand,
  };
  return history.some(item => item.operationId === attempt.operationId)
    ? history.map(item => item.operationId === attempt.operationId ? attempt : item) : [...history, attempt];
}

export function creationFingerprint(request: CreationRequest): string {
  return canonicalJson([modelingFingerprint(request.input), request.names, request.allowWaiting, "event-time-schema-and-model-info-v8", SYSTEM_EVENT_TIME_DOCSTRING]);
}
function records(result: KqlTableResult, required: string[]): Record<string, unknown>[] {
  const names = result.Columns.map(column => column.ColumnName);
  if (new Set(names).size !== names.length || required.some(name => !names.includes(name))) {
    throw new Error(`The database returned an unsupported metadata shape; expected ${required.join(", ")}.`);
  }
  return result.Rows.map(row => Object.fromEntries(names.map((name, index) => [name, row[index]])));
}
function json(value: unknown, location: string): unknown {
  if (typeof value !== "string") return value;
  try { return JSON.parse(value); } catch { throw new Error(`${location} returned malformed JSON.`); }
}
function schemaFromMetadata(result: KqlTableResult, expectedName: string): { columns: OutputColumn[]; columnDocs: Record<string, string>; docstring: string } {
  const rows = records(result, ["TableName", "Schema", "Folder", "DocString"]);
  if (rows.length !== 1 || rows[0].TableName !== expectedName) throw new Error("The database did not return exactly the requested output schema.");
  const schema = json(rows[0].Schema, "Table schema");
  if (!isJsonObject(schema) || !Array.isArray(schema.OrderedColumns)) throw new Error("The database did not return ordered columns.");
  const docs: Record<string, string> = Object.create(null);
  const columns = schema.OrderedColumns.map(column => {
    if (!isJsonObject(column) || typeof column.Name !== "string" || typeof column.CslType !== "string") {
      throw new Error("The database returned an unsupported ordered column.");
    }
    docs[column.Name] = typeof column.DocString === "string" ? column.DocString : "";
    return { name: column.Name, type: column.CslType };
  });
  return { columns, columnDocs: docs, docstring: typeof rows[0].DocString === "string" ? rows[0].DocString : "" };
}
export function assertOrderedSchema(result: KqlTableResult, expected: OutputColumn[]): void {
  const rows = records(result, ["ColumnName", "ColumnOrdinal", "ColumnType"]);
  if (rows.length !== expected.length || rows.some((row, index) =>
    row.ColumnOrdinal !== index || row.ColumnName !== expected[index].name || row.ColumnType !== expected[index].type)) {
    throw new Error("The actual query's ordered column names/types do not match the output table. No update policy was installed for this query.");
  }
}
function assertColumns(actual: OutputColumn[], expected: OutputColumn[]): void {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) throw new Error("The created table's ordered schema does not match the requested output.");
}
function policy(result: KqlTableResult): unknown[] {
  const rows = records(result, ["Policy"]);
  if (rows.length !== 1) throw new Error("The database returned an unsupported update-policy result.");
  const value = json(rows[0].Policy, "Update policy");
  if (value === null) return [];
  if (!Array.isArray(value)) throw new Error("The database returned an unsupported update policy.");
  return value;
}
function assertPolicy(actual: unknown[], command: Extract<CreationCommand, { kind: "policy" }>): void {
  if (actual.length !== 1 || !isJsonObject(actual[0]) ||
      actual[0].Source !== command.source || actual[0].Query !== command.query ||
      actual[0].IsEnabled !== true || actual[0].IsTransactional !== false ||
      actual[0].PropagateIngestionProperties !== false) {
    throw new Error("The installed update policy did not match the requested definition.");
  }
}
const querySchema = (query: string) => `${query}\n| take 0\n| getschema`;
const message = (error: unknown) => error instanceof Error ? error.message : String(error);

/** A single in-memory operation. No retry, backfill, destructive rollback or source-policy writes. */
export class ModeledDataCreator {
  private running = false;

  async run(request: CreationRequest, services: CreationServices, create: boolean,
    onUpdate: (result: CreationResult) => void = () => undefined): Promise<CreationResult> {
    if (this.running) throw new Error("A modeled-data operation is already running.");
    this.running = true;
    const current: CreationRequest = JSON.parse(JSON.stringify(request));
    const result: CreationResult = {
      operationId: crypto.randomUUID(), names: { ...current.names }, databaseName: current.input.source.databaseName,
      phase: "checking", fingerprint: creationFingerprint(current), steps: [], objects: [],
      awaitingData: current.coverage?.telemetry.payloads.length === 0,
    };
    const publish = () => onUpdate({ ...result, steps: result.steps.map(step => ({ ...step })), objects: result.objects.map(object => ({ ...object })) });
    const step = async <T>(label: string, action: () => Promise<T>): Promise<T> => {
      const status: CreationStep = { label, status: "running" }; result.steps.push(status); publish();
      try { const value = await action(); status.status = "passed"; publish(); return value; }
      catch (error: unknown) { status.status = "failed"; status.detail = message(error); publish(); throw error; }
    };
    const read = (csl: string) => services.request(csl, "mgmt");
    const showTable = async (name: string) => {
      const rows = records(await read(`.show table ${quoteKqlIdentifier(name)} details`),
        ["TableName", "TableId", "Folder", "StreamingIngestionPolicy"]);
      if (rows.length !== 1 || rows[0].TableName !== name || typeof rows[0].TableId !== "string" || !rows[0].TableId) {
        throw new Error(`Could not verify table identity for ${name}.`);
      }
      return rows[0];
    };
    const sourceIds = new Map<string, string>();
    const ownership = new Map<string, string>();
    const createdViews = new Set<string>();
    let mutationAttempted = false;
    try {
      await step("Validate current model, roles and processing scope", async () => {
        const errors = modelValidation(current.input.model);
        if (errors.length) throw new Error(errors.join(" "));
        if (!coverageReady(current.input, current.coverage, current.allowWaiting)) throw new Error("Check the current processing scope before creating modeled data.");
      });
      const fresh = await step("Recheck processing samples and native identities", () => services.inspect(current.input));
      if (!coverageReady(current.input, fresh, current.allowWaiting)) throw new Error("Matching telemetry changed. Recheck the current model coverage.");
      if (!fresh.telemetry.payloads.length && !verifiedEmptyCoverage(current.input, fresh)) {
        throw new Error("No matching telemetry was returned and the current raw-source proof is empty, invalid or incomplete. Verify both raw streams before creating modeled data.");
      }
      result.coverage = fresh; result.awaitingData = fresh.telemetry.payloads.length === 0;
      const plan = buildModeledDataPlan(current.input, fresh, current.allowWaiting, current.names);
      const needsTwins = plan.commands.some(command => command.kind === "view");
      const enrichesTwins = current.input.fields.some(field => field.included && field.role === "twin");
      const evidence = processingBindingSamples(fresh);
      await step("Verify native device and component bindings", async () => {
        const issues = nativeBindingIssues(evidence.telemetry, evidence.twin, needsTwins);
        if (issues.length) throw new Error(issues.join(" "));
      });
      const outputs = plan.commands.filter((command): command is Extract<CreationCommand, { kind: "table" | "view" }> =>
        command.kind === "table" || command.kind === "view");
      result.plannedOutputs = outputs.map(command => ({ name: command.target, kind: command.kind }));
      await step("Check database creation permissions", async () => {
        const roles = records(await read(`.show database ${quoteKqlIdentifier(current.input.source.databaseName)} principal roles`), ["Role"]);
        const allowed = roles.some(row => typeof row.Role === "string" &&
          (/\badmins?$/i.test(row.Role) || (!needsTwins && /\busers?$/i.test(row.Role))));
        if (!allowed) throw new Error(needsTwins
          ? "Database Admin permission is required to create the last-known-state materialized view. No roles were changed."
          : "Database User or Admin permission is required to create modeled data. No roles were changed.");
      });
      const assertUnused = async (name?: string) => {
        const [tables, views, functions] = await Promise.all([
          read(".show tables | project TableName"), read(".show materialized-views | project Name"), read(".show functions | project Name"),
        ]);
        const names = [
          ...records(tables, ["TableName"]).map(row => row.TableName),
          ...records(views, ["Name"]).map(row => row.Name), ...records(functions, ["Name"]).map(row => row.Name),
        ];
        if (names.some(value => typeof value !== "string" || !value)) throw new Error("The database returned an invalid object-name inventory.");
        const existing = new Set(names);
        const collision = outputs.find(output => (!name || output.target === name) && existing.has(output.target));
        if (collision) throw new Error(`"${collision.target}" already exists. Choose unused output names; existing objects and policies will not be changed.`);
      };
      await step("Check output names against tables, views and functions", () => assertUnused());
      await step("Verify raw schemas and unchanged source-policy prerequisites", async () => {
        for (const name of new Set(needsTwins
          ? [current.input.source.telemetryTable, current.input.source.twinTable] : [current.input.source.telemetryTable])) {
          const source = await showTable(name); sourceIds.set(name, String(source.TableId));
          const schema = schemaFromMetadata(await read(`.show table ${quoteKqlIdentifier(name)} schema as json`), name);
          for (const [column, type] of [["data", "string"], ["headers", "dynamic"], ["user_headers", "dynamic"]]) {
            if (!schema.columns.some(item => item.name === column && (item.type === type ||
                column === "data" && item.type === "dynamic"))) throw new Error(
              `Raw table ${name} must contain ${column}:${column === "data" ? "string or dynamic" : type}.`);
          }
          if (enrichesTwins) {
            if (name === current.input.source.telemetryTable) {
              const streaming = json(source.StreamingIngestionPolicy, `Effective streaming policy for ${name}`);
              if (!isJsonObject(streaming) || streaming.IsEnabled !== false) {
                throw new Error(`Lookup enrichment requires a verified disabled streaming-ingestion policy on ${name}. This operation will not alter the raw policy; resolve this prerequisite separately.`);
              }
            }
            if (policy(await read(`.show table ${quoteKqlIdentifier(name)} policy update`)).length) {
              throw new Error(`${name} has upstream update-policy dependencies. Cascading lookup/streaming prerequisites need review; no existing policy will be changed.`);
            }
          }
        }
      });
      for (const command of plan.commands) if (command.kind === "policy") {
        const target = outputs.find(output => output.target === command.target);
        if (!target) throw new Error("The generated policy has no output schema.");
        // The actual view does not exist yet. This verifies planned expressions; its real schema is rechecked before installation.
        const plannedQuery = needsTwins && command.target === plan.names.modeled
          ? command.query.replace(quoteKqlIdentifier(plan.names.twinState),
            "(datatable(deviceId:string, propertyName:string, event_time:datetime, propertyValue:dynamic)[])")
          : command.query;
        await step(`Validate ordered projection for ${command.target}`, async () =>
          assertOrderedSchema(await services.request(querySchema(plannedQuery), "query"), target.columns));
      }
      result.phase = "ready"; publish();
      if (!create) return result;
      result.phase = "creating"; publish();

      const assertOwned = async (name: string) => {
        const table = await showTable(name);
        if (!ownership.has(name) || table.TableId !== ownership.get(name)) {
          throw new Error(`Ownership of ${name} changed or could not be verified. No further changes will be made.`);
        }
      };
      const assertSources = async () => {
        for (const [name, id] of sourceIds) {
          const actual = await showTable(name);
          if (actual.TableId !== id) throw new Error(`Raw source ${name} was replaced during creation. Creation stopped.`);
          if (enrichesTwins) {
            const lookupSource = name === current.input.source.telemetryTable;
            const streaming = lookupSource ? json(actual.StreamingIngestionPolicy, `Effective streaming policy for ${name}`) : undefined;
            if ((lookupSource && (!isJsonObject(streaming) || streaming.IsEnabled !== false)) ||
                policy(await read(`.show table ${quoteKqlIdentifier(name)} policy update`)).length) {
              throw new Error(`Raw policy prerequisites changed on ${name}. No raw policy was altered; creation stopped.`);
            }
          }
        }
      };
      const verifyView = async (command: Extract<CreationCommand, { kind: "view" }>) => {
        const rows = records(await read(`.show materialized-view ${quoteKqlIdentifier(command.target)}`),
          ["Name", "SourceTable", "Query", "IsEnabled", "Folder", "DocString"]);
        if (!createdViews.has(command.target) || rows.length !== 1 || rows[0].Name !== command.target || rows[0].SourceTable !== command.source ||
            typeof rows[0].Query !== "string" || rows[0].Query.trim() !== command.query.trim() ||
            rows[0].IsEnabled !== true || (rows[0].Folder !== "" && rows[0].Folder !== null)) {
          throw new Error(`The materialized view ${command.target} did not match the requested definition.`);
        }
        return rows[0];
      };
      for (const command of plan.commands) {
        const assertCommandOwned = async () => {
          if (command.kind !== "viewDoc") { await assertOwned(command.target); return; }
          const view = outputs.find((output): output is Extract<CreationCommand, { kind: "view" }> =>
            output.kind === "view" && output.target === command.target);
          if (!view) throw new Error("The generated view docstring has no owned output.");
          await assertOwned(view.source); await verifyView(view);
        };
        await step(command.label, async () => {
          if (command.kind === "table") {
            await assertUnused(command.target);
            mutationAttempted = true; result.uncertainCommand = command.label;
            // .set fails on an existing name; .create table would silently adopt it.
            if (records(await read(command.csl), ["ExtentId", "RowCount"]).length) {
              throw new Error(`The creation of ${command.target} did not return an empty-ingestion result. Ownership is unconfirmed.`);
            }
            const created = await showTable(command.target);
            if (created.Folder !== "" && created.Folder !== null) {
              throw new Error(`The newly created table ${command.target} was not returned in the database root. Creation stopped without changing its folder.`);
            }
            ownership.set(command.target, String(created.TableId));
            result.objects.push({ name: command.target, kind: "table", tableId: String(created.TableId) });
            assertColumns(schemaFromMetadata(await read(`.show table ${quoteKqlIdentifier(command.target)} schema as json`), command.target).columns, command.columns);
            if (policy(await read(`.show table ${quoteKqlIdentifier(command.target)} policy update`)).length) {
              throw new Error(`The new table ${command.target} has an unexpected policy. It will not be overwritten.`);
            }
          } else if (command.kind === "view") {
            await assertUnused(command.target); await assertOwned(command.source);
            const ingestionRows = records(await read(`.show table ${quoteKqlIdentifier(command.source)} policy ingestiontime`), ["Policy"]);
            const ingestion = ingestionRows.length === 1 ? json(ingestionRows[0].Policy, "Ingestion-time policy") : undefined;
            if (ingestion !== true && (!isJsonObject(ingestion) || ingestion.IsEnabled !== true)) throw new Error("The new twin-change table requires a verified enabled ingestion-time policy. No raw policy was changed.");
            await assertOwned(command.source);
            mutationAttempted = true; result.uncertainCommand = command.label;
            await read(command.csl);
            createdViews.add(command.target);
            await verifyView(command); result.objects.push({ name: command.target, kind: "view" });
            assertOrderedSchema(await services.request(querySchema(quoteKqlIdentifier(command.target)), "query"), command.columns);
          } else {
            await assertCommandOwned();
            if (command.kind === "policy") {
              await assertSources();
              for (const view of outputs) if (view.kind === "view") {
                await assertOwned(view.source);
                await verifyView(view);
              }
              const target = outputs.find(output => output.target === command.target)!;
              assertOrderedSchema(await services.request(querySchema(command.query), "query"), target.columns);
              if (policy(await read(`.show table ${quoteKqlIdentifier(command.target)} policy update`)).length) {
                throw new Error(`An update policy appeared on ${command.target}. It will not be overwritten.`);
              }
            }
            await assertCommandOwned();
            mutationAttempted = true; result.uncertainCommand = command.label;
            await read(command.csl);
            if (command.kind === "policy") assertPolicy(policy(await read(`.show table ${quoteKqlIdentifier(command.target)} policy update`)), command);
            if (command.kind === "tableDoc" &&
                schemaFromMetadata(await read(`.show table ${quoteKqlIdentifier(command.target)} schema as json`), command.target).docstring !== command.docstring) {
              throw new Error(`Object docstring for ${command.target} could not be verified.`);
            }
            if (command.kind === "viewDoc") {
              const view = outputs.find((output): output is Extract<CreationCommand, { kind: "view" }> =>
                output.kind === "view" && output.target === command.target);
              if (!view || (await verifyView(view)).DocString !== command.docstring) {
                throw new Error(`Object docstring for ${command.target} could not be verified.`);
              }
            }
          }
          result.uncertainCommand = undefined;
        });
      }
      await step("Verify all created definitions and native semantic documentation", async () => {
        for (const command of outputs) {
          if (command.kind === "table") {
            await assertOwned(command.target);
            assertColumns(schemaFromMetadata(await read(`.show table ${quoteKqlIdentifier(command.target)} schema as json`), command.target).columns, command.columns);
          } else await verifyView(command);
        }
        const documented = schemaFromMetadata(await read(`.show table ${quoteKqlIdentifier(plan.names.modeled)} schema as json`), plan.names.modeled);
        if (documented.docstring !== plan.documentation.table ||
            Object.entries(plan.documentation.columns).some(([name, doc]) => documented.columnDocs[name] !== doc)) {
          throw new Error("Native table/column documentation could not be verified. Creation is incomplete.");
        }
        for (const command of plan.commands) if (command.kind === "policy") {
          assertPolicy(policy(await read(`.show table ${quoteKqlIdentifier(command.target)} policy update`)), command);
        }
        for (const command of plan.commands) {
          if (command.kind === "tableDoc" &&
              schemaFromMetadata(await read(`.show table ${quoteKqlIdentifier(command.target)} schema as json`), command.target).docstring !== command.docstring) {
            throw new Error(`Object docstring for ${command.target} could not be verified.`);
          }
          if (command.kind === "viewDoc") {
            const view = outputs.find((output): output is Extract<CreationCommand, { kind: "view" }> =>
              output.kind === "view" && output.target === command.target)!;
            if ((await verifyView(view)).DocString !== command.docstring) throw new Error(`Object docstring for ${command.target} could not be verified.`);
          }
        }
        await assertSources();
      });
      result.phase = "completed"; result.uncertainCommand = undefined; publish();
      return result;
    } catch (error: unknown) {
      result.phase = mutationAttempted ? "failed" : "blocked"; result.error = message(error); publish();
      return result;
    } finally {
      this.running = false;
    }
  }
}
