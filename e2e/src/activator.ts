import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { readFile, readdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import type { Page, Request, Response, Route } from "@playwright/test";
import { Cloud, object, text, redactSensitiveText } from "./cloud.js";
import { artifacts, requireGuid } from "./config.js";
import { FABRIC, connections, deleteFabricResource, workspaceItems } from "./cleanup.js";
import { operationFromHeaders } from "./fabric.js";
import { moveOwnedItem } from "./folders.js";
import { modelFixture } from "./models.js";
import { modeledNames } from "./modeling.js";
import { query, type DeviceFleet } from "./data.js";
import type { IngestionPath, PathId } from "./paths.js";
import { updatePlan, type RunPlan } from "./plan.js";
import { expectedModeledHubKeys } from "./scoping.js";
import { itemDefinition } from "./experience-verification.js";

export function activatorInputSchema(parts: readonly { path: string; value: unknown }[], destination: Record<string, unknown>) {
  const document = object(parts.find(part => part.path === "eventstream.json")?.value, "native Eventstream definition");
  assert.ok(Array.isArray(document.destinations) && document.destinations.length === 1, "[native-activator] Expected one authoring destination.");
  const authoring = object(document.destinations[0], "authoring destination");
  assert.equal(authoring.id, destination.id);
  assert.equal(authoring.type, "Activator");
  const properties = object(authoring.properties, "authoring destination properties");
  const runtime = object(destination.properties, "runtime destination properties");
  assert.equal(properties.workspaceId, runtime.workspaceId);
  assert.equal(properties.itemId, runtime.itemId);
  assert.ok(Array.isArray(authoring.inputSchemas) && authoring.inputSchemas.length === 1, "[native-activator] Expected one authoring input schema.");
  return object(authoring.inputSchemas[0], "Activator input schema");
}

type Kind = "Eventstream" | "Reflex" | "Connection";
interface Resource {
  kind: Kind;
  name: string;
  requested: boolean;
  id?: string;
  status?: number;
  operation?: string;
  deleted?: boolean;
  cluster?: string;
}
interface Journal {
  runId: string;
  workspaceId: string;
  attempt: string;
  path: PathId;
  baselineItems: string[];
  baselineConnections: string[];
  resources: Resource[];
}

export class ActivatorAttempt {
  private readonly requests = new WeakMap<Request, Resource>();
  private readonly pending = new Set<Promise<void>>();
  private readonly errors: Error[] = [];
  private writes = Promise.resolve();
  private incomingName: string | undefined;

  private constructor(
    readonly cloud: Cloud, readonly plan: RunPlan, readonly journal: Journal, readonly file: string,
    private readonly existingConnectionNames: Set<string>,
  ) {}

  static async begin(cloud: Cloud, plan: RunPlan, path: PathId): Promise<ActivatorAttempt> {
    const attempt = randomUUID(), prefix = plan.azure.resourceGroup.replace(/^fabric-/, "");
    const items = await workspaceItems(cloud, plan), beforeConnections = await connections(cloud);
    const journal: Journal = {
      runId: plan.runId, workspaceId: plan.target.workspaceId, attempt, path,
      baselineItems: items.map(item => requireGuid(text(item.id, "baseline item ID"), "baseline item ID")),
      baselineConnections: beforeConnections.map(item => requireGuid(text(item.id, "baseline connection ID"), "baseline connection ID")),
      resources: [
        { kind: "Eventstream", name: `${prefix}-${path}-act-stream-${attempt.slice(0, 8)}`, requested: false },
        { kind: "Reflex", name: `${prefix}-${path}-activator-${attempt.slice(0, 8)}`, requested: false },
      ],
    };
    const result = new ActivatorAttempt(cloud, plan, journal,
      join(artifacts, "runs", plan.runId, `activator-attempt-${attempt}.json`),
      new Set(beforeConnections.map(item => text(item.displayName, "connection name"))));
    await result.persist();
    return result;
  }

  get streamName(): string { return text(this.journal.resources.find(resource => resource.kind === "Eventstream")?.name, "planned stream"); }
  get activatorName(): string { return text(this.journal.resources.find(resource => resource.kind === "Reflex")?.name, "planned Activator"); }
  get activatorId(): string { return text(this.journal.resources.find(resource => resource.kind === "Reflex")?.id, "verified Activator ID"); }
  get incomingStreamName(): string { return text(this.incomingName, "verified Activator incoming stream name"); }

  private persist(): Promise<void> {
    const snapshot = JSON.stringify(this.journal, null, 2);
    this.writes = this.writes.then(() => writeFile(this.file, snapshot, { mode: 0o600 }));
    return this.writes;
  }

  private track(action: Promise<void>): void {
    const task = action.catch(error => {
      const failure = new Error(redactSensitiveText(error instanceof Error ? error.message : String(error)));
      this.errors.push(failure);
      console.error(`[activator] ${failure.message}`);
    });
    this.pending.add(task);
    void task.then(() => this.pending.delete(task));
  }

  assertHealthy(): void {
    if (this.errors.length) throw new AggregateError(this.errors, "[activator] Creation ownership/response tracking failed.");
  }

  async observe(page: Page): Promise<void> {
    const workspace = `/v1/workspaces/${this.plan.target.workspaceId}`;
    await page.route(`${FABRIC}/v1/**`, async (route: Route) => {
      const request = route.request(), url = new URL(request.url());
      if (request.method() !== "POST" || !["/v1/connections", `${workspace}/items`, `${workspace}/eventstreams`].includes(url.pathname)) {
        await route.continue();
        return;
      }
      try {
        const body = object(request.postDataJSON(), "Activator creation request");
        let resource: Resource;
        if (url.pathname === "/v1/connections") {
          const name = text(body.displayName, "new connection name");
          const details = object(body.connectionDetails, "connection details");
          const credentials = object(object(body.credentialDetails, "credential details").credentials, "connection credentials");
          assert.equal(body.connectivityType, "ShareableCloud");
          assert.equal(details.type, "AzureDataExplorer");
          assert.equal(credentials.credentialType, "WorkspaceIdentity");
          assert.ok(Array.isArray(details.parameters));
          assert.ok(details.parameters.some(value => {
            const parameter = object(value, "connection parameter");
            return parameter.name === "cluster" && parameter.value === this.plan.observed.queryServiceUri;
          }));
          if (!/^KQL-[A-Za-z0-9_-]+$/.test(name) || this.existingConnectionNames.has(name)) {
            throw new Error("[ownership] Refusing to create/reuse a pre-existing or unexpected global connection.");
          }
          if (this.journal.resources.some(value => value.kind === "Connection")) throw new Error("[ownership] More than one connection creation attempted.");
          resource = { kind: "Connection", name, requested: false, cluster: text(this.plan.observed.queryServiceUri, "owned Kusto cluster") };
          this.journal.resources.push(resource);
        } else {
          const kind = url.pathname.endsWith("/eventstreams") ? "Eventstream" : "Reflex";
          const planned = this.journal.resources.find(value => value.kind === kind);
          if (!planned) throw new Error("[ownership] Missing planned Activator resource.");
          resource = planned;
          assert.equal(body.displayName, resource.name);
          if (kind === "Reflex") assert.equal(body.type, "Reflex");
        }
        if (resource.requested) throw new Error("[ownership] Repeated Activator resource creation is not permitted.");
        resource.requested = true;
        this.requests.set(request, resource);
        // Persist ownership intent before letting the unchanged native request leave the browser.
        await this.persist();
        await route.continue();
      } catch (error) {
        this.track(Promise.reject(error));
        await route.abort("blockedbyclient");
      }
    });
    page.on("response", response => {
      const resource = this.requests.get(response.request());
      if (resource) this.track(this.capture(response, resource));
    });
  }

  private async capture(response: Response, resource: Resource): Promise<void> {
    resource.status = response.status();
    if (response.status() === 202) {
      resource.operation = operationFromHeaders(new Headers(await response.allHeaders()));
    } else if (response.ok()) {
      resource.id = requireGuid(text(object(await response.json(), "created resource").id, "created ID"), "created ID");
    }
    await this.persist();
  }

  private async settle(): Promise<void> {
    await Promise.all(this.pending);
    for (const resource of this.journal.resources) {
      if (!resource.operation || resource.id || resource.deleted) continue;
      const deadline = Date.now() + 180_000;
      let completed = false;
      while (Date.now() < deadline) {
        const state = object(await this.cloud.fabric(resource.operation), "creation operation");
        if (state.status === "Succeeded") {
          const result = object(await this.cloud.fabric(`${resource.operation}/result`), "created resource");
          resource.id = requireGuid(text(result.id, "created resource ID"), "created resource ID");
          completed = true;
          break;
        }
        if (state.status === "Failed" || state.status === "Canceled") { completed = true; break; }
        await pause(2000);
      }
      if (!completed) throw new Error("[cleanup] Activator creation is still pending; cleanup cannot yet be certified.");
    }
    await this.persist();
  }

  async reconcile(): Promise<void> {
    await this.settle();
    const items = await workspaceItems(this.cloud, this.plan), global = await connections(this.cloud);
    for (const resource of this.journal.resources) {
      if (!resource.requested) continue;
      const list = resource.kind === "Connection" ? global : items;
      const baseline = resource.kind === "Connection" ? this.journal.baselineConnections : this.journal.baselineItems;
      const matches = list.filter(item => item.displayName === resource.name && (resource.kind === "Connection" || item.type === resource.kind));
      if (matches.length > 1) throw new Error("[ownership] Ambiguous Activator cleanup resource name.");
      if (resource.deleted) {
        if (!resource.id) throw new Error("[ownership] A confirmed deletion has no recorded resource ID.");
        if (matches.length || list.some(item => item.id === resource.id)) {
          throw new Error("[ownership] A previously deleted Activator resource or its name reappeared; refusing adoption.");
        }
        continue;
      }
      const item = matches[0];
      if (item) {
        const id = requireGuid(text(item.id, "observed resource ID"), "observed resource ID");
        if (baseline.includes(id) || resource.id && resource.id !== id) throw new Error("[ownership] Activator resource collides with pre-existing ownership.");
        resource.id = id;
      } else if (resource.status === undefined) {
        throw new Error("[cleanup] A submitted creation has no native response or resource readback; retain the journal for cleanup.");
      }
    }
    await this.persist();
  }

  async verify(path: IngestionPath, fleet: DeviceFleet): Promise<void> {
    await this.reconcile();
    this.assertHealthy();
    const stream = this.journal.resources.find(resource => resource.kind === "Eventstream");
    const activator = this.journal.resources.find(resource => resource.kind === "Reflex");
    const connection = this.journal.resources.find(resource => resource.kind === "Connection");
    if (!stream?.id || !activator?.id || !connection?.id) throw new Error("[native-activator] Expected one new stream, Activator and connection.");
    await moveOwnedItem(this.cloud, this.plan, stream.id, stream.name);
    await moveOwnedItem(this.cloud, this.plan, activator.id, activator.name);
    const nativeConnection = object(await this.cloud.request(`${FABRIC}/v1/connections/${connection.id}`, FABRIC), "native connection");
    assert.equal(nativeConnection.displayName, connection.name);
    assert.equal(object(nativeConnection.connectionDetails, "native connection details").type, "AzureDataExplorer");
    const workspace = object(await this.cloud.request(`${FABRIC}/v1/workspaces/${this.plan.target.workspaceId}`, FABRIC), "workspace");
    assert.ok(workspace.workspaceIdentity, "[native-identity] Workspace identity was not enabled.");
    const topology = object(await this.cloud.fabric(`/workspaces/${this.plan.target.workspaceId}/eventstreams/${stream.id}/topology`), "Activator stream topology");
    assert.ok(Array.isArray(topology.sources) && topology.sources.length === 1);
    assert.ok(Array.isArray(topology.destinations) && topology.destinations.length === 1);
    const source = object(topology.sources[0], "Activator source"), destination = object(topology.destinations[0], "Activator destination");
    assert.equal(source.type, "AzureDataExplorer");
    const sourceProperties = object(source.properties, "source properties");
    assert.equal(sourceProperties.dataConnectionId, connection.id);
    assert.equal(sourceProperties.databaseName, this.plan.fabric.databaseName);
    assert.equal(sourceProperties.tableNames, modeledNames(path).modeled);
    assert.equal(destination.type, "Activator");
    const destinationProperties = object(destination.properties, "destination properties");
    assert.equal(destinationProperties.workspaceId, this.plan.target.workspaceId);
    assert.equal(destinationProperties.itemId, activator.id);
    const inputSchema = activatorInputSchema(await itemDefinition(this.cloud, this.plan, stream.id), destination);
    this.incomingName = text(inputSchema.name, "Activator incoming stream name");
    assert.equal(this.incomingName, `${stream.name}-stream`);
    const schema = object(inputSchema.schema, "stream schema");
    assert.ok(Array.isArray(schema.columns));
    const types = { string: "Nvarchar(max)", real: "Float", dynamic: "Record" };
    assert.deepEqual(schema.columns.map(value => {
      const column = object(value, "Activator column"); return [column.name, column.type];
    }), [["deviceId", "Nvarchar(max)"], ["event_time", "DateTime"], ...modelFixture(path.id).fields.map(field => [field.column, types[field.type]])]);
    const start = new Date().toISOString();
    await fleet.sendRound();
    const devices = fleet.devices.filter(device => expectedModeledHubKeys(path.id, "telemetry").includes(device.hubKey));
    assert.equal(devices.length, expectedModeledHubKeys(path.id, "telemetry").length);
    const deadline = Date.now() + 180_000;
    let rows: Record<string, unknown>[] = [];
    while (Date.now() < deadline) {
      fleet.assertHealthy();
      rows = await query(this.cloud, this.plan, `${modeledNames(path).modeled} | where event_time > datetime(${start}) | summarize rows=count() by deviceId`);
      if (devices.every(device => rows.some(row => row.deviceId === device.id && Number(row.rows) > 0))) break;
      await pause(5000);
    }
    for (const device of devices) {
      assert.ok(rows.some(row => row.deviceId === device.id && Number(row.rows) > 0), "[native-activator] No fresh source records for a selected device.");
    }
    await writeFile(join(artifacts, "runs", this.plan.runId, `activator-evidence-${this.journal.attempt}.json`),
      JSON.stringify({ topology, inputSchema, nativeConnection: { id: connection.id, name: connection.name }, sourceRows: rows,
        scope: "workload setup, native bindings/schema and fresh source data; not destination arrival or rule firing" }, null, 2));
    this.plan.verification[`${path.id}.activator`] = "native-setup-bindings-schema-and-fresh-source-verified";
    this.plan.observed[`${path.id}.activatorDeviceId`] = text(devices[0]?.id, "expected Activator device");
    this.plan.observed[`${path.id}.activatorId`] = this.activatorId;
    this.plan.observed[`${path.id}.activatorName`] = this.activatorName;
    await updatePlan(this.plan);
  }

  async cleanup(): Promise<void> {
    const errors: Error[] = [];
    try { await this.reconcile(); } catch (error) { errors.push(error instanceof Error ? error : new Error(String(error))); }
    for (const resource of this.journal.resources) {
      if (!resource.requested || !resource.id) continue;
      try {
        const baseline = resource.kind === "Connection" ? this.journal.baselineConnections : this.journal.baselineItems;
        if (baseline.includes(resource.id)) throw new Error("[ownership] Cannot delete a baseline resource.");
        if (resource.kind === "Connection") {
          assert.equal(resource.cluster, this.plan.observed.queryServiceUri);
          const native = await this.cloud.request(`${FABRIC}/v1/connections/${resource.id}`, FABRIC, "GET", undefined, true);
          if (native !== undefined) assertConnectionCluster(object(native, "owned connection"), text(resource.cluster, "connection cluster"));
        }
        const path = resource.kind === "Connection" ? `/connections/${resource.id}` : `/workspaces/${this.plan.target.workspaceId}/items/${resource.id}`;
        await deleteFabricResource(this.cloud, path, resource.name, resource.kind === "Connection" ? undefined : resource.kind);
        resource.deleted = true;
        console.log(`[cleanup] Activator ${resource.kind} confirmed absent: ${resource.name}`);
      } catch (error) { errors.push(error instanceof Error ? error : new Error(String(error))); }
    }
    await this.persist();
    this.plan.verification[`${this.journal.path}.activatorCleanup`] = errors.length ? "failed" : "all-attempt-resources-confirmed-absent";
    await updatePlan(this.plan);
    if (errors.length) throw new AggregateError(errors, "[cleanup] Activator cleanup incomplete; inspect the ownership journal.");
    this.assertHealthy();
  }

  async cleanupAfterScenario(): Promise<void> {
    if (!this.cloud.config.keepResources) {
      await this.cleanup();
      return;
    }
    await Promise.all(this.pending);
    await this.persist();
    this.plan.verification[`${this.journal.path}.activatorCleanup`] = "retained-by-request";
    await updatePlan(this.plan);
    console.log(`[retention] Keeping Activator resources for run ${this.plan.runId}. Charges may continue. Ownership journal: ${this.file}`);
    this.assertHealthy();
  }

  static async cleanupJournals(cloud: Cloud, plan: RunPlan): Promise<void> {
    const directory = join(artifacts, "runs", plan.runId);
    const errors: Error[] = [];
    for (const file of await readdir(directory)) {
      if (!/^activator-attempt-[0-9a-f-]{36}\.json$/.test(file)) continue;
      try {
        const value: unknown = JSON.parse(await readFile(join(directory, file), "utf8"));
        assertJournal(value, plan);
        const attempt = new ActivatorAttempt(cloud, plan, value, join(directory, file), new Set());
        await attempt.cleanup();
      } catch (error) { errors.push(error instanceof Error ? error : new Error(String(error))); }
    }
    if (errors.length) throw new AggregateError(errors, "[cleanup] Some Activator journals could not be reconciled.");
  }
}

function assertJournal(value: unknown, plan: RunPlan): asserts value is Journal {
  const journal = object(value, "Activator journal");
  assert.equal(journal.runId, plan.runId);
  assert.equal(journal.workspaceId, plan.target.workspaceId);
  requireGuid(text(journal.attempt, "attempt ID"), "attempt ID");
  assert.ok(journal.path === "regular" || journal.path === "combined" || journal.path === "linked");
  for (const key of ["baselineItems", "baselineConnections"]) {
    const values = journal[key];
    assert.ok(Array.isArray(values));
    for (const id of values) requireGuid(text(id, "baseline ID"), "baseline ID");
  }
  assert.ok(Array.isArray(journal.resources) && journal.resources.length <= 3);
  const prefix = `${plan.azure.resourceGroup.replace(/^fabric-/, "")}-${journal.path}-`;
  const suffix = text(journal.attempt, "attempt").slice(0, 8);
  const kinds = new Set<string>();
  for (const value of journal.resources) {
    const resource = object(value, "journal resource");
    assert.ok(resource.kind === "Eventstream" || resource.kind === "Reflex" || resource.kind === "Connection");
    const name = text(resource.name, "resource name");
    assert.ok(!kinds.has(String(resource.kind)), "[ownership] Duplicate cleanup resource kind.");
    kinds.add(String(resource.kind));
    assert.ok(resource.kind === "Connection" ? /^KQL-[A-Za-z0-9_-]+$/.test(name)
      : name === `${prefix}${resource.kind === "Eventstream" ? "act-stream" : "activator"}-${suffix}`);
    if (resource.kind === "Connection") assert.equal(resource.cluster, plan.observed.queryServiceUri);
    assert.equal(typeof resource.requested, "boolean");
    if (resource.id !== undefined) requireGuid(text(resource.id, "resource ID"), "resource ID");
    if (resource.deleted !== undefined) {
      assert.equal(typeof resource.deleted, "boolean");
      if (resource.deleted) assert.ok(resource.requested && resource.id, "[ownership] Invalid confirmed deletion receipt.");
    }
    if (resource.operation !== undefined && !/^\/operations\/[0-9a-f-]{36}$/.test(text(resource.operation, "operation"))) throw new Error("[ownership] Invalid cleanup operation.");
  }
  assert.ok(kinds.has("Eventstream") && kinds.has("Reflex"));
}

export function assertConnectionCluster(connection: Record<string, unknown>, cluster: string): void {
  const details = object(connection.connectionDetails, "native connection details");
  assert.equal(details.type, "AzureDataExplorer");
  if (Array.isArray(details.parameters)) {
    assert.ok(details.parameters.some(value => {
      const parameter = object(value, "connection parameter");
      return parameter.name === "cluster" && parameter.value === cluster;
    }), "[ownership] Connection no longer targets the owned cluster.");
  } else {
    const path = new URL(text(details.path, "native connection path"));
    assert.equal(path.origin, new URL(cluster).origin);
    assert.ok(path.pathname === "/" && !path.search && !path.hash && !path.username && !path.password);
  }
}
