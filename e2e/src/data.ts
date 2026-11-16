import { createHmac, randomBytes, randomUUID } from "node:crypto";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import { createRequire } from "node:module";
import assert from "node:assert/strict";
import type { Client as DeviceClient } from "azure-iot-device";
import { Cloud, object, text, safeServiceMessage } from "./cloud.js";
import { artifacts } from "./config.js";
import { updatePlan, type RunPlan } from "./plan.js";
import { queryRows } from "./kusto.js";
import { DPS_API, HUB_API } from "./linking.js";
import { ingestionPath, type IngestionPath } from "./paths.js";
import { DeviceOperationQueue } from "./device-operations.js";
import { initializeModelReference } from "./model-assignment.js";
import { captureProcessOwner } from "./process-owner.js";

export function loadSimulationSdk() {
  const require = createRequire(import.meta.url);
  const device: typeof import("azure-iot-device") = require("azure-iot-device");
  const mqtt: typeof import("azure-iot-device-mqtt") = require("azure-iot-device-mqtt");
  if (typeof device.Client?.fromConnectionString !== "function" || typeof device.Message !== "function" ||
      typeof mqtt.MqttWs !== "function") {
    throw new Error("[simulation] The installed IoT SDKs do not expose the expected CommonJS API.");
  }
  return { Client: device.Client, Message: device.Message, MqttWs: mqtt.MqttWs };
}

class SimulationTimeout extends Error {}

function deviceErrorCode(error: unknown): string {
  const known = new Set(["UnauthorizedError", "NotConnectedError", "TimeoutError", "OperationTimeout", "ETIMEDOUT", "ECONNRESET", "ENOTFOUND", "ECONNREFUSED", "Disconnected"]);
  const name = error instanceof Error ? error.name : "";
  return known.has(name) ? name : "DeviceSdkError";
}

async function bounded<T>(action: () => Promise<T>, label: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  try {
    return await Promise.race([
      action(),
      new Promise<never>((_, reject) => {
        timer = setTimeout(() => reject(new SimulationTimeout(`[simulation] ${label} timed out.`)), 30_000);
      }),
    ]);
  } catch (error) {
    if (error instanceof SimulationTimeout) throw error;
    throw new Error(`[simulation] ${label} failed: ${deviceErrorCode(error)}. SDK credential-bearing error details are not recorded.`);
  } finally {
    clearTimeout(timer);
  }
}

export async function query(cloud: Cloud, plan: RunPlan, csl: string): Promise<Record<string, unknown>[]> {
  const uri = new URL(text(plan.observed.queryServiceUri, "owned Kusto endpoint"));
  if (uri.protocol !== "https:" || ![".kusto.fabric.microsoft.com", ".kusto.windows.net"].some(suffix => uri.hostname.endsWith(suffix))) {
    throw new Error("[ownership] Unexpected Kusto endpoint.");
  }
  return queryRows(await cloud.request(`${uri.origin}/v1/rest/${csl.startsWith(".") ? "mgmt" : "query"}`,
    "https://kusto.kusto.windows.net", "POST", { db: plan.fabric.databaseName, csl }));
}

export async function assertRawSchemas(cloud: Cloud, plan: RunPlan, path: IngestionPath = ingestionPath(plan, "regular")): Promise<void> {
  for (const table of [path.telemetryTable, path.propertiesTable]) {
    const rows = await query(cloud, plan, `.show table ${table} schema as json | project Schema`);
    if (rows.length !== 1) throw new Error("[service-shape] Native Kusto returned an ambiguous raw table schema.");
    const schema = object(JSON.parse(text(rows[0]?.Schema, "raw schema JSON")), "raw schema");
    if (!Array.isArray(schema.OrderedColumns)) throw new Error("[service-shape] Raw table column metadata is missing.");
    for (const [name, type] of [["data", "string"], ["headers", "dynamic"], ["user_headers", "dynamic"]]) {
      if (!schema.OrderedColumns.some(value => {
        const column = object(value, "raw column");
        const actual = column.CslType ?? column.Type;
        return column.Name === name && (actual === type || actual === (type === "string" ? "System.String" : "System.Object"));
      })) throw new Error(`[product-or-deployed-contract] ${table} must expose ${name}:${type}.`);
    }
  }
}

export interface SimulatedDevice {
  id: string;
  hubKey: string;
  hubId: string;
  client: DeviceClient;
  fault: Error | undefined;
  telemetryAcks: number;
  propertyAcks: number;
  telemetry: Record<string, unknown>;
  reported: Record<string, unknown>;
  components: Record<string, Record<string, unknown>>;
}

export interface DeviceFleet {
  attempt: string;
  devices: readonly SimulatedDevice[];
  assertHealthy(): void;
  sendRound(): Promise<void>;
  reportProperties(): Promise<void>;
}

export async function runSimulationWork(
  run: () => Promise<void>, pump: Promise<void>, stop: () => void, onFault: (error: Error) => void,
): Promise<void> {
  let fault: Error | undefined;
  const observed = pump.catch(error => {
    fault = error instanceof Error ? error : new Error("[simulation] Device pump failed.");
    onFault(fault);
  });
  try {
    await run();
  } finally {
    stop();
    await observed;
  }
  if (fault) throw fault;
}

export async function withSimulation(
  cloud: Cloud, plan: RunPlan, run: (fleet: DeviceFleet) => Promise<void>, allHubs = false,
): Promise<void> {
  const { Client, Message, MqttWs } = loadSimulationSdk();
  const owner = await captureProcessOwner(plan.runId);
  if (!owner) throw new Error("[simulation] Cannot establish the simulator process ownership.");
  const attempt = randomUUID();
  plan.phase = "local-device-simulation-starting";
  plan.observed.dataAttemptId = attempt;
  plan.observed.nativeDataVerifiedAt = "";
  plan.observed.simulationState = "starting";
  plan.observed.simulationStartedAt = "";
  plan.observed.simulationStoppedAt = "";
  plan.observed.simulationOwner = JSON.stringify(owner);
  await updatePlan(plan);
  const devices: SimulatedDevice[] = [];
  const controller = new AbortController();
  let stopping = false;
  let pump: Promise<void> | undefined;
  let pumpFault: Error | undefined;
  const operations = new DeviceOperationQueue();
  const fleet: DeviceFleet = {
    attempt,
    devices,
    sendRound,
    reportProperties,
    assertHealthy() {
      if (stopping) throw new Error("[simulation] Device clients stopped before test completion.");
      if (pumpFault) throw pumpFault;
      for (const device of devices) {
        if (device.fault) throw new Error(`[simulation] Transport failed for ${device.id}: ${deviceErrorCode(device.fault)}`);
      }
    },
  };
  async function sendRound(): Promise<void> {
    fleet.assertHealthy();
    return operations.round(async sequence => {
      if (!stopping) await publishRound(sequence);
    });
  }
  async function publishRound(sequence: number): Promise<void> {
    fleet.assertHealthy();
    for (const device of devices) {
      if (stopping) return;
      const message = new Message(JSON.stringify({ ...device.telemetry, e2eSequence: sequence }));
      message.contentType = "application/json";
      message.contentEncoding = "utf-8";
      await bounded(() => device.client.sendEvent(message), `${device.id} telemetry acknowledgement`);
      device.telemetryAcks++;
      for (const [component, fields] of Object.entries(device.components)) {
        const componentMessage = new Message(JSON.stringify({ ...fields, e2eRunId: plan.runId, e2eAttemptId: attempt, e2eSequence: sequence }));
        componentMessage.contentType = "application/json";
        componentMessage.contentEncoding = "utf-8";
        componentMessage.properties.add("$.sub", component);
        await bounded(() => device.client.sendEvent(componentMessage), `${device.id}/${component} acknowledgement`);
        device.telemetryAcks++;
      }
    }
    if (!stopping) await publishProperties();
  }
  async function reportProperties(): Promise<void> {
    fleet.assertHealthy();
    return operations.properties(async () => {
      if (!stopping) await publishProperties();
    });
  }
  async function publishProperties(): Promise<void> {
    fleet.assertHealthy();
    for (const device of devices) {
      if (stopping) return;
      const twin = await bounded(() => device.client.getTwin(), `${device.id} twin`);
      await bounded(() => new Promise<void>((resolve, reject) => {
        twin.properties.reported.update({ ...device.reported, e2eSequence: device.propertyAcks + 1 },
          (error: unknown) => error ? reject(error) : resolve());
      }), `${device.id} reported-properties acknowledgement`);
      device.propertyAcks++;
    }
  }
  try {
    for (const [index, hub] of plan.azure.hubs.filter(entry => allHubs || entry.kind === "nonlinked").entries()) {
      const current = object(await cloud.arm(hub.id, HUB_API), "simulated Hub");
      if (object(current.tags, "simulation Hub tags")["e2e-run-id"] !== plan.runId) throw new Error("[ownership] Simulation reached an unowned Hub.");
      const hubProperties = object(current.properties, "Hub properties");
      const host = text(hubProperties.hostName, "Hub hostname");
      if (host.toLowerCase() !== `${hub.name}.azure-devices.net`) throw new Error("[ownership] Unexpected simulation Hub hostname.");
      const id = `${hub.name}-device`;
      const deviceHost = hub.namespaceKey !== null ? text(hubProperties.deviceHostName, "native device hostname") : host;
      const serviceHost = hub.namespaceKey !== null ? text(hubProperties.serviceHostName, "native service hostname") : host;
      if (!deviceHost.endsWith(".azure-devices.net") || !serviceHost.endsWith(".azure-devices.net")) throw new Error("[ownership] Invalid native device/service hostname.");
      const endpoint = `https://${serviceHost}/devices/${encodeURIComponent(id)}?api-version=${hub.namespaceKey !== null ? "2026-11-01-preview" : "2021-04-12"}`;
      if (hub.namespaceKey !== null) {
        const ns = plan.azure.namespaces.find(namespace => namespace.key === hub.namespaceKey);
        if (!ns) throw new Error("[ownership] Missing planned namespace.");
        const dps = object(await cloud.arm(ns.dps.id, DPS_API), "owned DPS");
        if (object(dps.tags, "DPS ownership")["e2e-run-id"] !== plan.runId) throw new Error("[ownership] Cannot use an external DPS.");
        const dpsProperties = object(dps.properties, "DPS properties");
        const dpsHost = text(dpsProperties.serviceOperationsHostName, "DPS hostname");
        const scope = text(dpsProperties.idScope, "DPS ID scope");
        const enrollmentUrl = `https://${dpsHost}/enrollments/${encodeURIComponent(id)}?api-version=2026-11-02-preview`;
        let enrollment = await cloud.request(enrollmentUrl, "https://azure-devices-provisioning.net", "GET", undefined, true);
        if (enrollment === undefined) enrollment = await cloud.request(enrollmentUrl, "https://azure-devices-provisioning.net", "PUT", {
          registrationId: id, deviceId: id, provisioningStatus: "enabled", allocationPolicy: "static", iotHubs: [host],
          deviceTypeRefs: [], capabilities: { iotEdge: false }, initialTwin: { tags: plan.azure.tags },
          attestation: { type: "symmetricKey", symmetricKey: {
            primaryKey: randomBytes(32).toString("base64"), secondaryKey: randomBytes(32).toString("base64"),
          } },
        }, false, undefined, "*");
        const entry = object(enrollment, "owned enrollment");
        if (entry.registrationId !== id || entry.deviceId !== id || entry.allocationPolicy !== "static" ||
            !Array.isArray(entry.iotHubs) || entry.iotHubs.length !== 1 || entry.iotHubs[0] !== host) throw new Error("[ownership] DPS enrollment targets a different device/Hub.");
        const readbackUrl = `https://${dpsHost}/registrations/${encodeURIComponent(id)}?api-version=2026-11-02-preview`;
        const previous = await cloud.request(readbackUrl, "https://azure-devices-provisioning.net", "GET", undefined, true);
        if (previous === undefined) {
        const secret = text(object(object(entry.attestation, "attestation").symmetricKey, "device key").primaryKey, "DPS device key");
        const encodedResource = encodeURIComponent(`${scope}/registrations/${id}`);
        const expires = Math.floor(Date.now() / 1000) + 900;
        const signature = createHmac("sha256", Buffer.from(secret, "base64")).update(`${encodedResource}\n${expires}`).digest("base64");
        const sas = `SharedAccessSignature sr=${encodedResource}&sig=${encodeURIComponent(signature)}&se=${expires}`;
        const registrationRoot = `https://global.azure-devices-provisioning.net/${encodeURIComponent(scope)}/registrations/${encodeURIComponent(id)}`;
        async function deviceRequest(path: string, method = "GET", body?: unknown) {
          const response = await fetch(`${registrationRoot}/${path}?api-version=2026-11-02-preview`, {
            method, headers: { Authorization: sas, "Content-Type": "application/json", Connection: "close" },
            ...(body === undefined ? {} : { body: JSON.stringify(body) }), signal: AbortSignal.timeout(60_000), redirect: "error",
          });
          const result = object(await response.json(), "DPS device response");
          if (!response.ok) throw new Error(`[simulation] DPS registration HTTP ${response.status}: ${safeServiceMessage(result.message ?? result.Message)}`);
          return result;
        }
          let result = await deviceRequest("register", "PUT", { registrationId: id });
          const deadline = Date.now() + 5 * 60_000;
          while (Date.now() < deadline && object(result.registrationState ?? {}, "DPS registration state").status !== "assigned") {
            const state = object(result.registrationState ?? {}, "DPS registration state");
            if (state.status === "failed" || result.status === "failed") throw new Error("[simulation] DPS assignment failed.");
            const operationId = text(result.operationId, "DPS operation ID");
            await pause(2000);
            result = await deviceRequest(`operations/${encodeURIComponent(operationId)}`);
          }
          const registration = object(result.registrationState, "assigned device");
          if (![host, deviceHost, serviceHost].includes(text(registration.assignedHub, "assigned Hub")) || registration.deviceId !== id || registration.status !== "assigned") {
            throw new Error("[topology] Native DPS registration did not assign the exact expected device/Hub.");
          }
        if (entry.provisioningStatus === "enabled") {
          const disabled = object(await cloud.request(enrollmentUrl, "https://azure-devices-provisioning.net", "PUT", {
            registrationId: id, deviceId: id, provisioningStatus: "disabled", allocationPolicy: "static", iotHubs: [host],
            deviceTypeRefs: entry.deviceTypeRefs, capabilities: entry.capabilities, initialTwin: entry.initialTwin, attestation: entry.attestation,
          }, false, text(entry.etag, "current enrollment ETag")), "disabled bootstrap enrollment");
          if (disabled.provisioningStatus !== "disabled" ||
              object(object(disabled.attestation, "attestation").symmetricKey, "preserved device key").primaryKey !== secret) {
            throw new Error("[security] Bootstrap disable did not preserve the assigned device's key.");
          }
        }
        plan.observed[`${hub.key}.bootstrapEnrollment`] = "disabled-after-assignment";
        } else {
          plan.observed[`${hub.key}.bootstrapEnrollment`] = `retained-existing-${String(entry.provisioningStatus)}`;
        }
        const native = object(await cloud.request(readbackUrl, "https://azure-devices-provisioning.net"), "DPS independent registration readback");
        if (![host, deviceHost, serviceHost].includes(text(native.assignedHub, "native assigned Hub")) || native.deviceId !== id || native.status !== "assigned") throw new Error("[topology] Independent DPS readback did not confirm the assignment.");
        plan.observed[`${hub.key}.dpsRegistration`] = "assigned-to-exact-owned-hub";
      }
      let identity = await cloud.request(endpoint, "https://iothubs.azure.net", "GET", undefined, true);
      if (identity === undefined) {
        if (hub.namespaceKey !== null) throw new Error("[service] DPS did not create the expected linked device identity.");
        identity = await cloud.request(endpoint, "https://iothubs.azure.net", "PUT", {
          deviceId: id, status: "enabled",
          authentication: { type: "sas", symmetricKey: {
            primaryKey: randomBytes(32).toString("base64"), secondaryKey: randomBytes(32).toString("base64"),
          } },
        });
      }
      const record = object(identity, "owned device identity");
      if (record.deviceId !== id || record.status !== "enabled") throw new Error("[ownership] Returned device identity does not match the owned run.");
      if (hub.namespaceKey !== null) {
        const ns = plan.azure.namespaces.find(namespace => namespace.key === hub.namespaceKey);
        if (!ns) throw new Error("[ownership] Missing namespace binding.");
        const native = object(record.adrDeviceProperties, "native ADR device");
        if (native.name !== id || object(native.deviceCapability, "native capability").name !== hub.name) throw new Error("[native-identity] Wrong native ADR device/capability binding.");
        const registryId = `${ns.id}/registryDevices/${id}`;
        const registry = object(await cloud.arm(registryId, "2026-11-02-preview"), "owned registry device");
        const props = object(registry.properties, "registry properties");
        if (props.uuid !== native.uuid || props.externalDeviceId !== id || props.enablementState !== "Enabled") throw new Error("[native-identity] Native/ADR UUID and external identity do not match.");
        const capability = object(object(await cloud.arm(`${registryId}/capabilities/${hub.name}`, "2026-11-02-preview"), "registry capability").properties, "capability properties");
        if (capability.messagingEndpointName !== hub.name || capability.capabilityType !== "Microsoft.IoTHub" || capability.provisioningState !== "Succeeded") throw new Error("[native-identity] Wrong namespace capability/Hub source binding.");
        plan.observed[`${hub.key}.adrDeviceUuid`] = text(native.uuid, "ADR device UUID");
        await initializeModelReference(cloud, plan, hub.key);
      }
      const auth = object(record.authentication, "device authentication");
      const key = text(object(auth.symmetricKey, "device symmetric key").primaryKey, "device primary key");
      const client = Client.fromConnectionString(`HostName=${deviceHost};DeviceId=${id};SharedAccessKey=${key}`, MqttWs);
      const device: SimulatedDevice = {
        id, hubKey: hub.key, hubId: hub.id, client, fault: undefined, telemetryAcks: 0, propertyAcks: 0,
        telemetry: {
          temperature: 23.75 + index, humidity: 45.5, pressure: 1012.3, status: "normal",
          gpsLocation: { lat: 47.6062, lon: -122.3321, floor: 2 },
          e2eRunId: plan.runId, e2eAttemptId: attempt,
        },
        reported: {
          firmwareVersion: "1.0.0", customerName: "E2E", tags: ["hvac", "zone-a", "monitored"],
          maintenanceSchedule: { intervalDays: 90, lastServiceDate: "2025-10-05" },
          thermostat: { __t: "c", setpoint: 22.0 + index * 0.25, scheduleEnabled: true },
          powerModule: { __t: "c", batteryHealth: "good", powerSource: "mains" },
          e2eRunId: plan.runId, e2eAttemptId: attempt,
        },
        components: {
          thermostat: { targetTemperature: 22.5 + index * 0.25, fanSpeed: 50 },
          powerModule: { powerConsumption: 75.5 + index, current: 0.5 },
        },
      };
      devices.push(device);
      client.on("error", (error: Error) => { device.fault = error; });
      client.on("disconnect", () => { device.fault = new Error("[simulation] Device transport disconnected."); });
      await bounded(() => client.open(), `${id} connect`);
      plan.observed[`${hub.key}DeviceId`] = id;
    }
    await sendRound();
    plan.observed.simulationStartedAt = new Date().toISOString();
    plan.observed.simulationState = "running-in-process";
    await updatePlan(plan);
    console.log("Local device clients connected and acknowledged sends. They remain running throughout the test; no ACR or containers are used.");
    pump = (async () => {
      while (!stopping) {
        await pause(10_000, undefined, { signal: controller.signal });
        if (!stopping) await sendRound();
      }
    })().catch(error => {
      if (error instanceof Error && error.name === "AbortError" && controller.signal.aborted) return;
      throw error;
    });
    await runSimulationWork(() => run(fleet), pump, () => {
      stopping = true;
      controller.abort();
    }, error => {
      pumpFault = error;
      console.error("[simulation] Device pump failed; waiting for the active scenario to settle before transport cleanup.");
    });
  } finally {
    stopping = true;
    controller.abort();
    try {
      try {
        await pump;
      } catch (error) {
        if (!(error instanceof Error && error.name === "AbortError" && controller.signal.aborted)) throw error;
      }
    } finally {
      const closed = await Promise.allSettled(devices.map(device => bounded(() => device.client.close(), `${device.id} transport close`)));
      const failures = closed.filter((result): result is PromiseRejectedResult => result.status === "rejected");
      if (failures.length) throw new AggregateError(failures.map(result => result.reason), "[simulation] Device transport shutdown could not be confirmed.");
      plan.observed.simulationState = "stopped";
      plan.observed.simulationStoppedAt = new Date().toISOString();
      await updatePlan(plan);
    }
  }
}

export async function verifyArrivingData(cloud: Cloud, plan: RunPlan, fleet: DeviceFleet, path: IngestionPath = ingestionPath(plan, "regular"), exclude: readonly SimulatedDevice[] = []): Promise<void> {
  const { attempt, devices } = fleet;
  assert.equal(devices.filter(device => path.hubKeys.includes(device.hubKey)).length, path.hubKeys.length,
    "[native-data] Every selected Hub must have a simulated device; an empty fleet cannot pass.");
  const evidencePath = join(artifacts, "runs", plan.runId, `raw-evidence-${path.id}-${attempt}.json`);
  plan.phase = "regular-hub-native-data-verifying";
  const verificationKey = path.id === "regular" ? "rawData" : `${path.id}.rawData`;
  plan.verification[verificationKey] = "in-progress";
  await updatePlan(plan);
  try {
    await assertRawSchemas(cloud, plan, path);
    const deadline = Date.now() + 10 * 60_000;
    let lastCounts = "";
    while (Date.now() < deadline) {
      fleet.assertHealthy();
      const filter = `| where data contains "${attempt}" | project data, headers, user_headers | take 1000`;
      const telemetry = await query(cloud, plan, `${path.telemetryTable} ${filter}`);
      const properties = await query(cloud, plan, `${path.propertiesTable} ${filter}`);
      if (exclude.length) {
        const ids = JSON.stringify(exclude.map(device => device.id));
        for (const table of [path.telemetryTable, path.propertiesTable]) {
          const count = await query(cloud, plan, `${table} | where data contains "${attempt}" | where tostring(user_headers.cloudEvents_subject) in (dynamic(${ids})) or tostring(user_headers.cloudEvents_deviceid) in (dynamic(${ids})) | count`);
          if (count[0]?.Count !== 0) throw new Error(`[native-exclusion] Regular negative-control rows were found in ${table}.`);
        }
      }
      const counts = `${telemetry.length} telemetry, ${properties.length} properties`;
      if (counts !== lastCounts) console.log(`[data] Current-attempt native rows: ${counts}`);
      lastCounts = counts;
      await writeFile(evidencePath, JSON.stringify({
        attemptId: attempt,
        path, excludedDeviceIds: exclude.map(device => device.id),
        capturedAt: new Date().toISOString(),
        acknowledgements: devices.map(device => ({
          deviceId: device.id, telemetry: device.telemetryAcks, reportedProperties: device.propertyAcks,
        })),
        telemetry, properties,
      }, null, 2), { mode: 0o600 });
      const values = (rows: Record<string, unknown>[], device: SimulatedDevice, component = ""): Record<string, unknown>[] =>
        rows.flatMap(row => {
          if (typeof row.data !== "string") throw new Error("[product-or-deployed-contract] Raw data must remain JSON text.");
          object(row.headers, "raw headers");
          const headers = object(row.user_headers, "raw user headers");
          const identity = [headers.cloudEvents_deviceid, headers.cloudEvents_subject]
            .find(value => typeof value === "string" && value.length > 0);
          if (typeof identity !== "string" || !identity) throw new Error("[product-or-deployed-contract] A test observation is missing CloudEvents device identity.");
          if (exclude.some(device => device.id === identity)) throw new Error("[native-exclusion] Regular device data appeared in linked-only outputs.");
          if (identity !== device.id) return [];
          if (typeof headers.cloudEvents_source !== "string" || headers.cloudEvents_source.toLowerCase() !== device.hubId.toLowerCase()) {
            throw new Error("[native-identity] CloudEvent source does not identify the expected owned Hub.");
          }
          if ((headers.cloudEvents_iothubdtsubject ?? "") !== component) return [];
          const observation = object(JSON.parse(row.data), "DO observation");
          if (observation.schema_id !== "com.do.telemetry.observation.v1" ||
              observation.schema_version !== "1.0" || observation.contract_version !== "1.0") {
            throw new Error(`[product-or-deployed-contract] Test data is not the expected versioned DO observation. Inspect ${evidencePath}; no legacy-payload fallback is used.`);
          }
          return [object(object(observation.payload, "DO payload").value, "DO value")];
        });
      let complete = true;
      for (const device of devices.filter(device => path.hubKeys.includes(device.hubKey))) {
        const measurements = values(telemetry, device);
        const lifecycle = values(properties, device);
        const measurement = measurements.find(value => value.e2eAttemptId === attempt);
        const reported = lifecycle.map(value => object(object(value.properties, "lifecycle properties").reported, "reported properties"))
          .find(value => value.e2eAttemptId === attempt);
        if (!measurement || !reported) { complete = false; continue; }
        for (const [name, value] of Object.entries(device.telemetry)) assert.deepEqual(measurement[name], value, `[native-data] ${device.id} telemetry ${name}`);
        for (const [name, value] of Object.entries(device.reported)) assert.deepEqual(reported[name], value, `[native-data] ${device.id} reported ${name}`);
        for (const [component, expected] of Object.entries(device.components)) {
          const samples = values(telemetry, device, component);
          const sample = samples.find(value => value.e2eAttemptId === attempt);
          if (!sample) { complete = false; continue; }
          for (const [name, value] of Object.entries(expected)) assert.deepEqual(sample[name], value, `[native-component] ${device.id}/${component}/${name}`);
        }
        assert.ok(device.telemetryAcks > 0 && device.propertyAcks > 0, "[simulation] Native send acknowledgements are required.");
      }
      if (complete) {
        fleet.assertHealthy();
        plan.phase = "regular-hub-ingestion-native-data-verified";
        plan.verification[verificationKey] = `native-kusto-DO-data-from-${path.hubKeys.length}-owned-devices`;
        plan.observed.nativeDataVerifiedAt = new Date().toISOString();
        plan.observed[`${path.id}.rawEvidenceFile`] = evidencePath.slice(artifacts.length + 1);
        await updatePlan(plan);
        console.log(`[data] ${path.id}: exact root/component telemetry and reported properties verified for ${path.hubKeys.length} devices. Evidence: ${evidencePath}`);
        return;
      }
      await pause(10_000);
    }
    throw new Error(`[native-data] Acknowledged test telemetry/properties did not arrive from both owned devices within ten minutes. Inspect ${evidencePath}.`);
  } catch (error) {
    plan.phase = "regular-hub-native-data-failed";
    plan.verification[verificationKey] = "failed";
    await updatePlan(plan);
    throw error;
  }
}
