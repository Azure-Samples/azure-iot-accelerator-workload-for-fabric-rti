// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { decodeDoSamples, type JsonObject, type JsonValue } from "../app/items/IoTSolutionItem/ModelAuthoring.ts";
import { parseDtdlModel } from "../app/items/IoTSolutionItem/DtdlModelParser.ts";
import { fromDtdlModel, inspectFieldEvidence, modelingFingerprint, nativeBindingIssues } from "../app/items/IoTSolutionItem/HubModeling.ts";
import type { AssignmentSource } from "../app/items/IoTSolutionItem/HubAssignment.ts";
import { creationRequest } from "./creationMock.ts";

export const mixedHubRoot = "/subscriptions/00000000-0000-0000-0000-000000000001/resourceGroups/mixed-fixture";
export const mixedNamespace = `${mixedHubRoot}/providers/Microsoft.DeviceRegistry/namespaces/mixed`;
export const mixedHubs = {
  a: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/linked-a`, host: "linked-a.azure-devices.net" },
  b: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/linked-b`, host: "linked-b.azure-devices.net" },
  u: { id: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/unlinked`, host: "unlinked.azure-devices.net" },
};
export function mixedAssignmentSource(telemetry: (keyof typeof mixedHubs)[] = ["a", "u"],
  twin: (keyof typeof mixedHubs)[] = telemetry): AssignmentSource {
  return { key: "mock-source",
    telemetry: { table: "TelemetryRawIoTHub_combined_hubs",
      docstring: `@iot.meta tableRole=iothub-telemetry-raw\n@iot.meta source=${telemetry.map(key => mixedHubs[key].host).join("; ")}` },
    twin: { table: "PropertiesRawIoTHub_combined_hubs",
      docstring: `@iot.meta tableRole=iothub-properties-raw\n@iot.meta source=${twin.map(key => mixedHubs[key].host).join("; ")}` } };
}

export function mixedAssignmentFixture() {
  const calls: { url: string; method: string; body?: string }[] = [], scopes: string[] = [];
  const devices = ["a", "b", "both", "unrelated", "wrong-kind"].map(name => ({
    name, id: `${mixedNamespace}/registryDevices/${name}`,
  }));
  const fetcher: typeof fetch = async (input, init) => {
    const url = String(input), parsed = new URL(url), path = parsed.pathname;
    calls.push({ url, method: init?.method || "GET", body: typeof init?.body === "string" ? init.body : undefined });
    if (parsed.origin !== "https://management.azure.com" || !["GET", "POST"].includes(init?.method || "GET"))
      throw new Error("Mixed-source fixture permits only synthetic ARM reads/ARG POST.");
    const reply = (body: unknown) => new Response(JSON.stringify(body), { headers: { "Content-Type": "application/json" } });
    if (url.includes("ResourceGraph")) {
      const query = JSON.parse(String(init?.body)).query as string;
      if (query.includes("registrydevices")) throw new Error("Device reference authority must use the canonical GET, not a device ARG query.");
      if (!query.includes("microsoft.devices/iothubs")) throw new Error("Unexpected DPS/unscoped fixture query.");
      return reply({ data: Object.values(mixedHubs).filter(hub => query.includes(JSON.stringify(hub.host)))
        .map(hub => ({ id: hub.id, name: hub.host.split(".")[0], hostName: hub.host })) });
    }
    for (const [key, hub] of Object.entries(mixedHubs)) if (path === hub.id) return reply({
      id: hub.id, name: hub.host.split(".")[0], properties: { hostName: hub.host,
        deviceRegistry: key === "u" ? null : { namespaceResourceId: mixedNamespace, linkingProperties: { state: "Succeeded" } } },
    });
    if (path === mixedNamespace) return reply({ id: mixedNamespace, properties: { messaging: { endpoints: {
      a: { endpointType: "Microsoft.Devices/IotHubs", resourceId: mixedHubs.a.id },
      b: { endpointType: "Microsoft.Devices/IotHubs", resourceId: mixedHubs.b.id },
      unrelated: { endpointType: "Microsoft.Devices/IotHubs", resourceId: `${mixedHubRoot}/providers/Microsoft.Devices/IotHubs/not-selected` },
    } } } });
    if (path === `${mixedNamespace}/registryDevices`) return reply({ value: devices.map(device => ({ id: device.id })) });
    const device = devices.find(value => path === value.id || path.startsWith(`${value.id}/`));
    if (!device) throw new Error("Unexpected mixed-source fixture resource.");
    if (path.endsWith("/capabilities")) return reply({ value: (device.name === "both" ? ["a", "b"]
      : device.name === "wrong-kind" ? ["a"] : [device.name]).map(endpoint => ({
        id: `${device.id}/capabilities/${endpoint}`,
        properties: { capabilityType: device.name === "wrong-kind" ? "Microsoft.SoftwareUpdate" : "Microsoft.IoTHub",
          messagingEndpointName: endpoint },
      })) });
    if (path.endsWith("/permissions")) return reply({ value: [{ actions: ["*"], notActions: [] }] });
    if (path === device.id) return reply({ id: device.id, name: device.name,
      type: "Microsoft.DeviceRegistry/namespaces/registryDevices", etag: `"mixed-${device.name}-1"`,
      properties: { deviceTypeRefs: [], uuid: "11111111-1111-4111-8111-111111111111",
        externalDeviceId: device.name, provisioningState: "Succeeded" } });
    throw new Error("Unexpected mixed-source fixture operation.");
  };
  return { fetcher, calls, scopes, token: async (scope: string) => { scopes.push(scope); return "synthetic-only"; } };
}

const complex: JsonObject = { "@type": "Object", fields: [{ name: "value", schema: "double" }, { name: "state", schema: "string" }] };
const array: JsonObject = { "@type": "Array", elementSchema: complex };
const capabilities = (): JsonObject[] => [
  { "@type": "Telemetry", name: "temperature", schema: "double" },
  { "@type": "Telemetry", name: "measurement", schema: complex },
  { "@type": "Telemetry", name: "samples", schema: array },
  { "@type": "Property", name: "state", schema: "string" },
  { "@type": "Property", name: "configuration", schema: complex },
  { "@type": "Property", name: "history", schema: array },
];
export const mixedDtdlDocument: JsonValue = [
  { "@id": "dtmi:copilot:tests:CombinedMachine;1", "@type": "Interface", extends: "dtmi:copilot:tests:CombinedBase;1",
    contents: [...capabilities(),
      { "@type": "Component", name: "motor", schema: "dtmi:copilot:tests:CombinedSensor;1" },
      { "@type": "Component", name: "power", schema: "dtmi:copilot:tests:CombinedSensor;1" },
      { "@type": "Command", name: "reset" }, { "@type": ["Property", "Cloud"], name: "customer", schema: "string" }] },
  { "@id": "dtmi:copilot:tests:CombinedBase;1", "@type": "Interface", contents: [] },
  { "@id": "dtmi:copilot:tests:CombinedSensor;1", "@type": "Interface", contents: capabilities() },
];
export const mixedTelemetryValues = { temperature: 22, measurement: { value: 1, state: "whole" },
  samples: [{ value: 2, state: "first" }, { value: 3, state: "second" }] };
export const mixedReportedValues = { state: "root", configuration: { value: 4, state: "root-whole" },
  history: [{ value: 5, state: "root-history" }],
  motor: { __t: "c", state: "motor", configuration: { value: 6, state: "motor-whole" }, history: [{ value: 7, state: "motor-history" }] },
  power: { __t: "c", state: "power", configuration: { value: 8, state: "power-whole" }, history: [{ value: 9, state: "power-history" }] } };
export function mixedRawSample(stream: "telemetry" | "twin", component = "", module = "") {
  return {
    data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
      event_time: "2026-10-23T21:00:00Z", event_kind: stream === "telemetry" ? "measurement" : "lifecycle",
      payload: { value: stream === "telemetry" ? mixedTelemetryValues : { properties: { reported: mixedReportedValues } } } }),
    headers: { IoTConnectionDeviceId: "fixture-device", IoTConnectionModuleId: module, IoTSubject: component },
    user_headers: { hubName: "linked-a" },
  };
}
export function mixedDtdlRequest() {
  const request = creationRequest();
  request.input.model = fromDtdlModel(parseDtdlModel(mixedDtdlDocument), mixedDtdlDocument);
  request.input.fields = request.input.model.fields;
  request.input.source.telemetryTable = "TelemetryRawIoTHub_combined_hubs";
  request.input.source.twinTable = "PropertiesRawIoTHub_combined_hubs";
  const telemetry = decodeDoSamples(["", "motor", "power"].map(component => mixedRawSample("telemetry", component)), "telemetry", 100);
  const twin = decodeDoSamples([mixedRawSample("twin")], "twin", 100);
  request.coverage = { fingerprint: modelingFingerprint(request.input), kind: "processing", telemetry, twin,
    sourceProbe: { telemetry, twin }, fields: inspectFieldEvidence(request.input.fields, telemetry, twin),
    nativeIssues: nativeBindingIssues(telemetry, twin), checkedAt: "2026-10-23T21:00:00Z" };
  request.names = { modeled: "Modeled_Mixed_Dtdl", twinChanges: "Changes_Mixed_Dtdl", twinState: "Lkv_Mixed_Dtdl" };
  return request;
}
