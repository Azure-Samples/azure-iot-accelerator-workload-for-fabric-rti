// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { decodeDoSamples, type JsonObject, type JsonValue } from "../app/items/IoTSolutionItem/ModelAuthoring.ts";
import { fromThingModel, inspectFieldEvidence, modelingFingerprint, nativeBindingIssues, setFieldRole } from "../app/items/IoTSolutionItem/HubModeling.ts";
import { resolveThingModel } from "../app/items/IoTSolutionItem/ThingModelParser.ts";
import { creationRequest } from "./creationMock.ts";
import type { CreationRequest } from "../app/items/IoTSolutionItem/ModeledDataCreation.ts";

export interface WotVariation {
  key: string;
  title: string;
  provenance: "agent-authored";
  document: JsonObject;
  expected: "supported" | "selection-conflict" | "no-data-fields" | "root-id-required" | "unsupported" | "rejected";
  fields?: { name: string; type: string; value: JsonValue; role?: "telemetry" | "twin" }[];
  error?: string;
}
const context = "https://www.w3.org/2022/wot/td/v1.1";
const model = (key: string, body: JsonObject): JsonObject => ({
  "@context": context, "@type": "tm:ThingModel", id: `urn:copilot:test:wot-variation:${key}`, ...body,
});
const variation = (key: string, title: string, body: JsonObject, fields: WotVariation["fields"],
  expected: WotVariation["expected"] = "supported"): WotVariation => ({
  key, title, provenance: "agent-authored", document: model(key, body), expected, fields,
});

// These fixed test IDs belong only to authored fixtures. No customer/public document is repaired or normalized.
export const wotVariations: WotVariation[] = [
  variation("number-unit", "Number with supplied unit and constraints", { properties: {
    pressure: { type: "number", title: "Pressure", description: "Line pressure.", unit: "kPa", minimum: 0, maximum: 500 },
  } }, [{ name: "pressure", type: "real", value: 12.5 }]),
  variation("primitive-types", "Integer boolean and string affordances", { properties: {
    count: { type: "integer" }, enabled: { type: "boolean" }, state: { type: "string" },
  } }, [{ name: "count", type: "long", value: 7 }, { name: "enabled", type: "bool", value: false, role: "twin" },
    { name: "state", type: "string", value: "fixture-state" }]),
  variation("complex-whole", "One nested object remains one column", { properties: {
    configuration: { type: "object", properties: {
      control: { type: "object", properties: { setpoint: { type: "number" }, mode: { type: "string" } } },
      alarms: { type: "array", items: { type: "string" } },
    } },
  } }, [{ name: "configuration", type: "dynamic", value: { control: { setpoint: 21, mode: "auto" }, alarms: ["hot"] }, role: "twin" }]),
  variation("arrays-whole", "Array of objects remains one whole value", { properties: {
    readings: { type: "array", items: { type: "object", properties: { value: { type: "number" }, valid: { type: "boolean" } } } },
  } }, [{ name: "readings", type: "dynamic", value: [{ value: 1, valid: true }, { value: 2, valid: false }] }]),
  variation("tuple-array", "Tuple DataSchema stays dynamic", { properties: {
    coordinate: { type: "array", items: [{ type: "number" }, { type: "number" }], minItems: 2, maxItems: 2 },
  } }, [{ name: "coordinate", type: "dynamic", value: [1.25, -2.5] }]),
  variation("nullable-oneof", "oneOf number and null alternatives", { properties: {
    optionalReading: { oneOf: [{ type: "number" }, { type: "null" }] },
  } }, [], "no-data-fields"),
  variation("coded-labels", "Exact model code labels and descriptions", { properties: {
    operatingMode: { type: "string", oneOf: [
      { const: "R", title: "Running", description: "Motor running." },
      { const: "S", title: "Stopped", description: "Motor stopped." },
    ] },
  } }, [], "no-data-fields"),
  variation("untyped", "Untyped declared property is not inferred from the sample", { properties: {
    observed: { description: "Unspecified declared type." },
  } }, [{ name: "observed", type: "dynamic", value: { nested: [1, "two", null] } }]),
  variation("whole-event", "One event data value can be a twin role", { events: {
    alarm: { title: "Alarm", data: { type: "object", properties: { code: { type: "integer" }, active: { type: "boolean" } } } },
  }, actions: { acknowledge: { input: { type: "integer" } } } },
  [{ name: "alarm", type: "dynamic", value: { code: 7, active: true }, role: "twin" }]),
  variation("role-neutral", "Affordance flags and custom kind do not assign Hub roles", {
    "@context": [context, { custom: "urn:copilot:test:", kind: "custom:kind" }],
    properties: { measurement: { type: "number", readOnly: true, observable: true, kind: "Twin" },
      desiredMode: { type: "string", writeOnly: true, kind: "Telemetry" } },
    events: { alarm: { data: { type: "boolean", readOnly: true } } }, actions: { reset: {} },
  }, [{ name: "measurement", type: "real", value: 20 }, { name: "desiredMode", type: "string", value: "auto", role: "twin" },
    { name: "alarm", type: "bool", value: true }]),
  variation("escaped-name", "JSON pointer escaping does not split a literal property name", { properties: {
    "pump/health~state": { type: "object", properties: { ok: { type: "boolean" } } },
  } }, [{ name: "pump/health~state", type: "dynamic", value: { ok: true } }]),
  variation("reserved-name", "Declared deviceId is distinct from native deviceId", { properties: {
    deviceId: { type: "string", description: "Model-declared label, not native transport identity." },
  } }, [{ name: "deviceId", type: "string", value: "fixture-label" }]),
  variation("localized", "Localized titles and distinct supplied document version", {
    title: "Fallback", titles: { en: "English model", "en-US": "US model" }, version: { model: "3.4.0" },
    properties: { temperature: { type: "number", titles: { en: "Temperature", "en-US": "US temperature" },
      descriptions: { en: "Measured temperature." }, unit: "degC", "@type": "urn:quantity:Temperature" } },
  }, [{ name: "temperature", type: "real", value: 22.5 }]),
  variation("column-collision", "Distinct model paths with the same normalized column need selection", { properties: {
    "pump.health": { type: "number" }, "pump-health": { type: "number" },
  } }, [{ name: "pump.health", type: "real", value: 1 }, { name: "pump-health", type: "real", value: 2 }], "selection-conflict"),
  variation("actions-only", "Actions-only TM has no supported data field candidates", { actions: { reboot: {}, stop: {} } }, [], "no-data-fields"),
  { key: "missing-root-id", title: "Optional WoT root ID is required by this workload", provenance: "agent-authored",
    document: { "@context": context, "@type": "tm:ThingModel", properties: { temperature: { type: "number" } } },
    expected: "root-id-required", error: "doesn't have an ID" },
  { ...variation("external-import", "Required external import is never fetched", { properties: {
    measurement: { "tm:ref": "https://example.org/model.tm.json#/properties/measurement" },
  } }, []), expected: "unsupported", error: "Required model file missing" },
  { ...variation("external-extends", "Required external inheritance is never fetched", {
    links: [{ rel: "tm:extends", href: "https://example.org/base.tm.json" }], properties: { own: { type: "number" } },
  }, []), expected: "unsupported", error: "Required model file missing" },
  { ...variation("submodel", "Submodel composition is valid but unsupported here", {
    links: [{ rel: "tm:submodel", href: "./motor.tm.json", instanceName: "motor" }],
  }, []), expected: "no-data-fields" },
  { ...variation("placeholder", "Unresolved template facts cannot expose effective fields", { properties: {
    reading: { type: "number", unit: "{{unit}}" },
  } }, []), expected: "unsupported", error: "placeholder" },
  { ...variation("json-schema-ref", "A DataSchema dependency is not implicitly resolved", { properties: {
    reading: { "$ref": "#/schemaDefinitions/reading" },
  }, schemaDefinitions: { reading: { type: "number" } } }, []), expected: "unsupported", error: "JSON Schema" },
  { ...variation("remapped-tm-prefix", "Remapped tm prefix cannot establish WoT authority", {
    "@context": [context, { tm: "https://example.org/not-wot#" }], properties: { reading: { type: "number" } },
  }, []), expected: "unsupported", error: "namespace" },
  { key: "thing-description", title: "A real TD is not promoted to a TM from its filename", provenance: "agent-authored",
    document: { "@context": context, "@type": "Thing", id: "urn:copilot:test:td", title: "TD",
      properties: { value: { type: "number", forms: [{ href: "https://example.org/value" }] } } },
    expected: "rejected", error: "not a WoT Thing Model" },
  { ...variation("invalid-dataschema", "Unsupported float is a schema rejection, not a valid type", {
    properties: { reading: { type: "float" } },
  }, []), expected: "rejected", error: "WoT 1.1 format" },
];

export const mountedWotVariationKeys = ["complex-whole", "arrays-whole", "primitive-types", "whole-event", "role-neutral", "localized"] as const;

export function wotVariationMockPayloads(request: CreationRequest): { telemetry: JsonObject; twin: JsonObject } {
  const telemetry = request.coverage.telemetry.payloads[0], twin = request.coverage.twin.payloads[0];
  if (!telemetry || !twin?.reported) throw new Error("A variation fixture needs decoded telemetry and reported-property samples.");
  return { telemetry: telemetry.value, twin: { properties: { reported: twin.reported } } };
}

export function wotVariationRequest(variation: WotVariation): CreationRequest {
  if (variation.expected !== "supported" || !variation.fields) throw new Error("Only supported authored fixtures have creation requests.");
  const request = creationRequest();
  request.input.model = fromThingModel(resolveThingModel(variation.document));
  request.input.fields = request.input.model.fields.map((field, index) =>
    setFieldRole(field, variation.fields![index].role || "telemetry"));
  request.input.scope = { mode: "data", operator: "and", telemetry: "", twin: "", reference: { id: request.input.model.id! } };
  const sample = (stream: "telemetry" | "twin") => {
    const values = Object.fromEntries(request.input.fields.flatMap((field, index) =>
      field.role === stream ? [[field.path[0], variation.fields![index].value]] : []));
    return decodeDoSamples([{
      data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0", contract_version: "1.0",
        event_kind: stream === "telemetry" ? "measurement" : "lifecycle", event_time: "2026-10-23T04:00:00Z",
        payload: { value: stream === "telemetry" ? values : { properties: { reported: values } }, memberTypeIRI: request.input.model.id } }),
      headers: { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" },
      user_headers: stream === "telemetry" ? { cloudEvents_subject: "device-1" }
        : { cloudEvents_deviceid: "device-1", cloudEvents_subject: "device-1" },
    }], stream, request.input.limit);
  };
  const telemetry = sample("telemetry"), twin = sample("twin");
  request.coverage = {
    fingerprint: modelingFingerprint(request.input), kind: "processing", telemetry, twin, sourceProbe: { telemetry, twin },
    nativeIssues: nativeBindingIssues(telemetry, twin),
    fields: inspectFieldEvidence(request.input.fields, telemetry, twin), checkedAt: "2026-10-23T04:00:00Z",
  };
  const suffix = variation.key.replace(/-/g, "_");
  request.names = { modeled: `Modeled_V_${suffix}`, twinChanges: `Changes_V_${suffix}`, twinState: `LKV_V_${suffix}` };
  return request;
}
