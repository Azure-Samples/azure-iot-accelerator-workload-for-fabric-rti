// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { JsonObject, JsonValue } from "../app/items/IoTSolutionItem/ModelAuthoring.ts";
import { decodeDoSamples } from "../app/items/IoTSolutionItem/ModelAuthoring.ts";
import { fromThingModel, inspectFieldEvidence, modelingFingerprint, nativeBindingIssues, setFieldRole } from "../app/items/IoTSolutionItem/HubModeling.ts";
import { resolveThingModel } from "../app/items/IoTSolutionItem/ThingModelParser.ts";
import { creationRequest } from "./creationMock.ts";

const header = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel" };
export function localPlainThingModel(count = 8000): JsonObject {
  return { ...header, id: "urn:fixture:plain-measurements", title: "Environment sensor array",
    properties: Object.fromEntries(Array.from({ length: count }, (_, index) => [`measurement${index}`, {
      type: "number", title: `Measurement ${index}`, description: "Observed sensor reading for operational monitoring.",
      unit: "degC", observable: true, minimum: -50, maximum: 150,
    }])) };
}

export function localReferenceWorkModel(attributeCount: number): JsonValue {
  const definitions: JsonObject[] = Array.from({ length: 30 }, (_, index) => ({ ...header,
    id: `urn:fixture:reference-work:${index}`, schemaDefinitions: { value: index === 0
      ? { type: "number", ...Object.fromEntries(Array.from({ length: attributeCount }, (_, key) => [`x:fact${key}`, true])) }
      : { "tm:ref": `urn:fixture:reference-work:${index - 1}#/schemaDefinitions/value`, title: `Reading ${index}` } },
  }));
  return [{ ...header, id: "urn:fixture:reference-work-root",
    properties: { reading: { "tm:ref": "urn:fixture:reference-work:29#/schemaDefinitions/value" } }, "x:boundary": [] },
    ...definitions];
}

export function localRepeatedReferenceModel(levels: number, fanout = 5): JsonObject {
  const definitions: JsonObject = { leaf: { type: "number" } };
  let previous = "leaf";
  for (let level = 0; level < levels; level++) {
    const name = `level${level}`;
    definitions[name] = { type: "object", properties: Object.fromEntries(Array.from({ length: fanout }, (_, index) =>
      [`field${index}`, { "tm:ref": `#/schemaDefinitions/${previous}` }])) };
    previous = name;
  }
  return { ...header, id: "urn:fixture:bounded-local-imports", schemaDefinitions: definitions,
    properties: { reading: { "tm:ref": `#/schemaDefinitions/${previous}` } } };
}

export const localDefinitions: JsonObject = { ...header, id: "https://example.org/local/definitions?version=1",
  title: "Definitions, not the selected device", version: { model: "base-version" }, schemaDefinitions: {
    "whole/value~type": { type: "object", properties: { value: { type: "number" }, label: { type: "string" } } },
    history: { type: "array", items: { type: "object", properties: { value: { type: "number" } } } },
    state: { type: "string", readOnly: true, observable: true },
  } };
const root = (id: string, location: string): JsonObject => ({ ...header, id, title: "Local model",
  version: { model: "root-version" }, properties: {
    measurement: { "tm:ref": `${location}#/schemaDefinitions/whole~1value~0type`, title: "Whole measurement" },
    history: { "tm:ref": `${location}#/schemaDefinitions/history` },
  }, events: { alarm: { data: { "tm:ref": `${location}#/schemaDefinitions/state` } } }, actions: { reset: {} } });
export const localSingleModel: JsonObject = { ...root("urn:copilot:test:local-object", ""),
  schemaDefinitions: localDefinitions.schemaDefinitions };
export const localBundleRoot: JsonObject = { ...root("urn:copilot:test:local-bundle", String(localDefinitions.id)),
  links: [{ rel: "tm:extends", href: localDefinitions.id! }] };
export const localBundle: JsonValue = [localDefinitions, localBundleRoot];
const middle: JsonObject = { ...header, id: "urn:copilot:test:local-middle", schemaDefinitions: {
  "whole/value~type": { "tm:ref": `${localDefinitions.id}#/schemaDefinitions/whole~1value~0type` },
  history: { "tm:ref": `${localDefinitions.id}#/schemaDefinitions/history` },
  state: { "tm:ref": `${localDefinitions.id}#/schemaDefinitions/state` },
} };
export const localForwardBundle: JsonValue = [root("urn:copilot:test:local-forward", String(middle.id)), localDefinitions, middle];
export const localPositiveFixtures: { key: string; document: JsonValue; rootId: string }[] = [
  { key: "single", document: localSingleModel, rootId: String(localSingleModel.id) },
  { key: "bundle", document: localBundle, rootId: String(localBundleRoot.id) },
  { key: "forward", document: localForwardBundle, rootId: "urn:copilot:test:local-forward" },
];
export const localNegativeFixtures: { key: string; document: JsonValue; error: string }[] = [
  { key: "ambiguous", document: [localSingleModel, { ...localSingleModel, id: "urn:copilot:test:unrelated" }], error: "multiple roots" },
  { key: "duplicate", document: [localSingleModel, localSingleModel], error: "Duplicate model ID" },
  { key: "missing", document: { ...localSingleModel, properties: { value: { "tm:ref": "#/schemaDefinitions/absent" } } }, error: "not found" },
  { key: "schema-ref", document: { ...localSingleModel, properties: { value: { "$ref": "#/schemaDefinitions/state" } } }, error: "JSON Schema $ref" },
  { key: "external", document: { ...localSingleModel, properties: { value: { "tm:ref": "https://example.org/not-in-file#/value" } } }, error: "Required model file missing" },
];
export const localFixtureValues: JsonObject = { measurement: { value: 23, label: "whole" }, history: [{ value: 1 }, { value: 2 }], alarm: "ready" };
export function localFixtureRequest(document: JsonValue) {
  const request = creationRequest();
  request.input.model = fromThingModel(resolveThingModel(document));
  request.input.fields = request.input.model.fields.map(field => setFieldRole(field, field.path[0] === "history" ? "twin" : "telemetry"));
  const sample = (stream: "telemetry" | "twin") => {
    const values = Object.fromEntries(request.input.fields.filter(field => field.role === stream).map(field => [field.path[0], localFixtureValues[field.path[0]]]));
    return decodeDoSamples([{ data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1", schema_version: "1.0",
      contract_version: "1.0", event_time: "2026-10-23T23:00:00Z", event_kind: stream === "telemetry" ? "measurement" : "lifecycle",
      payload: { value: stream === "telemetry" ? values : { properties: { reported: values } }, memberTypeIRI: request.input.model.id } }),
    headers: { IoTConnectionDeviceId: "local-fixture-device", IoTConnectionModuleId: "", IoTSubject: "" }, user_headers: {} }], stream, 100);
  };
  const telemetry = sample("telemetry"), twin = sample("twin");
  request.coverage = { fingerprint: modelingFingerprint(request.input), kind: "processing", telemetry, twin,
    sourceProbe: { telemetry, twin }, fields: inspectFieldEvidence(request.input.fields, telemetry, twin),
    nativeIssues: nativeBindingIssues(telemetry, twin), checkedAt: "2026-10-23T23:00:00Z" };
  return request;
}
