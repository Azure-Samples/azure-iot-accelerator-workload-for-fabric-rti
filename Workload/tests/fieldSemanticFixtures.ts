// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { JsonValue } from "../app/items/IoTSolutionItem/ModelAuthoring.ts";

export const tupleThingModel = {
  "@context": "https://www.w3.org/2022/wot/td/v1.1",
  "@type": "tm:ThingModel",
  id: "urn:example:ThermoProX200",
  title: "Commercial Greenhouse Climate Controller",
  schemaDefinitions: {
    TupleArray: {
      title: "Greenhouse bay location", type: "array", minItems: 3, maxItems: 3,
      items: [
        { type: "number", title: "Latitude" },
        { type: "number", title: "Longitude" },
        { type: "string", title: "Coordinate reference system", enum: ["WGS84", "local-grid"] },
      ],
    },
  },
  properties: {
    bayLocation: { "tm:ref": "#/schemaDefinitions/TupleArray" },
    serialNumber: { type: "string" },
  },
};

export const metadataThingModel = {
  "@context": ["https://www.w3.org/2022/wot/td/v1.1", { ex: "https://example.test/semantics/" }],
  "@type": "tm:ThingModel",
  id: "urn:fixture:declared-field-metadata",
  title: "Declared metadata device",
  description: "Model meaning retained with its native column documentation.",
  version: { model: "metadata-v1" },
  schemaDefinitions: {
    reading: {
      type: "number", "@type": "ex:Temperature", unit: "degC", format: "float",
      minimum: 0, exclusiveMinimum: -1, maximum: 100, exclusiveMaximum: 101, multipleOf: 0.1,
      const: 0, default: 0, enum: [0, 1, 100], readOnly: false, writeOnly: false,
      titles: { en: "Resolved reading", de: "Aufgelöster Messwert" },
      descriptions: { en: "Resolved local reading.", de: "Lokaler Messwert." },
    },
  },
  properties: {
    temperature: {
      "tm:ref": "#/schemaDefinitions/reading",
      title: "Temperature", description: "A \"quoted\" reading with \\ characters.\nSecond line.", observable: false,
    },
    state: { type: "boolean", title: "Operating state", const: false, default: false, enum: [false, true], readOnly: false },
    message: {
      type: "string", title: "Message", description: "", const: "", default: "",
      enum: ["", "on", "quote\"and\\slash", ...Array.from({ length: 24 }, (_, index) => `declared-code-${index}`)],
      minLength: 0, maxLength: 80, pattern: "^[a-z]*$",
      contentEncoding: "base64", contentMediaType: "text/plain", writeOnly: true,
    },
    config: {
      type: "object", title: "Configuration", description: "One whole configuration value.",
      default: null as JsonValue, const: { levels: [] as number[], enabled: false },
      properties: { levels: { type: "array", items: { type: "number", minimum: 0 }, minItems: 0, maxItems: 8 },
        enabled: { type: "boolean", default: false } },
      required: ["levels"],
    },
  },
  events: {
    alarm: { title: "Alarm", description: "Whole event value.", "@type": "ex:Alarm",
      data: { type: "object", "@type": "ex:AlarmData", properties: { code: { type: "integer", enum: [0, 1] } }, required: ["code"] } },
  },
};

export const metadataDtdlModel = [
  { "@id": "dtmi:fixture:SemanticBase;1", "@type": "Interface", contents: [
    { "@id": "dtmi:fixture:SemanticBase:temperature;1", "@type": ["Telemetry", "Temperature"], name: "temperature",
      displayName: { en: "Temperature", de: "Temperatur" }, description: "Current measured temperature",
      comment: "Sensor measurement.", schema: "double", unit: "degreeCelsius" },
  ] },
  { "@id": "dtmi:fixture:SemanticMotor;1", "@type": "Interface", contents: [
    { "@type": "Property", name: "configuration", displayName: "Motor configuration", writable: false,
      schema: { "@type": "Object", fields: [
        { name: "offset", displayName: "Offset", description: "Calibration offset", schema: "double" },
        { name: "enabled", schema: "boolean" },
      ] } },
  ] },
  { "@id": "dtmi:fixture:SemanticDevice;2", "@type": "Interface", displayName: "Semantic device",
    extends: "dtmi:fixture:SemanticBase;1", schemas: [
      { "@id": "dtmi:fixture:OperatingMode;1", "@type": "Enum", valueSchema: "integer", enumValues: [
        { name: "idle", enumValue: 0, displayName: "Idle", description: "Not running" },
        { name: "active", enumValue: 1, displayName: "Active" },
      ] },
    ], contents: [
      { "@type": "Component", name: "motor", schema: "dtmi:fixture:SemanticMotor;1" },
      { "@type": "Property", name: "mode", displayName: "Operating mode", writable: true, schema: "dtmi:fixture:OperatingMode;1" },
      { "@type": "Telemetry", name: "samples", schema: { "@type": "Array", elementSchema: { "@type": "Object", fields: [
        { name: "value", schema: "double" },
      ] } } },
      { "@type": "Property", name: "labels", schema: { "@type": "Map", mapKey: { name: "label", schema: "string" },
        mapValue: { name: "value", schema: "string" } } },
      { "@type": "Command", name: "reboot" },
      { "@type": ["Property", "Cloud"], name: "cloudOnly", schema: "string" },
    ] },
];
