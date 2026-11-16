// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { parseDtdlModel, UnsupportedDtdlModelFileError } from "./DtdlModelParser.ts";
import { fromDtdlModel, modelingFingerprint } from "./HubModeling.ts";
import { buildModeledDataPlan } from "./ModeledDataPlan.ts";
import { creationRequest } from "../../../tests/creationMock.ts";

test("unsupported files expose the actual DTDL parser cause without classifying supported-model failures as wrong kind", () => {
  for (const document of [{ notAModel: true }, [], "not a model"]) {
    assert.throws(() => parseDtdlModel(document), (error: unknown) => {
      assert.ok(error instanceof UnsupportedDtdlModelFileError);
      assert.equal(error.message, "Unsupported model file. Provide a DTDL interface, an array of interfaces, or an IoT Central device template export.");
      return true;
    });
  }
  assert.throws(() => parseDtdlModel({
    "@id": "dtmi:example:MissingSchema;1", "@type": "Interface",
    contents: [{ "@type": "Telemetry", name: "value" }],
  }), (error: unknown) => {
    assert.ok(error instanceof Error && !(error instanceof UnsupportedDtdlModelFileError));
    assert.match(error.message, /Unsupported or missing DTDL schema/);
    return true;
  });
});

test("parses root telemetry and reported properties", () => {
  const result = parseDtdlModel({
    "@id": "dtmi:example:Thermostat;1",
    "@type": "Interface",
    displayName: "Thermostat",
    contents: [
      {
        "@type": "Telemetry",
        name: "temperature",
        schema: "double",
        unit: "degreeCelsius",
      },
      {
        "@type": "Property",
        name: "firmwareVersion",
        schema: "string",
      },
      {
        "@type": ["Property", "Cloud"],
        name: "customerName",
        schema: "string",
      },
    ],
  });

  assert.deepEqual(
    result.telemetries.map(({ name, kustoType }) => ({ name, kustoType })),
    [{ name: "temperature", kustoType: "real" }]
  );
  assert.deepEqual(
    result.properties.map(({ name, kustoType }) => ({ name, kustoType })),
    [{ name: "firmwareVersion", kustoType: "string" }]
  );
});

test("prefixes capabilities resolved through component references", () => {
  const result = parseDtdlModel([
    {
      "@id": "dtmi:example:Machine;1",
      "@type": "Interface",
      contents: [
        {
          "@type": "Component",
          name: "motor",
          schema: "dtmi:example:Motor;1",
        },
      ],
    },
    {
      "@id": "dtmi:example:Motor;1",
      "@type": "Interface",
      contents: [
        {
          "@type": "Telemetry",
          name: "current",
          schema: "double",
        },
      ],
    },
  ]);

  assert.deepEqual(
    result.telemetries.map(({ name, leafName, component }) => ({
      name,
      leafName,
      component,
    })),
    [{ name: "motor_current", leafName: "current", component: "motor" }]
  );
});

test("preserves inherited roles and equal leaf names in separate components", () => {
  const result = parseDtdlModel([
    {
      "@id": "dtmi:example:Machine;1", "@type": "Interface",
      extends: "dtmi:example:Base;1",
      contents: [
        { "@type": "Component", name: "motor", schema: "dtmi:example:Sensor;1" },
        { "@type": "Component", name: "compressor", schema: "dtmi:example:Sensor;1" },
        { "@type": "Command", name: "reboot" },
      ],
    },
    {
      "@id": "dtmi:example:Base;1", "@type": "Interface",
      contents: [
        { "@type": "Property", name: "firmwareVersion", schema: "string" },
        { "@type": ["Property", "Cloud"], name: "customer", schema: "string" },
      ],
    },
    {
      "@id": "dtmi:example:Sensor;1", "@type": "Interface",
      contents: [{ "@type": "Telemetry", name: "temperature", schema: "double" }],
    },
  ]);
  assert.deepEqual(result.telemetries.map(field => [field.name, field.component, field.kind]), [
    ["motor_temperature", "motor", "Telemetry"], ["compressor_temperature", "compressor", "Telemetry"],
  ]);
  assert.deepEqual(result.properties.map(field => [field.name, field.kind]), [["firmwareVersion", "Property"]]);
});

test("continues accepting the IoT Central device-template export envelope", () => {
  const result = parseDtdlModel({
    "@type": ["ModelDefinition", "DeviceModel"],
    capabilityModel: {
      "@id": "dtmi:example:Thermostat;1", "@type": "Interface",
      contents: [{ "@type": "Telemetry", name: "temperature", schema: "double" }],
    },
  });
  assert.equal(result.modelId, "dtmi:example:Thermostat;1");
  assert.equal(result.telemetries[0].name, "temperature");
});

test("all six DTDL geospatial schemas remain whole dynamic values at root, inherited and component bindings", () => {
  const geospatial = ["point", "multiPoint", "lineString", "multiLineString", "polygon", "multiPolygon"];
  const rootId = "dtmi:example:GeoDevice;1", baseId = "dtmi:example:GeoBase;1", componentId = "dtmi:example:GeoSensor;1";
  const capabilities = geospatial.map(schema => ({
    "@type": "Telemetry", name: schema, schema, description: `Declared ${schema} value`,
  }));
  const document = [
    { "@id": rootId, "@type": "Interface", extends: baseId,
      contents: [...capabilities, { "@type": "Component", name: "sensor", schema: componentId }] },
    { "@id": baseId, "@type": "Interface",
      contents: geospatial.map(schema => ({ "@type": "Property", name: `last_${schema}`, schema })) },
    { "@id": componentId, "@type": "Interface", contents: capabilities },
  ];
  const parsed = parseDtdlModel(document);
  const model = fromDtdlModel(parsed, document);
  assert.equal(model.fields.length, 18);
  for (const schema of geospatial) {
    const root = parsed.telemetries.find(field => field.name === schema)!;
    const component = parsed.telemetries.find(field => field.name === `sensor_${schema}`)!;
    const inherited = parsed.properties.find(field => field.name === `last_${schema}`)!;
    assert.equal(root.declaredBy, rootId);
    assert.equal(component.component, "sensor");
    assert.equal(component.declaredBy, componentId);
    assert.equal(inherited.declaredBy, baseId);
    for (const field of [root, component, inherited]) {
      assert.equal(field.schema, schema);
      assert.equal(field.declaredSchema, schema);
      assert.equal(field.kustoType, "dynamic");
    }
  }
  const request = creationRequest();
  request.input = { ...request.input, model, fields: model.fields };
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
  const table = plan.commands.find(command => command.kind === "table" && command.target === request.names.modeled);
  assert.ok(table?.kind === "table");
  assert.deepEqual(table.columns.slice(2), model.fields.map(field => ({ name: field.column, type: "dynamic" })));
  assert.ok(model.fields.every(field => field.path.length === 1 && field.roleSource === "model" &&
    field.declaredMetadata?.schema === field.schema));
});

test("geospatial support preserves unknown-schema rejection and root/component output collisions", () => {
  assert.throws(() => parseDtdlModel({
    "@id": "dtmi:example:GeoUnknown;1", "@type": "Interface",
    contents: [{ "@type": "Telemetry", name: "location", schema: "point" },
      { "@type": "Telemetry", name: "unknown", schema: "unsupportedGeometry" }],
  }), /Unsupported or missing DTDL schema on unknown/);
  assert.throws(() => parseDtdlModel([
    { "@id": "dtmi:example:GeoCollision;1", "@type": "Interface",
      contents: [{ "@type": "Telemetry", name: "sensor_location", schema: "point" },
        { "@type": "Component", name: "sensor", schema: "dtmi:example:GeoComponent;1" }] },
    { "@id": "dtmi:example:GeoComponent;1", "@type": "Interface",
      contents: [{ "@type": "Property", name: "location", schema: "multiPolygon" }] },
  ]), /Distinct DTDL paths collide/);
});
