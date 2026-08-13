import assert from "node:assert/strict";
import test from "node:test";
import { parseDtdlModel } from "./DtdlModelParser.ts";

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
