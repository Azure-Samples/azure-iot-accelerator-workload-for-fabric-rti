import assert from "node:assert/strict";
import test from "node:test";
import { decodeJsonFromBase64, encodeJsonToBase64 } from "./base64.ts";

test("round-trips Unicode JSON values", () => {
  const value = {
    displayName: "温度センサー",
    description: "Détecteur de température 🌡️",
    nested: { unit: "°C" },
  };

  assert.deepEqual(decodeJsonFromBase64(encodeJsonToBase64(value)), value);
});

test("encodes plain ASCII definitions without changing their data", () => {
  const value = { name: "Thermostat", enabled: true, threshold: 42 };

  assert.deepEqual(decodeJsonFromBase64(encodeJsonToBase64(value)), value);
});
