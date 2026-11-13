import test from "node:test";
import assert from "node:assert/strict";
import { expectedModelValues } from "../src/modeling.js";
import { modelFixture } from "../src/models.js";

const device = {
  telemetry: { temperature: 23.75, gpsLocation: { lat: 47.6062, lon: -122.3321, floor: 2 } },
  reported: {
    firmwareVersion: "1.0.0",
    tags: ["hvac", "zone-a", "monitored"],
    maintenanceSchedule: { intervalDays: 90, lastServiceDate: "2025-10-05" },
    thermostat: { __t: "c", setpoint: 22 },
    powerModule: { __t: "c", batteryHealth: "good" },
  },
  components: {
    thermostat: { targetTemperature: 22.5 },
    powerModule: { powerConsumption: 75.5 },
  },
};

test("modeled expectations use exact sent root and component values, not table readbacks", () => {
  assert.deepEqual(expectedModelValues(device, modelFixture("regular").fields), {
    temperature: 23.75,
    gpsLocation: { lat: 47.6062, lon: -122.3321, floor: 2 },
    thermostat_targetTemperature: 22.5,
    powerModule_powerConsumption: 75.5,
    firmwareVersion: "1.0.0",
    maintenanceSchedule: { intervalDays: 90, lastServiceDate: "2025-10-05" },
    thermostat_setpoint: 22,
    powerModule_batteryHealth: "good",
  });
  const wot = expectedModelValues(device, modelFixture("combined").fields);
  assert.deepEqual(Object.keys(wot), ["temperature", "gpsLocation", "firmwareVersion", "tags", "maintenanceSchedule"]);
  assert.deepEqual(wot.tags, device.reported.tags);
});

test("missing selected simulator fields fail before native expectation construction", () => {
  assert.throws(() => expectedModelValues({ ...device, components: {} }, modelFixture("regular").fields),
    /Missing simulated value for selected field thermostat_targetTemperature/);
});
