import test from "node:test";
import assert from "node:assert/strict";
import { assertDashboardComposition, assertDashboardLayout, dashboardSelection, dashboardTipTitle } from "../src/dashboard.js";

test("native onboarding dismissal recognizes only the exact known tips, never errors containing similar words", () => {
  assert.equal(dashboardTipTitle("New \u00b7\nSearch this table\nDismiss"), "Search this table");
  assert.equal(dashboardTipTitle("New \u00b7\nMaximize this visual\nDismiss"), "Maximize this visual");
  assert.equal(dashboardTipTitle("Maximize this visual failed\nDismiss"), undefined);
  assert.equal(dashboardTipTitle("Unable to execute query\nDismiss"), undefined);
  assert.equal(dashboardTipTitle("Access denied\nMaximize this visual\nDismiss"), undefined);
  assert.equal(dashboardTipTitle("New \u00b7 Unknown tip\nDismiss"), undefined);
});

test("nonlinked and combined dashboards explicitly exercise telemetry and property components", () => {
  for (const path of ["regular", "combined"] as const) {
    const fields = dashboardSelection(path);
    assert.equal(fields.filter(field => field.visual === "line").length, 3);
    assert.equal(fields.filter(field => field.visual === "table").length, 1);
    assert.equal(fields.filter(field => field.visual === "card").length, 3);
    assert.deepEqual(fields.filter(field => field.component).map(field => field.name),
      ["thermostat_targetTemperature", "powerModule_powerConsumption", "thermostat_setpoint", "powerModule_batteryHealth"]);
    assert.equal(new Set(fields.map(field => field.name)).size, fields.length);
  }
});

test("linked dashboard values come from linked1's actual simulator index", () => {
  const fields = dashboardSelection("linked");
  assert.equal(fields.length, 3);
  assert.equal(fields[0]?.expected, 25.75);
  assert.ok(fields.every(field => field.component === undefined));
});

test("native tile composition is order-independent but still rejects missing, duplicate and wrong visuals", () => {
  const fields = dashboardSelection("combined").map(field => ({ title: field.name, visualType: field.visual }));
  const decorations = Array.from({ length: 4 }, () => ({ title: "", visualType: "designTile" }));
  const tiles = [...fields.toReversed(), ...decorations];
  assert.doesNotThrow(() => assertDashboardComposition(tiles, "combined"));
  assert.throws(() => assertDashboardComposition(tiles.slice(1), "combined"));
  assert.throws(() => assertDashboardComposition([fields[0]!, ...tiles.slice(1)], "combined"));
  const first = tiles[0];
  assert.ok(first);
  assert.throws(() => assertDashboardComposition([{ ...first, visualType: "bar" }, ...tiles.slice(1)], "combined"));
});

for (const count of [1, 3, 4, 7]) test(`native property cards use width six and wrap every three: ${count} properties`, () => {
  const selected = Array.from({ length: count }, (_, index) => ({ name: `property${index}`, visual: "card" as const }));
  const tiles = selected.map((field, index) => ({
    title: field.name, visualType: field.visual,
    layout: { x: (index % 3) * 6, y: 4 + Math.floor(index / 3) * 3, width: 6, height: 3 },
    visualOptions: { multiStat__textSize: "auto", multiStat__valueColumn: "Value", colorRulesDisabled: false, colorRules: [] },
  }));
  assert.doesNotThrow(() => assertDashboardLayout(tiles.toReversed(), selected));
  assert.doesNotThrow(() => assertDashboardLayout(tiles.map((tile, index) => ({
    ...tile, title: selected[count - 1 - index]!.name,
  })), selected));
  if (count > 1) assert.throws(() => assertDashboardLayout(tiles.map(tile => ({
    ...tile, layout: { ...tile.layout, x: 0, y: 4 },
  })), selected));
  assert.throws(() => assertDashboardLayout(tiles.map((tile, index) => ({
    ...tile, layout: { x: (index % 6) * 3, y: 4 + Math.floor(index / 6) * 3, width: 3, height: 3 },
  })), selected));
  assert.throws(() => assertDashboardLayout(tiles.map(tile => ({ ...tile, layout: { ...tile.layout, x: 18 } })), selected));
  assert.throws(() => assertDashboardLayout(tiles.map(tile => ({
    ...tile, visualOptions: { ...tile.visualOptions, multiStat__valueColumn: "Incorrect" },
  })), selected));
});

test("valid chart/table/card slots do not impose a fixture-specific field order", () => {
  const selected = dashboardSelection("regular");
  const positions = [
    ["powerModule_powerConsumption", "line", 0, 4, 9, 7],
    ["temperature", "line", 9, 4, 9, 7],
    ["thermostat_targetTemperature", "line", 0, 11, 9, 7],
    ["status", "table", 0, 19, 9, 7],
    ["powerModule_batteryHealth", "card", 0, 27, 6, 3],
    ["firmwareVersion", "card", 6, 27, 6, 3],
    ["thermostat_setpoint", "card", 12, 27, 6, 3],
  ] as const;
  const tiles = positions.map(([title, visualType, x, y, width, height]) => ({
    title, visualType, layout: { x, y, width, height },
    visualOptions: { multiStat__textSize: "auto", multiStat__valueColumn: "Value", colorRulesDisabled: false, colorRules: [] },
  }));
  assert.doesNotThrow(() => assertDashboardLayout(tiles, selected));
  assert.throws(() => assertDashboardLayout(tiles.map(tile => ({
    ...tile, layout: { ...tile.layout, y: 4 },
  })), selected));
});
