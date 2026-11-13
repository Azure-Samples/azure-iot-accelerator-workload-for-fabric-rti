import assert from "node:assert/strict";
import type { PathId } from "./paths.js";
import { object } from "./cloud.js";

export const dashboardTipTitles = ["Search this table", "Maximize this visual"] as const;

export function dashboardTipTitle(value: string): string | undefined {
  const title = value.replace(/\s+/g, " ").trim()
    .replace(/^New\s*\u00b7\s*/, "").replace(/\s*Dismiss$/, "").trim();
  return dashboardTipTitles.find(known => title === known);
}

export function dashboardSelection(path: PathId) {
  const fields: { name: string; visual: "line" | "table" | "card"; expected: string | number; component?: string }[] = [
    { name: "temperature", visual: "line", expected: path === "linked" ? 25.75 : 23.75 },
    { name: "status", visual: "table", expected: "normal" },
    { name: "firmwareVersion", visual: "card", expected: "1.0.0" },
  ];
  if (path !== "linked") fields.push(
    { name: "thermostat_targetTemperature", visual: "line", expected: 22.5, component: "thermostat" },
    { name: "powerModule_powerConsumption", visual: "line", expected: 75.5, component: "powerModule" },
    { name: "thermostat_setpoint", visual: "card", expected: 22, component: "thermostat" },
    { name: "powerModule_batteryHealth", visual: "card", expected: "good", component: "powerModule" },
  );
  return fields;
}

export function assertDashboardComposition(tiles: readonly { title: unknown; visualType: unknown }[], path: PathId): void {
  const selected = dashboardSelection(path);
  assert.equal(tiles.length, selected.length + 4);
  const actual = tiles.filter(tile => tile.visualType !== "designTile").map(tile => {
    assert.ok(typeof tile.title === "string" && typeof tile.visualType === "string", "[native-dashboard] Missing tile identity.");
    return JSON.stringify([tile.title, tile.visualType]);
  }).sort();
  const expected = selected.map(field => JSON.stringify([field.name, field.visual])).sort();
  assert.deepEqual(actual, expected, "[native-dashboard] Selected field/visual composition differs from the native dashboard.");
}

export function assertDashboardLayout(tiles: readonly Record<string, unknown>[],
  selected: readonly { name: string; visual: "line" | "table" | "card" }[]): void {
  let y = 3;
  for (const [visual, columns, width, height] of [["line", 2, 9, 7], ["table", 2, 9, 7], ["card", 3, 6, 3]] as const) {
    const fields = selected.filter(field => field.visual === visual);
    if (!fields.length) continue;
    y++;
    const positioned = fields.map(field => {
      const matches = tiles.filter(tile => tile.title === field.name && tile.visualType === visual);
      assert.equal(matches.length, 1, `[native-dashboard] Missing or duplicate layout for ${field.name}.`);
      const tile = matches[0];
      assert.ok(tile);
      const layout = object(tile.layout, "dashboard tile layout");
      const x = layout.x, positionY = layout.y;
      assert.ok(typeof x === "number" && typeof positionY === "number",
        "[native-dashboard] Missing numeric tile coordinates.");
      return { field, tile, layout, x, positionY };
    }).sort((left, right) => left.positionY - right.positionY || left.x - right.x);
    for (const [index, { field, tile, layout }] of positioned.entries()) {
      assert.deepEqual([layout.x, layout.y, layout.width, layout.height],
        [(index % columns) * width, y + Math.floor(index / columns) * height, width, height],
        `[native-dashboard] ${field.name} differs from the published section/card layout.`);
      if (visual === "card") {
        const options = object(tile.visualOptions, "property card options");
        assert.equal(options.multiStat__textSize, "auto");
        assert.equal(options.multiStat__valueColumn, "Value");
        assert.equal(options.colorRulesDisabled, false);
        assert.deepEqual(options.colorRules, []);
      }
    }
    y += Math.ceil(fields.length / columns) * height;
  }
}
