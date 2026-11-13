// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { generateDashboardDefinition, type DashboardField } from "./DashboardDefinitionGenerator.ts";
import { isJsonObject } from "./ModelAuthoring.ts";

const databaseId = "11111111-1111-4111-8111-111111111111";
const introHtml = '<div style="font-size: 24px;">Device dashboard</div><div>(Choose a device ID above to get started)</div>';

function object(value: unknown): Record<string, unknown> {
  assert.ok(isJsonObject(value));
  return value;
}

function objects(value: unknown): Record<string, unknown>[] {
  assert.ok(Array.isArray(value));
  return value.map(object);
}

function layout(tile: Record<string, unknown>) {
  const { x, y, width, height } = object(tile.layout);
  assert.ok(typeof x === "number" && typeof y === "number" && typeof width === "number" && typeof height === "number");
  return { x, y, width, height };
}

for (const [numeric, recent, properties] of [
  [0, 0, 0], [1, 0, 0], [0, 2, 0], [0, 0, 3], [0, 0, 4], [0, 0, 7], [2, 1, 6], [3, 3, 7], [2, 2, 12],
] as const) {
  test(`native dashboard sections lay out ${numeric} charts, ${recent} tables and ${properties} cards without empty sections`, () => {
    const fields = (prefix: string, count: number, isNumeric: boolean): DashboardField[] =>
      Array.from({ length: count }, (_, index) => ({ key: `${prefix}_${index}`, included: true, isNumeric }));
    const numericFields = fields("numeric", numeric, true), recentFields = fields("recent", recent, false);
    const propertyFields = fields("property", properties, false);
    const definition = object(JSON.parse(JSON.stringify(generateDashboardDefinition({
      title: "Native section fixture", telemetryTable: "Raw Telemetry", propertiesTable: "Raw Properties",
      telemetryFields: [{ key: "excluded_numeric", included: false, isNumeric: true }, ...numericFields,
        { key: "excluded_recent", included: false, isNumeric: false }, ...recentFields],
      propertyFields: [{ key: "excluded_property", included: false }, ...propertyFields],
      queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databaseId,
    }))));
    const tiles = objects(definition.tiles), queries = objects(definition.queries), pages = objects(definition.pages);
    const parameters = objects(definition.parameters), sources = objects(definition.dataSources);
    assert.equal(definition.$schema, "https://dataexplorer.azure.com/static/d/schema/88/dashboard.json");
    assert.equal(definition.schema_version, 88);
    assert.deepEqual(definition.embeddedApps, []);
    assert.equal(sources[0].databaseArtifactId, databaseId);
    assert.equal(sources[0].database, databaseId);
    assert.equal(pages.length, 1);
    assert.equal(tiles.every(tile => tile.pageId === pages[0].id), true);

    const ids = [definition.id, ...[tiles, queries, pages, parameters, sources].flatMap(items => items.map(item => item.id))];
    assert.equal(new Set(ids).size, ids.length);
    assert.ok(ids.every(id => typeof id === "string" && /^[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(id)));
    const { id: introId, ...intro } = tiles[0];
    assert.ok(introId);
    assert.deepEqual(intro, {
      title: "", hideTitle: true, visualType: "designTile", pageId: pages[0].id,
      layout: { x: 0, y: 0, width: 18, height: 3 },
      visualOptions: { designTile__kind: "textBox", textBox__html: introHtml },
    });

    const groups = [
      { fields: numericFields, label: "Time-series telemetry", visualType: "line", columns: 2, width: 9, height: 7 },
      { fields: recentFields, label: "Non-time-series telemetry (last 10 values)", visualType: "table", columns: 2, width: 9, height: 7 },
      { fields: propertyFields, label: "Device properties (last known value)", visualType: "card", columns: 3, width: 6, height: 3 },
    ];
    let position = 1, y = 3;
    for (const group of groups.filter(group => group.fields.length > 0)) {
      const { id: dividerId, ...divider } = tiles[position++];
      assert.ok(dividerId);
      assert.deepEqual(divider, {
        title: "", hideTitle: true, visualType: "designTile", pageId: pages[0].id,
        layout: { x: 0, y, width: 24, height: 1 },
        visualOptions: { designTile__kind: "divider", divider__label: group.label },
      });
      y += 1;
      for (const [index, field] of group.fields.entries()) {
        const tile = tiles[position++];
        assert.equal(tile.title, field.key);
        assert.equal(tile.visualType, group.visualType);
        assert.deepEqual(layout(tile), {
          x: (index % group.columns) * group.width, y: y + Math.floor(index / group.columns) * group.height,
          width: group.width, height: group.height,
        });
      }
      y += Math.ceil(group.fields.length / group.columns) * group.height;
    }
    assert.equal(position, tiles.length);
    for (let first = 0; first < tiles.length; first++) for (let second = first + 1; second < tiles.length; second++) {
      const a = layout(tiles[first]), b = layout(tiles[second]);
      assert.ok(!(a.x < b.x + b.width && a.x + a.width > b.x && a.y < b.y + b.height && a.y + a.height > b.y),
        `Tiles ${first} and ${second} overlap.`);
    }

    const decorations = tiles.filter(tile => tile.visualType === "designTile");
    const dataTiles = tiles.filter(tile => tile.visualType !== "designTile");
    assert.ok(dataTiles.every(tile => layout(tile).x >= 0 && layout(tile).x + layout(tile).width <= 18));
    assert.ok(dataTiles.filter(tile => tile.visualType === "card").every(tile =>
      JSON.stringify(tile.visualOptions) === JSON.stringify({
        multiStat__textSize: "auto", multiStat__valueColumn: "Value", colorRulesDisabled: false, colorRules: [],
      })));
    assert.equal(decorations.length, 1 + groups.filter(group => group.fields.length > 0).length);
    assert.ok(decorations.every(tile => !Object.hasOwn(tile, "queryRef") && !Object.hasOwn(tile, "dataSource")));
    assert.equal(queries.length, 1 + numeric + recent + properties);
    assert.deepEqual(dataTiles.map(tile => object(tile.queryRef).queryId), queries.slice(1).map(query => query.id));
    assert.ok(queries.every(query => object(query.dataSource).dataSourceId === sources[0].id));
    assert.equal(object(object(parameters[1].dataSource).queryRef).queryId, queries[0].id);
  });
}
