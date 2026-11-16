import test from "node:test";
import assert from "node:assert/strict";
import { queryRows } from "../src/kusto.js";

const primary = { Columns: [{ ColumnName: "data" }], Rows: [['{"value":23.75}']] };
const status = (severity: number, code: number) => ({
  Columns: [{ ColumnName: "Severity" }, { ColumnName: "StatusCode" }],
  Rows: [[severity, code]],
});

test("native Kusto primary data is returned without treating status rows as telemetry", () => {
  assert.deepEqual(queryRows({ Tables: [primary, status(4, 0)] }), [{ data: '{"value":23.75}' }]);
});

test("native Kusto in-band failures cannot produce a success-shaped empty result", () => {
  assert.throws(() => queryRows({ Tables: [primary, status(2, 0)] }), /in-band query failure/);
  assert.throws(() => queryRows({ Tables: [primary, status(4, 400)] }), /status 400/);
  assert.throws(() => queryRows({ Tables: [primary], Exceptions: ["failure"] }), /query exceptions/);
});

test("missing or malformed native results fail explicitly", () => {
  assert.throws(() => queryRows({ Tables: [] }), /no result table/);
  assert.throws(() => queryRows({ Tables: [{ Columns: primary.Columns, Rows: [[]] }] }), /invalid row shape/);
});
