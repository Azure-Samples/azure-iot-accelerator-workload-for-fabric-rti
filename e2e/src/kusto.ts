import { object, text } from "./cloud.js";

export function queryRows(value: unknown): Record<string, unknown>[] {
  const result = object(value, "native Kusto result");
  if (Array.isArray(result.Exceptions) && result.Exceptions.length) {
    throw new Error("[service] Native Kusto reported query exceptions.");
  }
  if (!Array.isArray(result.Tables) || result.Tables.length < 1) {
    throw new Error("[service-shape] Native Kusto returned no result table.");
  }
  const tables = result.Tables.map(table => object(table, "native Kusto table"));
  for (const table of tables) {
    if (!Array.isArray(table.Columns) || !Array.isArray(table.Rows)) {
      throw new Error("[service-shape] Native Kusto table is missing columns or rows.");
    }
    const columns = table.Columns.map(column => text(object(column, "Kusto column").ColumnName, "Kusto column name"));
    const severity = columns.indexOf("Severity");
    const status = columns.indexOf("StatusCode");
    if (severity !== -1 && status !== -1) {
      for (const row of table.Rows) {
        if (!Array.isArray(row)) throw new Error("[service-shape] Invalid Kusto status row.");
        if (typeof row[severity] === "number" && row[severity] <= 2 ||
            typeof row[status] === "number" && row[status] !== 0) {
          throw new Error(`[service] Native Kusto returned an in-band query failure (status ${String(row[status])}).`);
        }
      }
    }
  }
  const table = tables[0];
  if (!table || !Array.isArray(table.Columns) || !Array.isArray(table.Rows)) {
    throw new Error("[service-shape] Native Kusto primary result is missing.");
  }
  const columns = table.Columns.map(column => text(object(column, "Kusto column").ColumnName, "Kusto column name"));
  return table.Rows.map(row => {
    if (!Array.isArray(row) || row.length !== columns.length) {
      throw new Error("[service-shape] Native Kusto returned an invalid row shape.");
    }
    return Object.fromEntries(columns.map((column, index) => [column, row[index]]));
  });
}
