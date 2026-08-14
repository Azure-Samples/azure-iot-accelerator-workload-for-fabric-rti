// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export function buildLatestValueQuery(
  tableName: string,
  columnName: string
): string {
  return `${tableName} | where isnotnull(${columnName}) | summarize arg_max(enqueuedTime, *) by deviceId | project deviceId, enqueuedTime, ${columnName}`;
}

export function buildDataAgentQueryGuidelines(rowKind?: string): string {
  const sparseRowGuidance =
    rowKind === "property-update"
      ? "- Reported-property update rows are sparse. A null property means that property was not included in that update.\n"
      : "- Telemetry events can be sparse: a device message may contain only a subset of its telemetry measurements. A null telemetry value means the measurement was absent from that event, not that the device has never reported it.\n";

  return (
    `## Query Guidelines\n` +
    `- Always filter by \`deviceId\` when the user asks about a specific device.\n` +
    `- Use \`enqueuedTime\` for time-range filters (e.g., \`where enqueuedTime >= ago(1h)\`).\n` +
    `- For aggregations over time, use \`summarize ... by bin(enqueuedTime, interval)\`.\n` +
    `- To find the latest complete event row, use \`arg_max(enqueuedTime, *)\` grouped by \`deviceId\`.\n` +
    `- To find the latest value of a specific telemetry or property column, first filter with \`where isnotnull(columnName)\`, then use \`arg_max(enqueuedTime, *)\` grouped by \`deviceId\`.\n` +
    `- If the user asks for the latest values of multiple telemetry columns, calculate each column's latest non-null value independently because those measurements may occur in different events.\n` +
    sparseRowGuidance +
    `- Property columns may be null if the device has never reported that property.\n`
  );
}
