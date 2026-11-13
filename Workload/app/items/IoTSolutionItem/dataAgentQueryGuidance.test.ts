// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  buildDataAgentQueryGuidelines,
  buildIndependentLatestValuesQuery,
  buildLatestValueQuery,
} from "./dataAgentQueryGuidance.ts";

test("filters null measurements before selecting the latest value", () => {
  assert.equal(
    buildLatestValueQuery("ModeledData_Thermostat", "temperature"),
    "['ModeledData_Thermostat'] | where isnotnull(['temperature']) | summarize arg_max(event_time, *) by deviceId | project deviceId, event_time, ['temperature']"
  );
});

test("independent latest values reuse per-column filters and preserve native types and each timestamp with a full outer join", () => {
  const query = buildIndependentLatestValuesQuery("Selected's table", { name: "root value", type: "real" },
    { name: "chosen'\\[]", type: "string" });
  assert.equal(query, [
    "['Selected\\'s table'] | where isnotnull(['root value']) | summarize arg_max(event_time, *) by deviceId | project deviceId, event_time, ['root value']",
    "| project deviceId, Value1 = ['root value'], event_time1 = event_time",
    "| join kind=fullouter (",
    "  ['Selected\\'s table'] | where isnotempty(['chosen\\'\\\\[]']) | summarize arg_max(event_time, *) by deviceId | project deviceId, event_time, ['chosen\\'\\\\[]']",
    "  | project deviceId, Value2 = ['chosen\\'\\\\[]'], event_time2 = event_time",
    ") on deviceId",
    "| project deviceId = coalesce(deviceId, deviceId1), Value1, event_time1, Value2, event_time2",
  ].join("\n"));
  assert.doesNotMatch(query, /join kind=(?:inner|leftouter)|where .* and |to(?:string|dynamic|real|double|decimal)\(/);
  assert.equal((query.match(/arg_max\(event_time, \*\)/g) || []).length, 2);
});

test("explains sparse telemetry and latest-value semantics", () => {
  const guidance = buildDataAgentQueryGuidelines("telemetry-event");

  assert.match(guidance, /Telemetry events can be sparse/);
  assert.match(guidance, /where isnotnull\(columnName\)/);
  assert.match(guidance, /calculate each column's latest non-null value independently/);
  assert.match(guidance, /does not distinguish omission from an explicit null/);
});

test("explains sparse property-update rows", () => {
  const guidance = buildDataAgentQueryGuidelines("property-update");

  assert.match(guidance, /Reported-property update rows are sparse/);
  assert.match(guidance, /does not distinguish an omitted property from an explicit null/);
});
