// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  buildDataAgentQueryGuidelines,
  buildLatestValueQuery,
} from "./dataAgentQueryGuidance.ts";

test("filters null measurements before selecting the latest value", () => {
  assert.equal(
    buildLatestValueQuery("ModeledData_Thermostat", "temperature"),
    "ModeledData_Thermostat | where isnotnull(temperature) | summarize arg_max(enqueuedTime, *) by deviceId | project deviceId, enqueuedTime, temperature"
  );
});

test("explains sparse telemetry and latest-value semantics", () => {
  const guidance = buildDataAgentQueryGuidelines("telemetry-event");

  assert.match(guidance, /Telemetry events can be sparse/);
  assert.match(guidance, /where isnotnull\(columnName\)/);
  assert.match(guidance, /calculate each column's latest non-null value independently/);
});

test("explains sparse property-update rows", () => {
  const guidance = buildDataAgentQueryGuidelines("property-update");

  assert.match(guidance, /Reported-property update rows are sparse/);
});
