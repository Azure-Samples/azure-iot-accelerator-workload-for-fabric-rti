// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { RAW_EVENT_NORMALIZATION_KQL } from "./RawEventKql.ts";

test("normalizes legacy and native Eventstream row formats", () => {
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /parse_json\(data\)/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /specversion/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /__rawEvent\.id/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /__rawEvent\.source/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /__rawEvent\.type/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /__rawEvent\.data/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /base64_decode_tostring/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /__rawEvent\.deviceid/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /__rawEvent\["time"\]/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /__rawEvent\.iothubdtsubject/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /headers\.IoTConnectionDeviceId/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /headers\.IoTEnqueueTime/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /headers\.IoTSubject/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /userProperties/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /cloudEvents_specversion/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /cloudEvents_id/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /cloudEvents_source/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /cloudEvents_type/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /cloudEvents_deviceid/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /cloudEvents_time/);
  assert.match(RAW_EVENT_NORMALIZATION_KQL, /cloudEvents_iothubdtsubject/);
});
