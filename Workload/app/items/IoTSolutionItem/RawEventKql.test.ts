// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { RAW_EVENT_PROJECTION_KQL } from "./RawEventKql.ts";

const statements = RAW_EVENT_PROJECTION_KQL.split("\n");

test("reads the JSON body and removes only top-level Eventstream metadata", () => {
  assert.equal(
    statements[0],
    '| extend data = bag_remove_keys(parse_json(data), dynamic(["EventProcessedUtcTime", "PartitionId", "EventEnqueuedUtcTime", "EventHub", "User"]))'
  );
  assert.doesNotMatch(statements[0], /properties|reported|__t|deviceType/);
});

test("uses twin device metadata with the telemetry subject as fallback", () => {
  assert.equal(
    statements[1],
    "| extend deviceId = coalesce(tostring(user_headers.cloudEvents_deviceid), tostring(user_headers.cloudEvents_subject))"
  );
});

test("uses CloudEvent time before the twin operation timestamp", () => {
  assert.equal(
    statements[2],
    "| extend enqueuedTime = coalesce(todatetime(user_headers.cloudEvents_time), todatetime(user_headers.cloudEvents_operationtimestamp))"
  );
});

test("reads component identity only from CloudEvent application properties", () => {
  assert.equal(
    statements[3],
    "| extend component = tostring(user_headers.cloudEvents_iothubdtsubject)"
  );
  assert.doesNotMatch(RAW_EVENT_PROJECTION_KQL, /deviceType|headers\.IoT|data\.User/);
});

test("does not model records without the new device and time metadata", () => {
  assert.equal(
    statements[4],
    "| where isnotempty(deviceId) and isnotnull(enqueuedTime)"
  );
});

test("has no structured CloudEvent, legacy-header, or base64 decoder", () => {
  assert.equal(statements.length, 5);
  assert.doesNotMatch(
    RAW_EVENT_PROJECTION_KQL,
    /data_base64|data_json|base64_decode|__rawEvent|__isNativeEvent|IoTConnectionDeviceId|IoTEnqueueTime|IoTSubject|bag_merge|data\.data/
  );
});
