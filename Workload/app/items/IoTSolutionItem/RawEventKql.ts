// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Read IoT Hub binary CloudEvents: JSON body in data, AMQP application properties
 * in user_headers. Remove Eventstream's injected body metadata before field
 * discovery. Rows without device identity or event time cannot be modeled.
 */
export const RAW_EVENT_PROJECTION_KQL = [
  '| extend data = bag_remove_keys(parse_json(data), dynamic(["EventProcessedUtcTime", "PartitionId", "EventEnqueuedUtcTime", "EventHub", "User"]))',
  "| extend deviceId = coalesce(tostring(user_headers.cloudEvents_deviceid), tostring(user_headers.cloudEvents_subject))",
  "| extend enqueuedTime = coalesce(todatetime(user_headers.cloudEvents_time), todatetime(user_headers.cloudEvents_operationtimestamp))",
  "| extend component = tostring(user_headers.cloudEvents_iothubdtsubject)",
  "| where isnotempty(deviceId) and isnotnull(enqueuedTime)",
].join("\n");
