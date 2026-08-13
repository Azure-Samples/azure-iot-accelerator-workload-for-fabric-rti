/**
 * Normalize each raw Eventstream row into the legacy logical contract used by
 * the workload:
 *   - data: the device telemetry or twin-change payload
 *   - headers.IoTConnectionDeviceId: device ID
 *   - headers.IoTEnqueueTime: source event time
 *   - headers.IoTSubject: component name, or empty for root telemetry
 *
 * Native IoT Hub Eventstream routes emit CloudEvents. Telemetry is either in
 * `data` or JSON encoded in `data_base64`; twin changes are in `data`.
 */
export const RAW_EVENT_NORMALIZATION_KQL = [
  "| extend __rawEvent = parse_json(data)",
  '| extend __isNativeEvent = tostring(__rawEvent.specversion) == "1.0" and isnotempty(tostring(__rawEvent.id)) and isnotempty(tostring(__rawEvent.source)) and isnotempty(tostring(__rawEvent.type))',
  "| extend __nativePayload = iff(isnotnull(__rawEvent.data), __rawEvent.data, parse_json(base64_decode_tostring(tostring(__rawEvent.data_base64))))",
  "| extend data = iff(__isNativeEvent, __nativePayload, __rawEvent)",
  "| extend __deviceId = iff(__isNativeEvent, coalesce(tostring(__rawEvent.deviceid), tostring(__rawEvent.subject)), tostring(headers.IoTConnectionDeviceId))",
  '| extend __enqueuedTime = iff(__isNativeEvent, coalesce(tostring(__rawEvent["time"]), tostring(__rawEvent.operationtimestamp), tostring(__rawEvent.EventEnqueuedUtcTime)), tostring(headers.IoTEnqueueTime))',
  '| extend __subject = iff(__isNativeEvent, tostring(__rawEvent.iothubdtsubject), tostring(headers.IoTSubject))',
  '| extend headers = bag_merge(bag_pack("IoTConnectionDeviceId", __deviceId, "IoTEnqueueTime", __enqueuedTime, "IoTSubject", __subject), headers)',
  "| project-away __rawEvent, __isNativeEvent, __nativePayload, __deviceId, __enqueuedTime, __subject",
].join("\n");
