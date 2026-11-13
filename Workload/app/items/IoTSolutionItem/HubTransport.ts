// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import type { JsonObject, SampleStream } from "./ModelAuthoring.ts";

export interface NativeBinding {
  profile: "eventhub" | "cloudEvents" | "unknown";
  identity?: string;
  component?: string;
  issues: string[];
}

// These are different transport contracts, not alternate semantic identities.
export const HUB_METADATA = {
  eventhub: { device: "IoTConnectionDeviceId", module: "IoTConnectionModuleId", component: "IoTSubject" },
  cloudEvents: { subject: "cloudEvents_subject", device: "cloudEvents_deviceid",
    module: "cloudEvents_moduleid", component: "cloudEvents_iothubdtsubject" },
} as const;
const own = (value: JsonObject, key: string) => Object.prototype.hasOwnProperty.call(value, key);
const string = (value: unknown) => typeof value === "string" && value ? value : undefined;

export function bindNativeMetadata(headers: JsonObject, user: JsonObject, stream: SampleStream): NativeBinding {
  const eh = HUB_METADATA.eventhub, ce = HUB_METADATA.cloudEvents;
  const ceIdentity = stream === "telemetry" ? ce.subject : ce.device;
  const cloudEvents = own(user, ceIdentity);
  const empty = (value: unknown) => value === undefined || value === null || value === "";
  // An unused EventHub profile can be materialized as blank strings by Eventstream.
  const emptyEventhub = headers[eh.device] === "" && empty(headers[eh.module]) && empty(headers[eh.component]) &&
    !!string(user[ceIdentity]);
  const eventhub = !emptyEventhub && (own(headers, eh.device) || own(headers, eh.module) || own(headers, eh.component));
  const profile = eventhub ? "eventhub" : cloudEvents ? "cloudEvents" : "unknown";
  const issues: string[] = [];
  const device = eventhub ? string(headers[eh.device]) : string(user[ce.device]);
  const module = eventhub ? headers[eh.module] : user[ce.module];
  let identity = eventhub || stream === "twin" ? device && (module ? `${device}/${module}` : device)
    : string(user[ce.subject]);
  if (module !== undefined && module !== null && typeof module !== "string") {
    identity = undefined;
    issues.push("Module identity metadata is not a string.");
  }
  const rawComponent = stream === "telemetry" ? eventhub ? headers[eh.component] : user[ce.component] : undefined;
  const component = string(rawComponent);
  if (rawComponent !== undefined && rawComponent !== null && typeof rawComponent !== "string") {
    issues.push("Telemetry component metadata is not a string.");
  }
  if (profile === "unknown") issues.push("The sample does not contain a supported native transport identity.");
  if (!identity) issues.push("Native device identity is missing.");
  if ((eventhub || emptyEventhub) && stream === "twin" && string(user.deviceId) && user.deviceId !== device) {
    identity = undefined;
    issues.push("Twin device identity disagrees with the transport metadata.");
  }
  if (eventhub && cloudEvents) {
    const other = stream === "telemetry" ? string(user[ce.subject])
      : string(user[ce.device]) && (user[ce.module] ? `${user[ce.device]}/${user[ce.module]}` : string(user[ce.device]));
    if ((other && other !== identity) || (stream === "telemetry" && string(user[ce.component]) && user[ce.component] !== component)) {
      identity = undefined;
      issues.push("The two transport metadata profiles disagree. Check the selected source.");
    }
  }
  return { profile, identity, component, issues };
}

/** Use the same profile precedence and keys as the decoded sample binding. */
export function nativeProjection(stream: SampleStream): string[] {
  const eh = HUB_METADATA.eventhub, ce = HUB_METADATA.cloudEvents;
  const key = (bag: string, name: string) => `${bag}[${JSON.stringify(name)}]`;
  const has = (bag: string, name: string) => `array_index_of(bag_keys(${bag}), ${JSON.stringify(name)}) >= 0`;
  const text = (value: string) => `iff(gettype(${value}) == "string", tostring(${value}), "")`;
  const ehPresence = [eh.device, eh.module, eh.component].map(name => has("headers", name)).join(" or ");
  const emptyString = (value: string) => `(gettype(${value}) == "string" and tostring(${value}) == "")`;
  const empty = (value: string) => `(isnull(${value}) or ${emptyString(value)})`;
  const ceIdentity = key("user_headers", stream === "telemetry" ? ce.subject : ce.device);
  const emptyEventhub = `${emptyString(key("headers", eh.device))} and ${empty(key("headers", eh.module))} and ` +
    `${empty(key("headers", eh.component))} and gettype(${ceIdentity}) == "string" and not(isempty(tostring(${ceIdentity})))`;
  const ehProfile = `(${ehPresence}) and not(__emptyEventHubProfile)`;
  const id = (device: string, module: string) =>
    `iff(isempty(tostring(${device})), "", iff(isempty(tostring(${module})), tostring(${device}), strcat(tostring(${device}), "/", tostring(${module}))))`;
  const ehId = id(text(key("headers", eh.device)), key("headers", eh.module));
  const ceId = stream === "telemetry" ? text(key("user_headers", ce.subject))
    : id(text(key("user_headers", ce.device)), key("user_headers", ce.module));
  const module = `iff(__eventHubProfile, ${key("headers", eh.module)}, ${key("user_headers", ce.module)})`;
  const component = stream === "telemetry"
    ? `iff(__eventHubProfile, ${text(key("headers", eh.component))}, ${text(key("user_headers", ce.component))})` : '""';
  const rawComponent = `iff(__eventHubProfile, ${key("headers", eh.component)}, ${key("user_headers", ce.component)})`;
  const componentValid = stream === "telemetry"
    ? ` and (isnull(${rawComponent}) or gettype(${rawComponent}) == "string")` : "";
  const twinDevice = text(key("user_headers", "deviceId"));
  const twinAgreement = stream === "twin"
    ? ` and (not(__eventHubProfile or __emptyEventHubProfile) or (isempty(${twinDevice}) or ` +
      `${twinDevice} == iff(__eventHubProfile, ${text(key("headers", eh.device))}, ${text(key("user_headers", ce.device))})))` : "";
  const otherComponent = stream === "telemetry"
    ? ` and (isempty(${text(key("user_headers", ce.component))}) or ${text(key("user_headers", ce.component))} == component)` : "";
  return [
    `| extend __emptyEventHubProfile = ${emptyEventhub}`,
    `| extend __eventHubProfile = (${ehProfile})`,
    `| extend __nativeId = iff(__eventHubProfile, ${ehId}, ${ceId}), __otherId = ${ceId}, __nativeModule = ${module}`,
    `| extend component = ${component}`,
    `| extend deviceId = iff((isnull(__nativeModule) or gettype(__nativeModule) == "string")${componentValid}${twinAgreement} and ` +
      `(not(__eventHubProfile) or ((isempty(__otherId) or __nativeId == __otherId)${otherComponent})), __nativeId, "")`,
  ];
}
