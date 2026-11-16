// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const IOT_HUB_ROUTING_API_VERSION = "2026-10-01-preview";

interface NamedResource {
  name: string;
  [key: string]: unknown;
}

interface RoutingEndpoints {
  eventStreams?: NamedResource[];
  eventHubs?: NamedResource[];
  [key: string]: unknown;
}

export interface IoTHubRoutingConfiguration {
  endpoints?: RoutingEndpoints;
  routes?: NamedResource[];
  [key: string]: unknown;
}

interface IoTHubRoutingResource {
  etag?: unknown;
  properties: {
    routing?: IoTHubRoutingConfiguration;
    [key: string]: unknown;
  };
  [key: string]: unknown;
}

export interface NativeEventStreamEndpoint {
  name: string;
  endpointUri: string;
  entityPath: string;
  workspaceId: string;
  eventStreamId: string;
  sourceId: string;
  userAssignedIdentity: string | null;
}

export interface NativeEventStreamRoutes {
  telemetryEndpoint: NativeEventStreamEndpoint;
  propertiesEndpoint: NativeEventStreamEndpoint;
  telemetryRouteName: string;
  propertiesRouteName: string;
}

function withoutNames(
  resources: NamedResource[] | undefined,
  names: Set<string>
): NamedResource[] {
  return (resources || []).filter((resource) => !names.has(resource.name));
}

function toEndpoint(endpoint: NativeEventStreamEndpoint, current?: NamedResource): NamedResource {
  return {
    ...current,
    name: endpoint.name,
    endpointUri: endpoint.endpointUri,
    entityPath: endpoint.entityPath,
    authenticationType: "identityBased",
    identity: {
      userAssignedIdentity: endpoint.userAssignedIdentity,
    },
    workspaceId: endpoint.workspaceId,
    eventStreamId: endpoint.eventStreamId,
    sourceId: endpoint.sourceId,
    messagePayloadFormat: "DOObservationV1",
  };
}

export function buildNativeEventStreamRouting(
  current: IoTHubRoutingConfiguration | undefined,
  input: NativeEventStreamRoutes
): IoTHubRoutingConfiguration {
  const routing = current || {};
  const endpoints = routing.endpoints || {};
  const endpointNames = new Set([
    input.telemetryEndpoint.name,
    input.propertiesEndpoint.name,
  ]);
  const routeNames = new Set([
    input.telemetryRouteName,
    input.propertiesRouteName,
  ]);

  return {
    ...routing,
    endpoints: {
      ...endpoints,
      eventStreams: [
        ...withoutNames(endpoints.eventStreams, endpointNames),
        toEndpoint(input.telemetryEndpoint, endpoints.eventStreams?.find(endpoint => endpoint.name === input.telemetryEndpoint.name)),
        toEndpoint(input.propertiesEndpoint, endpoints.eventStreams?.find(endpoint => endpoint.name === input.propertiesEndpoint.name)),
      ],
    },
    routes: [
      ...withoutNames(routing.routes, routeNames),
      {
        ...routing.routes?.find(route => route.name === input.telemetryRouteName),
        name: input.telemetryRouteName,
        source: "DeviceMessages",
        condition: "true",
        endpointNames: [input.telemetryEndpoint.name],
        isEnabled: true,
      },
      {
        ...routing.routes?.find(route => route.name === input.propertiesRouteName),
        name: input.propertiesRouteName,
        source: "TwinChangeEvents",
        condition: "true",
        endpointNames: [input.propertiesEndpoint.name],
        isEnabled: true,
      },
    ],
  };
}

export function buildNativeEventStreamHubUpdate(
  current: IoTHubRoutingResource,
  input: NativeEventStreamRoutes
): { body: IoTHubRoutingResource; ifMatch: string } {
  const etag = current.etag;
  if (typeof etag !== "string" || !etag || /[\u0000-\u0020\u007f,]/.test(etag) ||
      etag === "*" || etag === '"*"' || etag.includes('"') && !/^(W\/)?"[^"]+"$/.test(etag)) {
    throw new Error("The current IoT Hub ETag is missing or invalid. Refresh the Hub before configuring routing.");
  }
  if (!current.properties || typeof current.properties !== "object" || Array.isArray(current.properties)) {
    throw new Error("The current IoT Hub properties are unavailable. Refresh the Hub before configuring routing.");
  }
  return {
    ifMatch: etag,
    body: {
      ...current,
      properties: {
        ...current.properties,
        routing: buildNativeEventStreamRouting(current.properties.routing, input),
      },
    },
  };
}
