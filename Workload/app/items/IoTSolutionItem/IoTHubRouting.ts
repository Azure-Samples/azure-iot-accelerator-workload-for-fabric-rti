// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const IOT_HUB_ROUTING_API_VERSION = "2026-05-01-preview";

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

function toEndpoint(endpoint: NativeEventStreamEndpoint): NamedResource {
  return {
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
        toEndpoint(input.telemetryEndpoint),
        toEndpoint(input.propertiesEndpoint),
      ],
    },
    routes: [
      ...withoutNames(routing.routes, routeNames),
      {
        name: input.telemetryRouteName,
        source: "DeviceMessages",
        condition: "true",
        endpointNames: [input.telemetryEndpoint.name],
        isEnabled: true,
      },
      {
        name: input.propertiesRouteName,
        source: "TwinChangeEvents",
        condition: "true",
        endpointNames: [input.propertiesEndpoint.name],
        isEnabled: true,
      },
    ],
  };
}
