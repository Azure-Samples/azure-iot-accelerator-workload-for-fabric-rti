// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  buildNativeEventStreamRouting,
  IOT_HUB_ROUTING_API_VERSION,
} from "./IoTHubRouting.ts";

test("builds native Eventstream endpoints and preserves unrelated routing", () => {
  const routing = buildNativeEventStreamRouting(
    {
      endpoints: {
        eventHubs: [
          {
            name: "fabric-telemetry-old",
            endpointUri: "sb://telemetry.servicebus.windows.net",
            entityPath: "telemetry_eh",
          },
          { name: "unrelated-event-hub", endpointUri: "keep" },
        ],
        eventStreams: [
          { name: "fabric-properties-abcd", endpointUri: "stale" },
          { name: "unrelated-eventstream", endpointUri: "keep" },
        ],
        storageContainers: [],
      },
      routes: [
        {
          name: "fabric-telemetry-route-old",
          source: "DeviceMessages",
          endpointNames: ["fabric-telemetry-old"],
        },
        { name: "unrelated-route", source: "DeviceMessages" },
      ],
    },
    {
      telemetryEndpoint: {
        name: "fabric-telemetry-abcd",
        endpointUri: "sb://telemetry.servicebus.windows.net",
        entityPath: "telemetry_eh",
        workspaceId: "workspace-id",
        eventStreamId: "telemetry-stream-id",
        sourceId: "telemetry-source-id",
        userAssignedIdentity: null,
      },
      propertiesEndpoint: {
        name: "fabric-properties-abcd",
        endpointUri: "sb://properties.servicebus.windows.net",
        entityPath: "properties_eh",
        workspaceId: "workspace-id",
        eventStreamId: "properties-stream-id",
        sourceId: "properties-source-id",
        userAssignedIdentity: "/subscriptions/sub/resourceGroups/rg/providers/Microsoft.ManagedIdentity/userAssignedIdentities/uami",
      },
      telemetryRouteName: "fabric-telemetry-route-abcd",
      propertiesRouteName: "fabric-properties-route-abcd",
    }
  );

  assert.equal(IOT_HUB_ROUTING_API_VERSION, "2026-05-01-preview");
  assert.deepEqual(
    routing.endpoints?.eventHubs?.map((endpoint) => endpoint.name),
    ["fabric-telemetry-old", "unrelated-event-hub"]
  );
  assert.deepEqual(
    routing.endpoints?.eventStreams?.map((endpoint) => endpoint.name),
    [
      "unrelated-eventstream",
      "fabric-telemetry-abcd",
      "fabric-properties-abcd",
    ]
  );
  assert.deepEqual(routing.endpoints?.eventStreams?.[1], {
    name: "fabric-telemetry-abcd",
    endpointUri: "sb://telemetry.servicebus.windows.net",
    entityPath: "telemetry_eh",
    authenticationType: "identityBased",
    identity: { userAssignedIdentity: null },
    workspaceId: "workspace-id",
    eventStreamId: "telemetry-stream-id",
    sourceId: "telemetry-source-id",
  });
  assert.deepEqual(
    routing.routes?.map((route) => route.name),
    [
      "fabric-telemetry-route-old",
      "unrelated-route",
      "fabric-telemetry-route-abcd",
      "fabric-properties-route-abcd",
    ]
  );
});
