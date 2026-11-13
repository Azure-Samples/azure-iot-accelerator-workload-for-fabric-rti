// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  buildNativeEventStreamRouting,
  buildNativeEventStreamHubUpdate,
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

  assert.equal(IOT_HUB_ROUTING_API_VERSION, "2026-10-01-preview");
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
    messagePayloadFormat: "DOObservationV1",
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
  assert.equal(routing.endpoints?.eventStreams?.[2].messagePayloadFormat, "DOObservationV1");
  assert.deepEqual(routing.endpoints?.eventStreams?.[2].identity, {
    userAssignedIdentity: "/subscriptions/sub/resourceGroups/rg/providers/Microsoft.ManagedIdentity/userAssignedIdentities/uami",
  });
});

function input(): Parameters<typeof buildNativeEventStreamRouting>[1] {
  const endpoint = {
    name: "owned-telemetry", endpointUri: "sb://telemetry.example", entityPath: "telemetry",
    workspaceId: "workspace", eventStreamId: "telemetry-stream", sourceId: "telemetry-source", userAssignedIdentity: null,
  };
  return {
    telemetryEndpoint: endpoint,
    propertiesEndpoint: { ...endpoint, name: "owned-properties", endpointUri: "sb://properties.example",
      entityPath: "properties", eventStreamId: "properties-stream", sourceId: "properties-source" },
    telemetryRouteName: "owned-telemetry-route", propertiesRouteName: "owned-properties-route",
  };
}

test("DO observation mode belongs only to both owned Eventstream endpoints and is idempotent without changing route sources", () => {
  const request = input(), first = buildNativeEventStreamRouting(undefined, request);
  assert.deepEqual(first.endpoints?.eventStreams?.map(endpoint => endpoint.messagePayloadFormat), ["DOObservationV1", "DOObservationV1"]);
  assert.deepEqual(first.routes, [
    { name: request.telemetryRouteName, source: "DeviceMessages", condition: "true",
      endpointNames: [request.telemetryEndpoint.name], isEnabled: true },
    { name: request.propertiesRouteName, source: "TwinChangeEvents", condition: "true",
      endpointNames: [request.propertiesEndpoint.name], isEnabled: true },
  ]);
  assert.deepEqual(buildNativeEventStreamRouting(first, request), first);
  assert.ok(first.routes?.every(route => !("messagePayloadFormat" in route) && !("dataSchema" in route)));
});

test("conditional full Hub update preserves unrelated Hub state, routing arrays, credentials, schemas and unknown endpoint fields", () => {
  const hub = {
    id: "/subscriptions/sub/resourceGroups/rg/providers/Microsoft.Devices/IotHubs/hub", name: "hub", etag: '"current-version"',
    location: "westus", sku: { name: "S1", capacity: 1 }, tags: { keep: "value" },
    identity: { type: "SystemAssigned, UserAssigned", principalId: "existing-principal",
      userAssignedIdentities: { "/subscriptions/sub/resourceGroups/rg/providers/Microsoft.ManagedIdentity/userAssignedIdentities/existing": {} } },
    properties: {
      deviceRegistry: { namespaceResourceId: "existing-namespace" }, publicNetworkAccess: "Disabled",
      disableLocalAuth: true, networkRuleSets: { defaultAction: "Deny" }, unknownSetting: { keep: false },
      routing: {
        endpoints: {
          eventHubs: [{ name: "existing-eventhub", connectionString: "Endpoint=sb://synthetic;SharedAccessKey=synthetic-fixture-only" }],
          serviceBusQueues: [{ name: "existing-queue", unknownCredentialCarrier: { retain: "synthetic" } }],
          storageContainers: [{ name: "existing-storage", authenticationType: "identityBased" }],
          eventStreams: [
            { name: "unrelated-stream", messagePayloadFormat: "None", identity: { userAssignedIdentity: "keep" } },
            { name: "owned-telemetry", messagePayloadFormat: "None", unknownProviderField: { keep: true } },
            { name: "owned-properties", messagePayloadFormat: "None", unknownProviderField: "preserve" },
          ],
        },
        routes: [
          { name: "unrelated-route", source: "DeviceMessages", condition: "level > 1", isEnabled: false,
            endpointNames: ["existing-eventhub"], dataSchema: { reference: "keep-exact-version" } },
          { name: "owned-telemetry-route", dataSchema: { reference: "owned-existing-schema" }, providerField: true },
        ],
        enrichments: [{ key: "keep", value: "$twin.tags.value", endpointNames: ["existing-eventhub"] }],
        fallbackRoute: { name: "$fallback", source: "DeviceMessages", isEnabled: false, endpointNames: ["events"] },
        unknownRoutingField: { value: 0 },
      },
    },
  };
  const original = structuredClone(hub), update = buildNativeEventStreamHubUpdate(hub, input());
  assert.deepEqual(hub, original);
  assert.equal(update.ifMatch, hub.etag);
  assert.equal(update.body.etag, hub.etag);
  const { properties: beforeProperties, ...beforeRoot } = hub;
  const { properties: afterProperties, ...afterRoot } = update.body;
  assert.deepEqual(afterRoot, beforeRoot);
  const { routing: beforeRouting, ...beforeNonrouting } = beforeProperties;
  const { routing: afterRouting, ...afterNonrouting } = afterProperties;
  assert.deepEqual(afterNonrouting, beforeNonrouting);
  assert.deepEqual(afterRouting?.endpoints?.eventHubs, beforeRouting.endpoints.eventHubs);
  assert.deepEqual(afterRouting?.endpoints?.serviceBusQueues, beforeRouting.endpoints.serviceBusQueues);
  assert.deepEqual(afterRouting?.endpoints?.storageContainers, beforeRouting.endpoints.storageContainers);
  assert.deepEqual(afterRouting?.endpoints?.eventStreams?.[0], beforeRouting.endpoints.eventStreams[0]);
  assert.deepEqual(afterRouting?.enrichments, beforeRouting.enrichments);
  assert.deepEqual(afterRouting?.fallbackRoute, beforeRouting.fallbackRoute);
  assert.deepEqual(afterRouting?.unknownRoutingField, beforeRouting.unknownRoutingField);
  assert.deepEqual(afterRouting?.routes?.[0], beforeRouting.routes[0]);
  assert.deepEqual(afterRouting?.routes?.[1].dataSchema, beforeRouting.routes[1].dataSchema);
  assert.equal(afterRouting?.endpoints?.eventStreams?.[1].unknownProviderField !== undefined, true);
  assert.equal(afterRouting?.endpoints?.eventStreams?.[2].unknownProviderField, "preserve");
  assert.ok(!("messagePayloadFormat" in update.body.properties));
  assert.ok(!("enableDO" in update.body.properties));
});

test("full Hub routing update requires an exact current ETag with no wildcard or unconditional fallback", () => {
  for (const etag of [undefined, null, "", "*", '"*"', " ", "bad\nvalue", '"unterminated', "first,second"]) {
    assert.throws(() => buildNativeEventStreamHubUpdate({ etag, properties: {} }, input()), /ETag/);
  }
  for (const etag of ['"current-version"', "opaque-version=="]) {
    const update = buildNativeEventStreamHubUpdate({ etag, properties: {} }, input());
    assert.equal(update.ifMatch, etag);
  }
  assert.throws(() => buildNativeEventStreamHubUpdate({ etag: '"current"', properties: null }, input()), /properties/);
});
