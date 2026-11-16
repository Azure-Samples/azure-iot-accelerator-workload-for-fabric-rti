import test from "node:test";
import assert from "node:assert/strict";
import { assertRawStreamDestination, routingTargetsReady } from "../src/routing.js";
import { runConfig } from "../src/config.js";
import { createPlan } from "../src/plan.js";
import { ingestionPath } from "../src/paths.js";

const config = runConfig({
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "published",
});

test("native routing readiness requires both exact current DO endpoints, not an accepted UI write or an older endpoint", () => {
  const ids = ["telemetry-owned", "properties-owned"];
  const endpoint = (id: string, format = "DOObservationV1") =>
    ({ eventStreamId: id, workspaceId: config.workspaceId, messagePayloadFormat: format });
  assert.equal(routingTargetsReady({}, config.workspaceId, ids), false);
  assert.equal(routingTargetsReady({ endpoints: { eventStreams: [endpoint(ids[0]!)] } }, config.workspaceId, ids), false);
  assert.equal(routingTargetsReady({ endpoints: { eventStreams: ids.map(id => endpoint(id, "CloudEvents")) } }, config.workspaceId, ids), false);
  assert.equal(routingTargetsReady({ endpoints: { eventStreams: ids.map(id => ({ ...endpoint(id), workspaceId: "foreign" })) } }, config.workspaceId, ids), false);
  assert.equal(routingTargetsReady({ endpoints: { eventStreams: ids.map(id => endpoint(id)) } }, config.workspaceId, ids), true);
  assert.equal(routingTargetsReady({ endpoints: { eventStreams: [...ids.map(id => endpoint(id)), endpoint(ids[0]!)] } }, config.workspaceId, ids), false);
  assert.throws(() => routingTargetsReady({}, config.workspaceId, []), /distinct/);
  assert.throws(() => routingTargetsReady({ endpoints: { eventStreams: {} } }, config.workspaceId, ids), /not an array/);
});

test("a raw Eventstream must bind exclusively to the manifest database, workspace, table and processed ingestion mode", () => {
  const plan = createPlan(config), path = ingestionPath(plan, "combined");
  plan.observed.databaseId = "00000000-0000-4000-8000-000000000009";
  const destination = { type: "Eventhouse", properties: {
    workspaceId: config.workspaceId, itemId: plan.observed.databaseId, tableName: path.propertiesTable,
    dataIngestionMode: "ProcessedIngestion",
  } };
  assertRawStreamDestination({ destinations: [destination] }, plan, path.propertiesTable);
  for (const changes of [
    { workspaceId: "foreign" }, { itemId: "foreign" }, { tableName: "foreign" }, { dataIngestionMode: "DirectIngestion" },
  ]) assert.throws(() => assertRawStreamDestination({
    destinations: [{ ...destination, properties: { ...destination.properties, ...changes } }],
  }, plan, path.propertiesTable), /manifest/);
  assert.throws(() => assertRawStreamDestination({ destinations: [] }, plan, path.propertiesTable), /one exact/);
  assert.throws(() => assertRawStreamDestination({ destinations: [destination, destination] }, plan, path.propertiesTable), /one exact/);
  assert.throws(() => assertRawStreamDestination({ destinations: [{ ...destination, type: "EventHub" }] }, plan, path.propertiesTable), /manifest/);
});
