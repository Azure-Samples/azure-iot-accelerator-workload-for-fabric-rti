import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import { Cloud, object, text } from "./cloud.js";
import { artifacts, requireGuid } from "./config.js";
import { query } from "./data.js";
import { moveOwnedItem } from "./folders.js";
import { operationFromHeaders } from "./fabric.js";
import { modelFixture, type SelectedField } from "./models.js";
import { modeledNames } from "./modeling.js";
import type { IngestionPath } from "./paths.js";
import { updatePlan, type RunPlan } from "./plan.js";
import { assertDashboardComposition, assertDashboardLayout, dashboardSelection } from "./dashboard.js";
import { expectedModeledHubKeys } from "./scoping.js";

export function assertFirmwareAnswer(answer: string): void {
  const versions = answer.match(/\b\d+\.\d+\.\d+(?:\.\d+)*(?:[-+][A-Za-z0-9.-]+)?\b/g) ?? [];
  assert.ok(versions.length === 1 && versions[0] === "1.0.0",
    "[native-agent] The newest response must contain the exact fixture firmware version 1.0.0, without another version.");
}

export function assertDataAgentExamples(value: unknown, table: string, fields: readonly SelectedField[]) {
  assert.ok(Array.isArray(value), "[native-agent] Missing fewshot examples.");
  const actual = value.map(value => {
    const example = object(value, "native agent example");
    return { question: text(example.question, "example question"), query: text(example.query, "example query") };
  });
  const identifier = (value: string) => `['${value.replace(/\\/g, "\\\\").replace(/'/g, "\\'")}']`;
  const source = identifier(table);
  const telemetry = fields.filter(field => field.role === "telemetry");
  const numeric = telemetry.find(field => field.type === "real");
  const latest = numeric ?? fields[0], property = fields.find(field => field.role === "twin");
  const latestQuery = (field: SelectedField) =>
    `${source} | where ${field.type === "string" ? "isnotempty" : "isnotnull"}(${identifier(field.column)}) | summarize arg_max(event_time, *) by deviceId | project deviceId, event_time, ${identifier(field.column)}`;
  const expected = [
    { question: "What devices are in the system?", query: `${source} | distinct deviceId` },
    { question: "Show me the latest modeled records for a device (replace YOUR_DEVICE_ID)",
      query: `${source} | where deviceId == "YOUR_DEVICE_ID" | top 10 by event_time desc` },
  ];
  if (numeric) expected.push({
    question: `What is the average ${numeric.column} over the last hour?`,
    query: `${source} | where event_time >= ago(1h) and isnotnull(${identifier(numeric.column)}) | summarize avg(${identifier(numeric.column)}) by bin(event_time, 5m) | render timechart`,
  });
  if (latest) expected.push({
    question: `Show the latest ${latest.type === "string" ? "nonempty" : "non-null"} ${latest.column} for each device`,
    query: latestQuery(latest),
  });
  if (property) expected.push({
    question: `Show the latest available ${property.type === "string" ? "nonempty" : "non-null"} reported ${property.column} for each device`,
    query: latestQuery(property),
  });
  expected.push({
    question: "How many readings per device in the last 24 hours?",
    query: `${source} | where event_time >= ago(24h) | summarize Count=count() by deviceId | order by Count desc`,
  });
  const originalCount = expected.length;
  const component = telemetry.find(field => field.component);
  if (component) expected.push({
    question: `Show the latest ${component.type === "string" ? "nonempty" : "non-null"} ${component.column} telemetry from component ${component.component} for each device`,
    query: latestQuery(component),
  });
  const first = telemetry.find(field => !field.component) ?? telemetry[0];
  const second = telemetry.find(field => field !== first && field.component) ?? telemetry.find(field => field !== first);
  if (first && second) expected.push({
    question: `Show the latest available ${first.column} and ${second.column} for each device, with each measurement's own timestamp`,
    query: [
      latestQuery(first),
      `| project deviceId, Value1 = ${identifier(first.column)}, event_time1 = event_time`,
      "| join kind=fullouter (",
      `  ${latestQuery(second)}`,
      `  | project deviceId, Value2 = ${identifier(second.column)}, event_time2 = event_time`,
      ") on deviceId",
      "| project deviceId = coalesce(deviceId, deviceId1), Value1, event_time1, Value2, event_time2",
    ].join("\n"),
  });
  assert.deepEqual(actual, expected, "[native-agent] Original examples or appended documented latest-value examples differ.");
  return actual.slice(originalCount);
}

export async function itemDefinition(cloud: Cloud, plan: RunPlan, id: string) {
  const response = await cloud.send(`https://api.fabric.microsoft.com/v1/workspaces/${plan.target.workspaceId}/items/${id}/getDefinition`,
    "https://api.fabric.microsoft.com", "POST");
  let result = response.body;
  if (response.status === 202) {
    const operation = operationFromHeaders(response.headers);
    const deadline = Date.now() + 180_000;
    let completed = false;
    while (Date.now() < deadline) {
      const state = object(await cloud.fabric(operation), "item definition operation");
      if (state.status === "Succeeded") { completed = true; break; }
      if (state.status === "Failed" || state.status === "Canceled") throw new Error("[service] Native definition read failed.");
      await pause(2000);
    }
    if (!completed) throw new Error("[service] Native definition read timed out.");
    result = await cloud.fabric(`${operation}/result`);
  }
  const parts = object(object(result, "item definition").definition, "definition parts").parts;
  if (!Array.isArray(parts)) throw new Error("[service-shape] Native item definition has no parts.");
  return parts.map(value => {
    const part = object(value, "definition part");
    if (part.payloadType !== "InlineBase64") throw new Error("[service-shape] Expected an inline native definition part.");
    return { path: text(part.path, "part path"), value: JSON.parse(Buffer.from(text(part.payload, "part payload"), "base64").toString("utf8")) as unknown };
  });
}

async function createdItem(cloud: Cloud, plan: RunPlan, name: string, type: string) {
  const matches: Record<string, unknown>[] = [];
  let continuation: string | undefined;
  do {
    const page = object(await cloud.fabric(`/workspaces/${plan.target.workspaceId}/items${continuation ? `?continuationToken=${encodeURIComponent(continuation)}` : ""}`), "workspace items");
    if (!Array.isArray(page.value)) throw new Error("[service-shape] Missing workspace item inventory.");
    matches.push(...page.value.map(value => object(value, "Fabric item")).filter(item => item.displayName === name && item.type === type));
    continuation = typeof page.continuationToken === "string" ? page.continuationToken : undefined;
  } while (continuation);
  if (matches.length !== 1) throw new Error(`[native-artifact] Expected one UI-created ${type} named ${name}; found ${matches.length}.`);
  const id = requireGuid(text(matches[0]?.id, "created item ID"), "created item ID");
  await moveOwnedItem(cloud, plan, id, name);
  return id;
}

export async function verifyDashboard(cloud: Cloud, plan: RunPlan, path: IngestionPath, name: string): Promise<void> {
  const id = await createdItem(cloud, plan, name, "KQLDashboard");
  const parts = await itemDefinition(cloud, plan, id);
  const document = object(parts.find(part => part.path === "RealTimeDashboard.json")?.value, "native dashboard");
  const sources = document.dataSources, queries = document.queries, tiles = document.tiles;
  if (!Array.isArray(sources) || !Array.isArray(queries) || !Array.isArray(tiles)) throw new Error("[native-dashboard] Missing sources, queries or tiles.");
  const selected = dashboardSelection(path.id);
  assertDashboardComposition(tiles.map(value => {
    const tile = object(value, "tile");
    return { title: tile.title, visualType: tile.visualType };
  }), path.id);
  assertDashboardLayout(tiles.map(value => object(value, "tile")), selected);
  assert.equal(sources.length, 1);
  for (const value of sources) {
    const source = object(value, "dashboard source");
    assert.equal(source.databaseArtifactId, plan.observed.databaseId);
    assert.equal(source.clusterUri, plan.observed.queryServiceUri);
  }
  const nativeQueries = queries.map(value => text(object(value, "dashboard query").text, "query text"));
  assert.ok(nativeQueries.some(value => value.includes(path.telemetryTable)));
  assert.ok(nativeQueries.some(value => value.includes(path.propertiesTable)));
  for (const field of ["temperature", "status", "firmwareVersion"]) assert.ok(nativeQueries.some(value => value.includes(field)));
  const allRows = await query(cloud, plan, `${path.telemetryTable} | where data contains "${plan.runId}" | project deviceId=tostring(user_headers.cloudEvents_subject) | distinct deviceId`);
  assert.ok(allRows.length >= path.hubKeys.length, "[native-dashboard] The dashboard's actual telemetry source is missing fixture devices.");
  const hub = plan.azure.hubs.find(value => value.key === path.hubKeys[0]);
  assert.ok(hub);
  const deviceId = `${hub.name}-device`;
  const tileData: Record<string, unknown> = {};
  for (const field of selected) {
    const tile = object(tiles.find(value => object(value, "tile").title === field.name), "selected dashboard tile");
    const queryId = object(tile.queryRef, "tile query reference").queryId;
    const csl = text(queries.map(value => object(value, "dashboard query")).find(value => value.id === queryId)?.text, "actual tile query");
    if (field.component && field.visual === "line") assert.ok(csl.includes(`== "${field.component}"`), "[native-dashboard] Missing component source binding.");
    if (field.component && field.visual === "card") assert.ok(csl.includes(`rp["${field.component}"]`), "[native-dashboard] Missing component property binding.");
    const rows = await query(cloud, plan, `let _startTime=ago(1h); let _endTime=now(); let _deviceId=${JSON.stringify(deviceId)};\n${csl}`);
    assert.ok(rows.length > 0, `[native-dashboard] The actual ${field.name} tile query returned no data.`);
    assert.ok(rows.some(row => row.Value === field.expected || typeof field.expected === "number" && Number(row.Value) === field.expected),
      `[native-dashboard] Actual tile query ${field.name} did not return the sent fixture value.`);
    tileData[field.name] = rows;
  }
  await writeFile(join(artifacts, "runs", plan.runId, `dashboard-evidence-${path.id}.json`), JSON.stringify({ id, name, path, definition: parts, sourceDevices: allRows, deviceId, tileData }, null, 2));
  plan.observed[`${path.id}.dashboardId`] = id;
  plan.verification[`${path.id}.dashboard`] = "native-definition-components-and-actual-tile-query-data-verified";
  await updatePlan(plan);
}

export async function verifyDataAgent(cloud: Cloud, plan: RunPlan, path: IngestionPath, name: string): Promise<void> {
  const names = modeledNames(path), fixture = modelFixture(path.id);
  const id = await createdItem(cloud, plan, name, "DataAgent");
  const parts = await itemDefinition(cloud, plan, id);
  let targetedExamples: ReturnType<typeof assertDataAgentExamples> = [];
  for (const stage of ["draft", "published"]) {
    const source = object(parts.find(part => part.path.includes(`/${stage}/`) && part.path.endsWith("/datasource.json"))?.value, "native agent data source");
    assert.equal(source.workspaceId, plan.target.workspaceId);
    assert.equal(source.artifactId, plan.observed.databaseId);
    assert.ok(Array.isArray(source.elements));
    const group = object(source.elements[0], "agent table grouping");
    assert.ok(Array.isArray(group.children));
    assert.equal(group.children.length, 1);
    const table = object(group.children[0], "agent selected table");
    assert.equal(table.display_name, names.modeled);
    assert.equal(table.is_selected, true);
    assert.ok(Array.isArray(table.children));
    assert.deepEqual(table.children.map(value => object(value, "agent column").display_name),
      ["deviceId", "event_time", ...fixture.fields.map(field => field.column)]);
    assert.ok(text(source.dataSourceInstructions, "native semantic source instructions").includes(fixture.id));
    const stageConfig = object(parts.find(part => part.path === `Files/Config/${stage}/stage_config.json`)?.value, "stage config");
    assert.ok(text(stageConfig.aiInstructions, "agent instructions").length > 100);
    const examples = object(parts.find(part => part.path.includes(`/${stage}/`) && part.path.endsWith("/fewshots.json"))?.value, "agent examples");
    targetedExamples = assertDataAgentExamples(examples.fewShots, names.modeled, fixture.fields);
  }
  plan.observed[`${path.id}.dataAgentId`] = id;
  const rows = await query(cloud, plan, `${names.modeled} | where isnotempty(firmwareVersion) | summarize arg_max(event_time, firmwareVersion) by deviceId | take 100`);
  const expectedIds = new Set(plan.azure.hubs.filter(hub => expectedModeledHubKeys(path.id, "telemetry").includes(hub.key))
    .map(hub => `${hub.name}-device`));
  const device = rows.find(row => row.firmwareVersion === "1.0.0" && typeof row.deviceId === "string" && expectedIds.has(row.deviceId));
  assert.ok(device, "[native-agent] No known fixture firmware is present in the actual modeled grounding table.");
  const exampleData = [];
  for (const example of targetedExamples) {
    const rows = await query(cloud, plan, example.query);
    assert.ok(rows.length > 0, "[native-agent] The actual latest-value example returned no modeled data.");
    assert.ok(rows.every(row => typeof row.deviceId === "string" && expectedIds.has(row.deviceId)),
      "[native-agent] Latest-value examples returned a device outside the modeled scope.");
    const current = rows.find(row => row.deviceId === device.deviceId);
    assert.ok(current, "[native-agent] The grounded device is missing from the generated example.");
    if (example.query.includes("join kind=fullouter")) {
      assert.ok(current.Value1 !== null && current.Value1 !== undefined && current.Value2 !== null && current.Value2 !== undefined,
        "[native-agent] The fixture's two independently sent measurements were not retained.");
      for (const key of ["event_time1", "event_time2"]) {
        assert.ok(typeof current[key] === "string" && Number.isFinite(Date.parse(current[key])),
          `[native-agent] Missing native observation timestamp ${key}.`);
      }
    } else assert.ok(typeof current.event_time === "string" && Number.isFinite(Date.parse(current.event_time)),
      "[native-agent] Missing component's native observation timestamp.");
    exampleData.push({ ...example, rows });
  }
  await writeFile(join(artifacts, "runs", plan.runId, `data-agent-evidence-${path.id}.json`), JSON.stringify({
    id, name, path, definition: parts, documentedLatestExamples: exampleData,
  }, null, 2));
  plan.observed[`${path.id}.agentDeviceId`] = text(device.deviceId, "grounded fixture device");
  plan.verification[`${path.id}.dataAgent`] = "ui-created-native-draft-published-grounding-and-semantics-verified";
  await updatePlan(plan);
}
