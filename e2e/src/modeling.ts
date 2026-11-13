import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { join } from "node:path";
import { setTimeout as pause } from "node:timers/promises";
import { isDeepStrictEqual } from "node:util";
import { Cloud, object, text } from "./cloud.js";
import { artifacts } from "./config.js";
import { query, verifyArrivingData, type DeviceFleet, type SimulatedDevice } from "./data.js";
import { modelFixture, type SelectedField } from "./models.js";
import { ingestionPath, type IngestionPath } from "./paths.js";
import { updatePlan, type RunPlan } from "./plan.js";
import { assertScopedOutput, expectedModeledHubKeys, matchesScope, scopePredicate, scopeSamplesReady } from "./scoping.js";
import { referenceRowsReady } from "./model-assignment.js";

export function modeledNames(path: IngestionPath) {
  return { modeled: `Modeled_${path.telemetryTable}`, normalized: `Normalized_${path.propertiesTable}`, lkv: `LKV_${path.propertiesTable}` };
}

export function expectedModelValues(
  device: Pick<SimulatedDevice, "telemetry" | "reported" | "components">,
  fields: readonly SelectedField[],
): Record<string, unknown> {
  return Object.fromEntries(fields.map(field => {
    const value = field.role === "telemetry"
      ? field.component ? device.components[field.component]?.[field.property] : device.telemetry[field.property]
      : field.component ? object(device.reported[field.component], "component state")[field.property] : device.reported[field.property];
    if (value === undefined) throw new Error(`[fixture] Missing simulated value for selected field ${field.column}.`);
    return [field.column, value];
  }));
}

export async function verifyIncomingReferences(cloud: Cloud, plan: RunPlan, path: IngestionPath, fleet: DeviceFleet): Promise<void> {
  assert.equal(plan.verification.modelAssignment, "confirmed", "[prerequisite] Confirm native model assignment before checking enrichment.");
  const devices = fleet.devices.filter(device => ["linked1", "linked2", "linked3"].includes(device.hubKey));
  assert.equal(devices.length, 3);
  const start = new Date().toISOString(), deadline = Date.now() + 5 * 60_000;
  const file = join(artifacts, "runs", plan.runId, `model-reference-evidence-${path.id}-${fleet.attempt}.json`);
  await fleet.sendRound();
  while (Date.now() < deadline) {
    fleet.assertHealthy();
    const observed: Record<string, Record<string, unknown>[]> = {};
    for (const [stream, table, identity] of [
      ["telemetry", path.telemetryTable, "cloudEvents_subject"],
      ["properties", path.propertiesTable, "cloudEvents_deviceid"],
    ] as const) {
      observed[stream] = await query(cloud, plan, `${table} | where data contains "${fleet.attempt}"
| extend observation=parse_json(data), deviceId=tostring(user_headers.${identity})
| extend event_time=todatetime(observation.event_time), modelReference=tostring(observation.payload.memberTypeIRI)
| where event_time > datetime(${start}) | summarize arg_max(event_time, modelReference) by deviceId`);
    }
    await writeFile(file, JSON.stringify({ start, devices: devices.map(device => ({ id: device.id, hubKey: device.hubKey })), observed }, null, 2));
    if (referenceRowsReady(observed.telemetry ?? [], devices) && referenceRowsReady(observed.properties ?? [], devices)) {
      plan.verification[`${path.id}.incomingModelReferences`] = "native-DO-model-references-confirmed-for-three-linked-devices";
      await updatePlan(plan);
      return;
    }
    await pause(5000);
  }
  throw new Error(`[native-model-reference] Fresh telemetry/properties did not carry the three expected model references. Evidence: ${file}`);
}

export async function verifyModeled(cloud: Cloud, plan: RunPlan, path: IngestionPath, fleet: DeviceFleet): Promise<void> {
  const fixture = modelFixture(path.id), names = modeledNames(path);
  const devices = fleet.devices.filter(device => expectedModeledHubKeys(path.id, "telemetry").includes(device.hubKey));
  const propertyDevices = fleet.devices.filter(device => expectedModeledHubKeys(path.id, "twin").includes(device.hubKey));
  assert.equal(devices.length, expectedModeledHubKeys(path.id, "telemetry").length);
  assert.equal(propertyDevices.length, expectedModeledHubKeys(path.id, "twin").length);
  const supplied = Object.fromEntries(propertyDevices.map(device => [device.id, expectedModelValues(device, fixture.fields)]));
  // Independently apply Kusto's dynamic JSON semantics to sent values, never to output-derived expectations.
  const expectedRows = await query(cloud, plan, `print expected=parse_json(${JSON.stringify(JSON.stringify(supplied))})`);
  const expected = object(expectedRows[0]?.expected, "native expected dynamic values");
  function matches(device: SimulatedDevice, field: SelectedField, actual: unknown): boolean {
    return isDeepStrictEqual(actual, object(expected[device.id], "expected device values")[field.column]);
  }
  const evidenceFile = join(artifacts, "runs", plan.runId, `modeled-evidence-${path.id}-${fleet.attempt}.json`);
  const schema = await query(cloud, plan, `.show table ${names.modeled} schema as json | project Schema`);
  const metadata = object(JSON.parse(text(schema[0]?.Schema, "modeled schema")), "schema");
  if (!Array.isArray(metadata.OrderedColumns)) throw new Error("[native-schema] Missing ordered modeled columns.");
  assert.deepEqual(metadata.OrderedColumns.map(value => {
    const column = object(value, "column");
    return [column.Name, column.CslType];
  }), [["deviceId", "string"], ["event_time", "datetime"], ...fixture.fields.map(field => [field.column, field.type])]);
  const proofs: Record<string, unknown> = {};
  for (const [name, role] of [[names.modeled, "table"], [names.normalized, "table"], [names.lkv, "materialized-view"]]) {
    const rows = await query(cloud, plan, role === "table" ? `.show table ${name} schema as json` : `.show materialized-view ${name}`);
    const doc = text(rows[0]?.DocString, `${name} docstring`);
    const modelInfo = /^@iot\.meta modelSemanticInfo=(.*)$/m.exec(doc)?.[1];
    if (!modelInfo) throw new Error("[native-semantics] Missing modelSemanticInfo.");
    const declared = object(JSON.parse(modelInfo), "model semantics");
    assert.equal(declared[fixture.format === "dtdl" ? "@id" : "id"], fixture.id);
    const expectedSource = name === names.normalized ? path.propertiesTable : name === names.lkv ? names.normalized : path.telemetryTable;
    assert.ok(doc.includes(`/${plan.fabric.databaseName}/${expectedSource}`), "[native-bindings] Semantic source binding is wrong.");
    if (name === names.modeled) assert.ok(doc.includes(`/${plan.fabric.databaseName}/${names.lkv}`), "[native-bindings] Missing LKV source binding.");
    proofs[name ?? ""] = rows;
  }
  for (const [target, source] of [[names.normalized, path.propertiesTable], [names.modeled, path.telemetryTable]] as const) {
    const result = await query(cloud, plan, `.show table ${target} policy update`);
    const policies: unknown = JSON.parse(text(result[0]?.Policy, "native policy"));
    if (!Array.isArray(policies) || policies.length !== 1) throw new Error("[native-policy] Expected exactly one policy per output.");
    const policy = object(policies[0], "policy");
    assert.equal(policy.IsEnabled, true);
    assert.equal(policy.Source, source);
    assert.equal(policy.IsTransactional, false);
    assert.equal(policy.PropagateIngestionProperties, false);
    assert.ok(text(policy.Query, "policy query").includes(source), "[native-policy] Policy query does not bind the intended raw source.");
    assert.ok(text(policy.Query, "policy query").includes(scopePredicate(path.id, target === names.modeled ? "telemetry" : "twin")),
      "[native-policy] The native update policy did not preserve the chosen model reference/condition combination.");
    proofs[`${target}.policy`] = policy;
  }
  const view = await query(cloud, plan, `.show materialized-view ${names.lkv}`);
  assert.equal(view[0]?.SourceTable, names.normalized);
  assert.equal(view[0]?.IsEnabled, true);
  assert.equal(view[0]?.IsHealthy, true);
  assert.match(text(view[0]?.Query, "LKV query"), /ingestion_time\(\) > ago\(30d\)/);
  assert.match(text(view[0]?.Query, "LKV query"), /arg_max\(event_time,\s*propertyValue\)/);
  const normalized = await query(cloud, plan, `.show table ${names.normalized} schema as json | project Schema`);
  const normalizedSchema = object(JSON.parse(text(normalized[0]?.Schema, "normalized schema")), "normalized schema");
  assert.ok(Array.isArray(normalizedSchema.OrderedColumns));
  assert.deepEqual(normalizedSchema.OrderedColumns.map(value => {
    const column = object(value, "normalized column");
    return [column.Name, column.CslType];
  }), [["deviceId", "string"], ["propertyName", "string"], ["propertyValue", "dynamic"], ["event_time", "datetime"]]);
  // Read native ordered schema with column documentation rather than accepting UI-only semantic labels.
  for (const field of fixture.fields) {
    const column = metadata.OrderedColumns.map(value => object(value, "column")).find(value => value.Name === field.column);
    const doc = text(column?.DocString, `fieldSemanticInfo for ${field.column}`);
    const factsText = /^@iot\.meta fieldSemanticInfo=(.*)$/m.exec(doc)?.[1];
    if (!factsText) throw new Error("[native-semantics] Missing fieldSemanticInfo on a selected modeled column.");
    const facts = object(JSON.parse(factsText), "field semantics");
    if (facts.name !== field.column && facts.name !== field.property) {
      throw new Error("[native-semantics] Selected column semantic name does not match its fixture field.");
    }
  }
  // Property-first fresh updates make the LKV state available before the matching telemetry round.
  const stateStart = new Date().toISOString();
  await fleet.reportProperties();
  const stateDeadline = Date.now() + 5 * 60_000;
  let propertyStateReady = false;
  let state: Record<string, unknown>[] = [];
  while (Date.now() < stateDeadline) {
    fleet.assertHealthy();
    state = await query(cloud, plan, `${names.lkv} | where event_time > datetime(${stateStart}) | take 1000`);
    propertyStateReady = propertyDevices.every(device =>
      fixture.fields.filter(field => field.role === "twin").every(field => state.some(row =>
        row.deviceId === device.id && row.propertyName === (field.component ? `${field.component}_${field.property}` : field.property) &&
        matches(device, field, row.propertyValue))));
    if (propertyStateReady) break;
    await pause(5000);
  }
  if (!propertyStateReady) {
    await writeFile(evidenceFile, JSON.stringify({ path, names, expected, stateStart, state, proofs }, null, 2));
    throw new Error("[native-state] Fresh expected property LKVs were not observed before telemetry verification. Expected and observed state retained in modeled evidence.");
  }
  console.log(`[modeling] ${path.id}: fresh expected LKVs observed; send verification telemetry`);
  const start = new Date().toISOString();
  await fleet.sendRound();
  const deadline = Date.now() + 5 * 60_000;
  let data: Record<string, unknown>[] = [];
  while (Date.now() < deadline) {
    fleet.assertHealthy();
    data = await query(cloud, plan, `${names.modeled} | where event_time > datetime(${start}) | take 1000`);
    let complete = true;
    for (const device of devices) {
      for (const component of fixture.format === "dtdl" ? ["", "thermostat", "powerModule"] : [""]) {
        const relevant = fixture.fields.filter(field => field.role === "twin" || (field.component ?? "") === component);
        if (!data.some(row => row.deviceId === device.id && relevant.every(field => matches(device, field, row[field.column])))) complete = false;
      }
    }
    if (complete) break;
    await pause(5000);
  }
  await writeFile(evidenceFile, JSON.stringify({ path, fixture: { id: fixture.id, hash: fixture.hash }, names, proofs, expected, stateStart, state, data }, null, 2));
  for (const device of devices) {
    const rows = data.filter(row => row.deviceId === device.id);
    assert.ok(rows.length > 0, `[native-data] No fresh modeled records for ${device.id}.`);
    for (const component of fixture.format === "dtdl" ? ["", "thermostat", "powerModule"] : [""]) {
      const relevant = fixture.fields.filter(field => field.role === "twin" || (field.component ?? "") === component);
      assert.ok(rows.some(row => relevant.every(field =>
        matches(device, field, row[field.column]))),
      `[native-data] No single fresh ${device.id}/${component || "root"} row has the exact selected telemetry and complete property enrichment.`);
    }
  }
  const rawTimes = await query(cloud, plan, `${path.telemetryTable} | where data contains "${fleet.attempt}" | extend observation=parse_json(data) | extend observedTime=todatetime(observation.event_time) | where observedTime > datetime(${start}) | project observedTime | take 1000`);
  const actualClocks = new Set(rawTimes.map(row => row.observedTime));
  for (const row of data) assert.ok(actualClocks.has(row.event_time), "[native-clock] Modeled event_time must retain the actual DO observation time, not ingestion/enqueue time.");
  await verifyScopedData(cloud, plan, path, fleet);
  if (path.id === "linked") {
    await verifyArrivingData(cloud, plan, fleet, ingestionPath(plan, "combined"));
    const negativeIds = JSON.stringify(fleet.devices.filter(device => !path.hubKeys.includes(device.hubKey)).map(device => device.id));
    const negative = await query(cloud, plan, `${names.modeled} | where event_time > datetime(${start}) | where deviceId in(dynamic(${negativeIds})) | count`);
    assert.equal(negative[0]?.Count, 0, "[native-exclusion] Regular devices appeared in linked-only modeled outputs.");
  }
  plan.verification[`${path.id}.modeling`] = "native-model-schema-policies-semantics-and-fresh-data-verified";
  await updatePlan(plan);
  console.log(`[modeling] ${path.id}: native schema, semantic metadata, policies, observation clock and enriched values verified`);
}

export async function verifyScopedData(cloud: Cloud, plan: RunPlan, path: IngestionPath, fleet: DeviceFleet): Promise<void> {
  const names = modeledNames(path), start = new Date().toISOString();
  await fleet.sendRound();
  await fleet.sendRound();
  const evidenceFile = join(artifacts, "runs", plan.runId, `scope-evidence-${path.id}-${fleet.attempt}.json`);
  const evidence: Record<string, unknown> = { start, attempt: fleet.attempt, windowClock: "native-observation" };
  const sourceQuery = (stream: "telemetry" | "twin", end?: string) =>
    `${stream === "telemetry" ? path.telemetryTable : path.propertiesTable} | where data contains "${fleet.attempt}" | extend observation=parse_json(data)
| extend event_time=todatetime(observation.event_time), deviceId=tostring(user_headers.${stream === "telemetry" ? "cloudEvents_subject" : "cloudEvents_deviceid"})
| where ${end ? `event_time between (datetime(${start}) .. datetime(${end}))` : `event_time > datetime(${start})`}
| project deviceId, event_time, reference=tostring(observation.payload.memberTypeIRI),
sequence=tolong(observation.payload.value.${stream === "telemetry" ? "e2eSequence" : "properties.reported.e2eSequence"}) | take 5000`;
  const cohortDeadline = Date.now() + 180_000;
  while (true) {
    fleet.assertHealthy();
    const telemetry = await query(cloud, plan, sourceQuery("telemetry"));
    const twin = await query(cloud, plan, sourceQuery("twin"));
    const ready = { telemetry: scopeSamplesReady(path.id, "telemetry", telemetry), twin: scopeSamplesReady(path.id, "twin", twin) };
    evidence.sourceCohorts = { ready, telemetry, twin };
    if (ready.telemetry && ready.twin) break;
    if (Date.now() >= cohortDeadline) {
      await writeFile(evidenceFile, JSON.stringify(evidence, null, 2));
      throw new Error(`[native-scope] Current-attempt raw positive/negative cohorts did not arrive. Evidence: ${evidenceFile}`);
    }
    console.log(`[modeling] ${path.id}: waiting for native telemetry/property source cohorts before freezing the observation window.`);
    await pause(5000);
  }
  const watermark = await query(cloud, plan, `union ${path.telemetryTable}, ${path.propertiesTable}
| where data contains "${fleet.attempt}" | extend observation=parse_json(data)
| extend event_time=todatetime(observation.event_time) | where event_time > datetime(${start})
| summarize nativeEnd=max(event_time) | project nativeEnd=tostring(nativeEnd)`);
  const end = text(watermark[0]?.nativeEnd, "native scope observation watermark");
  if (!Number.isFinite(Date.parse(end)) || Date.parse(end) < Date.parse(start)) {
    throw new Error("[service-shape] Native scope observation watermark is not a valid current observation time.");
  }
  evidence.end = end;
  await writeFile(evidenceFile, JSON.stringify(evidence, null, 2));
  for (const stream of ["telemetry", "twin"] as const) {
    const target = stream === "telemetry" ? names.modeled : names.normalized;
    const deadline = Date.now() + 180_000;
    let raw: Record<string, unknown>[] = [], output: Record<string, unknown>[] = [];
    while (Date.now() < deadline) {
      fleet.assertHealthy();
      raw = await query(cloud, plan, sourceQuery(stream, end));
      output = await query(cloud, plan, `${target} | where event_time between (datetime(${start}) .. datetime(${end}))${stream === "twin" ? ' | where propertyName == "firmwareVersion"' : ""} | summarize rows=count() by deviceId, event_time | take 5000`);
      const expected = raw.filter(row => matchesScope(path.id, stream, row.reference, row.sequence));
      const expectedCounts = new Map<string, number>();
      for (const row of expected) {
        const key = JSON.stringify([row.deviceId, row.event_time]);
        expectedCounts.set(key, (expectedCounts.get(key) ?? 0) + 1);
      }
      if (scopeSamplesReady(path.id, stream, raw) &&
          output.length === expectedCounts.size &&
          output.every(row => row.rows === expectedCounts.get(JSON.stringify([row.deviceId, row.event_time])))) break;
      await pause(5000);
    }
    evidence[stream] = { predicate: scopePredicate(path.id, stream), raw, output };
    await writeFile(evidenceFile, JSON.stringify(evidence, null, 2));
    assertScopedOutput(path, stream, raw, output);
  }
}
