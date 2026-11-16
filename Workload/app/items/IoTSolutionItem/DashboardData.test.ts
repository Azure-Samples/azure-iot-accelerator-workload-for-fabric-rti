// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { dashboardFieldQuery, dashboardRawQuery, readDashboardFields, type DashboardField } from "./DashboardData.ts";
import { generateDashboardDefinition, toBase64 } from "./DashboardDefinitionGenerator.ts";
import { isJsonObject } from "./ModelAuthoring.ts";
import { DashboardDateTime, dashboardClock, dashboardRow, replayDashboardQuery, type FixtureRow } from "../../../tests/dashboardFixtures.ts";

function fields(rows: FixtureRow[], stream: "telemetry" | "twin") {
  const result = replayDashboardQuery(dashboardFieldQuery("Raw", stream), { Raw: rows });
  return stream === "telemetry" ? readDashboardFields(result, stream) : readDashboardFields(result, stream);
}
function queries(telemetryFields: DashboardField[], propertyFields: DashboardField[] = [], telemetryTable = "Telemetry", propertiesTable = "Twins") {
  const definition = generateDashboardDefinition({ title: "Device dashboard", telemetryTable, propertiesTable,
    telemetryFields, propertyFields, queryServiceUri: "https://mock.kusto.fabric.microsoft.com", databaseId: "db" });
  assert.ok(isJsonObject(definition) && Array.isArray(definition.queries) && Array.isArray(definition.tiles));
  const values = definition.queries.map(query => {
    assert.ok(isJsonObject(query) && typeof query.text === "string");
    return query.text;
  });
  return { definition, values };
}
const range = { _deviceId: "device-1", _startTime: dashboardClock - 3600000 * 24, _endTime: dashboardClock };

test("dashboard DO telemetry discovers real whole-valued readings and classifies only numeric values", () => {
  const value = { temperature: 21.5, count: 0, active: false, text: "22", setting: { gain: 1 }, samples: [1, 2],
    nullable: null, "": 9, " \t": 9, $metadata: {}, $version: 1, __t: "c" };
  const result = fields([dashboardRow(value)], "telemetry");
  assert.deepEqual(result.map(field => [field.key, field.isNumeric]).sort(), [
    ["active", false], ["count", true], ["nullable", false], ["samples", false], ["setting", false], ["temperature", true], ["text", false],
  ]);
  assert.ok(result.every(field => field.leaf === field.key && !field.component));
  assert.ok(!result.some(field => ["payload", "schema_id", "entity_id", "contract_version", "event_time"].includes(field.key)));
});

test("dashboard reported discovery includes root and PnP component fields without metadata or ordinary-object flattening", () => {
  const result = fields([dashboardRow({ properties: { reported: { firmware: "1.0", setting: { nested: 4 }, samples: [1],
    motor_a: { __t: "c", target: 22, active: true, $metadata: {}, $version: 2, "": "bad" }, $metadata: {}, $version: 3, "": 1 } } },
  { stream: "twin" })], "twin");
  assert.deepEqual(result.map(field => [field.key, field.component || "", field.leaf]), [
    ["firmware", "", "firmware"], ["motor_a_active", "motor_a", "active"], ["motor_a_target", "motor_a", "target"],
    ["samples", "", "samples"], ["setting", "", "setting"],
  ]);
});

test("dashboard CloudEvents and empty native placeholders retain expected module and telemetry-component identity with DO time", () => {
  for (const profile of ["cloud", "empty-native"] as const) for (const module of ["", "module-1"]) {
    const row = dashboardRow({ temperature: 7 }, { profile, module, component: "motor_a" });
    if (isJsonObject(row.user_headers)) row.user_headers.cloudEvents_time = "0001-01-01T00:00:00Z";
    const discovered = fields([row], "telemetry");
    assert.equal(discovered[0].component, "motor_a");
    const generated = queries(discovered).values;
    const identity = module ? "device-1/module-1" : "device-1";
    assert.deepEqual(replayDashboardQuery(generated[0], { Telemetry: [row] }, range), [{ deviceId: identity }]);
    assert.deepEqual(replayDashboardQuery(generated[1], { Telemetry: [row] }, { ...range, _deviceId: identity }),
      [{ event_time: Date.parse("2026-10-29T11:00:00Z"), Value: 7 }]);
    const twin = dashboardRow({ properties: { reported: { firmware: "2.0" } } }, { stream: "twin", profile, module });
    const cards = queries([], fields([twin], "twin")).values;
    assert.deepEqual(replayDashboardQuery(cards[1], { Twins: [twin] }, { ...range, _deviceId: identity }),
      [{ event_time: Date.parse("2026-10-29T11:00:00Z"), Value: "2.0" }]);
  }
});

test("dashboard component/root underscore collisions and special keys retain exact identities in every generated query", () => {
  const special = `quote'["\\|;value`;
  const rows = [dashboardRow({ a_b_c: 1, [special]: 4 }), dashboardRow({ b_c: 2 }, { component: "a" }),
    dashboardRow({ c: 3 }, { component: "a_b" })];
  const discovered = fields(rows, "telemetry");
  assert.equal(new Set(discovered.map(field => field.key)).size, 4);
  assert.deepEqual(discovered.map(field => [field.component || "", field.leaf]).sort(), [
    ["", "a_b_c"], ["", special], ["a", "b_c"], ["a_b", "c"],
  ].sort());
  const table = "Raw's table", generated = queries(discovered, [], table);
  for (let index = 0; index < discovered.length; index++) {
    const field = discovered[index], expected = !field.component ? field.leaf === "a_b_c" ? 1 : 4 : field.component === "a" ? 2 : 3;
    const result = replayDashboardQuery(generated.values[index + 1], { [table]: rows }, range);
    assert.deepEqual(result.map(row => row.Value), [expected]);
  }
  const component = `motor'_part`, leaf = `target'\\[]`;
  const twin = dashboardRow({ properties: { reported: { [component]: { __t: "c", [leaf]: 8 } } } }, { stream: "twin" });
  const propertyQueries = queries([], fields([twin], "twin"), "Telemetry", table).values;
  assert.deepEqual(replayDashboardQuery(propertyQueries[1], { [table]: [twin] }, range).map(row => row.Value), [8]);
});

test("dashboard charts use direct native numeric casts and recent tables retain whole values", () => {
  const values = [3, "22", true, { nested: 5 }, [7, 8], 0];
  const rows = values.map((value, index) => dashboardRow({ reading: value }, { time: `2026-10-29T1${index}:00:00Z` }));
  const numeric = queries([{ key: "reading", leaf: "reading", included: true, isNumeric: true }]).values[1];
  assert.ok(numeric.includes('todecimal(data.payload.value["reading"])') && !numeric.includes("gettype("));
  assert.deepEqual(replayDashboardQuery(numeric, { Telemetry: [rows[0], rows[5]] }, { ...range, _endTime: dashboardClock + 86400000 })
    .map(row => row.Value), [3, 0]);
  const recent = queries([{ key: "reading", leaf: "reading", included: true, isNumeric: false }]).values[1];
  assert.deepEqual(replayDashboardQuery(recent, { Telemetry: rows }, { ...range, _endTime: dashboardClock + 86400000 })
    .map(row => row.Value), [...values].reverse());
});

test("dashboard all generated paths use the same normalized source and retain device/time/LKV windows", () => {
  const telemetry = [dashboardRow({ temperature: 5, state: "ok" }),
    dashboardRow({ temperature: 99 }, { component: "motor", device: "device-2" }),
    { ...dashboardRow({ temperature: 101 }), ingestedAt: dashboardClock - 31 * 86400000 }];
  const twin = [dashboardRow({ properties: { reported: { firmware: "old", motor: { __t: "c", target: 3 } } } },
    { stream: "twin", time: "2026-10-20T11:00:00Z" }),
  dashboardRow({ properties: { reported: { firmware: "new", motor: { __t: "c", target: 8 } } } }, { stream: "twin" })];
  const generated = queries(fields(telemetry, "telemetry"), fields(twin, "twin"));
  assert.deepEqual(replayDashboardQuery(generated.values[0], { Telemetry: telemetry }, range), [{ deviceId: "device-1" }, { deviceId: "device-2" }]);
  const results = generated.values.slice(1).map(query => replayDashboardQuery(query, { Telemetry: telemetry, Twins: twin }, range));
  assert.ok(results.some(rows => rows[0]?.Value === 5));
  assert.ok(results.some(rows => rows[0]?.Value === "ok"));
  assert.ok(results.some(rows => rows[0]?.Value === "new"));
  assert.ok(results.some(rows => rows[0]?.Value === 8));
  assert.ok(!results.some(rows => rows.some(row => row.Value === 99 || row.Value === 101)));
  const text = JSON.stringify(generated.definition);
  assert.equal(new TextDecoder().decode(Uint8Array.from(atob(toBase64(text)), char => char.charCodeAt(0))), text);
  assert.equal(generated.definition.schema_version, 88);
});

test("dashboard DO rows require no model reference or Hub-link evidence and reject legacy raw and mixed rows", () => {
  const value = { temperature: 4, payload: { legitimate: 2 }, state: "ok" };
  const observation = dashboardRow(value), expected = fields([observation], "telemetry");
  assert.equal(expected.find(field => field.key === "payload")!.isNumeric, false);
  for (const reference of [undefined, null, "urn:fixture:linked-model"]) {
    const linked = dashboardRow(value);
    const envelope = JSON.parse(String(linked.data));
    if (reference !== undefined) envelope.payload.memberTypeIRI = reference;
    linked.data = JSON.stringify(envelope);
    assert.deepEqual(fields([linked], "telemetry"), expected);
    assert.deepEqual(replayDashboardQuery(queries(expected).values[0], { Telemetry: [linked] }, range), [{ deviceId: "device-1" }]);
  }
  for (const stream of ["telemetry", "twin"] as const) {
    const legacy = dashboardRow(stream === "telemetry" ? value : { properties: { reported: { firmware: "unsupported" } } }, { do: false, stream });
    const supported = stream === "telemetry" ? observation : dashboardRow({ properties: { reported: { firmware: "1" } } }, { stream });
    assert.throws(() => fields([legacy], stream), /Unsupported DO observation/);
    assert.throws(() => fields([supported, legacy], stream), /Unsupported DO observation/);
  }
});

test("dashboard unknown malformed or batched observations never fall back to envelope readings", () => {
  for (const variant of ["version", "schema", "no-id", "no-contract", "kind", "time", "payload", "array", "json"]) {
    const row = dashboardRow({ temperature: 1 });
    const data: unknown = JSON.parse(String(row.data));
    assert.ok(isJsonObject(data));
    if (variant === "version") data.schema_version = "2.0";
    if (variant === "schema") data.schema_id = "com.do.unknown.v99";
    if (variant === "no-id") delete data.schema_id;
    if (variant === "no-contract") delete data.contract_version;
    if (variant === "kind") data.event_kind = "lifecycle";
    if (variant === "time") data.event_time = "bad";
    if (variant === "payload") data.payload = { value: [] };
    row.data = variant === "array" ? JSON.stringify([data]) : variant === "json" ? "{broken" : JSON.stringify(data);
    assert.throws(() => fields([row], "telemetry"), /Telemetry:/, variant);
  }
});

test("dashboard malformed metadata and duplicate/blank field identities are explicit errors rather than blank rows", () => {
  const row = { Key: "temperature", Leaf: "temperature", Component: "", IsNumericType: true, Error: "" };
  for (const invalid of [{ ...row, Leaf: null }, { ...row, Key: "" }, { ...row, Key: "wrong-name" }, { ...row, IsNumericType: "true" }]) {
    assert.throws(() => readDashboardFields([invalid], "telemetry"), /dashboard/);
  }
  assert.throws(() => readDashboardFields([row, row], "telemetry"), /duplicate/);
  assert.throws(() => readDashboardFields([{ Error: "Unsupported schema" }], "twin"), /Properties: Unsupported schema/);
});

test("dashboard typed dynamic ISO timestamps retain seven-digit precision in inspection and every generated time path", () => {
  const times = [
    "2026-10-29T18:20:48.2940000Z",
    "2026-10-30T18:28:37.6520000Z", "2026-10-30T18:28:37.6600000Z",
    "2026-10-29T18:20:48.2941234Z", "2026-10-29T11:20:48.2941234-07:00",
    "2026-10-29T20:50:48.2941234+02:30", "2026-10-29T18:20:48Z", "2026-10-29T18:20:48.294Z",
  ];
  const bounds = { _startTime: Date.parse("2026-10-28T00:00:00Z"), _endTime: Date.parse("2026-10-31T00:00:00Z") };
  for (const profile of ["cloud", "empty-native"] as const) for (const time of times) for (const dynamic of [false, true]) {
    const telemetry = dashboardRow({ temperature: 7, state: "ok" }, { profile, time, component: "motor", module: "module-1" });
    const twin = dashboardRow({ properties: { reported: { motor: { __t: "c", target: 8 } } } }, { stream: "twin", profile, time, module: "module-1" });
    for (const row of [telemetry, twin]) {
      if (dynamic) row.data = JSON.parse(String(row.data));
      if (isJsonObject(row.headers)) row.headers.IoTEnqueueTime = "0001-01-01T00:00:00";
      if (isJsonObject(row.user_headers)) row.user_headers.cloudEvents_time = "0001-01-01T00:00:00";
    }
    const tables = { Telemetry: [telemetry], Twins: [twin] };
    const tFields = readDashboardFields(replayDashboardQuery(dashboardFieldQuery("Telemetry", "telemetry"), tables, {}, dashboardClock, true), "telemetry");
    const pFields = readDashboardFields(replayDashboardQuery(dashboardFieldQuery("Twins", "twin"), tables, {}, dashboardClock, true), "twin");
    assert.deepEqual(readDashboardFields(replayDashboardQuery(dashboardFieldQuery("Telemetry", "telemetry"), tables), "telemetry"), tFields);
    assert.deepEqual(readDashboardFields(replayDashboardQuery(dashboardFieldQuery("Twins", "twin"), tables), "twin"), pFields);
    assert.deepEqual(tFields.map(field => field.key), ["motor_state", "motor_temperature"]);
    assert.deepEqual(pFields.map(field => field.key), ["motor_target"]);
    const generated = queries(tFields, pFields).values;
    assert.deepEqual(replayDashboardQuery(generated[0], tables, bounds, dashboardClock, true), [{ deviceId: "device-1/module-1" }]);
    for (const query of generated.slice(1)) {
      const result = replayDashboardQuery(query, tables, { ...bounds, _deviceId: "device-1/module-1" }, dashboardClock, true);
      assert.equal(result.length, 1);
      const actual = result[0].event_time;
      assert.ok(actual instanceof DashboardDateTime);
      assert.equal(actual.iso, DashboardDateTime.parse(time)!.utcKey, "The fixture or generated query clipped fractional digits while normalizing to UTC.");
      assert.equal(actual.utcKey, DashboardDateTime.parse(time)!.utcKey);
      assert.ok(query.includes("event_time = todatetime(data.event_time)") && !query.includes("IoTEnqueueTime") &&
        !query.includes("coalesce(todatetime"), "Generated dashboard time must come only from the actual DO observation.");
    }
  }
  assert.equal(DashboardDateTime.parse(times[3])!.utcKey, "2026-10-29T18:20:48.2941234Z");
  assert.equal(DashboardDateTime.parse(times[4])!.utcKey, DashboardDateTime.parse(times[3])!.utcKey);
  assert.equal(DashboardDateTime.parse(times[5])!.utcKey, DashboardDateTime.parse(times[3])!.utcKey);
});

test("dashboard unconvertible missing null or wrong-type event times keep explicit diagnostics without placeholder-time fallback", () => {
  for (const stream of ["telemetry", "twin"] as const) for (const time of [undefined, null, "", "not-an-instant", "2026-02-30T00:00:00Z", 123, false, {}, []]) {
    const row = dashboardRow(stream === "telemetry" ? { temperature: 2 } : { properties: { reported: { target: 8 } } }, { stream });
    const envelope = JSON.parse(String(row.data));
    if (time === undefined) delete envelope.event_time; else envelope.event_time = time;
    row.data = JSON.stringify(envelope);
    const result = replayDashboardQuery(dashboardFieldQuery("Raw", stream), { Raw: [row] }, {}, dashboardClock, true);
    assert.ok(result.some(value => value.Error === "The DO observation has an invalid event_time."));
    assert.throws(() => stream === "telemetry" ? readDashboardFields(result, stream) : readDashboardFields(result, stream), /invalid event_time/);
  }
});

test("runtime queries are short contract-based KQL and setup does not certify unsupported native-only identities", () => {
  for (const stream of ["telemetry", "twin"] as const) {
    const row = dashboardRow(stream === "telemetry" ? { temperature: 7 } : { properties: { reported: { target: 8 } } },
      { profile: "native", stream });
    assert.throws(() => fields([row], stream), /CloudEvents/);
  }
  const generated = queries([
    { key: "temperature", leaf: "temperature", included: true, isNumeric: true },
    { key: "motor_status", component: "motor", leaf: "status", included: true, isNumeric: false },
  ], [{ key: "motor_target", component: "motor", leaf: "target", included: true }]).values;
  for (const [index, query] of generated.entries()) {
    assert.ok(query.split("\n").length <= 16);
    assert.ok(!/\blet\b|mv-expand|pack_array|dashboard_error|schema_id|schema_version|contract_version|event_kind|memberTypeIRI|__eventHubProfile|__native/.test(query));
    assert.ok(!query.includes("parse_json(tostring(") && !query.includes("headers["));
    assert.equal((query.match(/parse_json\(/g) || []).length, index === 0 ? 0 : 1);
    assert.ok(query.includes("ingestion_time() between (['_startTime'] .. ['_endTime'])") && !query.includes("ago(30d)"));
    if (index) {
      assert.ok(query.indexOf("where deviceId == _deviceId") < query.indexOf("parse_json("));
      if (query.includes("cloudEvents_iothubdtsubject")) assert.ok(query.indexOf("cloudEvents_iothubdtsubject") < query.indexOf("parse_json("));
      assert.ok(query.includes("event_time = todatetime(data.event_time)"));
    }
  }
  const base = dashboardRawQuery("Raw's table", "telemetry", "");
  assert.ok(base.startsWith("['Raw\\'s table']") && base.includes('cloudEvents_iothubdtsubject) == ""'));
  for (const stream of ["telemetry", "twin"] as const) {
    const setup = dashboardFieldQuery("Raw", stream);
    assert.ok(setup.includes("ingestion_time() >= ago(30d)") && !setup.includes("_startTime") && !setup.includes("_endTime"));
  }
});

test("supplied linked and regular Hub timestamp shapes keep the existing DO time source distinct from payload and transport enqueue", () => {
  const examples = [
    { device: "linked-hub-1-device-1", event: "2026-10-30T18:28:37.4670000Z", enqueue: "2026-10-30T18:28:37.5780000Z", reference: "urn:fixture:linked" },
    { device: "regular-hub-2-device-3", event: "2026-10-30T18:28:37.4330000Z", enqueue: "2026-10-30T18:28:37.6010000Z", reference: null },
  ];
  for (const example of examples) {
    const row = dashboardRow({ temperature: 7, timestamp: "2026-10-30T18:28:37.382Z" },
      { profile: "empty-native", device: example.device, time: example.event });
    const body = JSON.parse(String(row.data));
    body.EventEnqueuedUtcTime = example.enqueue; body.EventProcessedUtcTime = example.enqueue;
    body.EventHub = { EventEnqueuedUtcTime: example.enqueue }; body.User = { source: "fixture" };
    body.payload.memberTypeIRI = example.reference;
    row.data = JSON.stringify(body);
    row.ingestedAt = Date.parse(example.enqueue);
    assert.ok(isJsonObject(row.headers) && isJsonObject(row.user_headers));
    row.headers.EventEnqueuedUtcTime = example.enqueue; row.headers.IoTEnqueueTime = "0001-01-01T00:00:00.0000000Z";
    const tables = { Telemetry: [row] };
    const inspected = readDashboardFields(replayDashboardQuery(dashboardFieldQuery("Telemetry", "telemetry"), tables, {}, dashboardClock, true), "telemetry");
    assert.deepEqual(inspected.map(field => field.key), ["temperature", "timestamp"]);
    const generated = queries(inspected).values;
    const values = replayDashboardQuery(generated[1], tables, { _deviceId: example.device,
      _startTime: Date.parse("2026-10-30T18:00:00Z"), _endTime: Date.parse("2026-10-30T19:00:00Z") }, dashboardClock, true);
    assert.ok(values[0].event_time instanceof DashboardDateTime);
    assert.equal(values[0].event_time.iso, example.event);
    assert.notEqual(values[0].event_time.iso, example.enqueue);
    assert.notEqual(values[0].event_time.iso, "2026-10-30T18:28:37.3820000Z");
  }
});

test("supplied lifecycle shapes retain reported components whole objects and arrays without switching to operation or enqueue clocks", () => {
  const examples = [
    { device: "linked-hub-1-device-1", event: "2026-10-30T18:33:23.4660000Z", operation: "2026-10-30T18:33:23.3494524Z",
      enqueue: "2026-10-30T18:33:23.7870000Z", reported: { powerModule: { powerSource: "mains", batteryHealth: "poor", chargeCycles: 726, __t: "c" } },
      expected: ["powerModule_batteryHealth", "powerModule_chargeCycles", "powerModule_powerSource"] },
    { device: "regular-hub-1-device-2", event: "2026-10-30T18:33:23.3300000Z", operation: "2026-10-30T18:33:23.2225511Z",
      enqueue: "2026-10-30T18:33:23.5770000Z", reported: { firmware: "1.0", thermostat: { target: 22.5, __t: "c" },
        maintenanceSchedule: { interval: "monthly", enabled: false }, tags: ["regular", "fixture"] },
      expected: ["firmware", "maintenanceSchedule", "tags", "thermostat_target"] },
  ];
  for (const example of examples) {
    const row = dashboardRow({ properties: { reported: { ...example.reported, $metadata: { $lastUpdated: example.operation }, $version: 1 } } },
      { stream: "twin", profile: "empty-native", device: example.device, time: example.event });
    const envelope = JSON.parse(String(row.data));
    envelope.payload.memberTypeIRI = null; row.data = JSON.stringify(envelope);
    row.ingestedAt = Date.parse(example.enqueue);
    assert.ok(isJsonObject(row.headers) && isJsonObject(row.user_headers));
    row.headers.EventEnqueuedUtcTime = example.enqueue; row.headers.IoTEnqueueTime = "0001-01-01T00:00:00.0000000Z";
    row.user_headers.cloudEvents_operationtimestamp = example.operation;
    const tables = { Twins: [row] };
    const selected = readDashboardFields(replayDashboardQuery(dashboardFieldQuery("Twins", "twin"), tables, {}, dashboardClock, true), "twin");
    assert.deepEqual(selected.map(field => field.key), example.expected);
    const results = queries([], selected).values.slice(1).map(query => replayDashboardQuery(query, tables, {
      _deviceId: example.device, _startTime: Date.parse("2026-10-30T18:00:00Z"), _endTime: Date.parse("2026-10-30T19:00:00Z"),
    }, dashboardClock, true));
    for (const result of results) {
      assert.equal(result.length, 1); assert.ok(result[0].event_time instanceof DashboardDateTime);
      assert.equal(result[0].event_time.iso, example.event);
      assert.notEqual(result[0].event_time.iso, example.operation); assert.notEqual(result[0].event_time.iso, example.enqueue);
    }
    if (example.device.startsWith("regular")) {
      assert.deepEqual(results[selected.findIndex(field => field.key === "maintenanceSchedule")][0].Value, { interval: "monthly", enabled: false });
      assert.deepEqual(results[selected.findIndex(field => field.key === "tags")][0].Value, ["regular", "fixture"]);
    }
  }
});

test("all runtime paths use the selected ingestion window with inclusive boundaries and independent telemetry event time", () => {
  const start = Date.parse("2026-10-30T10:00:00Z"), end = Date.parse("2026-10-30T11:00:00Z");
  const row = (device: string, value: unknown, ingestedAt: number, event: string) =>
    ({ ...dashboardRow(value, { device, time: event }), ingestedAt });
  const telemetry = [
    row("device-1", { temperature: 1, status: "start" }, start, "2026-10-30T10:15:00Z"),
    row("device-1", { temperature: 2, status: "end" }, end, "2026-10-30T10:45:00Z"),
    row("too-early", { temperature: 99 }, start - 1, "2026-10-30T10:30:00Z"),
    row("too-late", { temperature: 100 }, end + 1, "2026-10-30T10:30:00Z"),
    row("old-event-new-arrival", { temperature: 101 }, start + 1, "2026-10-20T10:30:00Z"),
  ];
  const property = (reported: unknown, ingestedAt: number, event: string) =>
    ({ ...dashboardRow({ properties: { reported } }, { stream: "twin", time: event }), ingestedAt });
  const twin = [
    property({ firmware: "inside", settings: { enabled: false }, tags: ["whole"], motor: { __t: "c", target: 7 } }, start, "2026-10-20T11:00:00Z"),
    property({ firmware: "", motor: { __t: "c", target: 8 } }, end, "2026-10-30T10:50:00Z"),
    property({ firmware: "outside-newer", outsideOnly: "must-not-fall-back" }, end + 1, "2026-10-30T10:55:00Z"),
  ];
  const generated = queries([
    { key: "temperature", leaf: "temperature", included: true, isNumeric: true },
    { key: "status", leaf: "status", included: true, isNumeric: false },
  ], [
    { key: "firmware", leaf: "firmware", included: true }, { key: "settings", leaf: "settings", included: true },
    { key: "tags", leaf: "tags", included: true }, { key: "motor_target", component: "motor", leaf: "target", included: true },
    { key: "outsideOnly", leaf: "outsideOnly", included: true },
  ]);
  const params = { _deviceId: "device-1", _startTime: start, _endTime: end }, tables = { Telemetry: telemetry, Twins: twin };
  const results = generated.values.map(query => replayDashboardQuery(query, tables, params, end, true));
  assert.deepEqual(results[0], [{ deviceId: "device-1" }, { deviceId: "old-event-new-arrival" }]);
  assert.deepEqual(results[1].map(row => row.Value), [1, 2]);
  assert.deepEqual(results[2].map(row => row.Value), ["end", "start"]);
  assert.deepEqual(results[3].map(row => row.Value), ["inside"]);
  assert.deepEqual(results[4][0].Value, { enabled: false }); assert.deepEqual(results[5][0].Value, ["whole"]);
  assert.deepEqual(results[6].map(row => row.Value), [8]); assert.deepEqual(results[7], []);
  const oldDevice = replayDashboardQuery(generated.values[1], tables, { ...params, _deviceId: "old-event-new-arrival" }, end, true);
  assert.deepEqual(oldDevice, [], "Telemetry must retain its event_time window as well as ingestion time.");
  for (const result of generated.definition.queries) {
    assert.ok(isJsonObject(result) && Array.isArray(result.usedVariables));
    assert.ok(result.usedVariables.includes("_startTime") && result.usedVariables.includes("_endTime"));
    if (result === generated.definition.queries[0]) assert.ok(!result.usedVariables.includes("_deviceId"), "Device picker dependencies must not form a selected-device cycle.");
  }
  const past = { ...params, _startTime: Date.parse("2026-10-20T10:00:00Z"), _endTime: Date.parse("2026-10-20T12:00:00Z") };
  assert.deepEqual(replayDashboardQuery(generated.values[0], tables, past, end, true), []);
  assert.deepEqual(replayDashboardQuery(generated.values[3], tables, past, end, true), []);
});

test("wide runtime ingestion windows can include data older than thirty days without a hidden cap", () => {
  const end = Date.parse("2026-10-30T12:00:00Z"), old = end - 60 * 86400000;
  const time = new Date(old).toISOString();
  const tables = {
    Telemetry: [{ ...dashboardRow({ temperature: 4, status: "old" }, { time }), ingestedAt: old }],
    Twins: [{ ...dashboardRow({ properties: { reported: { firmware: "old-state" } } }, { stream: "twin", time }), ingestedAt: old }],
  };
  const generated = queries([
    { key: "temperature", included: true, isNumeric: true }, { key: "status", included: true, isNumeric: false },
  ], [{ key: "firmware", included: true }]).values;
  const parameters = { _deviceId: "device-1", _startTime: end - 90 * 86400000, _endTime: end };
  const results = generated.map(query => replayDashboardQuery(query, tables, parameters, end, true));
  assert.deepEqual(results[0], [{ deviceId: "device-1" }]);
  assert.deepEqual(results.slice(1).map(rows => rows[0].Value), [4, "old", "old-state"]);
  for (const query of generated) assert.ok(!query.includes("ago(30d)") && !query.includes("dashboard_error"));
});
