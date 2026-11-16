// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  decodeDoSamples, inferPayloadSchema, inferRootPreparationSchema, parseKqlTableResult, quoteKqlIdentifier,
  readAuthoringSamples, executeAuthoringRead, KqlRequestError,
  filterTablesByRole, parseIotMeta,
} from "./ModelAuthoring.ts";

// Synthetic fixtures for the reviewed DOObservationV1 contract; no production data.
const observation = (value: unknown, kind = "measurement") => ({
  entity_id: "", entity_type_id: "", event_time: "2026-09-21T12:34:56.0000000Z",
  event_kind: kind, contract_version: "1.0", schema_id: "com.do.telemetry.observation.v1",
  schema_version: "1.0", group_context_id: null, sequence_id: null,
  payload: { value, memberIRI: "", memberTypeIRI: "" },
});
const row = (value: unknown) => ({ data: JSON.stringify(value), headers: {}, user_headers: {} });
const source = {
  queryServiceUri: "https://example.kusto.fabric.microsoft.com", databaseName: "samples",
  telemetryTable: "Manual Telemetry", twinTable: "Manual Twins",
};
const result = (rows: unknown[][]) => ({
  Tables: [{
    Columns: [
      { ColumnName: "data", ColumnType: "string" },
      { ColumnName: "headers", ColumnType: "dynamic" },
      { ColumnName: "user_headers", ColumnType: "dynamic" },
    ], Rows: rows,
  }],
});

test("reuses V1 docstring tag parsing and exact telemetry/properties role matches", () => {
  const tables = ["Telemetry", "Properties", "Manual", "OtherRole", "Missing"];
  const docs = {
    Telemetry: "Measurements.\n@iot.meta tableRole=iothub-telemetry-raw\n@iot.meta source=hub",
    Properties: "@iot.meta tableRole=iothub-properties-raw",
    Manual: "Untagged manually created table.",
    OtherRole: "@iot.meta tableRole=iothub-telemetry-raw-extra",
  };
  assert.deepEqual(parseIotMeta(docs.Telemetry), { tableRole: "iothub-telemetry-raw", source: "hub" });
  assert.deepEqual(filterTablesByRole(tables, docs, "iothub-telemetry-raw"), ["Telemetry"]);
  assert.deepEqual(filterTablesByRole(tables, docs, "iothub-properties-raw"), ["Properties"]);
  assert.deepEqual(filterTablesByRole(tables, {}, "iothub-telemetry-raw"), []);
});
test("decodes observation arrays and individual intact observations, never transport wrappers", () => {
  const sample = decodeDoSamples([
    row([observation({ temperature: 22.5 }), observation({ compressor: { temperature: 24 } })]),
    row(observation({ pressure: 10 })),
  ], "telemetry", 100);
  assert.equal(sample.rawRows, 2);
  assert.equal(sample.observations, 3);
  assert.deepEqual(sample.payloads.map(item => item.value), [
    { temperature: 22.5 }, { compressor: { temperature: 24 } }, { pressure: 10 },
  ]);
  const schema = inferPayloadSchema(sample.payloads.map(item => item.value));
  assert.deepEqual(Object.keys(schema.properties!), ["compressor", "pressure", "temperature"]);
  assert.equal(schema.properties!.temperature.type, "number");
  assert.equal(schema.properties!.compressor.properties!.temperature.type, "integer");
});

test("bounds decoded observations as well as raw rows", () => {
  const many = Array.from({ length: 120 }, (_, temperature) => observation({ temperature }));
  const sample = decodeDoSamples([row(many)], "telemetry", 100);
  assert.equal(sample.payloads.length, 100);
  assert.throws(() => decodeDoSamples([], "telemetry", 101), /sample depth/);
});

test("extracts reported twin values, strips protocol metadata, and keeps component paths separate", () => {
  const sample = decodeDoSamples([row([observation({
    deviceId: "not a payload field",
    properties: { reported: {
      $version: 5, $metadata: {}, temperature: 21,
      motor: { __t: "c", temperature: 25, $metadata: {} },
      objectValue: { __t: "user data", temperature: 40 },
    } },
  }, "lifecycle")])], "twin", 100);
  assert.deepEqual(sample.payloads[0].value, {
    temperature: 21, motor: { temperature: 25 }, objectValue: { __t: "user data", temperature: 40 },
  });
});

test("does not equate arbitrary lifecycle bodies with twin reported-property samples", () => {
  const sample = decodeDoSamples([row([observation({ deviceId: "example" }, "lifecycle")])], "twin", 100);
  assert.equal(sample.payloads.length, 0);
  assert.equal(sample.skippedWithoutReportedProperties, 1);
  assert.throws(() => inferPayloadSchema([]), /No object payload samples/);
});

test("root-only preparation excludes bound component telemetry and marked reported objects without changing source samples", () => {
  const nativeRow = (value: unknown, component = "", module = "") => ({
    ...row(value), headers: { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: module, IoTSubject: component },
  });
  const telemetry = decodeDoSamples([
    nativeRow(observation({ temperature: 22.5, reading: 0, enabled: false, nullable: null, empty: "", values: [] })),
    nativeRow(observation({ targetTemperature: 25, fanSpeed: 3, compressorLoad: 0.8, reading: "component-value" }), "thermostat"),
    { ...row(observation({ powerConsumption: 40, current: 2, voltageRipple: 0.01 })),
      user_headers: { cloudEvents_subject: "device-1", cloudEvents_iothubdtsubject: "powerModule" } },
    nativeRow(observation({ moduleReading: 1 }), "", "module-1"),
  ], "telemetry", 100);
  const reported = {
    $version: 2, firmware: "private-version", maintenanceSchedule: { enabled: false, day: 0, nullable: null, note: "" },
    tags: ["private-tag"], thermostat: { __t: "c", targetTemperature: 26 },
    powerModule: { __t: "c", current: 3 },
  };
  const twin = decodeDoSamples([nativeRow(observation({ properties: { reported } }, "lifecycle"))], "twin", 100);
  const before = JSON.stringify({ telemetry, twin });
  const oldSchema = inferPayloadSchema([...telemetry.payloads, ...twin.payloads].map(sample => sample.value));
  assert.deepEqual(oldSchema.properties!.reading.type, ["integer", "string"]);
  assert.ok(oldSchema.properties!.thermostat && oldSchema.properties!.current);
  const schema = inferRootPreparationSchema(telemetry, twin);
  assert.deepEqual(schema, inferPayloadSchema([
    telemetry.payloads[0].value, telemetry.payloads[3].value,
    { firmware: reported.firmware, maintenanceSchedule: reported.maintenanceSchedule, tags: reported.tags },
  ]));
  for (const key of ["targetTemperature", "fanSpeed", "compressorLoad", "powerConsumption", "current", "voltageRipple", "thermostat", "powerModule"]) {
    assert.ok(!Object.prototype.hasOwnProperty.call(schema.properties, key), key);
  }
  assert.equal(schema.properties!.reading.type, "integer");
  assert.equal(schema.properties!.moduleReading.type, "integer");
  assert.equal(telemetry.payloads[3].native!.identity, "device-1/module-1");
  assert.equal(telemetry.payloads[3].component, undefined);
  assert.equal(schema.properties!.enabled.type, "boolean");
  assert.equal(schema.properties!.nullable.type, "null");
  assert.deepEqual(schema.properties!.values, { type: "array", items: {} });
  assert.equal(JSON.stringify({ telemetry, twin }), before);
  assert.equal(telemetry.payloads.length, 4);
  assert.deepEqual(twin.payloads[0].reported, reported);
  assert.ok(!JSON.stringify(schema).includes("private-version") && !JSON.stringify(schema).includes("private-tag"));
});

test("root-only preparation keeps same-named ordinary objects arrays and false zero null empty values", () => {
  const reported = {
    thermostat: { enabled: false, count: 0, nullable: null, text: "", list: [], nested: { __t: "c", userValue: 1 } },
    powerModule: { __t: "not-a-component", values: [0, false, null, ""] },
    ordinaryArray: [{ __t: "c", ordinaryValue: 1 }], markerNotString: { __t: true, value: 0 },
  };
  const twin = decodeDoSamples([row(observation({ properties: { reported } }, "lifecycle"))], "twin", 100);
  assert.deepEqual(inferRootPreparationSchema(decodeDoSamples([], "telemetry", 100), twin), inferPayloadSchema([reported]));
  assert.deepEqual(twin.payloads[0].reported, reported);
});

test("root-only preparation reports no eligible fields and invalid provenance rather than inventing a root schema", () => {
  const emptyTelemetry = decodeDoSamples([], "telemetry", 100), emptyTwin = decodeDoSamples([], "twin", 100);
  const component = decodeDoSamples([{
    ...row(observation({ componentOnly: 1 })), headers: { IoTConnectionDeviceId: "device-1", IoTSubject: "component" },
  }], "telemetry", 100);
  const markerOnly = decodeDoSamples([row(observation({
    properties: { reported: { $version: 1, thermostat: { __t: "c", value: 2 } } },
  }, "lifecycle"))], "twin", 100);
  for (const [telemetry, twin] of [[component, emptyTwin], [emptyTelemetry, markerOnly], [component, markerOnly]]) {
    assert.throws(() => inferRootPreparationSchema(telemetry, twin), /sampled records contain no root-level fields/);
  }
  assert.throws(() => inferRootPreparationSchema(emptyTelemetry, emptyTwin), /No object payload samples/);
  assert.deepEqual(inferPayloadSchema([{}]).properties, {});
  const noOriginal = { ...markerOnly, payloads: markerOnly.payloads.map(({ reported: _reported, ...sample }) => sample) };
  assert.throws(() => inferRootPreparationSchema(emptyTelemetry, noOriginal), {
    message: "The reported-property sample is missing its original reported payload. Choose Validate again before preparing a model.",
  });
  const malformed = decodeDoSamples([{
    ...row(observation({ unknownBinding: 1 })), headers: { IoTConnectionDeviceId: "device-1", IoTSubject: 7 },
  }], "telemetry", 100);
  assert.ok(malformed.payloads[0].native!.issues.includes("Telemetry component metadata is not a string."));
  assert.throws(() => inferRootPreparationSchema(malformed, emptyTwin), /Telemetry component metadata is not a string/);
  assert.throws(() => decodeDoSamples([{ ...row(observation({ value: 1 })), headers: [] }], "telemetry", 100), /metadata must be an object/);
});

test("rejects legacy, malformed, unsupported and nonobject data rather than guessing", () => {
  for (const stream of ["telemetry", "twin"] as const) {
    assert.throws(() => decodeDoSamples([row({ temperature: 12 })], stream, 100), {
      message: `${stream === "telemetry" ? "Telemetry" : "Twin-change"} row 1 is not a supported DOObservationV1 observation (contract/schema version 1.0). ` +
        "Matching raw columns alone do not establish compatibility. Run Ingest device data from your IoT Hub to receive new DOObservationV1 events in compatible raw tables, then select those tables and choose Validate again.",
    });
  }
  assert.throws(() => decodeDoSamples([{ ...row([]), data: "{bad" }], "telemetry", 100), /invalid JSON/);
  assert.throws(() => decodeDoSamples([row([])], "telemetry", 100), /empty observation array/);
  for (const value of [null, [], 12, "text", true]) {
    assert.throws(() => decodeDoSamples([row([observation(value)])], "telemetry", 100), /non-object payload/);
  }
  assert.throws(() => decodeDoSamples([row([{ ...observation({}), contract_version: "2.0" }])], "telemetry", 100), /not a supported/);
  assert.throws(() => decodeDoSamples([row([observation({}, "lifecycle")])], "telemetry", 100), /event kind/);
  assert.throws(() => decodeDoSamples([{ ...row([]), data: " ".repeat(2 * 1024 * 1024 + 1) }], "telemetry", 100), /2 MB/);
});

test("export contains only standard observed types/nesting/alternatives and no sample values or invented facts", () => {
  const schema = inferPayloadSchema([
    { state: "private-sample-value", reading: 2, nested: { enabled: true }, values: [], variant: null },
    { state: 3, reading: 2.5, nested: { enabled: null }, values: [1, "x"], variant: { x: 4 } },
  ]);
  assert.deepEqual(schema, {
    $schema: "https://json-schema.org/draft/2020-12/schema", type: "object",
    properties: {
      nested: { type: "object", properties: { enabled: { type: ["boolean", "null"] } } },
      reading: { type: "number" },
      state: { type: ["integer", "string"] },
      values: { type: "array", items: { type: ["integer", "string"] } },
      variant: { type: ["null", "object"], properties: { x: { type: "integer" } } },
    },
  });
  const text = JSON.stringify(schema);
  for (const forbidden of ["private-sample-value", "required", "enum", "default", "unit", "source", "sample"]) {
    assert.ok(!text.includes(forbidden), forbidden);
  }
});

test("handles unusual payload keys without prototype pollution or collisions", () => {
  const input = JSON.parse('{"__proto__":{"polluted":true},"constructor":1,"a.b":2,"a":{"b":"x"}}');
  const schema = inferPayloadSchema([input]);
  assert.ok(Object.prototype.hasOwnProperty.call(schema.properties, "__proto__"));
  assert.equal(schema.properties!["a.b"].type, "integer");
  assert.equal(schema.properties!.a.properties!.b.type, "string");
  assert.equal(({} as Record<string, unknown>).polluted, undefined);
});

test("validates exact query-result shapes and surfaces partial failures", () => {
  assert.deepEqual(parseKqlTableResult(result([])).Rows, []);
  for (const invalid of [{}, { Tables: [] }, { error: "failure" }, { ...result([]), Exceptions: ["partial failure"] },
    { ...result([]), Exceptions: "malformed failure" }]) {
    assert.throws(() => parseKqlTableResult(invalid));
  }
  assert.throws(() => parseKqlTableResult(result([["too short"]])), /shape/);
  assert.equal(quoteKqlIdentifier("a'b\\c"), "['a\\'b\\\\c']");
  assert.throws(() => quoteKqlIdentifier("\n"), /valid KQL/);
});

test("reads only bounded selected tables, allows one empty stream and performs no service mutations", async () => {
  const calls: { url: string; csl: string }[] = [];
  const fetcher: typeof fetch = async (url, init) => {
    const body = JSON.parse(String(init?.body));
    calls.push({ url: String(url), csl: body.csl });
    return new Response(JSON.stringify(result(
      body.csl.includes("Manual Telemetry") ? [[JSON.stringify([observation({ temperature: 22.5 })]), {}, {}]] : []
    )), { status: 200 });
  };
  const sample = await readAuthoringSamples(source, "mock-token", 100, undefined, fetcher);
  assert.equal(sample.telemetry.payloads.length, 1);
  assert.equal(sample.twin.payloads.length, 0);
  assert.equal(calls.length, 2);
  assert.ok(calls.every(call => call.url.endsWith("/v1/rest/query") && call.csl.includes("| top 100 ")));
  assert.ok(calls.every(call => !/\.(create|alter|set-or-append|ingest)\b/i.test(call.csl)));
});

test("public authoring reads use the same root-only schema while retaining both original component streams", async () => {
  const telemetryRows = [
    [JSON.stringify(observation({ rootValue: 1 })), { IoTConnectionDeviceId: "device-1", IoTSubject: "" }, {}],
    [JSON.stringify(observation({ componentValue: "not-root" })), { IoTConnectionDeviceId: "device-1", IoTSubject: "motor" }, {}],
  ];
  const reported = { rootState: false, motor: { __t: "c", componentState: 1 }, ordinary: { value: 0 } };
  const twinRows = [[JSON.stringify(observation({ properties: { reported } }, "lifecycle")), {}, {}]];
  const original = JSON.stringify({ telemetryRows, twinRows }), queries: string[] = [];
  const fetcher: typeof fetch = async (_url, init) => {
    const body = JSON.parse(String(init?.body)); queries.push(body.csl);
    return new Response(JSON.stringify(result(body.csl.includes("Manual Telemetry") ? telemetryRows : twinRows)));
  };
  const sample = await readAuthoringSamples(source, "mock-token", 100, undefined, fetcher);
  assert.deepEqual(sample.schema, inferPayloadSchema([{ rootValue: 1 }, { rootState: false, ordinary: { value: 0 } }]));
  assert.deepEqual(sample.schema, inferRootPreparationSchema(sample.telemetry, sample.twin));
  assert.equal(sample.telemetry.payloads[1].component, "motor");
  assert.deepEqual(sample.twin.payloads[0].reported, reported);
  assert.equal(queries.length, 2);
  assert.ok(queries.every(query => query.includes("| top 100 ") && !query.includes(".create")));
  assert.equal(JSON.stringify({ telemetryRows, twinRows }), original);
});

test("all-empty, permission, partial-result and invalid-column cases never yield an export", async () => {
  const empty: typeof fetch = async () => new Response(JSON.stringify(result([])));
  await assert.rejects(readAuthoringSamples(source, "mock", 100, undefined, empty), /No object payload/);
  const denied: typeof fetch = async () => new Response("", { status: 403 });
  await assert.rejects(readAuthoringSamples(source, "mock", 100, undefined, denied), /denied \(403\)/);
  const partial: typeof fetch = async () => new Response(JSON.stringify({ ...result([]), Exceptions: ["limit"] }));
  await assert.rejects(readAuthoringSamples(source, "mock", 100, undefined, partial), /did not complete/);
  const wrongColumns: typeof fetch = async () => {
    const value = result([]);
    value.Tables[0].Columns[0].ColumnType = "long";
    return new Response(JSON.stringify(value));
  };
  await assert.rejects(readAuthoringSamples(source, "mock", 100, undefined, wrongColumns), /data:string/);
  await assert.rejects(executeAuthoringRead({ ...source, queryServiceUri: "http://example.com" }, "mock", "read", "query", undefined, empty), /address/);
});

test("carries cancellation through both reads", async () => {
  const controller = new AbortController();
  controller.abort();
  const cancelled: typeof fetch = async (_url, init) => {
    assert.equal(init?.signal, controller.signal);
    throw new DOMException("Cancelled", "AbortError");
  };
  await assert.rejects(readAuthoringSamples(source, "mock", 100, controller.signal, cancelled), /Cancelled/);
});

test("query and management requests supply required Accept without changing body path auth or signal", async () => {
  for (const kind of ["query", "mgmt"] as const) {
    const controller = new AbortController(), csl = "exact request text\n| getschema";
    let calls = 0;
    const fetcher: typeof fetch = async (url, init) => {
      calls++;
      assert.equal(String(url), `${source.queryServiceUri}/v1/rest/${kind}`);
      assert.equal(init?.method, "POST"); assert.equal(init?.signal, controller.signal);
      const headers = new Headers(init?.headers);
      assert.equal(headers.get("Accept"), "application/json");
      assert.equal(headers.get("Content-Type"), "application/json");
      assert.equal(headers.get("Authorization"), "Bearer mock-token");
      assert.deepEqual(JSON.parse(String(init?.body)), { db: source.databaseName, csl });
      return new Response(JSON.stringify(result([])));
    };
    await executeAuthoringRead(source, "mock-token", csl, kind, controller.signal, fetcher);
    assert.equal(calls, 1);
  }
});

test("non-OK native plaintext exposes only its code message and safe ActivityId", async () => {
  const activityId = "00000000-0000-4000-8000-000000000001";
  const body = "General_BadRequest: Request is invalid and cannot be executed.\r\nError details:\r\n" +
    `ClientRequestId='unspecified;00000000-0000-4000-8000-000000000002', ActivityId='${activityId}', Timestamp='2026-10-20T19:22:18Z'.`;
  const fetcher: typeof fetch = async () => new Response(body, { status: 400, headers: { "Content-Type": "text/plain" } });
  await assert.rejects(executeAuthoringRead(source, "mock-token", "unchanged", "query", undefined, fetcher), (error: unknown) => {
    assert.ok(error instanceof KqlRequestError);
    assert.equal(error.status, 400); assert.equal(error.code, "General_BadRequest");
    assert.equal(error.nativeMessage, "Request is invalid and cannot be executed.");
    assert.equal(error.activityId, activityId);
    assert.ok(error.message.includes("failed (400)") && error.message.includes(activityId));
    assert.ok(!error.message.includes("ClientRequestId") && !error.message.includes("Timestamp") && !error.message.includes("Check permissions"));
    return true;
  });
});

test("native JSON uses allowlisted detail fields and redacts credential/header echoes without retaining raw context", async () => {
  const token = "mock-secret-token";
  const fetcher: typeof fetch = async () => new Response(JSON.stringify({ error: {
    code: "General_BadRequest", message: "Generic outer failure",
    "@message": `Specific native cause. Bearer ${token}\n"Authorization":"Basic basic-secret"\n` +
      '"Cookie":"cookie-secret"\nclient_secret="secondary-secret"\n"token":"other-token-secret"\n',
    "@context": { headers: { Authorization: `Bearer ${token}`, Cookie: "context-cookie-secret" }, Query: "raw-query-secret", Principal: "principal-secret" },
    stack: "stack-secret",
  } }), { status: 400 });
  await assert.rejects(executeAuthoringRead(source, token, "original-query", "query", undefined, fetcher), (error: unknown) => {
    assert.ok(error instanceof KqlRequestError);
    assert.equal(error.code, "General_BadRequest");
    assert.ok(error.nativeMessage?.includes("Specific native cause."));
    const exposed = error.message + JSON.stringify(error);
    for (const secret of [token, "basic-secret", "cookie-secret", "secondary-secret", "other-token-secret", "context-cookie-secret", "raw-query-secret", "principal-secret", "stack-secret"]) {
      assert.ok(!exposed.includes(secret), "A non-allowlisted or credential-bearing value escaped.");
    }
    assert.ok(!("body" in error) && !("headers" in error) && !("cause" in error));
    return true;
  });
});

test("native failure body and message limits cancel oversized streams rather than exposing full bodies", async () => {
  let cancelled = false, pulls = 0;
  const stream = new ReadableStream<Uint8Array>({
    pull(controller) { pulls++; controller.enqueue(new TextEncoder().encode("x".repeat(2048))); },
    cancel() { cancelled = true; },
  });
  const fetcher: typeof fetch = async () => new Response(stream, { status: 400 });
  await assert.rejects(executeAuthoringRead(source, "mock-token", "read", "query", undefined, fetcher), (error: unknown) => {
    assert.ok(error instanceof KqlRequestError);
    assert.equal(error.status, 400); assert.equal(error.nativeMessage, undefined);
    assert.ok(error.message.includes("8 KiB") && error.message.length < 1000);
    return true;
  });
  assert.ok(cancelled && pulls <= 6, "The failure reader did not stop near its byte bound.");
  const long: typeof fetch = async () => new Response(`General_BadRequest: ${"d".repeat(20000)}`, { status: 400 });
  await assert.rejects(executeAuthoringRead(source, "mock-token", "read", "query", undefined, long), (error: unknown) => {
    assert.ok(error instanceof KqlRequestError);
    assert.ok(error.nativeMessage!.length <= 612 && error.nativeMessage!.endsWith("[truncated]") && error.message.length < 1000);
    return true;
  });
});

test("empty unrecognized malformed and unreadable non-OK bodies remain explicit HTTP failures", async () => {
  for (const body of ["", "<html>private-body-secret</html>", '{"error":', '{"error":{"code":42,"message":{"unsafe":"object"}}}']) {
    const fetcher: typeof fetch = async () => new Response(body, { status: 500 });
    await assert.rejects(executeAuthoringRead(source, "mock-token", "read", "query", undefined, fetcher), (error: unknown) => {
      assert.ok(error instanceof KqlRequestError);
      assert.equal(error.status, 500); assert.equal(error.nativeMessage, undefined);
      assert.ok(error.message.includes("500") && !error.message.includes("private-body-secret"));
      return true;
    });
  }
  const fetcher: typeof fetch = async () => new Response(new ReadableStream<Uint8Array>({
    start(controller) { controller.error(new Error("private-reader-secret")); },
  }), { status: 400 });
  await assert.rejects(executeAuthoringRead(source, "mock-token", "read", "query", undefined, fetcher), (error: unknown) => {
    assert.ok(error instanceof KqlRequestError);
    assert.equal(error.status, 400); assert.ok(error.message.includes("could not be read") && !error.message.includes("private-reader-secret"));
    return true;
  });
});

test("401 and 403 preserve access-denied handling with no auth retry or success fallback", async () => {
  for (const status of [401, 403]) {
    let calls = 0;
    const fetcher: typeof fetch = async () => {
      calls++;
      return new Response(JSON.stringify({ error: { code: "AccessDenied", message: "Native access denied." } }), { status });
    };
    await assert.rejects(executeAuthoringRead(source, "mock-token", "read", "query", undefined, fetcher), (error: unknown) => {
      assert.ok(error instanceof KqlRequestError);
      assert.equal(error.status, status); assert.equal(error.code, "AccessDenied");
      assert.ok(error.message.includes(`denied (${status})`) && error.message.includes("sign in again"));
      return true;
    });
    assert.equal(calls, 1);
  }
});

test("abort while reading a native error body stays cancellation rather than a wrapped HTTP detail", async () => {
  const controller = new AbortController();
  const fetcher: typeof fetch = async (_url, init) => {
    assert.equal(init?.signal, controller.signal);
    const stream = new ReadableStream<Uint8Array>({
      start(body) { controller.signal.addEventListener("abort", () => body.error(controller.signal.reason), { once: true }); },
    });
    return new Response(stream, { status: 400 });
  };
  const pending = executeAuthoringRead(source, "mock-token", "read", "query", controller.signal, fetcher);
  await Promise.resolve();
  controller.abort(new DOMException("Cancelled error-body read", "AbortError"));
  await assert.rejects(pending, (error: unknown) => {
    assert.ok(error instanceof Error && error.name === "AbortError" && !(error instanceof KqlRequestError));
    return true;
  });
});
