// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import fs from "node:fs";
import { canonicalJson, decodeDoSamples, isJsonObject, type JsonObject, type SampleStream } from "./ModelAuthoring.ts";
import { applyRoleHints, coverageReady, fieldValidation, fieldValue, fromDtdlModel, fromThingModel, inspectFieldEvidence, modelingFingerprint,
  omittedObservedFields, setFieldRole } from "./HubModeling.ts";
import { resolveThingModel } from "./ThingModelParser.ts";
import { parseDtdlModel } from "./DtdlModelParser.ts";
import { buildModeledDataPlan, semanticDocumentation } from "./ModeledDataPlan.ts";
import { creationFingerprint, ModeledDataCreator, rememberCreationAttempt, type CreationRequest } from "./ModeledDataCreation.ts";
import { createCreationMock, creationRequest } from "../../../tests/creationMock.ts";
import { buildDataAgentDataSource } from "./DataAgentContext.ts";
import { bindNativeMetadata, nativeProjection } from "./HubTransport.ts";
import { mixedDtdlRequest, mixedRawSample, mixedReportedValues, mixedTelemetryValues } from "../../../tests/mixedSourceDtdlFixtures.ts";

const propagationOutcomes: Record<string, unknown>[] = [];
process.once("exit", () => {
  if (process.env.PROPAGATION_VARIATION_RESULTS) fs.writeFileSync(process.env.PROPAGATION_VARIATION_RESULTS,
    JSON.stringify({ method: "Synthetic replay of the actual emitted KQL subset; not live Kusto, Hub or ADR propagation.",
      count: propagationOutcomes.length, cases: propagationOutcomes }, null, 2) + "\n");
});

// Interpret the expression actually emitted into the normalizer, not a second field-presence implementation.
function evaluate(expression: string, bindings: Record<string, unknown>): unknown {
  const tokens = expression.match(/"(?:[^"\\]|\\.)*"|-?\d+(?:\.\d+)?|>=|==|!=|[a-zA-Z_]\w*|[()[\],.]/g) || [];
  let index = 0;
  const take = (expected: string) => {
    assert.equal(tokens[index++], expected, `Unexpected normalizer syntax in ${expression}`);
  };
  const binary = (operator: string, left: unknown, right: unknown): unknown => {
    if (operator === "or") return left === true || right === true ? true : left === null || right === null ? null : false;
    if (operator === "and") return left === false || right === false ? false : left === null || right === null ? null : true;
    if (left === null || right === null) return null;
    if (operator === "==") return left === right;
    if (operator === "!=") return left !== right;
    assert.ok(typeof left === "number" && typeof right === "number");
    return left >= right;
  };
  const call = (name: string, args: unknown[]): unknown => {
    switch (name) {
      case "dynamic": return args[0];
      case "tostring": return args[0] === null ? "" : String(args[0]);
      case "isnull": return args[0] === null;
      case "isempty": return args[0] === null || args[0] === "";
      case "not": return args[0] === null ? null : !args[0];
      case "gettype": return isJsonObject(args[0]) ? "dictionary" : Array.isArray(args[0]) ? "array" : typeof args[0];
      case "iff": return args[0] === true ? args[1] : args[2];
      case "strcat": return args.map(value => value === null ? "" : String(value)).join("");
      case "todynamic": return args[0];
      case "bag_keys": return isJsonObject(args[0]) ? Object.keys(args[0]) : null;
      case "array_index_of": return Array.isArray(args[0]) ? args[0].indexOf(args[1]) : null;
      case "pack_array": return args;
      case "bag_pack": {
        assert.equal(args.length % 2, 0);
        const entries: [string, unknown][] = [];
        for (let position = 0; position < args.length; position += 2) {
          assert.equal(typeof args[position], "string");
          entries.push([String(args[position]), args[position + 1]]);
        }
        return Object.fromEntries(entries);
      }
      default: throw new Error(`Untested emitted Kusto function: ${name}`);
    }
  };
  const primary = (): unknown => {
    const token = tokens[index++];
    assert.ok(token, "Missing normalizer token.");
    let value: unknown;
    if (token === "(") { value = or(); take(")"); }
    else if (token.startsWith('"')) value = JSON.parse(token);
    else if (/^-?\d/.test(token)) value = Number(token);
    else if (["true", "false", "null"].includes(token)) value = JSON.parse(token);
    else if (tokens[index] === "(") {
      index++;
      const args: unknown[] = [];
      if (tokens[index] !== ")") {
        do { args.push(or()); } while (tokens[index] === "," && ++index);
      }
      take(")"); value = call(token, args);
    } else {
      assert.ok(Object.prototype.hasOwnProperty.call(bindings, token), `Unknown emitted binding ${token}`);
      value = bindings[token];
    }
    while (tokens[index] === "[" || tokens[index] === ".") {
      const access = tokens[index++];
      const key = access === "[" ? or() : tokens[index++];
      if (access === "[") take("]");
      value = isJsonObject(value) && typeof key === "string" && Object.prototype.hasOwnProperty.call(value, key) ? value[key] : null;
    }
    return value;
  };
  const comparison = (): unknown => {
    const left = primary();
    return [">=", "==", "!="].includes(tokens[index]) ? binary(tokens[index++], left, primary()) : left;
  };
  const and = (): unknown => {
    let value = comparison();
    while (tokens[index] === "and") { index++; value = binary("and", value, comparison()); }
    return value;
  };
  const or = (): unknown => {
    let value = and();
    while (tokens[index] === "or") { index++; value = binary("or", value, and()); }
    return value;
  };
  const value = or();
  assert.equal(index, tokens.length, "The emitted normalizer contains unsupported syntax.");
  return value;
}

function projectedNative(stream: SampleStream, headers: JsonObject, user: JsonObject): Record<string, unknown> {
  const bindings: Record<string, unknown> = { headers, user_headers: user };
  for (const line of nativeProjection(stream)) {
    for (const assignment of line.slice("| extend ".length).split(/, (?=__\w+ =)/)) {
      const separator = assignment.indexOf(" = ");
      bindings[assignment.slice(0, separator)] = evaluate(assignment.slice(separator + 3), bindings);
    }
  }
  return bindings;
}

test("emitted native binding rejects malformed primary components in both profiles without excluding roots or valid components", () => {
  for (const profile of ["eventhub", "cloudEvents"] as const) {
    for (const component of [undefined, null, "", "motor", 42, { name: "motor" }, [], false]) {
      const headers: JsonObject = profile === "eventhub" ? { IoTConnectionDeviceId: "device-1" } : {};
      const user: JsonObject = profile === "cloudEvents" ? { cloudEvents_subject: "device-1" } : {};
      if (component !== undefined) {
        if (profile === "eventhub") headers.IoTSubject = component;
        else user.cloudEvents_iothubdtsubject = component;
      }
      const bound = bindNativeMetadata(headers, user, "telemetry");
      const bindings = projectedNative("telemetry", headers, user);
      const valid = component === undefined || component === null || typeof component === "string";
      assert.equal(bindings.deviceId, valid ? "device-1" : "", `${profile}: ${JSON.stringify(component)}`);
      assert.equal(bindings.deviceId, bound.issues.length ? "" : bound.identity);
      assert.equal(bindings.component, typeof component === "string" ? component : "");
    }
  }
  assert.ok(!nativeProjection("twin").join("\n").includes('gettype(iff(__eventHubProfile, headers["IoTSubject"]'),
    "Twin messages must not gain a telemetry-component prerequisite.");
});

test("blank-string EventHub placeholders use the existing native CloudEvents device/module/component profile in sampler and emitted KQL", () => {
  for (const stream of ["telemetry", "twin"] as const) for (const module of ["", "module-1"]) {
    for (const component of [undefined, null, "", "motor"]) {
      const identity = module ? `device-1/${module}` : "device-1";
      const headers = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" };
      const user: JsonObject = stream === "telemetry" ? { cloudEvents_subject: identity, cloudEvents_moduleid: module,
        ...(component !== undefined ? { cloudEvents_iothubdtsubject: component } : {}) }
        : { cloudEvents_subject: identity, cloudEvents_deviceid: "device-1", cloudEvents_moduleid: module, deviceId: "device-1" };
      const bound = bindNativeMetadata(headers, user, stream), projected = projectedNative(stream, headers, user);
      assert.equal(bound.profile, "cloudEvents"); assert.deepEqual(bound.issues, []);
      assert.equal(bound.identity, identity); assert.equal(projected.deviceId, identity);
      assert.equal(projected.__eventHubProfile, false); assert.equal(projected.__emptyEventHubProfile, true);
      assert.equal(projected.component, stream === "telemetry" && typeof component === "string" ? component : "");
      assert.equal(bound.component, stream === "telemetry" && component ? component : undefined);
    }
  }
});

test("DO sample decoding keeps whole payloads and verified native CE bindings with blank EventHub placeholders", () => {
  for (const stream of ["telemetry", "twin"] as const) {
    const headers = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" };
    const user: JsonObject = stream === "telemetry" ? { cloudEvents_subject: "device-1", cloudEvents_iothubdtsubject: "motor" }
      : { cloudEvents_subject: "device-1", cloudEvents_deviceid: "device-1" };
    const payload = stream === "telemetry" ? { reading: 22 } : { properties: { reported: { state: "on" } } };
    const decoded = decodeDoSamples([{
      data: JSON.stringify({ entity_id: "urn:semantic:not-native", schema_id: "com.do.telemetry.observation.v1",
        schema_version: "1.0", contract_version: "1.0", event_time: "2026-10-22T19:00:00Z",
        event_kind: stream === "telemetry" ? "measurement" : "lifecycle", payload: { value: payload } }),
      headers, user_headers: user,
    }], stream, 100);
    const sample = decoded.payloads[0];
    assert.equal(sample.native?.profile, "cloudEvents"); assert.equal(sample.native?.identity, "device-1");
    assert.deepEqual(sample.native?.issues, []);
    assert.equal(sample.component, stream === "telemetry" ? "motor" : undefined);
    assert.deepEqual(sample.value, stream === "telemetry" ? { reading: 22 } : { state: "on" });
    assert.equal(projectedNative(stream, headers, user).deviceId, sample.native?.identity);
  }
});

test("the empty-EH exception is limited to a string-empty device and optional empty/null/absent module and component, not other missing/null rules", () => {
  for (const stream of ["telemetry", "twin"] as const) {
    const user: JsonObject = stream === "telemetry" ? { cloudEvents_subject: "device-1" } : { cloudEvents_deviceid: "device-1" };
    for (const module of [undefined, null, ""]) for (const component of [undefined, null, ""]) {
      const headers: JsonObject = { IoTConnectionDeviceId: "",
        ...(module !== undefined ? { IoTConnectionModuleId: module } : {}),
        ...(component !== undefined ? { IoTSubject: component } : {}) };
      assert.equal(bindNativeMetadata(headers, user, stream).identity, "device-1");
      assert.equal(projectedNative(stream, headers, user).deviceId, "device-1");
    }
    for (const device of [undefined, null, 0, false, [], {}, " "]) {
      const headers: JsonObject = { ...(device !== undefined ? { IoTConnectionDeviceId: device } : {}),
        IoTConnectionModuleId: "", IoTSubject: "" };
      const bound = bindNativeMetadata(headers, user, stream);
      assert.equal(bound.profile, "eventhub"); assert.equal(bound.identity, undefined);
      assert.equal(projectedNative(stream, headers, user).deviceId, "");
    }
    assert.equal(bindNativeMetadata({}, user, stream).identity, "device-1");
    assert.equal(projectedNative(stream, {}, user).deviceId, "device-1");
  }
});

test("nonempty or malformed orphan EventHub module/component metadata cannot be discarded through CloudEvents fallback", () => {
  for (const stream of ["telemetry", "twin"] as const) for (const field of ["IoTConnectionModuleId", "IoTSubject"]) {
    const user: JsonObject = stream === "telemetry" ? { cloudEvents_subject: "device-1" } : { cloudEvents_deviceid: "device-1" };
    for (const value of ["orphan", 7, false, [], {}]) {
      const headers: JsonObject = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "", [field]: value };
      const bound = bindNativeMetadata(headers, user, stream);
      assert.equal(bound.profile, "eventhub"); assert.equal(bound.identity, undefined); assert.ok(bound.issues.length);
      assert.equal(projectedNative(stream, headers, user).deviceId, "");
    }
  }
});

test("valid EventHub precedence and contradictory profile/module/component/twin identities remain fail-closed in both paths", () => {
  for (const stream of ["telemetry", "twin"] as const) {
    const headers = { IoTConnectionDeviceId: "device-1", IoTConnectionModuleId: "module-1", IoTSubject: "motor" };
    const user: JsonObject = stream === "telemetry"
      ? { cloudEvents_subject: "device-1/module-1", cloudEvents_moduleid: "module-1", cloudEvents_iothubdtsubject: "motor" }
      : { cloudEvents_deviceid: "device-1", cloudEvents_moduleid: "module-1", deviceId: "device-1" };
    assert.equal(bindNativeMetadata(headers, user, stream).profile, "eventhub");
    assert.equal(projectedNative(stream, headers, user).deviceId, "device-1/module-1");
    const conflict = stream === "telemetry" ? { ...user, cloudEvents_subject: "other/module-1" } : { ...user, cloudEvents_moduleid: "other" };
    assert.equal(bindNativeMetadata(headers, conflict, stream).identity, undefined);
    assert.equal(projectedNative(stream, headers, conflict).deviceId, "");
  }
  const telemetryUser = { cloudEvents_subject: "device-1", cloudEvents_iothubdtsubject: "other" };
  const headers = { IoTConnectionDeviceId: "device-1", IoTSubject: "motor" };
  assert.equal(bindNativeMetadata(headers, telemetryUser, "telemetry").identity, undefined);
  assert.equal(projectedNative("telemetry", headers, telemetryUser).deviceId, "");
  for (const device of ["device-1", ""]) {
    const twinHeaders = { IoTConnectionDeviceId: device, IoTConnectionModuleId: "", IoTSubject: "" };
    const user = { cloudEvents_deviceid: "device-1", deviceId: "other" };
    assert.equal(bindNativeMetadata(twinHeaders, user, "twin").identity, undefined);
    assert.equal(projectedNative("twin", twinHeaders, user).deviceId, "");
  }
});

test("blank-EH fallback cannot hide malformed CloudEvents module/components or invent identity from deviceref or payload fields", () => {
  const headers = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" };
  for (const stream of ["telemetry", "twin"] as const) for (const malformed of [8, false, {}, []]) {
    const user: JsonObject = stream === "telemetry" ? { cloudEvents_subject: "device-1", cloudEvents_moduleid: malformed }
      : { cloudEvents_deviceid: "device-1", cloudEvents_moduleid: malformed };
    assert.ok(bindNativeMetadata(headers, user, stream).issues.length);
    assert.equal(projectedNative(stream, headers, user).deviceId, "");
  }
  for (const malformed of [8, false, {}, []]) {
    const user = { cloudEvents_subject: "device-1", cloudEvents_iothubdtsubject: malformed };
    assert.ok(bindNativeMetadata(headers, user, "telemetry").issues.length);
    assert.equal(projectedNative("telemetry", headers, user).deviceId, "");
  }
  for (const stream of ["telemetry", "twin"] as const) for (const identity of [undefined, null, "", 0, false, {}, []]) {
    const key = stream === "telemetry" ? "cloudEvents_subject" : "cloudEvents_deviceid";
    const user: JsonObject = { cloudEvents_deviceref: "urn:untrusted:device", entity_id: "not-native",
      ...(identity !== undefined ? { [key]: identity } : {}) };
    assert.equal(bindNativeMetadata(headers, user, stream).identity, undefined);
    assert.equal(projectedNative(stream, headers, user).deviceId, "");
  }
  assert.ok(!nativeProjection("telemetry").join("\n").includes("deviceref"));
  assert.ok(!nativeProjection("twin").join("\n").includes("entity_id"));
});

function nestedRequest(component = false): CreationRequest {
  const request = creationRequest();
  const model = fromThingModel(resolveThingModel({
    "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel", id: "urn:nested:model",
    properties: { configuration: { type: "object", properties: { setpoint: { type: "number" }, mode: { type: "string" } } } },
  }));
  request.input.model = model;
  request.input.fields = model.fields.map(field => ({
    ...setFieldRole(field, "twin"), ...(component ? { component: "thermostat" } : {}),
  }));
  request.coverage.fingerprint = modelingFingerprint(request.input);
  return request;
}

function replayEmitted(request: CreationRequest, patches: { payload: JsonObject; opType?: string; deviceId?: string; time?: number; ingestedAt?: number }[],
  forDevice = "device-1") {
  const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
  const command = plan.commands.find(command => command.kind === "policy" && command.target === request.names.twinChanges)!;
  if (command.kind !== "policy") throw new Error("Missing twin normalizer.");
  const lines = command.query.split("\n"), position = lines.indexOf("| extend rp = payload");
  const rawScope = lines.indexOf("| extend data = parse_json(tostring(data))");
  assert.ok(rawScope >= 0 && lines[rawScope + 1].startsWith("| where "));
  const scope = lines[rawScope + 1].slice("| where ".length);
  assert.equal(lines[position + 1], "| extend keys = bag_keys(rp)");
  assert.equal(lines[position + 2], "| mv-expand key = keys to typeof(string)");
  assert.equal(lines[position + 3], '| where key !startswith "$" and key != "iothub-enqueuedtime" and key != "iothub-connection-device-id"');
  const componentLine = lines.find(line => line.startsWith("| extend isComponent = "));
  if (componentLine) assert.equal(lines.slice(position + 4).join("\n"), [
    "| extend val = rp[key]",
    '| extend isComponent = (gettype(val) == "dictionary" and tostring(val["__t"]) == "c")',
    "| mv-expand subkey = iff(isComponent, bag_keys(val), pack_array(key)) to typeof(string)",
    '| where subkey != "__t"',
    '| project deviceId, propertyName = iff(isComponent, strcat(key, "_", subkey), key), propertyValue = iff(isComponent, val[subkey], val), event_time',
  ].join("\n"));
  const project = lines.at(-1)!;
  const nameExpression = project.split("propertyName = ")[1].split(", propertyValue = ")[0];
  const valueExpression = project.split("propertyValue = ")[1].replace(/, event_time$/, "");
  const view = plan.commands.find(command => command.kind === "view")!;
  if (view.kind !== "view") throw new Error("Missing state view.");
  assert.match(view.query, /summarize arg_max\(event_time, propertyValue\) by deviceId, propertyName$/);
  assert.match(view.query, /\| where ingestion_time\(\) > ago\(30d\)/);
  const state = new Map<string, { deviceId: string; propertyName: string; time: number; value: unknown }>();
  const day = 86400000, now = 40 * day;
  for (const [index, patch] of patches.entries()) {
    if ((patch.ingestedAt ?? now) <= now - 30 * day) continue;
    const bindings = { data: { payload: { value: { properties: { reported: patch.payload } }, memberTypeIRI: request.input.model.id } },
      payload: patch.payload, user_headers: { opType: patch.opType || "updateTwin" },
      deviceId: patch.deviceId || "device-1", memberTypeIRI: request.input.model.id };
    if (evaluate(scope, bindings) !== true) continue;
    for (const key of Object.keys(patch.payload).filter(key => !key.startsWith("$") &&
      !["iothub-enqueuedtime", "iothub-connection-device-id"].includes(key))) {
      const item: Record<string, unknown> = { ...bindings, rp: patch.payload, key };
      let subkeys: unknown[] = [key];
      if (componentLine) {
        item.val = evaluate("rp[key]", item);
        item.isComponent = evaluate(componentLine.slice("| extend isComponent = ".length), item);
        const expression = lines.find(line => line.startsWith("| mv-expand subkey = "))!
          .slice("| mv-expand subkey = ".length).replace(/ to typeof\(string\)$/, "");
        const values = evaluate(expression, item);
        assert.ok(Array.isArray(values)); subkeys = values;
      }
      for (const subkey of subkeys) {
        if (componentLine && subkey === "__t") continue;
        item.subkey = subkey;
        const name = evaluate(nameExpression, item);
        assert.equal(typeof name, "string");
        const group = JSON.stringify([bindings.deviceId, name]);
        const time = patch.time ?? index;
        if (!state.has(group) || state.get(group)!.time < time) {
          state.set(group, { deviceId: bindings.deviceId, propertyName: String(name), time, value: evaluate(valueExpression, item) });
        }
      }
    }
  }
  return new Map([...state.values()].filter(row => row.deviceId === forDevice).map(row => [row.propertyName, row.value]));
}

function replayModeledProjection(request: CreationRequest, state: Map<string, unknown>, payload: JsonObject, component = "") {
  const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
  const policy = plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
  if (policy.kind !== "policy") throw new Error("Missing modeled policy.");
  const line = policy.query.split("\n").at(-1)!;
  assert.match(line, /^\| project deviceId, event_time, /);
  const output = Object.fromEntries([...line.matchAll(/\['([^']+)'\]=(.+?)(?=, \['|$)/g)].map(match =>
    [match[1], evaluate(match[2], { __state: Object.fromEntries(state), payload, component })]));
  assert.equal(Object.keys(output).length, request.input.fields.filter(field => field.included).length);
  return { output, query: policy.query, queries: plan.commands.flatMap(command =>
    command.kind === "policy" || command.kind === "view" ? [{ kind: command.kind, target: command.target, query: command.query }] : []) };
}

test("mixed DTDL emitted projections retain exact component LKV keys and whole object array values without cross-module enrichment", () => {
  const request = mixedDtdlRequest();
  request.input.fields = request.input.fields.filter(field => field.kustoType === "dynamic");
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const patches = [
    { payload: mixedReportedValues, deviceId: "device-1", time: 1 },
    { payload: { configuration: { value: 99 }, motor: { __t: "c", configuration: null } }, deviceId: "device-1/module-1", time: 2 },
  ];
  const state = replayEmitted(request, patches), moduleState = replayEmitted(request, patches, "device-1/module-1");
  assert.deepEqual(state.get("configuration"), mixedReportedValues.configuration);
  assert.deepEqual(state.get("motor_configuration"), mixedReportedValues.motor.configuration);
  assert.deepEqual(state.get("power_history"), mixedReportedValues.power.history);
  assert.equal(moduleState.get("motor_configuration"), null);
  assert.deepEqual(moduleState.get("configuration"), { value: 99 });
  assert.equal(moduleState.has("power_history"), false);
  assert.equal(state.has("motor_configuration_value"), false);
  for (const component of ["", "motor", "power", "unknown"]) {
    const modeled = replayModeledProjection(request, state, mixedTelemetryValues, component);
    for (const [prefix, expected] of [["", mixedReportedValues], ["motor_", mixedReportedValues.motor], ["power_", mixedReportedValues.power]] as const) {
      assert.deepEqual(modeled.output[`${prefix}configuration`], expected.configuration);
      assert.deepEqual(modeled.output[`${prefix}history`], expected.history);
      assert.deepEqual(modeled.output[`${prefix}measurement`], component === prefix.replace(/_$/, "") ? mixedTelemetryValues.measurement : null);
      assert.deepEqual(modeled.output[`${prefix}samples`], component === prefix.replace(/_$/, "") ? mixedTelemetryValues.samples : null);
    }
  }
  propagationOutcomes.push({ name: "mixed DTDL native LKV keys/whole arrays/objects/module partitions", input: patches,
    stateKeys: [...state.keys()], moduleStateKeys: [...moduleState.keys()], passed: true });
});

test("mixed DTDL actual emitted native projection agrees with decoded device module and component identities", () => {
  for (const component of ["", "motor", "power", "unknown"]) for (const module of ["", "module-1"]) {
    for (const stream of ["telemetry", "twin"] as const) {
      const row = mixedRawSample(stream, component, module);
      const decoded = decodeDoSamples([row], stream, 100).payloads[0];
      const projected = projectedNative(stream, row.headers, row.user_headers);
      assert.equal(projected.deviceId, decoded.native!.identity);
      assert.equal(projected.component, stream === "telemetry" ? component : "");
    }
  }
  const row = mixedRawSample("telemetry", "motor", "module-1");
  for (const user of [{ cloudEvents_subject: "other/module-1" },
    { cloudEvents_subject: "fixture-device/module-1", cloudEvents_iothubdtsubject: "power" }]) {
    assert.equal(projectedNative("telemetry", row.headers, user).deviceId, "");
    assert.equal(bindNativeMetadata(row.headers, user, "telemetry").identity, undefined);
  }
  propagationOutcomes.push({ name: "mixed DTDL emitted native/decoded parity", bindingPairs: 16, conflictPairs: 2, passed: true });
});

test("propagation replay preserves whole property replacement, explicit null, absent properties and out-of-order values through the emitted modeled lookup", () => {
  const request = nestedRequest();
  const cases: { name: string; patches: Parameters<typeof replayEmitted>[1]; expected: unknown }[] = [
    { name: "whole replacement", patches: [{ payload: { configuration: { setpoint: 1, mode: "auto" } }, time: 1 },
      { payload: { configuration: { setpoint: 2 } }, time: 2 }], expected: { setpoint: 2 } },
    { name: "explicit null", patches: [{ payload: { configuration: { setpoint: 1 } }, time: 1 },
      { payload: { configuration: null }, time: 2 }], expected: null },
    { name: "unmentioned property", patches: [{ payload: { configuration: { mode: "auto" } }, time: 1 },
      { payload: { unrelated: false }, time: 2 }], expected: { mode: "auto" } },
    { name: "out-of-order event", patches: [{ payload: { configuration: { setpoint: 3 } }, time: 3 },
      { payload: { configuration: { setpoint: 1 } }, time: 1 }], expected: { setpoint: 3 } },
    { name: "replaceTwin remains whole", patches: [{ payload: { configuration: { setpoint: 1, mode: "auto" } }, time: 1 },
      { payload: { configuration: { mode: "manual" } }, opType: "replaceTwin", time: 2 }], expected: { mode: "manual" } },
  ];
  for (const scenario of cases) {
    const original = canonicalJson(scenario.patches), state = replayEmitted(request, scenario.patches);
    assert.equal(state.has("configuration"), true, scenario.name);
    assert.deepEqual(state.get("configuration"), scenario.expected, scenario.name);
    const modeled = replayModeledProjection(request, state, {});
    assert.deepEqual(modeled.output.configuration, scenario.expected, scenario.name);
    assert.equal(canonicalJson(scenario.patches), original, "The replay must not merge/mutate received whole values.");
    propagationOutcomes.push({ name: scenario.name, input: scenario.patches, expected: scenario.expected,
      actual: modeled.output.configuration, passed: true, emittedQueries: modeled.queries });
  }
});

test("propagation replay applies the exact30-day ingestion cutoff and isolates native device/module state partitions", () => {
  const request = nestedRequest(), day = 86400000;
  const patches = [
    { payload: { configuration: { label: "expired-high-event-time" } }, deviceId: "device-1", time: 99, ingestedAt: 10 * day },
    { payload: { configuration: { label: "recent-low-event-time" } }, deviceId: "device-1", time: 1, ingestedAt: 10 * day + 1 },
    { payload: { configuration: { label: "second-device" } }, deviceId: "device-2", time: 2 },
    { payload: { configuration: { label: "module" } }, deviceId: "device-1/module-1", time: 3 },
  ];
  for (const [deviceId, expected] of [["device-1", "recent-low-event-time"], ["device-2", "second-device"], ["device-1/module-1", "module"]]) {
    const state = replayEmitted(request, patches, deviceId);
    assert.deepEqual(state.get("configuration"), { label: expected });
    const modeled = replayModeledProjection(request, state, {});
    assert.deepEqual(modeled.output.configuration, { label: expected });
    propagationOutcomes.push({ name: `30-day cutoff / partition ${deviceId}`, input: patches, selectedDevice: deviceId,
      expected: { label: expected }, actual: modeled.output.configuration, passed: true, emittedQueries: modeled.queries });
  }
});

test("propagation replay retains native component keys, whole nested values and intended twin scope before enriching telemetry", () => {
  const request = nestedRequest(true);
  request.input.scope.twin = 'user_headers.opType != "replaceTwin"';
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const patches = [
    { payload: { configuration: { mode: "root" }, thermostat: { __t: "c", configuration: { setpoint: 1, mode: "old" } } }, time: 1 },
    { payload: { thermostat: { __t: "c", configuration: { mode: "new" } } }, time: 2 },
    { payload: { thermostat: { __t: "c", configuration: { mode: "excluded" } } }, time: 3, opType: "replaceTwin" },
  ];
  const state = replayEmitted(request, patches);
  assert.deepEqual(state.get("configuration"), { mode: "root" });
  assert.deepEqual(state.get("thermostat_configuration"), { mode: "new" });
  const modeled = replayModeledProjection(request, state, {}, "thermostat");
  assert.deepEqual(modeled.output.configuration, { mode: "new" });
  assert.match(modeled.query, /lookup kind=leftouter .* on deviceId/);
  assert.ok(!/asof|event_time\s*<=/.test(modeled.query), "Last-known enrichment must not become event-time reconstruction.");
  propagationOutcomes.push({ name: "native components and intended twin scope", input: patches, expected: { mode: "new" },
    actual: modeled.output.configuration, retainedRootValue: state.get("configuration"), passed: true,
    fixtureBoundary: "Existing explicit native-component binding fixture; no inference that nested WoT object fields are components.",
    scope: request.input.scope, emittedQueries: modeled.queries });
});

test("reference-scoped propagation excludes other-model property rows instead of silently enriching across model references", () => {
  const request = nestedRequest();
  request.input.scope = { mode: "reference", operator: "and", reference: { id: "urn:copilot:test:other-model" },
    telemetry: "", twin: "" };
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const patches = [{ payload: { configuration: { setpoint: 42 } }, time: 1 }];
  const state = replayEmitted(request, patches);
  assert.equal(state.size, 0);
  const modeled = replayModeledProjection(request, state, {});
  assert.equal(modeled.output.configuration, null);
  assert.ok(modeled.query.includes('tostring(data["payload"]["memberTypeIRI"]) == "urn:copilot:test:other-model"'));
  propagationOutcomes.push({ name: "other-model property scope", input: patches, expected: null, actual: modeled.output.configuration,
    scope: request.input.scope, passed: true, emittedQueries: modeled.queries });
});

test("inherited meaning changes creation credit without replacing root-declared metadata or original model JSON", () => {
  const root = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel", id: "urn:root",
    links: [{ rel: "tm:extends", href: "urn:base" }] };
  const original = JSON.stringify(root);
  const requests = ["Original", "Changed"].map(meaning => {
    const request = creationRequest();
    request.input.model = fromThingModel(resolveThingModel(root, [{ fileName: "base.json", document: {
      "@context": root["@context"], "@type": "tm:ThingModel", id: "urn:base", title: meaning, description: meaning,
      properties: { temperature: { type: "number" } },
    } }]));
    request.input.fields = request.input.model.fields.map(field => setFieldRole(field, "telemetry"));
    return request;
  });
  assert.deepEqual(requests[0].input.fields, requests[1].input.fields);
  assert.equal(JSON.stringify(requests[0].input.model.document), original);
  assert.equal(JSON.stringify(requests[1].input.model.document), original);
  assert.equal(semanticDocumentation(requests[0].input.model, requests[0].input.fields).table,
    semanticDocumentation(requests[1].input.model, requests[1].input.fields).table);
  assert.equal(requests[0].input.model.declaredModelMetadata?.id, "urn:root");
  assert.equal(requests[0].input.model.declaredModelMetadata?.description, undefined);
  assert.notEqual(creationFingerprint(requests[0]), creationFingerprint(requests[1]));
  assert.equal(creationFingerprint(requests[0]), creationFingerprint({ ...requests[0],
    names: { twinState: requests[0].names.twinState, modeled: requests[0].names.modeled, twinChanges: requests[0].names.twinChanges } }));
});

test("V1 property LKV retains the latest whole received value, not reconstructed nested state", () => {
  const request = nestedRequest();
  assert.equal(request.input.fields.length, 1);
  assert.equal(request.input.fields[0].kustoType, "dynamic");
  const first = { payload: { configuration: { setpoint: 20, mode: "auto" }, unselected: [1, 2] }, time: 1 };
  const partial = replayEmitted(request, [first, { payload: { configuration: { setpoint: 21 } }, time: 3 },
    { payload: { configuration: { mode: "older" } }, time: 2 }]);
  assert.deepEqual(partial.get("configuration"), { setpoint: 21 });
  assert.deepEqual(partial.get("unselected"), [1, 2]);
  const omitted = replayEmitted(request, [first, { payload: { unrelated: true }, opType: "replaceTwin" }]);
  assert.deepEqual(omitted.get("configuration"), first.payload.configuration);
  assert.equal(replayEmitted(request, [first, { payload: { configuration: null }, time: 2 }]).get("configuration"), null);
  const day = 86400000;
  assert.equal(replayEmitted(request, [{ ...first, ingestedAt: 10 * day }]).size, 0);
  assert.deepEqual(replayEmitted(request, [{ ...first, ingestedAt: 10 * day + 1 }]).get("configuration"), first.payload.configuration);
});

test("V1 component envelopes separate capability names without unwrapping complex property values", () => {
  const request = nestedRequest(true);
  const root = { configuration: { setpoint: 1, mode: "root" },
    thermostat: { __t: "c", configuration: { setpoint: 20, mode: "auto", nested: { __t: "c", value: [1] } } } };
  const latest = { thermostat: { __t: "c", configuration: { setpoint: 21 } } };
  const state = replayEmitted(request, [{ payload: root }, { payload: latest }]);
  assert.deepEqual(state.get("configuration"), root.configuration);
  assert.deepEqual(state.get("thermostat_configuration"), { setpoint: 21 });
  assert.deepEqual(replayEmitted(request, [{ payload: root }]).get("thermostat_configuration"), root.thermostat.configuration);
  const row = decodeDoSamples([{ data: JSON.stringify({ schema_id: "com.do.telemetry.observation.v1",
    schema_version: "1.0", contract_version: "1.0", event_kind: "lifecycle", event_time: "2026-10-01",
    payload: { value: { properties: { reported: root } } } }), headers: {}, user_headers: {} }], "twin", 100).payloads[0];
  assert.deepEqual(fieldValue(row, request.input.fields[0], "twin"), { found: true, value: root.thermostat.configuration });
});

test("V1 intended twin scope still excludes otherwise newer whole values", () => {
  const request = nestedRequest();
  request.input.scope.twin = 'deviceId == "device-1"';
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const field = request.input.fields[0];
  const state = replayEmitted(request, [{ payload: { configuration: { setpoint: 20 } } },
    { payload: { configuration: null }, deviceId: "outside-scope" }]);
  assert.deepEqual(state.get(field.path[0]), { setpoint: 20 });
});

test("V1 declared Object Map Array values are accepted whole dynamic properties through checked creation", async () => {
  for (const schema of [
    { "@type": "Object", fields: [{ name: "setpoint", schema: "double" }, { name: "mode", schema: "string" }] },
    { "@type": "Map", mapKey: { name: "key", schema: "string" }, mapValue: { name: "value", schema: "double" } },
    { "@type": "Array", elementSchema: "double" },
  ]) {
    const document = { "@context": "dtmi:dtdl:context;2", "@id": "dtmi:correction:State;1", "@type": "Interface",
      contents: [{ "@type": "Property", name: "configuration", schema }] };
    const request = creationRequest();
    request.input.model = fromDtdlModel(parseDtdlModel(document), document);
    request.input.fields = request.input.model.fields;
    request.coverage.fingerprint = modelingFingerprint(request.input);
    const mock = createCreationMock(request);
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "completed");
    assert.deepEqual(mock.tables.get(request.names.modeled)!.columns.map(column => [column.name, column.type]),
      [["deviceId", "string"], ["event_time", "datetime"], ["configuration", "dynamic"]]);
    assert.equal(request.input.fields[0].included, true);
    assert.equal(request.input.fields[0].role, "twin");
  }
});

test("V1 whole complex telemetry projects unchanged with no child columns and retains nested model facts", () => {
  const schemas: JsonObject[] = [
    { type: "object", description: "Whole configuration.", properties: { mode: { type: "string", enum: ["auto", "manual"] } } },
    { type: "object", additionalProperties: { type: "number", unit: "degC" } },
    { type: "array", items: { type: "object", properties: { value: { type: "number" } } } },
  ];
  const values = [{ mode: "auto", nested: { untouched: [1, null, { value: 2 }] } }, { arbitrary: 1, another: 2 }, [{ value: 3 }, { value: 4 }]];
  for (const [index, schema] of schemas.entries()) {
    const request = creationRequest();
    const document = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
      id: "urn:fixture:whole-telemetry", properties: { measurement: schema } };
    request.input.model = fromThingModel(resolveThingModel(document));
    request.input.fields = request.input.model.fields.map(field => setFieldRole(field, "telemetry"));
    request.coverage.fingerprint = modelingFingerprint(request.input);
    const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
    const policy = plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
    if (policy.kind !== "policy") throw new Error("Missing modeled policy.");
    const expression = policy.query.split("\n").at(-1)!.split("['measurement']=")[1];
    const value = values[index];
    assert.deepEqual(evaluate(expression, { payload: { measurement: value }, component: "" }), value);
    assert.equal(request.input.fields.length, 1); assert.equal(request.input.fields[0].kustoType, "dynamic");
    assert.deepEqual(request.input.fields[0].schema, schema);
    assert.deepEqual(request.input.model.document, document);
    assert.deepEqual(plan.dataAgent.examples.map(example => example.question),
      ["What devices are in the system?", "Show me the latest modeled records for a device (replace YOUR_DEVICE_ID)",
        "Show the latest non-null measurement for each device", "How many readings per device in the last 24 hours?"]);
    assert.ok(plan.dataAgent.sourceInstructions.includes("`measurement` (dynamic)"));
    assert.ok(!plan.documentation.columns.mode);
    const output = plan.commands.find(command => command.kind === "table" && command.target === request.names.modeled)!;
    if (output.kind !== "table") throw new Error("Missing native modeled columns.");
    const datasource = buildDataAgentDataSource({ database: "Database", table: request.names.modeled,
      docstring: plan.documentation.table, columns: output.columns }, "workspace", "database", "table");
    assert.deepEqual(datasource.elements[0].children[0].children.map(column => [column.id, column.data_type, column.children]),
      [["deviceId", "System.String", []], ["event_time", "System.DateTime", []], ["measurement", "System.Object", []]]);
  }
});

test("V1 topology EAV columns property keys 30-day LKV column order and policy flags match the baseline contract", () => {
  for (const roles of [["telemetry", "telemetry"], ["twin", "twin"], ["twin", "telemetry"]] as const) {
    const request = creationRequest();
    request.input.fields = request.input.model.fields.map((field, index) => setFieldRole(field, roles[index]));
    request.coverage.fingerprint = modelingFingerprint(request.input);
    const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
    const objects = plan.commands.filter(command => command.kind === "table" || command.kind === "view");
    assert.deepEqual(objects.map(command => [command.kind, command.target]),
      [["table", request.names.twinChanges], ["view", request.names.twinState], ["table", request.names.modeled]]);
    assert.deepEqual(objects[0].kind === "table" ? objects[0].columns : [], [
      { name: "deviceId", type: "string" }, { name: "propertyName", type: "string" },
      { name: "propertyValue", type: "dynamic" }, { name: "event_time", type: "datetime" },
    ]);
    assert.equal(objects[1].kind === "view" ? objects[1].query : "",
      `['${request.names.twinChanges}']\n| where ingestion_time() > ago(30d)\n| summarize arg_max(event_time, propertyValue) by deviceId, propertyName`);
    const expected = [...request.input.fields.filter(field => field.role === "telemetry"), ...request.input.fields.filter(field => field.role === "twin")];
    assert.deepEqual(objects[2].kind === "table" ? objects[2].columns : [], [
      { name: "deviceId", type: "string" }, { name: "event_time", type: "datetime" },
      ...expected.map(field => ({ name: field.column, type: field.kustoType })),
    ]);
    const policies = plan.commands.filter(command => command.kind === "policy");
    assert.equal(policies.length, 2);
    for (const command of policies) {
      const actual = JSON.parse(JSON.parse(command.csl.split(" policy update ")[1]));
      assert.deepEqual(actual, [{ IsEnabled: true, Source: command.source, Query: command.query,
        IsTransactional: false, PropagateIngestionProperties: false }]);
    }
    const normalizer = policies.find(command => command.target === request.names.twinChanges)!;
    assert.equal(normalizer.query.split("| extend rp = payload\n")[1], [
      "| extend keys = bag_keys(rp)", "| mv-expand key = keys to typeof(string)",
      '| where key !startswith "$" and key != "iothub-enqueuedtime" and key != "iothub-connection-device-id"',
      "| project deviceId, propertyName = key, propertyValue = rp[key], event_time",
    ].join("\n"));
    assert.ok(!normalizer.query.includes("/properties/") && !normalizer.query.includes("__selected") && !normalizer.query.includes("replaceTwin"));
    assert.deepEqual(plan.commands.filter(command =>
      command.kind === "tableDoc" || command.kind === "viewDoc" || command.kind === "columnDocs")
      .map(command => [command.kind, command.target]).sort(), [
      ["tableDoc", request.names.twinChanges], ["viewDoc", request.names.twinState],
      ["tableDoc", request.names.modeled], ["columnDocs", request.names.modeled],
    ].sort());
  }
});

test("whole-event binding uses payload.alarm and retains one schema without flattening or inferring a role", () => {
  const data: JsonObject = { type: "object", properties: {
    code: { type: "integer", enum: [1, 2] }, details: { type: "object", properties: { source: { type: "string" } } },
  } };
  const document = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
    id: "urn:fixture:whole-alarm",
    events: { alarm: { title: "Alarm", description: "Whole alarm value.", sourceKind: "telemetry", data } } };
  const model = fromThingModel(resolveThingModel(document));
  assert.equal(model.fields.length, 1);
  const field = model.fields[0];
  assert.deepEqual(field.path, ["alarm"]); assert.equal(field.column, "alarm");
  assert.equal(field.kustoType, "dynamic"); assert.equal(field.role, undefined); assert.equal(field.component, undefined);
  assert.deepEqual(field.schema, data); assert.deepEqual(model.document, document);
  const value = { code: 1, details: { source: "device" } };
  const sample = { value: { alarm: value, code: 99 }, native: { profile: "eventhub" as const, identity: "device-1", issues: [] } };
  assert.deepEqual(fieldValue(sample, field, "telemetry"), { found: true, value });
  assert.deepEqual(inspectFieldEvidence([field], { rawRows: 1, observations: 1, skippedWithoutReportedProperties: 0, payloads: [sample] },
    { rawRows: 0, observations: 0, skippedWithoutReportedProperties: 0, payloads: [] })[0].state, "telemetry");
  assert.deepEqual(omittedObservedFields([setFieldRole(field, "telemetry")], {
    telemetry: { rawRows: 1, observations: 1, skippedWithoutReportedProperties: 0, payloads: [sample] },
    twin: { rawRows: 0, observations: 0, skippedWithoutReportedProperties: 0, payloads: [] },
  }).telemetry, ["/code"]);
  const request = creationRequest();
  request.input.model = model; request.input.fields = [setFieldRole(field, "telemetry")];
  request.coverage.fingerprint = modelingFingerprint(request.input);
  const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
  const command = plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
  if (command.kind !== "policy") throw new Error("Missing event projection.");
  assert.deepEqual(evaluate(command.query.split("\n").at(-1)!.split("['alarm']=")[1], { payload: sample.value, component: "" }), value);
  const collision = fromThingModel(resolveThingModel({ ...document, properties: { alarm: { type: "number" } } }));
  assert.ok(fieldValidation(collision.fields.map(item => setFieldRole(item, "telemetry"))).some(error => error.includes('map to "alarm"')));
});

test("Thing Model properties and events have equal role authority, whole values in either stream, and no action fields", () => {
  const schema: JsonObject = { type: "object", properties: { nested: { type: "object", properties: { value: { type: "number" } } } } };
  const document = { "@context": "https://www.w3.org/2022/wot/td/v1.1", "@type": "tm:ThingModel",
    id: "urn:fixture:equal-affordance-roles",
    properties: { measurement: { ...schema, readOnly: true, observable: true, sourceKind: "property" } },
    events: { alarm: { sourceKind: "telemetry", data: schema } },
    actions: { reboot: { input: { type: "string" }, output: { type: "boolean" } } } };
  const model = fromThingModel(resolveThingModel(document));
  assert.deepEqual(model.fields.map(field => [field.column, field.role]), [["measurement", undefined], ["alarm", undefined]]);
  assert.deepEqual(model.document, document);
  const telemetry = { rawRows: 1, observations: 1, skippedWithoutReportedProperties: 0,
    payloads: [{ value: { measurement: { nested: { value: 4 } } }, native: { profile: "eventhub" as const, identity: "d1", issues: [] } }] };
  const twin = { rawRows: 1, observations: 1, skippedWithoutReportedProperties: 0,
    payloads: [{ value: { alarm: { nested: { value: 2 } } } }] };
  const evidence = inspectFieldEvidence(model.fields, telemetry, twin);
  const hinted = applyRoleHints(model.fields, evidence);
  assert.deepEqual(hinted.map(field => [field.column, field.role, field.roleSource]),
    [["measurement", "telemetry", "hint"], ["alarm", "twin", "hint"]]);
  const request = creationRequest();
  request.input.model = model; request.input.fields = hinted;
  request.coverage = { ...request.coverage, telemetry, twin, fields: evidence, fingerprint: modelingFingerprint(request.input) };
  const plan = buildModeledDataPlan(request.input, request.coverage, false, request.names);
  const modeled = plan.commands.find(command => command.kind === "table" && command.target === request.names.modeled)!;
  assert.deepEqual(modeled.kind === "table" ? modeled.columns.map(column => [column.name, column.type]) : [],
    [["deviceId", "string"], ["event_time", "datetime"], ["measurement", "dynamic"], ["alarm", "dynamic"]]);
  const policy = plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
  assert.ok(policy.kind === "policy" && policy.query.includes('__state["alarm"]'));
  assert.ok(!JSON.stringify(plan.commands).includes("reboot"));
  const manual = hinted.map(field => setFieldRole(field, field.column === "alarm" ? "telemetry" : "twin"));
  assert.deepEqual(applyRoleHints(manual, evidence).map(field => field.role), ["twin", "telemetry"]);
  assert.deepEqual(model.fields.map(field => field.schema), [document.properties.measurement, schema]);
});

test("V1 architecture correction invalidates old scalar-only coverage and completion identity", () => {
  const request = creationRequest();
  const old = canonicalJson({ source: request.input.source, model: request.input.model, fields: request.input.fields,
    scope: request.input.scope, limit: request.input.limit });
  assert.notEqual(old, modelingFingerprint(request.input));
  assert.equal(coverageReady(request.input, { ...request.coverage, fingerprint: old }, false), false);
  assert.notEqual(canonicalJson([old, request.names, request.allowWaiting]), creationFingerprint(request));
});

test("every incomplete attempt retains names/database/confirmed and uncertain effects through later blocked checks", async () => {
  const request = creationRequest(), backend = createCreationMock(request);
  backend.options.documentationMismatch = true;
  const first = await new ModeledDataCreator().run(request, backend.service, true);
  assert.equal(first.phase, "failed"); assert.ok(first.objects.length);
  const secondRequest = { ...request, names: { modeled: "Output_B", twinChanges: "Changes_B", twinState: "State_B" } };
  const secondBackend = createCreationMock(secondRequest);
  secondBackend.options.fail = csl => csl.startsWith(".set ") ? new Error("Unknown outcome") : undefined;
  const second = await new ModeledDataCreator().run(secondRequest, secondBackend.service, true);
  assert.equal(second.phase, "failed"); assert.equal(second.objects.length, 0); assert.ok(second.uncertainCommand);
  const history = rememberCreationAttempt(rememberCreationAttempt([], first), second);
  const blocked = await new ModeledDataCreator().run(request, backend.service, false);
  assert.equal(blocked.phase, "blocked");
  assert.deepEqual(rememberCreationAttempt(history, blocked), history);
  assert.equal(history.length, 2);
  assert.equal(history[0].names.modeled, request.names.modeled);
  assert.equal(history[1].names.modeled, secondRequest.names.modeled);
  assert.equal(history[1].databaseName, request.input.source.databaseName);
  assert.equal(history[1].uncertainCommand, second.uncertainCommand);
  assert.ok(!("coverage" in history[0]));
});
