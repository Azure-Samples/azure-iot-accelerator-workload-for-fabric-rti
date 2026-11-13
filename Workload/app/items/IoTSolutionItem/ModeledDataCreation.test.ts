// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { assertOrderedSchema, creationFingerprint, ModeledDataCreator, rememberCreationAttempt } from "./ModeledDataCreation.ts";
import { executeKqlRequest, parseKqlTableResult } from "./ModelAuthoring.ts";
import { modelingFingerprint, nativeIdentity } from "./HubModeling.ts";
import { createCreationMock, creationRequest, mockResult } from "../../../tests/creationMock.ts";

test("plain helper object docs are real owned writes with readback while only ModeledData receives column semantics", async () => {
  const request = creationRequest(), mock = createCreationMock(request);
  const priorCredit = JSON.parse(creationFingerprint(request)).slice(0, 3);
  assert.notEqual(creationFingerprint(request), JSON.stringify(priorCredit));
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "completed");
  assert.equal(mock.writes.length, 9);
  assert.equal(result.plannedOutputs?.length, 3);
  assert.match(mock.tables.get(request.names.twinChanges)!.doc, /^Normalized device property changes in Entity-Attribute-Value \(EAV\) format/);
  assert.match(mock.views.get(request.names.twinState)!.doc, /^Property last known values/);
  assert.deepEqual(mock.tables.get(request.names.twinChanges)!.columnDocs, {});
  assert.equal(mock.writes.filter(command => command.includes("column-docstrings")).length, 1);
  assert.ok(mock.writes.find(command => command.includes("column-docstrings"))!.includes(request.names.modeled));
  assert.ok(mock.calls.some(call => call.csl.startsWith(".show materialized-view ") && call.csl.includes(request.names.twinState)));
});
test("helper object docstring readback failure cannot claim new creation completion", async () => {
  const request = creationRequest(), mock = createCreationMock(request);
  mock.options.before = command => {
    if (command.startsWith(".alter table ['Modeled_Output'] policy update")) mock.tables.get(request.names.twinChanges)!.doc = "External change";
  };
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "failed");
  assert.match(result.error!, /Object docstring for Changes_Output could not be verified/);
});

test("creation checks are read-only and successful creation requires definition/policy/doc readback", async () => {
  const request = creationRequest(), mock = createCreationMock(request), creator = new ModeledDataCreator();
  const checked = await creator.run(request, mock.service, false);
  assert.equal(checked.phase, "ready"); assert.equal(mock.writes.length, 0);
  const result = await creator.run(request, mock.service, true);
  assert.equal(result.phase, "completed"); assert.equal(result.objects.length, 3);
  assert.deepEqual(Object.keys(mock.service).sort(), ["inspect", "request"]);
  assert.ok([...request.coverage.telemetry.payloads, ...request.coverage.twin.payloads]
    .every(sample => !sample.userHeaders?.hubName));
  assert.ok(result.steps.every(step => step.status === "passed"));
  assert.ok(mock.calls.some(call => call.kind === "query" && call.csl.includes("['State_Output']") && call.csl.endsWith("| getschema")));
  const modeled = mock.tables.get(request.names.modeled)!;
  assert.deepEqual(modeled.columns.map(column => column.name), ["deviceId", "event_time", "temperature", "state"]);
  assert.ok(modeled.doc.startsWith("Modeled device telemetry enriched with property LKVs"));
  assert.ok(modeled.columnDocs.state.includes("twin property (last-known whole received value)"));
  assert.ok(!mock.writes.some(csl => /(?:drop|set-or|policy streamingingestion|policy retention|backfill=true)/.test(csl)));
  assert.ok(mock.writes.filter(csl => csl.startsWith(".alter")).every(csl => !csl.includes("['Telemetry'] policy ") && !csl.includes("['Twins'] policy ")));
});

test("shared schema-preflight and installed policy use the same legal V1 let declaration and lookup reference", async () => {
  const request = creationRequest(), mock = createCreationMock(request), creator = new ModeledDataCreator();
  const checked = await creator.run(request, mock.service, false);
  assert.equal(checked.phase, "ready"); assert.equal(mock.writes.length, 0);
  const command = mock.plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
  if (command.kind !== "policy") throw new Error("Missing modeled policy.");
  assert.ok(command.query.startsWith(`let lkv_bag = ['${request.names.twinState}'] | summarize __state=make_bag(bag_pack(propertyName, propertyValue)) by deviceId;`));
  assert.ok(command.query.includes("| lookup kind=leftouter lkv_bag on deviceId"));
  const planned = command.query.replace(`['${request.names.twinState}']`,
    "(datatable(deviceId:string, propertyName:string, event_time:datetime, propertyValue:dynamic)[])");
  assert.ok(mock.calls.some(call => call.kind === "query" && call.csl === `${planned}\n| take 0\n| getschema`));
  for (const query of [command.query, planned]) {
    const bindings = [...query.matchAll(/\blet\s+([A-Za-z_]\w*)\s*=/g)].map(match => match[1]);
    assert.deepEqual(bindings, ["lkv_bag"]);
    assert.ok(bindings.every(name => !name.startsWith("__") && !name.endsWith("__")));
    assert.ok(!query.includes("__lastKnown"));
  }
  const created = await creator.run(request, mock.service, true);
  assert.equal(created.phase, "completed"); assert.equal(created.objects.length, 3);
  assert.deepEqual(mock.tables.get(request.names.modeled)!.policy, [{ IsEnabled: true,
    Source: request.input.source.telemetryTable, Query: command.query, IsTransactional: false, PropagateIngestionProperties: false }]);
});

test("query-local LKV alias never shadows selected raw or helper/output table names", () => {
  for (const target of ["telemetry", "properties", "normalized", "lkv", "modeled", "multiple"] as const) {
    const request = creationRequest();
    if (target === "telemetry" || target === "multiple") request.input.source.telemetryTable = "lkv_bag";
    if (target === "properties") request.input.source.twinTable = "lkv_bag";
    if (target === "normalized") request.names.twinChanges = "lkv_bag";
    if (target === "lkv") request.names.twinState = "lkv_bag";
    if (target === "modeled") request.names.modeled = "lkv_bag";
    if (target === "multiple") {
      request.input.source.twinTable = "lkv_bag_1"; request.names.twinState = "lkv_bag_2";
    }
    request.coverage.fingerprint = modelingFingerprint(request.input);
    const mock = createCreationMock(request);
    const policy = mock.plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
    if (policy.kind !== "policy") throw new Error("Missing modeled policy.");
    const alias = /^let ([A-Za-z_]\w*) = /.exec(policy.query)![1];
    assert.equal(alias, target === "multiple" ? "lkv_bag_3" : "lkv_bag_1");
    assert.ok(![request.input.source.telemetryTable, request.input.source.twinTable, ...Object.values(request.names)].includes(alias));
    assert.ok(policy.query.includes(`| lookup kind=leftouter ${alias} on deviceId`));
    assert.ok(policy.query.includes(`= ['${request.names.twinState}'] | summarize`));
    assert.ok(policy.query.includes(`\n['${request.input.source.telemetryTable}']\n`));
  }
});

test("ordinary HTTP400 native details reach the blocked schema-preflight result without any mutation", async () => {
  const request = creationRequest(), mock = createCreationMock(request);
  const activityId = "00000000-0000-4000-8000-000000000001";
  const result = await new ModeledDataCreator().run(request, {
    inspect: mock.service.inspect,
    request: (csl, kind) => executeKqlRequest(request.input.source, "mock-token", csl, kind, undefined, async (_url, init) => {
      assert.equal(new Headers(init?.headers).get("Accept"), "application/json");
      assert.deepEqual(JSON.parse(String(init?.body)), { db: request.input.source.databaseName, csl });
      if (csl.includes("let lkv_bag") && csl.endsWith("| getschema")) {
        return new Response(`General_BadRequest: Request is invalid and cannot be executed.\r\nError details:\r\nActivityId='${activityId}'.`,
          { status: 400 });
      }
      const value = await mock.service.request(csl, kind);
      return new Response(JSON.stringify({ Tables: [value] }));
    }),
  }, false);
  assert.equal(result.phase, "blocked"); assert.equal(mock.writes.length, 0);
  assert.equal(result.objects.length, 0); assert.equal(result.uncertainCommand, undefined);
  assert.ok(result.error?.includes("General_BadRequest") && result.error.includes(activityId));
  const failed = result.steps.find(step => step.status === "failed");
  assert.equal(failed?.label, `Validate ordered projection for ${request.names.modeled}`);
  assert.ok(failed?.detail?.includes("Request is invalid and cannot be executed."));
});

test("lookup trigger disabled with streaming-enabled properties passes read-only check and all owned-creation revalidation", async () => {
  const request = creationRequest(), mock = createCreationMock(request), creator = new ModeledDataCreator();
  mock.tables.get(request.input.source.twinTable)!.streaming = true;
  const modeledPolicy = mock.plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
  const propertiesPolicy = mock.plan.commands.find(command => command.kind === "policy" && command.target === request.names.twinChanges)!;
  assert.ok(modeledPolicy.kind === "policy" && modeledPolicy.source === request.input.source.telemetryTable &&
    modeledPolicy.query.includes("| lookup "));
  assert.ok(propertiesPolicy.kind === "policy" && propertiesPolicy.source === request.input.source.twinTable &&
    !/\|\s*(?:lookup|join)\b/.test(propertiesPolicy.query));
  const checked = await creator.run(request, mock.service, false);
  assert.equal(checked.phase, "ready"); assert.equal(mock.writes.length, 0);
  const policyReads = new Map<string, number>();
  for (const name of [request.input.source.telemetryTable, request.input.source.twinTable]) {
    assert.ok(mock.calls.some(call => call.csl === `.show table ['${name}'] schema as json`));
    policyReads.set(name, mock.calls.filter(call => call.csl === `.show table ['${name}'] policy update`).length);
    assert.ok(policyReads.get(name)! > 0);
  }
  const result = await creator.run(request, mock.service, true);
  assert.equal(result.phase, "completed"); assert.equal(result.objects.length, 3);
  for (const name of policyReads.keys()) {
    assert.ok(mock.calls.filter(call => call.csl === `.show table ['${name}'] policy update`).length > policyReads.get(name)! + 1,
      "The source inventory was skipped during mutation revalidation.");
  }
  assert.equal(mock.writes.filter(csl => csl.includes(" policy update ")).length, 2);
  assert.ok(!mock.writes.some(csl => csl.startsWith(".alter table ['Telemetry']") ||
    csl.startsWith(".alter table ['Twins']") || csl.includes("policy streamingingestion")));
  assert.equal(mock.tables.get(request.input.source.twinTable)!.streaming, true);
});

test("root output collisions and identical or different-schema creation races never authorize alterations", async () => {
  for (const mode of ["table-identical", "table-different", "function", "race-identical", "race-different"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    const target = mode.startsWith("race") ? request.names.twinChanges : request.names.modeled;
    const planned = mock.plan.commands.find(command => command.kind === "table" && command.target === target)!;
    if (planned.kind !== "table") throw new Error("Missing table.");
    const foreign = { ...mock.tables.get("Telemetry")!, id: "foreign", folder: "",
      columns: mode.endsWith("identical") ? planned.columns : [{ name: "foreign", type: "string" }] };
    if (mode.startsWith("table")) mock.tables.set(target, foreign);
    if (mode === "function") mock.functions.push(request.names.modeled);
    if (mode.startsWith("race")) mock.options.before = csl => {
      if (csl.startsWith(".set ['Changes_Output']")) mock.tables.set(target, foreign);
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.notEqual(result.phase, "completed");
    assert.equal(result.objects.length, 0);
    assert.ok(!mock.writes.some(csl => csl.startsWith(".alter")));
    assert.match(result.error!, /already exists/);
    if (mode !== "function") assert.deepEqual(mock.tables.get(target), foreign);
    if (mode.startsWith("race")) assert.ok(result.uncertainCommand);
    else assert.equal(mock.writes.length, 0);
  }
});

test("new outputs use exact fail-if-exists empty-table plans and default-root views without Folder writes", async () => {
  for (const twins of [true, false]) {
    const request = creationRequest(false, twins), mock = createCreationMock(request);
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "completed"); assert.equal(result.objects.length, 3);
    assert.equal(mock.writes.length, 9);
    assert.ok(mock.writes.every(csl => !/\bfolder\s*=|IoTAccelerator\/|ifnotexists|\.set-or-/.test(csl)));
    for (const command of mock.plan.commands) {
      assert.ok(mock.writes.includes(command.csl), "The creator changed the reviewed plan.");
      if (command.kind === "table") {
        assert.match(command.csl, /^\.set \['\w+'\] <\|\ndatatable\([\s\S]*\)\[\]$/);
        const table = mock.tables.get(command.target)!;
        assert.equal(table.folder, "");
        assert.deepEqual(table.columns, command.columns);
        assert.equal(result.objects.find(object => object.name === command.target)?.tableId, table.id);
      }
      if (command.kind === "view") {
        assert.match(command.csl, /^\.create materialized-view with \(backfill=false\)/);
        assert.equal(mock.views.get(command.target)!.folder, "");
      }
    }
  }
});

test("empty creation requires the native receipt; no-op partial or uncertain effects never claim ownership or retry", async () => {
  for (const variant of ["wrong-shape", "nonempty", "no-table", "partial", "transport-after-create"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    const original = mock.service.request;
    let attempts = 0;
    mock.service.request = async (csl, kind) => {
      if (!csl.startsWith(".set ")) return original(csl, kind);
      attempts++;
      if (variant !== "no-table") await original(csl, kind);
      if (variant === "wrong-shape") return mockResult(["TableName"], [[request.names.twinChanges]]);
      if (variant === "nonempty") return mockResult(["ExtentId", "RowCount"], [["unexpected-extent", 1]]);
      if (variant === "transport-after-create") throw new Error("Connection closed; outcome unknown");
      if (variant === "partial") return parseKqlTableResult({ Tables: [
        { TableName: "PrimaryResult", Columns: [{ ColumnName: "ExtentId" }, { ColumnName: "RowCount" }], Rows: [] },
        { TableName: "QueryStatus", Columns: [{ ColumnName: "Severity" }, { ColumnName: "StatusCode" }], Rows: [[2, -1]] },
      ] });
      return mockResult(["ExtentId", "RowCount"], []);
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed", variant);
    assert.equal(result.objects.length, 0);
    assert.ok(result.error && result.uncertainCommand);
    assert.equal(attempts, 1);
    assert.ok(!mock.writes.some(csl => /^\.alter|^\.drop/.test(csl)));
    assert.equal(rememberCreationAttempt([], result)[0].uncertainCommand, result.uncertainCommand);
    if (variant !== "no-table") {
      const later = await new ModeledDataCreator().run(request, mock.service, true);
      assert.equal(later.phase, "blocked"); assert.equal(later.objects.length, 0);
      assert.equal(attempts, 1, "A later attempt adopted or retried the uncertain table.");
    }
  }
});

test("identical root materialized-view collisions and races cannot authorize view documentation", async () => {
  for (const race of [false, true]) {
    const request = creationRequest(), mock = createCreationMock(request);
    const command = mock.plan.commands.find(command => command.kind === "view")!;
    if (command.kind !== "view") throw new Error("Missing view.");
    const foreign = { command, folder: "", doc: "Unrelated existing view" };
    if (race) mock.options.before = csl => {
      if (csl.startsWith(".create materialized-view")) mock.views.set(command.target, foreign);
    };
    else mock.views.set(command.target, foreign);
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, race ? "failed" : "blocked");
    assert.ok(!result.objects.some(object => object.kind === "view"));
    assert.ok(!mock.writes.some(csl => /^\.alter|^\.drop/.test(csl)));
    assert.deepEqual(mock.views.get(command.target), foreign);
  }
});

test("current root view definitions and helper TableId remain mandatory after fail-if-exists creation", async () => {
  for (const variant of ["query", "source", "enabled", "folder", "helper-id"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    const original = mock.service.request;
    mock.service.request = async (csl, kind) => {
      const response = await original(csl, kind);
      if (csl.startsWith(".show materialized-view ")) {
        if (variant === "helper-id") mock.tables.get(request.names.twinChanges)!.id = "foreign-replacement";
        else {
          const column = { query: "Query", source: "SourceTable", enabled: "IsEnabled", folder: "Folder" }[variant];
          response.Rows[0][response.Columns.findIndex(item => item.ColumnName === column)] =
            variant === "enabled" ? false : "Unexpected external value";
        }
      }
      return response;
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed", variant);
    assert.ok(!mock.writes.some(csl => csl.startsWith(".alter")));
    assert.match(result.error!, variant === "helper-id" ? /Ownership/ : /requested definition/);
  }
});

test("enabled unknown or malformed lookup-trigger streaming and upstream policies block without any write", async () => {
  for (const variant of ["enabled", "unknown", "malformed-json", "malformed-boolean", "cascade"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    mock.tables.get("Twins")!.streaming = true;
    const table = mock.tables.get("Telemetry")!;
    if (variant === "cascade") table.policy = [{ Source: "AnotherSource" }];
    else if (variant === "enabled" || variant === "unknown") table.streaming = variant === "enabled" ? true : null;
    else {
      const original = mock.service.request;
      mock.service.request = async (csl, kind) => {
        const result = await original(csl, kind);
        if (csl === ".show table ['Telemetry'] details") {
          result.Rows[0][result.Columns.findIndex(column => column.ColumnName === "StreamingIngestionPolicy")] =
            variant === "malformed-json" ? "{invalid" : '{"IsEnabled":"false"}';
        }
        return result;
      };
    }
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "blocked"); assert.equal(mock.writes.length, 0);
    assert.match(result.error!, variant === "cascade" ? /upstream/ : /streaming/);
  }
});

test("both raw schemas identities and nonempty or unresolved upstream inventories remain mandatory", async () => {
  for (const name of ["Telemetry", "Twins"]) {
    for (const variant of ["identity", "schema", "cascade", "unresolved"] as const) {
      const request = creationRequest(), mock = createCreationMock(request);
      mock.tables.get("Twins")!.streaming = true;
      const table = mock.tables.get(name)!;
      if (variant === "identity") table.id = "";
      if (variant === "schema") table.columns = table.columns.filter(column => column.name !== "data");
      if (variant === "cascade") table.policy = [{ Source: "UnreviewedUpstream" }];
      if (variant === "unresolved") {
        const original = mock.service.request;
        mock.service.request = async (csl, kind) => csl === `.show table ['${name}'] policy update`
          ? mockResult(["Policy"], [['{"not":"an array"}']]) : original(csl, kind);
      }
      const result = await new ModeledDataCreator().run(request, mock.service, true);
      assert.equal(result.phase, "blocked", `${name}/${variant}: ${result.error}`);
      assert.equal(mock.writes.length, 0);
    }
  }
});

test("lookup-trigger streaming changes or malformed values still stop the later boundary with confirmed history", async () => {
  for (const variant of ["enabled", "unknown", "malformed"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    mock.tables.get("Twins")!.streaming = true;
    let changed = false;
    mock.options.before = csl => {
      if (!csl.startsWith(".alter table ['Modeled_Output'] column-docstrings")) return;
      changed = true;
      if (variant !== "malformed") mock.tables.get("Telemetry")!.streaming = variant === "enabled" ? true : null;
    };
    const original = mock.service.request;
    mock.service.request = async (csl, kind) => {
      const result = await original(csl, kind);
      if (changed && variant === "malformed" && csl === ".show table ['Telemetry'] details") {
        result.Rows[0][result.Columns.findIndex(column => column.ColumnName === "StreamingIngestionPolicy")] = "{invalid";
      }
      return result;
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed"); assert.equal(result.objects.length, 3);
    assert.ok(!mock.writes.some(csl => csl.includes(" policy update ")));
    const history = rememberCreationAttempt([], result);
    assert.equal(history.length, 1); assert.equal(history[0].objects.length, 3);
    assert.equal(history[0].phase, "failed");
  }
});

test("both raw identity and incoming-policy races still stop policy installation without losing partial effects", async () => {
  for (const name of ["Telemetry", "Twins"]) {
    for (const variant of ["identity", "cascade", "unresolved"] as const) {
      const request = creationRequest(), mock = createCreationMock(request);
      mock.tables.get("Twins")!.streaming = true;
      let changed = false;
      mock.options.before = csl => {
        if (!csl.startsWith(".alter table ['Modeled_Output'] column-docstrings")) return;
        changed = true;
        if (variant === "identity") mock.tables.get(name)!.id = "replacement-source";
        if (variant === "cascade") mock.tables.get(name)!.policy = [{ Source: "NewUpstream" }];
      };
      const original = mock.service.request;
      mock.service.request = async (csl, kind) => changed && variant === "unresolved" && csl === `.show table ['${name}'] policy update`
        ? mockResult(["Policy"], [["{invalid"]]) : original(csl, kind);
      const result = await new ModeledDataCreator().run(request, mock.service, true);
      assert.equal(result.phase, "failed", `${name}/${variant}: ${result.error}`);
      assert.ok(result.objects.length && !mock.writes.some(csl => csl.includes(" policy update ")));
      assert.deepEqual(rememberCreationAttempt([], result)[0].objects, result.objects);
    }
  }
});

test("properties streaming can become enabled between preflight and source revalidation without raw alteration", async () => {
  const request = creationRequest(), mock = createCreationMock(request);
  mock.options.before = csl => {
    if (csl.startsWith(".alter table ['Modeled_Output'] column-docstrings")) mock.tables.get("Twins")!.streaming = true;
  };
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "completed");
  assert.equal(mock.tables.get("Twins")!.streaming, true);
  assert.ok(!mock.writes.some(csl => csl.startsWith(".alter table ['Twins']")));
});

test("missing permission, malformed metadata and ordered/equal-type permutations are not success", async () => {
  for (const variant of ["permissions", "metadata", "schema"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    if (variant === "permissions") mock.options.role = "Database Database Viewer";
    if (variant === "metadata") mock.options.fail = csl => csl.includes(" principal roles") ? new Error("Denied (403)") : undefined;
    if (variant === "schema") mock.options.schemaMismatch = request.names.modeled;
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "blocked"); assert.equal(mock.writes.length, 0);
  }
});

test("actual projection schema rejects same-type name permutations and malformed metadata before writes", async () => {
  assert.throws(() => assertOrderedSchema(mockResult(["ColumnName", "ColumnOrdinal", "ColumnType"],
    [["second", 0, "string"], ["first", 1, "string"]]),
  [{ name: "first", type: "string" }, { name: "second", type: "string" }]), /ordered column/);
  const request = creationRequest(), mock = createCreationMock(request);
  const original = mock.service.request;
  mock.service.request = async (csl, kind) => csl === ".show functions | project Name"
    ? mockResult(["Unexpected"], [[]]) : original(csl, kind);
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "blocked"); assert.equal(mock.writes.length, 0);
});

test("actual post-helper projection mismatch or a newly appeared output policy cannot be installed", async () => {
  for (const variant of ["schema", "policy"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    mock.options.before = csl => {
      if (variant === "schema" && csl.startsWith(".create materialized-view")) mock.options.schemaMismatch = request.names.modeled;
      if (variant === "policy" && csl.startsWith(".alter table ['Modeled_Output'] column-docstrings")) {
        mock.tables.get(request.names.modeled)!.policy = [{ Source: "ForeignSource", Query: "ForeignQuery" }];
      }
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed");
    assert.ok(!mock.writes.some(csl => csl.startsWith(".alter table ['Modeled_Output'] policy update")));
  }
});

test("foreign policy/identity replacement during mutation stops with explicit partial state and no rollback", async () => {
  const request = creationRequest(), mock = createCreationMock(request);
  mock.options.before = csl => {
    if (csl.startsWith(".alter table ['Modeled_Output'] docstring")) mock.tables.get(request.names.modeled)!.id = "replaced-after-check";
  };
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "failed");
  assert.ok(result.objects.length);
  assert.ok(!mock.writes.some(csl => csl.includes(" policy update ")));
  assert.ok(!mock.writes.some(csl => csl.startsWith(".drop")));
});

test("changed raw identity or policy stops before installing new update policies", async () => {
  for (const variant of ["identity", "streaming"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    mock.options.before = csl => {
      if (!csl.startsWith(".alter table ['Modeled_Output'] column-docstrings")) return;
      if (variant === "identity") mock.tables.get("Telemetry")!.id = "replaced-source";
      if (variant === "streaming") mock.tables.get("Telemetry")!.streaming = true;
    };
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed"); assert.ok(result.objects.length);
    assert.ok(!mock.writes.some(csl => csl.includes(" policy update ")));
  }
});

test("HTTP200 management partial status failure and failed readback never complete creation", async () => {
  for (const variant of ["status", "policy", "docs"] as const) {
    const request = creationRequest(), mock = createCreationMock(request);
    if (variant === "status") mock.options.before = csl => {
      if (csl.startsWith(".alter table ")) parseKqlTableResult({ Tables: [
        { TableName: "PrimaryResult", Columns: [], Rows: [] },
        { TableName: "QueryStatus", Columns: [{ ColumnName: "Severity" }, { ColumnName: "StatusCode" }], Rows: [[2, -1]] },
      ] });
    };
    if (variant === "policy") mock.options.policyReadbackMismatch = true;
    if (variant === "docs") mock.options.documentationMismatch = true;
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "failed");
    assert.ok(result.error && result.objects.length);
  }
});

test("network-uncertain create is reported as incomplete, never retried or destructively cleaned up", async () => {
  const request = creationRequest(), mock = createCreationMock(request);
  mock.options.fail = csl => csl.startsWith(".set ") ? new Error("Connection closed; outcome unknown") : undefined;
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "failed");
  assert.ok(result.uncertainCommand);
  assert.equal(mock.calls.filter(call => call.csl.startsWith(".set ")).length, 1);
  assert.ok(!mock.calls.some(call => call.csl.startsWith(".drop")));
});

test("explicit waiting uses only source-probe metadata, not preview evidence as processing matches", async () => {
  const request = creationRequest(true), mock = createCreationMock(request);
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "completed"); assert.equal(result.awaitingData, true);
  assert.equal(result.coverage!.telemetry.payloads.length, 0);
  assert.ok(mock.writes.every(csl => !csl.includes("set-or-append") && !csl.includes("backfill=true")));
  const notAcknowledged = { ...request, allowWaiting: false };
  assert.equal((await new ModeledDataCreator().run(notAcknowledged, mock.service, true)).phase, "blocked");
});

test("a device or module retains the same correlation key across different Hubs", async () => {
  for (const module of ["", "module-1"]) {
    const request = creationRequest();
    const telemetry = request.coverage.telemetry.payloads[0], twin = request.coverage.twin.payloads[0];
    telemetry.userHeaders = { hubName: "telemetry-hub" };
    twin.userHeaders = { hubName: "twin-hub" };
    telemetry.headers!.IoTConnectionModuleId = module;
    twin.headers!.IoTConnectionModuleId = module;
    const identity = module ? `device-1/${module}` : "device-1";
    assert.equal(nativeIdentity(telemetry, "telemetry"), identity);
    assert.equal(nativeIdentity(twin, "twin"), identity);
    const mock = createCreationMock(request);
    assert.equal((await new ModeledDataCreator().run(request, mock.service, true)).phase, "completed");
    const command = mock.plan.commands.find(command => command.kind === "policy" && command.target === request.names.modeled)!;
    if (command.kind !== "policy") throw new Error("Missing modeled-data policy.");
    assert.match(command.query, /lookup kind=leftouter lkv_bag on deviceId/);
    assert.ok(!command.query.includes("hubName") && !command.query.includes("sourceHubId"));
  }
});

test("missing, malformed or conflicting native identities and incompatible modules still block before writes", async () => {
  for (const variant of ["missing", "device-type", "module-type", "profiles", "module-mismatch", "component-type"] as const) {
    const request = creationRequest();
    const telemetry = request.coverage.telemetry.payloads[0], twin = request.coverage.twin.payloads[0];
    if (variant === "missing") telemetry.headers = {};
    if (variant === "device-type") telemetry.headers!.IoTConnectionDeviceId = 42;
    if (variant === "module-type") telemetry.headers!.IoTConnectionModuleId = 42;
    if (variant === "profiles") telemetry.userHeaders = { cloudEvents_subject: "another-device" };
    if (variant === "component-type") telemetry.headers!.IoTSubject = 42;
    if (variant === "module-mismatch") {
      telemetry.headers!.IoTConnectionModuleId = "module-1";
      twin.headers!.IoTConnectionModuleId = "module-2";
    }
    const mock = createCreationMock(request);
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "blocked"); assert.equal(mock.writes.length, 0);
    assert.equal(result.steps.find(step => step.label === "Verify native device and component bindings")?.status, "failed");
  }
});

test("blank-string EventHub placeholders no longer block verified CloudEvents native device/module bindings at the creation boundary", async () => {
  for (const module of ["", "module-1"]) {
    const request = creationRequest(), identity = module ? `device-1/${module}` : "device-1";
    const telemetry = request.coverage.telemetry.payloads[0], twin = request.coverage.twin.payloads[0];
    telemetry.headers = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" };
    twin.headers = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" };
    telemetry.userHeaders = { cloudEvents_subject: identity, cloudEvents_moduleid: module };
    twin.userHeaders = { cloudEvents_deviceid: "device-1", cloudEvents_subject: identity, cloudEvents_moduleid: module };
    assert.equal(nativeIdentity(telemetry, "telemetry"), identity);
    assert.equal(nativeIdentity(twin, "twin"), identity);
    const mock = createCreationMock(request);
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "completed");
    assert.equal(result.steps.find(step => step.label === "Verify native device and component bindings")?.status, "passed");
    assert.equal(result.objects.length, 3);
    assert.ok(mock.plan.commands.filter(command => command.kind === "policy")
      .every(command => command.kind === "policy" && command.query.includes("__emptyEventHubProfile")));
  }
});

test("blank-EH exception still blocks malformed/conflicting fallback metadata before any modeled-data mutation", async () => {
  for (const variant of ["eh-device-null", "eh-module-type", "eh-component-nonempty", "ce-module-type", "ce-component-type",
    "twin-conflict", "module-mismatch"] as const) {
    const request = creationRequest(), telemetry = request.coverage.telemetry.payloads[0], twin = request.coverage.twin.payloads[0];
    telemetry.headers = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" };
    twin.headers = { IoTConnectionDeviceId: "", IoTConnectionModuleId: "", IoTSubject: "" };
    telemetry.userHeaders = { cloudEvents_subject: "device-1", cloudEvents_moduleid: "" };
    twin.userHeaders = { cloudEvents_deviceid: "device-1", cloudEvents_moduleid: "" };
    if (variant === "eh-device-null") telemetry.headers.IoTConnectionDeviceId = null;
    if (variant === "eh-module-type") telemetry.headers.IoTConnectionModuleId = 42;
    if (variant === "eh-component-nonempty") telemetry.headers.IoTSubject = "orphan-component";
    if (variant === "ce-module-type") telemetry.userHeaders.cloudEvents_moduleid = 42;
    if (variant === "ce-component-type") telemetry.userHeaders.cloudEvents_iothubdtsubject = 42;
    if (variant === "twin-conflict") twin.userHeaders.deviceId = "other-device";
    if (variant === "module-mismatch") {
      telemetry.userHeaders.cloudEvents_subject = "device-1/module-1";
      telemetry.userHeaders.cloudEvents_moduleid = "module-1";
      twin.userHeaders.cloudEvents_moduleid = "module-2";
    }
    const mock = createCreationMock(request);
    const result = await new ModeledDataCreator().run(request, mock.service, true);
    assert.equal(result.phase, "blocked", variant); assert.equal(mock.writes.length, 0);
    assert.equal(result.steps.find(step => step.label === "Verify native device and component bindings")?.status, "failed");
  }
});

test("V1 telemetry-only outputs retain the property table/view/policy architecture without a lookup streaming prerequisite", async () => {
  const request = creationRequest(false, false);
  request.coverage.telemetry.payloads.forEach(sample => { sample.userHeaders = {}; });
  const mock = createCreationMock(request);
  mock.tables.get(request.input.source.telemetryTable)!.streaming = true;
  mock.tables.get(request.input.source.twinTable)!.streaming = true;
  const result = await new ModeledDataCreator().run(request, mock.service, true);
  assert.equal(result.phase, "completed"); assert.equal(result.objects.length, 3);
  assert.ok(mock.calls.some(call => call.csl.startsWith(".show table ['Twins']")));
  assert.ok(mock.writes.some(csl => csl.includes(".create materialized-view")));
  assert.ok(!mock.writes.some(csl => /backfill\s*=\s*true/.test(csl)));
});

test("absent versus null twin fields retain presence-based normalization, not null-based omission", () => {
  const request = creationRequest(), mock = createCreationMock(request);
  const command = mock.plan.commands.find(command => command.kind === "policy" && command.target === request.names.twinChanges)!;
  if (command.kind !== "policy") throw new Error("Missing twin policy.");
  assert.ok(command.query.includes("bag_keys(rp)") && command.query.includes("propertyValue = rp[key]"));
  assert.ok(!command.query.includes("isnotnull(rp[key])"));
  assert.ok(!command.query.includes("__selected") && !command.query.includes("replaceTwin"));
});

test("duplicate create requests on the same operation are rejected before a second mutation", async () => {
  const request = creationRequest(), mock = createCreationMock(request), creator = new ModeledDataCreator();
  let release: () => void = () => undefined;
  mock.options.inspect = () => new Promise(resolve => { release = () => resolve(request.coverage); });
  const first = creator.run(request, mock.service, true);
  await new Promise(resolve => setTimeout(resolve, 0));
  await assert.rejects(creator.run(request, mock.service, true), /already running/);
  release(); assert.equal((await first).phase, "completed");
  assert.equal(mock.writes.filter(csl => csl.startsWith(".set ['Modeled_Output']")).length, 1);
});
