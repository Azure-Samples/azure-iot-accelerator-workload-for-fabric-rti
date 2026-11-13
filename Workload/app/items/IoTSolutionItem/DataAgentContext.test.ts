// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { readFileSync } from "node:fs";
import { buildDataAgentContext, buildDataAgentDataSource, dataAgentExamplesMarkdown, defaultDataAgentName, type DocumentedSource } from "./DataAgentContext.ts";
import { columnDocstring, columnSemanticAnnotation, FIELD_SEMANTIC_INFO_PREFIX } from "./FieldSemantics.ts";
import { dataAgentSourceFixture } from "../../../tests/dataAgentFixtures.ts";
import { declaredModelName, isJsonObject, modelResourceNameBase } from "./ModelAuthoring.ts";
import { replayDashboardQuery } from "../../../tests/dashboardFixtures.ts";

function nativeTables(value: unknown) {
  assert.ok(isJsonObject(value) && Array.isArray(value.elements) && value.elements.length === 1);
  const group = value.elements[0];
  assert.ok(isJsonObject(group) && group.type === "table_grouping" && group.display_name === "Tables" &&
    group.is_selected === false && Array.isArray(group.children) && group.children.length === 1);
  const table = group.children[0];
  assert.ok(isJsonObject(table) && table.type === "kusto.table" && table.is_selected === true && Array.isArray(table.children));
  assert.ok(table.children.every(column => isJsonObject(column) && column.type === "kusto.column" &&
    column.is_selected === true && Array.isArray(column.children) && column.children.length === 0));
  return { group, table };
}

const system = [
  { name: "deviceId", type: "string", docstring: columnDocstring({ name: "deviceId", description: "IoT Hub device Id" }) },
  { name: "event_time", type: "datetime", docstring: columnDocstring({ name: "event_time", description: "Event timestamp" }) },
];
const source = (columns: DocumentedSource["columns"], docstring = "@iot.meta tableRole=modeled-data"): DocumentedSource => ({
  database: "Fixture database", table: "Fixture's table", docstring, columns: [...system, ...columns],
});

for (const kind of ["wot", "dtdl"] as const) test(`V1-style ${kind} inspected native docs retain six real sample pairs and add independent latest measurements`, () => {
  const fixture = dataAgentSourceFixture(kind), before = JSON.stringify(fixture), context = buildDataAgentContext(fixture);
  const datasource = buildDataAgentDataSource(fixture, "workspace", "database", "table");
  assert.equal(context.agentInstructions, "You are an IoT data assistant. You help users explore and analyze IoT device data stored in a KQL (Kusto Query Language) database.");
  const headings = [...context.sourceInstructions.matchAll(/^#{2,3} (.*)$/gm)].map(match => match[1]);
  assert.deepEqual(headings, ["Data Model", "Device Model", "Key Columns", "Telemetry Columns (real-time measurements)",
    "Property Columns (device-reported metadata, enriched via last-known-value lookup)", "Query Guidelines"]);
  assert.ok(context.sourceInstructions.includes("- `deviceId` (string): IoT Hub device Id\n") &&
    context.sourceInstructions.includes("- `event_time` (datetime): Event timestamp\n\n  - Primary time axis"));
  assert.ok(!context.sourceInstructions.includes("not the IoT Hub enqueue time.") && !context.sourceInstructions.includes("Declared ID:"));
  assert.ok(!context.sourceInstructions.includes("@iot.meta") && !context.agentInstructions.includes("temperature"));
  assert.match(context.sourceInstructions, /not an event-time\/as-of snapshot/);
  assert.match(context.sourceInstructions, /Model name translations:/);
  assert.equal(context.examples.length, 7);
  assert.ok(context.examples[1].query.includes('deviceId == "YOUR_DEVICE_ID"') && context.examples[1].question.includes("replace YOUR_DEVICE_ID"));
  assert.ok(context.examples.every(example => !example.query.includes("```") && !example.query.includes("<device-id>") && !example.query.includes("&lt;")));
  const markdown = dataAgentExamplesMarkdown(context.examples);
  assert.ok(markdown.includes(`Q: ${context.examples[1].question}\n\n\`\`\`kql\n${context.examples[1].query}\n\`\`\`\n\n`));
  assert.match(context.examples[2].query, /event_time >= ago\(1h\).*isnotnull.*avg\(.*bin\(event_time, 5m\) \| render timechart$/);
  assert.ok(context.examples[3].query.indexOf("where isnotnull(") < context.examples[3].query.indexOf("arg_max(event_time"));
  assert.match(context.examples[4].question, /reported/);
  assert.match(context.examples[5].query, /ago\(24h\).*Count=count\(\).*by deviceId/);
  assert.ok(context.examples.every(example => example.query.startsWith(`['${fixture.table}']`) && !example.query.includes("enqueuedTime")));
  assert.equal(datasource.dataSourceInstructions, context.sourceInstructions);
  nativeTables(datasource);
  const table = datasource.elements[0].children[0];
  assert.equal(table.display_name, fixture.table);
  assert.ok(table.children.every(column => !column.description?.includes("@iot.meta") && column.type === "kusto.column"));
  assert.equal(table.children.find(column => column.id === "event_time")!.description, "Event timestamp");
  if (kind === "wot") {
    assert.match(context.sourceInstructions, /Model version: metadata-v1/);
    for (const value of ['WoT type: "number"', 'Semantic @type: "ex:Temperature"', "Minimum: 0", "Default: false",
      "Constant: false", "Observable: false", "Default: null", '"levels":{"type":"array"', "unit: degC"]) {
      assert.ok(context.sourceInstructions.includes(value), value);
    }
    assert.ok(!context.sourceInstructions.includes("DTDL schema:"));
  } else {
    assert.match(context.sourceInstructions, /Model ID: dtmi:fixture:SemanticDevice;2/);
    assert.match(context.sourceInstructions, /Model version: 2/);
    for (const value of ['DTDL schema: "double"', "Writable: false", '"enumValue":0', '"@type":"Array"', '"@type":"Map"',
      "component: motor", "unit: degreeCelsius"]) assert.ok(context.sourceInstructions.includes(value), value);
    assert.ok(!context.sourceInstructions.includes("WoT type:"));
  }
  assert.equal(JSON.stringify(fixture), before);
});

test("targeted latest examples use inspected component and telemetry facts, never name guesses or property roles", () => {
  const fixture = source([
    { name: "root reading", type: "real", docstring: columnDocstring({ name: "reading", role: "telemetry" }) },
    { name: "selected'\\[]", type: "string", docstring: columnDocstring({ name: "motor.status", component: "motor", role: "telemetry" }) },
    { name: "motor_guessed", type: "real" },
    { name: "reported", type: "real", docstring: columnDocstring({ name: "motor.reported", component: "motor", role: "twin property (last-known whole received value)" }) },
    { name: "other reading", type: "real", docstring: columnDocstring({ name: "other", role: "telemetry" }) },
  ]);
  const before = JSON.stringify(fixture), context = buildDataAgentContext(fixture);
  const component = context.examples.find(example => example.question.includes("telemetry from component motor"));
  assert.ok(component);
  assert.equal(component.query, "['Fixture\\'s table'] | where isnotempty(['selected\\'\\\\[]']) | summarize arg_max(event_time, *) by deviceId | project deviceId, event_time, ['selected\\'\\\\[]']");
  const pair = context.examples.find(example => example.question.includes("each measurement's own timestamp"));
  assert.ok(pair);
  assert.match(pair.question, /root reading and selected/);
  assert.doesNotMatch(pair.query, /motor_guessed|reported|other reading|motor_status/);
  assert.equal(context.examples.filter(example => example.question.includes("telemetry from component")).length, 1);
  assert.equal(context.examples.filter(example => example.question.includes("each measurement's own timestamp")).length, 1);
  const propertyOnly = buildDataAgentContext(source([fixture.columns.find(column => column.name === "reported")!]));
  assert.ok(!propertyOnly.examples.some(example => /telemetry from component|each measurement's own timestamp/.test(example.question)));
  const undocumented = buildDataAgentContext(source([{ name: "motor_temperature", type: "real" }, { name: "motor_pressure", type: "real" }]));
  assert.ok(!undocumented.examples.some(example => /telemetry from component|each measurement's own timestamp/.test(example.question)));
  assert.equal(JSON.stringify(fixture), before);
});

test("generated root and component latest examples preserve different event times, sparse newest rows and one-sided observations", () => {
  const first = "root reading", second = "chosen'\\[]";
  const fixture = source([
    { name: first, type: "real", docstring: columnDocstring({ name: "reading", role: "telemetry" }) },
    { name: second, type: "string", docstring: columnDocstring({ name: "component.status", role: "telemetry", component: "component" }) },
  ]);
  const context = buildDataAgentContext(fixture);
  const pair = context.examples.find(example => example.question.includes("each measurement's own timestamp"))!;
  const component = context.examples.find(example => example.question.includes("telemetry from component"))!;
  assert.ok(pair && component);
  const t1 = "2026-11-07T00:00:01Z", t2 = "2026-11-07T00:00:02Z", t3 = "2026-11-07T00:00:03Z";
  const row = (deviceId: string, event_time: string, value: number | null, status: string | null) =>
    ({ deviceId, event_time, [first]: value, [second]: status });
  const tables = { [fixture.table]: [
    row("both", t1, 0, ""), row("both", t2, null, "001"), row("both", t3, null, ""),
    row("first-only", t1, 42, null), row("first-only", t3, null, ""),
    row("second-only", t2, null, "ready"), row("second-only", t3, null, ""),
    row("neither", t3, null, ""),
  ] };
  const result = replayDashboardQuery(pair.query, tables);
  assert.deepEqual(result.find(item => item.deviceId === "both"),
    { deviceId: "both", Value1: 0, event_time1: t1, Value2: "001", event_time2: t2 });
  assert.deepEqual(result.find(item => item.deviceId === "first-only"),
    { deviceId: "first-only", Value1: 42, event_time1: t1, Value2: "", event_time2: null });
  assert.deepEqual(result.find(item => item.deviceId === "second-only"),
    { deviceId: "second-only", Value1: null, event_time1: null, Value2: "ready", event_time2: t2 });
  assert.equal(result.length, 3);
  assert.deepEqual(replayDashboardQuery(component.query, tables).map(item => [item.deviceId, item.event_time, item[second]]),
    [["both", t2, "001"], ["second-only", t2, "ready"]]);
  const absentSecond = { [fixture.table]: [row("first-only", t1, 0, ""), row("first-only", t3, null, "")] };
  assert.deepEqual(replayDashboardQuery(pair.query, absentSecond),
    [{ deviceId: "first-only", Value1: 0, event_time1: t1, Value2: "", event_time2: null }]);
  const absentFirst = { [fixture.table]: [row("second-only", t1, null, "0"), row("second-only", t3, null, "")] };
  assert.deepEqual(replayDashboardQuery(pair.query, absentFirst),
    [{ deviceId: "second-only", Value1: null, event_time1: null, Value2: "0", event_time2: t1 }]);
  assert.deepEqual(replayDashboardQuery(pair.query, { [fixture.table]: [] }), []);
});

test("targeted latest examples retain whole dynamic telemetry and false without scalar or string coercions", () => {
  const fixture = source([
    { name: "whole object", type: "dynamic", docstring: columnDocstring({ name: "object", role: "telemetry" }) },
    { name: "component flag", type: "bool", docstring: columnDocstring({ name: "component.flag", component: "component", role: "telemetry" }) },
  ]);
  const query = buildDataAgentContext(fixture).examples.find(example => example.question.includes("each measurement's own timestamp"))!.query;
  for (const whole of [{ nested: [0, false, null, { value: "001" }] }, [0, false, null, { value: "001" }]]) {
    assert.deepEqual(replayDashboardQuery(query, { [fixture.table]: [
      { deviceId: "device", event_time: 1, "whole object": whole, "component flag": null },
      { deviceId: "device", event_time: 2, "whole object": null, "component flag": false },
      { deviceId: "device", event_time: 3, "whole object": null, "component flag": null },
    ] }), [{ deviceId: "device", Value1: whole, event_time1: 1, Value2: false, event_time2: 2 }]);
  }
  assert.doesNotMatch(query, /tostring|todynamic|todouble|todecimal/);
});

test("generated examples remain the collection wired into both persisted Data Agent fewshot configurations", () => {
  const context = buildDataAgentContext(dataAgentSourceFixture("wot"));
  const fewShots = context.examples.map(example => ({ id: crypto.randomUUID(), ...example }));
  const serialized = JSON.parse(JSON.stringify({ fewShots }));
  assert.deepEqual(serialized.fewShots.map(({ question, query }: { question: string; query: string }) => ({ question, query })), context.examples);
  assert.ok(serialized.fewShots.some((example: { query: string }) => example.query.includes("join kind=fullouter")));
  const setup = readFileSync(new URL("./steps/DataAgentSetupStep.tsx", import.meta.url), "utf8").replace(/\r\n/g, "\n");
  assert.match(setup, /\(sourceContext\?\.examples \|\| \[\]\)\.map\(example => \(\{ id: crypto\.randomUUID\(\), \.\.\.example \}\)\)/);
  assert.match(setup, /const fewShotsConfig = \{[\s\S]*?fewShots: fewShots,/);
  for (const stage of ["draft", "published"]) assert.ok(setup.includes(
    `path: \`Files/Config/${stage}/\${dsName}/fewshots.json\`,\n          payload: toB64(fewShotsConfig),`));
  assert.match(setup, /aiInstructions: sourceContext\.agentInstructions,/);
});

test("V1-style roles never follow numeric/string types semantic quantity classes or access flags", () => {
  const fixture = source([
    { name: "reading'\\[]", type: "bool", docstring: columnDocstring({ name: "reading", type: "boolean", role: "telemetry", default: false }) },
    { name: "not_a_role", type: "real", docstring: columnDocstring({ name: "not_a_role", "@type": "Telemetry", readOnly: true }) },
    { name: "reported_numeric", type: "real", docstring: columnDocstring({ name: "reported_numeric", role: "twin property (last-known whole received value)" }) },
    { name: "text_telemetry", type: "string", docstring: columnDocstring({ name: "text_telemetry", type: "string", role: "telemetry" }) },
  ]);
  const context = buildDataAgentContext(fixture), body = context.sourceInstructions;
  assert.match(body, /### Additional Columns \(role not documented\)\n\n- `not_a_role` \(real\)/);
  assert.ok(body.indexOf("`reported_numeric`") > body.indexOf("### Property Columns"));
  assert.ok(body.indexOf("`text_telemetry`") < body.indexOf("### Property Columns"));
  assert.ok(!context.examples.some(example => example.query.includes("avg(")));
  assert.ok(context.examples.some(example => example.question.includes("reported reported_numeric")));
  assert.ok(context.examples.every(example => example.query.startsWith("['Fixture\\'s table']")));
  assert.ok(context.examples.some(example => example.query.includes("['reading\\'\\\\[]']")));
});

test("V1-style native constraints and localization preserve zero false null empty and complete nested values", () => {
  const facts = { name: "state", role: "telemetry", type: "object", title: "Original", titles: { "en-US": "Selected", de: "Zustand" },
    description: "", default: null, const: false, enum: [0, false, null, "", [], {}], readOnly: false,
    properties: { samples: { type: "array", items: { type: "number", minimum: 0 }, minItems: 0 } }, required: [] };
  const fixture = source([{ name: "state", type: "dynamic", docstring: columnDocstring(facts) }]);
  const text = buildDataAgentContext(fixture).sourceInstructions;
  for (const value of ['Selected \u2014 ""', 'Declared title: "Original"', 'Enum: [0,false,null,"",[],{}]', "Default: null",
    "Constant: false", "Read-only: false", 'Titles: {"en-US":"Selected","de":"Zustand"}',
    'Properties: {"samples":{"type":"array","items":{"type":"number","minimum":0},"minItems":0}}', "Required members: []"]) assert.ok(text.includes(value), value);
});

test("V1-style incomplete absent legacy and malformed metadata keep honest boundaries", () => {
  const bounded = source([{ name: "partial", type: "dynamic", docstring: FIELD_SEMANTIC_INFO_PREFIX +
    '{"name":"partial","role":"telemetry","truncated":true}' }], '@iot.meta modelSemanticInfo={"id":"urn:fixture:no-name","truncated":true}');
  const text = buildDataAgentContext(bounded).sourceInstructions;
  assert.match(text, /Model ID: urn:fixture:no-name/);
  assert.ok(!text.includes("Model name:") && !text.includes("Model version:") && !text.includes("WoT type:"));
  assert.match(text, /Metadata: incomplete \(truncated=true\)/);
  assert.match(text, /Do not reconstruct missing declarations/);
  const missing = buildDataAgentContext(source([{ name: "unknown", type: "string" }]));
  assert.ok(!missing.sourceInstructions.includes("### Device Model") && !missing.sourceInstructions.includes("### Telemetry Columns") &&
    !missing.sourceInstructions.includes("### Property Columns") && !missing.examples.some(example => /average|reported/.test(example.question)));
  assert.equal(buildDataAgentContext(source([])).examples.length, 3);
  const legacy = source([{ name: "temperature", type: "real", docstring: "Declared legacy reading.\nTelemetry.\n" +
    columnSemanticAnnotation({ kind: "wot" }, "temperature", { kind: "wot", schema: { type: "number", minimum: 0 }, annotations: {} }) }]);
  assert.match(buildDataAgentContext(legacy).sourceInstructions, /WoT type: "number"\n  - Minimum: 0/);
  const flat = source([{ name: "flat", type: "dynamic", docstring: 'name: "flat"\nrole: "telemetry"\nDeclared type: "array"\nitems:\n\ttype: "number"' }]);
  assert.match(buildDataAgentContext(flat).sourceInstructions, /Items: {"type":"number"}/);
  const updates = buildDataAgentContext({ ...legacy, docstring: "@iot.meta rowKind=property-update" });
  assert.match(updates.sourceInstructions, /Each row represents a reported property update/);
  assert.ok(!updates.sourceInstructions.includes("not an event-time/as-of snapshot"));
  for (const docstring of [FIELD_SEMANTIC_INFO_PREFIX + "{bad", FIELD_SEMANTIC_INFO_PREFIX + '{"name":"x","role":"Property"}',
    columnDocstring({ name: "x" }) + "\n" + columnDocstring({ name: "x" })]) {
    assert.throws(() => buildDataAgentContext(source([{ name: "x", type: "real", docstring }])));
  }
  assert.throws(() => buildDataAgentContext(source([], "@iot.meta modelSemanticInfo={bad")));
  assert.throws(() => buildDataAgentContext({ ...source([]), columns: [{ name: "enqueuedTime", type: "datetime" }] }), /event_time:datetime/);
});

test("V1-style native string latest examples explicitly filter nonempty values without pretending string telemetry is a property", () => {
  const fixture = source([{ name: "status", type: "string", docstring: columnDocstring({ name: "status", role: "telemetry", enum: ["", "ready"] }) }]);
  const context = buildDataAgentContext(fixture);
  assert.ok(context.examples.some(example => example.question.includes("nonempty status") && example.query.includes("where isnotempty(['status'])")));
  assert.ok(!context.examples.some(example => example.question.includes("reported")));
  assert.match(context.sourceInstructions, /empty string as a declared code unless the source documents it/);
});

test("agent default names use declared locale preferences or the table, never model IDs", () => {
  for (const [info, expected] of [
    [{ id: "urn:fixture:not-a-name", title: "Device title", titles: { en: "English", "en-US": "US device" } }, "US device"],
    [{ "@id": "dtmi:fixture:NotAName;2", displayName: { de: "Erstes Modell", en: "English device" } }, "English device"],
    [{ "@id": "dtmi:fixture:NotAName;2", displayName: { fr: "Premier appareil", "en-US": "Later device" } }, "Premier appareil"],
    [{ id: "urn:fixture:no-title" }, undefined],
    [{ id: "urn:fixture:empty-title", title: "" }, ""],
  ] as const) {
    const doc = "@iot.meta modelSemanticInfo=" + JSON.stringify(info);
    assert.equal(declaredModelName(doc), expected);
    assert.equal(defaultDataAgentName(doc, "Selected Table", "suffix"),
      `DataAgent_${(expected || "Selected Table").replace(/[^a-zA-Z0-9_]/g, "_")}_suffix`);
    assert.equal(modelResourceNameBase(doc, "Selected Table"), (expected || "Selected Table").replace(/[^a-zA-Z0-9_]/g, "_"));
    assert.ok(!defaultDataAgentName(doc, "Selected Table", "suffix").includes("urn_fixture") &&
      !defaultDataAgentName(doc, "Selected Table", "suffix").includes("dtmi_fixture"));
  }
  assert.equal(defaultDataAgentName("", "Selected Table", "suffix"), "DataAgent_Selected_Table_suffix");
  assert.throws(() => modelResourceNameBase("@iot.meta modelSemanticInfo={bad", "Selected Table"), /Malformed/);
  assert.throws(() => modelResourceNameBase("@iot.meta modelId=not-a-canonical-id", "Selected Table"), /Malformed/);
});

test("readable column facts have Markdown paragraph/list boundaries and omit only declaration IDs", () => {
  const doc = '@iot.meta modelSemanticInfo={"@id":"dtmi:fixture:Model;2","displayName":"Fixture model"}\n@iot.meta tableRole=modeled-data';
  const fixture = source([
    { name: "firmware", type: "string", docstring: columnDocstring({
      name: "firmware", displayName: "Firmware Version", schema: "string", "@id": "dtmi:fixture:FirmwareDeclaration;2",
      role: "twin property (last-known whole received value)", default: "dtmi:fixture:LiteralDefault;1",
      enum: ["dtmi:fixture:LiteralEnum;1", ""], const: { "@id": "urn:fixture:literal-data" },
    }) },
    { name: "mode", type: "string", docstring: columnDocstring({
      name: "mode", schema: { "@id": "dtmi:fixture:EnumDeclaration;1", "@type": "Enum", valueSchema: "string",
        enumValues: [{ name: "value", enumValue: "dtmi:fixture:LiteralCode;1" }] },
    }) },
  ], doc);
  const before = JSON.stringify(fixture), context = buildDataAgentContext(fixture);
  const datasource = buildDataAgentDataSource(fixture, "workspace", "database", "table");
  assert.ok(context.sourceInstructions.includes('Firmware Version\n\n  - DTDL schema: "string"\n'));
  const description = datasource.elements[0].children[0].children.find(column => column.id === "firmware")!.description!;
  assert.ok(description.startsWith('Firmware Version\n\n- DTDL schema: "string"\n'));
  for (const id of ["FirmwareDeclaration", "EnumDeclaration"]) assert.ok(!context.sourceInstructions.includes(id));
  for (const literal of ["LiteralDefault", "LiteralEnum", "LiteralCode", "urn:fixture:literal-data", "dtmi:fixture:Model;2"]) {
    assert.ok(context.sourceInstructions.includes(literal), literal);
  }
  assert.equal(JSON.stringify(fixture), before);
});

test("native Tables conformance rejects flat or duplicate-root tables while preserving selected raw-name columns and opaque IDs", () => {
  for (const kind of ["wot", "dtdl"] as const) {
    const fixture = dataAgentSourceFixture(kind), tableId = crypto.randomUUID();
    const definition = buildDataAgentDataSource(fixture, "workspace", "database", tableId);
    const { group, table } = nativeTables(JSON.parse(JSON.stringify(definition)));
    assert.equal(table.id, tableId);
    assert.notEqual(group.id, table.id);
    const columns = definition.elements[0].children[0].children;
    assert.deepEqual(columns.map(column => [column.id, column.display_name]), fixture.columns.map(column => [column.name, column.name]));
    const ids = [group.id, table.id, ...columns.map(column => column.id)];
    assert.equal(new Set(ids).size, ids.length);
    assert.equal(columns.find(column => column.id === "event_time")!.data_type, "System.DateTime");
    assert.ok(columns.every((column, index) => (column.data_type === "System.SByte") === (fixture.columns[index].type === "bool")));
    assert.throws(() => nativeTables({ ...definition, elements: [definition.elements[0].children[0]] }));
    assert.throws(() => nativeTables({ ...definition, elements: [definition.elements[0].children[0], definition.elements[0]] }));
  }
});
