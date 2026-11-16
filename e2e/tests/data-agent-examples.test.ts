import test from "node:test";
import assert from "node:assert/strict";
import { assertDataAgentExamples, assertFirmwareAnswer } from "../src/experience-verification.js";
import { modelFixture } from "../src/models.js";

test("the newest agent firmware response must contain the exact version, not a substring or conflicting value", () => {
  for (const answer of ["1.0.0", "Firmware: 1.0.0", "The firmware version is `1.0.0`."]) assertFirmwareAnswer(answer);
  for (const answer of ["11.0.0", "1.0.01", "1.0.0.1", "1.0.0-beta", "Not 1.0.0; use 11.0.0", "No result"]) {
    assert.throws(() => assertFirmwareAnswer(answer), /exact fixture firmware/);
  }
});

function examples(component: boolean) {
  const table = "['Modeled_Fixture']";
  const latest = (column: string, string = false) =>
    `${table} | where ${string ? "isnotempty" : "isnotnull"}(['${column}']) | summarize arg_max(event_time, *) by deviceId | project deviceId, event_time, ['${column}']`;
  const examples = [
    { question: "What devices are in the system?", query: `${table} | distinct deviceId` },
    { question: "Show me the latest modeled records for a device (replace YOUR_DEVICE_ID)",
      query: `${table} | where deviceId == "YOUR_DEVICE_ID" | top 10 by event_time desc` },
    { question: "What is the average temperature over the last hour?",
      query: `${table} | where event_time >= ago(1h) and isnotnull(['temperature']) | summarize avg(['temperature']) by bin(event_time, 5m) | render timechart` },
    { question: "Show the latest non-null temperature for each device", query: latest("temperature") },
    { question: "Show the latest available nonempty reported firmwareVersion for each device", query: latest("firmwareVersion", true) },
    { question: "How many readings per device in the last 24 hours?",
      query: `${table} | where event_time >= ago(24h) | summarize Count=count() by deviceId | order by Count desc` },
  ];
  if (component) examples.push({
    question: "Show the latest non-null thermostat_targetTemperature telemetry from component thermostat for each device",
    query: latest("thermostat_targetTemperature"),
  });
  const second = component ? "thermostat_targetTemperature" : "gpsLocation";
  examples.push({
    question: `Show the latest available temperature and ${second} for each device, with each measurement's own timestamp`,
    query: [
      latest("temperature"),
      "| project deviceId, Value1 = ['temperature'], event_time1 = event_time",
      "| join kind=fullouter (",
      `  ${latest(second)}`,
      `  | project deviceId, Value2 = ['${second}'], event_time2 = event_time`,
      ") on deviceId",
      "| project deviceId = coalesce(deviceId, deviceId1), Value1, event_time1, Value2, event_time2",
    ].join("\n"),
  });
  return examples;
}

for (const path of ["regular", "combined", "linked"] as const) test(`${path} native examples retain originals and append documented independent latest values`, () => {
  const sample = examples(path === "regular");
  const before = JSON.stringify(sample);
  const appended = assertDataAgentExamples(sample, "Modeled_Fixture", modelFixture(path).fields);
  assert.equal(appended.length, path === "regular" ? 2 : 1);
  assert.equal((appended.at(-1)!.query.match(/arg_max\(event_time, \*\)/g) ?? []).length, 2);
  assert.match(appended.at(-1)!.query, /join kind=fullouter/);
  assert.doesNotMatch(appended.at(-1)!.query, /tostring|todynamic|todouble|todecimal/);
  assert.equal(JSON.stringify(sample), before);
  assert.throws(() => assertDataAgentExamples(sample.slice(0, 6), "Modeled_Fixture", modelFixture(path).fields));
  for (const changed of ["join kind=inner", "join kind=leftouter", "join kind=fullouter | where isnotnull(Value1) and isnotnull(Value2)"]) {
    assert.throws(() => assertDataAgentExamples(sample.map(example => ({
      ...example, query: example.query.replace("join kind=fullouter", changed),
    })), "Modeled_Fixture", modelFixture(path).fields));
  }
});

test("an underscore name alone cannot introduce a component example", () => {
  const fields = modelFixture("regular").fields.map(field => {
    const copy = { ...field };
    delete copy.component;
    return copy;
  });
  assert.throws(() => assertDataAgentExamples(examples(true), "Modeled_Fixture", fields));
});
