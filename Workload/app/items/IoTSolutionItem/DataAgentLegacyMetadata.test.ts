// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import { buildDataAgentContext, buildDataAgentDataSource, normalizeDocumentedSource, type DocumentedSource } from "./DataAgentContext.ts";

test("one legacy normalization feeds preview serialized datasource and examples without overriding native meanings", () => {
  const source: DocumentedSource = { database: "Database", table: "Legacy", docstring: [
    "Device operating measurements.",
    '- temperature — "Temperature" (telemetry, double, unit: degC): Ambient temperature.',
    "- mode (property, string): Operating mode.",
    "- native (property, string): Obsolete meaning.",
  ].join("\n"), columns: [
    { name: "deviceId", type: "string" }, { name: "event_time", type: "datetime" },
    { name: "temperature", type: "real" }, { name: "mode", type: "string" },
    { name: "native", type: "string", docstring: "Native authoritative meaning.\nTelemetry." },
    { name: "undocumented", type: "string" },
  ] };
  const before = JSON.stringify(source);
  const normalized = normalizeDocumentedSource(source);
  const preview = buildDataAgentContext(normalized);
  const datasource = JSON.parse(JSON.stringify(buildDataAgentDataSource(normalized, "workspace", "database", "table")));
  assert.equal(datasource.dataSourceInstructions, preview.sourceInstructions);
  assert.equal(datasource.userDescription, preview.sourceDescription);
  assert.match(datasource.elements[0].children[0].children.find(column => column.id === "mode").description,
    /Operating mode\./);
  assert.match(datasource.dataSourceInstructions, /not an event-time\/as-of snapshot/);
  assert.equal(datasource.elements[0].children[0].children.find(column => column.id === "native").description, "Native authoritative meaning.");
  assert.match(preview.sourceInstructions, /### Telemetry Columns[\s\S]*`native` \(string\): Native authoritative meaning\./);
  assert.equal(datasource.elements[0].children[0].children.find(column => column.id === "undocumented").description, null);
  assert.ok(preview.examples.some(example => example.query.includes("isnotnull(['temperature'])")));
  assert.equal(JSON.stringify(source), before);
  assert.deepEqual(normalizeDocumentedSource(normalized), normalized);
  const updates = normalizeDocumentedSource({ ...source, docstring: source.docstring + "\n@iot.meta rowKind=property-update" });
  assert.match(updates.columns.find(column => column.name === "mode")!.docstring!, /Reported property value\./);
  assert.doesNotMatch(buildDataAgentContext(updates).sourceInstructions, /not an event-time\/as-of snapshot/);
});
