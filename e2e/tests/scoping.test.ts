import test from "node:test";
import assert from "node:assert/strict";
import { assertScopedOutput, expectedModeledHubKeys, matchesScope, scenarioScope, scopePredicate, scopeSamplesReady } from "../src/scoping.js";
import { modelFixture } from "../src/models.js";
import { ingestionPath } from "../src/paths.js";
import { createPlan } from "../src/plan.js";
import { runConfig } from "../src/config.js";

const config = runConfig({
  E2E_TENANT_ID: "00000000-0000-4000-8000-000000000001",
  E2E_WORKSPACE_ID: "00000000-0000-4000-8000-000000000002",
  E2E_SUBSCRIPTION_ID: "00000000-0000-4000-8000-000000000003",
  E2E_LOCATION: "australiaeast", E2E_DEPLOYMENT: "development",
});
const reference = modelFixture("combined").id;
const path = ingestionPath(createPlan(config), "regular");

for (const stream of ["telemetry", "twin"] as const) {
  test(`nonlinked ${stream} condition admits some events and excludes others`, () => {
    assert.equal(matchesScope("regular", stream, undefined, 2), true);
    assert.equal(matchesScope("regular", stream, reference, 3), false);
    assert.equal(matchesScope("regular", stream, reference, undefined), false);
    assert.equal(matchesScope("regular", stream, reference, "2"), false);
    assert.ok(scopePredicate("regular", stream).includes(stream === "telemetry"
      ? "data.payload.value.e2eSequence" : "data.payload.value.properties.reported.e2eSequence"));
  });
}

test("combined AND requires this reference even when the telemetry condition passes", () => {
  for (const ref of [undefined, "urn:another", reference]) for (const sequence of [1, 2]) {
    assert.equal(matchesScope("combined", "telemetry", ref, sequence), ref === reference && sequence === 2);
    assert.equal(matchesScope("combined", "twin", ref, sequence), ref === reference);
  }
  assert.match(scopePredicate("combined", "telemetry"), /\) and \(/);
  assert.equal(scenarioScope("combined").twin, "");
});

test("linked OR admits this model or matching properties, including the other-model device", () => {
  for (const ref of [undefined, "urn:another", reference]) for (const sequence of [1, 2]) {
    assert.equal(matchesScope("linked", "twin", ref, sequence), ref === reference || sequence === 2);
    assert.equal(matchesScope("linked", "telemetry", ref, sequence), ref === reference);
  }
  assert.match(scopePredicate("linked", "twin"), /\) or \(/);
  assert.deepEqual(expectedModeledHubKeys("combined", "telemetry"), ["linked1", "linked2"]);
  assert.deepEqual(expectedModeledHubKeys("linked", "twin"), ["linked1", "linked2", "linked3"]);
  assert.deepEqual(expectedModeledHubKeys("linked", "telemetry"), ["linked1", "linked2"]);
});

const positive = { deviceId: "owned-device", event_time: "2026-11-03T00:00:02Z", sequence: 2, rows: 1 };
const negative = { deviceId: "owned-device", event_time: "2026-11-03T00:00:03Z", sequence: 3, rows: 1 };
test("native scoping proof requires data and verifies both inclusion and exclusion against raw observations", () => {
  assert.doesNotThrow(() => assertScopedOutput(path, "telemetry", [positive, negative], [positive]));
  assert.throws(() => assertScopedOutput(path, "telemetry", [], []), /must both contain fresh data/);
  assert.throws(() => assertScopedOutput(path, "telemetry", [positive, negative], []), /must both contain fresh data/);
  assert.throws(() => assertScopedOutput(path, "telemetry", [negative], [negative]), /Missing positive\/negative/);
  assert.throws(() => assertScopedOutput(path, "telemetry", [positive], [positive]), /Missing positive\/negative/);
  assert.throws(() => assertScopedOutput(path, "telemetry", [positive, negative], [positive, negative]), /no matching accepted/);
  assert.throws(() => assertScopedOutput(path, "telemetry", [positive, negative], [{ ...positive, event_time: "old" }]), /no matching accepted/);
  assert.throws(() => assertScopedOutput(path, "telemetry", [positive, negative, { ...positive, event_time: "new" }], [positive]), /row count differs/);
  assert.doesNotThrow(() => assertScopedOutput(path, "telemetry", [positive, { ...positive, sequence: 3 }], [positive]));
  assert.doesNotThrow(() => assertScopedOutput(path, "telemetry", [positive, positive, negative], [{ ...positive, rows: 2 }]));
  assert.throws(() => assertScopedOutput(path, "telemetry", [positive, positive, negative], [positive]), /row count differs/);
  assert.throws(() => assertScopedOutput(path, "telemetry", [positive, positive, negative], [{ ...positive, rows: 3 }]), /row count differs/);
});

test("reference AND/OR verification cannot pass without exercising the additional condition independently", () => {
  const rows = [
    { reference, sequence: 2 },
    { reference, sequence: 3 },
    { reference: "urn:other", sequence: 2 },
    { reference: "urn:other", sequence: 3 },
  ];
  assert.equal(scopeSamplesReady("combined", "telemetry", rows), true);
  assert.equal(scopeSamplesReady("combined", "telemetry", rows.filter(row => row.reference !== reference || row.sequence === 2)), false);
  assert.equal(scopeSamplesReady("linked", "twin", rows), true);
  assert.equal(scopeSamplesReady("linked", "twin", rows.filter(row => row.reference === reference)), false);
  assert.equal(scopeSamplesReady("linked", "twin", rows.filter(row => row.reference !== reference || row.sequence === 2)), false);
});
