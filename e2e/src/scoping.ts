import assert from "node:assert/strict";
import type { IngestionPath, PathId } from "./paths.js";
import { modelFixture } from "./models.js";

export type Stream = "telemetry" | "twin";
const telemetryCondition = "tolong(data.payload.value.e2eSequence) % 2 == 0";
const propertyCondition = "tolong(data.payload.value.properties.reported.e2eSequence) % 2 == 0";

export function scenarioScope(path: PathId) {
  return {
    mode: path === "regular" ? "data" : "reference",
    operator: path === "linked" ? "or" : "and",
    telemetry: path === "linked" ? "" : telemetryCondition,
    twin: path === "combined" ? "" : propertyCondition,
  } as const;
}

export function scopePredicate(path: PathId, stream: Stream): string {
  const scope = scenarioScope(path), condition = scope[stream];
  if (scope.mode === "data") return condition;
  const reference = `tostring(data["payload"]["memberTypeIRI"]) == "${modelFixture(path).id}"`;
  return condition ? `(${reference}) ${scope.operator} (${condition})` : `(${reference})`;
}

export function matchesScope(path: PathId, stream: Stream, reference: unknown, sequence: unknown): boolean {
  const scope = scenarioScope(path);
  const condition = typeof sequence === "number" && Number.isInteger(sequence) && sequence % 2 === 0;
  if (scope.mode === "data") return condition;
  const matches = reference === modelFixture(path).id;
  if (!scope[stream]) return matches;
  return scope.operator === "and" ? matches && condition : matches || condition;
}

export function expectedModeledHubKeys(path: PathId, stream: Stream): string[] {
  if (path === "regular") return ["regular1", "regular2"];
  return path === "linked" && stream === "twin" ? ["linked1", "linked2", "linked3"] : ["linked1", "linked2"];
}

export function scopeSamplesReady(path: PathId, stream: Stream, raw: Record<string, unknown>[]): boolean {
  const scope = scenarioScope(path);
  const has = (reference: boolean | undefined, even: boolean | undefined) => raw.some(row =>
    (reference === undefined || (row.reference === modelFixture(path).id) === reference) &&
    (even === undefined || typeof row.sequence === "number" && (row.sequence % 2 === 0) === even));
  if (scope.mode === "data") return has(undefined, true) && has(undefined, false);
  if (!scope[stream]) return has(true, undefined) && has(false, undefined);
  return scope.operator === "and"
    ? has(true, true) && has(true, false) && has(false, true)
    : has(true, false) && has(false, true) && has(false, false);
}

export function assertScopedOutput(
  path: IngestionPath, stream: Stream, raw: Record<string, unknown>[], output: Record<string, unknown>[],
): void {
  assert.ok(raw.length > 0 && output.length > 0, "[native-scope] Raw and typed outputs must both contain fresh data.");
  assert.ok(scopeSamplesReady(path.id, stream, raw), "[native-scope] Missing positive/negative observations needed to exercise the chosen reference/condition combination.");
  const keys = new Map<string, number>();
  for (const row of raw) {
    assert.ok(typeof row.deviceId === "string" && row.deviceId && typeof row.event_time === "string" && row.event_time,
      "[native-scope] Raw observation identity and clock are required.");
    const key = JSON.stringify([row.deviceId, row.event_time]);
    const accepted = matchesScope(path.id, stream, row.reference, row.sequence);
    keys.set(key, (keys.get(key) ?? 0) + Number(accepted));
  }
  const actual = new Map<string, number>();
  for (const row of output) {
    const key = JSON.stringify([row.deviceId, row.event_time]);
    assert.ok(typeof row.rows === "number" && Number.isInteger(row.rows) && row.rows > 0, "[native-scope] Missing typed output row count.");
    assert.ok((keys.get(key) ?? 0) > 0, "[native-scope] An output row has no matching accepted current-attempt raw observation.");
    actual.set(key, (actual.get(key) ?? 0) + row.rows);
  }
  for (const [key, accepted] of keys) {
    assert.equal(actual.get(key) ?? 0, accepted,
      "[native-scope] Typed output row count differs from accepted raw observations.");
  }
}
