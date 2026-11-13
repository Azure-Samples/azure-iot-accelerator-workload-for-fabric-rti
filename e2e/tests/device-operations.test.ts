import test from "node:test";
import assert from "node:assert/strict";
import { DeviceOperationQueue } from "../src/device-operations.js";

test("background and explicit rounds never overlap and retain adjacent odd/even IDs across root/component acknowledgements", async () => {
  const queue = new DeviceOperationQueue(), events: string[] = [], sequences: number[] = [];
  let release: () => void = () => assert.fail("Missing first-round release");
  const blocked = new Promise<void>(resolve => { release = resolve; });
  let active = 0, maximum = 0, acknowledgements = 51;
  const round = (name: string, wait = false) => queue.round(async sequence => {
    maximum = Math.max(maximum, ++active);
    events.push(`${name}:start`);
    if (wait) await blocked;
    for (const component of ["root", "thermostat", "powerModule"]) {
      sequences.push(sequence);
      acknowledgements++;
      events.push(`${name}:${component}`);
      await Promise.resolve();
    }
    active--;
    events.push(`${name}:end`);
  });
  const background = round("background", true), explicit = round("explicit");
  const properties = queue.properties(async () => {
    maximum = Math.max(maximum, ++active);
    events.push("properties");
    active--;
  });
  const next = round("next");
  await Promise.resolve();
  assert.deepEqual(events, ["background:start"]);
  release();
  await Promise.all([background, explicit, properties, next]);
  assert.equal(maximum, 1);
  assert.equal(acknowledgements, 60);
  assert.deepEqual(sequences, [1, 1, 1, 2, 2, 2, 3, 3, 3]);
  assert.ok(events.indexOf("background:end") < events.indexOf("explicit:start"));
  assert.ok(events.indexOf("explicit:end") < events.indexOf("properties"));
  assert.ok(events.indexOf("properties") < events.indexOf("next:start"));
});

test("a failed device operation rejects queued operations without dispatching more sends or hiding the failure", async () => {
  const queue = new DeviceOperationQueue(), failure = new Error("Synthetic send failure");
  let dispatched = false;
  const first = queue.round(async () => { throw failure; });
  const queued = queue.properties(async () => { dispatched = true; });
  await assert.rejects(first, error => error === failure);
  await assert.rejects(queued, error => error === failure);
  assert.equal(dispatched, false);
});
