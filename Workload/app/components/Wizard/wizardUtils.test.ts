// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  adjacentWizardStep,
  completeWizardSafely,
  getPersistableWizardContext,
  WIZARD_COMPLETION_ERROR,
} from "./wizardUtils.ts";

test("forward and backward navigation preserve every registered step and stop at the wizard boundaries", () => {
  assert.equal(adjacentWizardStep(0, 1, 5), 1);
  assert.equal(adjacentWizardStep(2, -1, 5), 1);
  assert.equal(adjacentWizardStep(2, 1, 5), 3);
  assert.equal(adjacentWizardStep(0, -1, 5), undefined);
  assert.equal(adjacentWizardStep(4, 1, 5), undefined);
  assert.equal(adjacentWizardStep(0, 1, 0), undefined);
});

test("excludes connection strings from persisted wizard context", () => {
  const context = {
    telemetryEndpointConnectionString: "Endpoint=sb://secret",
    propertiesendpointconnectionstring: "Endpoint=sb://other-secret",
    actResourcesCreating: true,
    telemetryEndpointNamespace: "namespace.servicebus.windows.net",
  };

  assert.deepEqual(getPersistableWizardContext(context), {
    telemetryEndpointNamespace: "namespace.servicebus.windows.net",
  });

  test("excludes analyzed raw payload samples from all persisted wizard contexts", () => {
    const context = { hubSourceAnalysis: { telemetry: { payloads: [{ value: { privateSample: "fixture-only" } }] },
      twin: { payloads: [] } }, hubSourceSelectionRevision: 2, modelDbName: "Database" };
    const persisted = getPersistableWizardContext(context);
    assert.deepEqual(persisted, { hubSourceSelectionRevision: 2, modelDbName: "Database" });
    assert.ok(!JSON.stringify(persisted).includes("privateSample"));
  });
});

test("returns no error when wizard completion succeeds", async () => {
  const error = await completeWizardSafely(async () => undefined, { ready: true });

  assert.equal(error, undefined);
});

test("assignment intent stays ephemeral even when a wizard persists other context", () => {
  assert.deepEqual(getPersistableWizardContext({ hubAssignmentIntent: { owner: "fixture", mode: "skip" },
    hubAssignmentHubSummary: { owner: "fixture", value: { hubs: [] } },
    modelDbName: "Database" }), { modelDbName: "Database" });
});

test("returns an actionable error when wizard completion fails", async () => {
  const originalConsoleError = console.error;
  console.error = () => undefined;
  try {
    const error = await completeWizardSafely(
      async () => {
        throw new Error("save failed");
      },
      { ready: true }
    );

    assert.equal(error, WIZARD_COMPLETION_ERROR);
  } finally {
    console.error = originalConsoleError;
  }
});
