// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  completeWizardSafely,
  getPersistableWizardContext,
  WIZARD_COMPLETION_ERROR,
} from "./wizardUtils.ts";

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
});

test("returns no error when wizard completion succeeds", async () => {
  const error = await completeWizardSafely(async () => undefined, { ready: true });

  assert.equal(error, undefined);
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
