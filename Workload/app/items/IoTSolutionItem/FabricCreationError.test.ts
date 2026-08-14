// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import assert from "node:assert/strict";
import test from "node:test";
import {
  getFabricCreationErrorMessage,
  hasFabricErrorCode,
  parseFabricErrorPayload,
} from "./FabricCreationError.ts";

test("maps the Fabric duplicate display-name error to an actionable message", () => {
  const message = getFabricCreationErrorMessage(
    409,
    {
      requestId: "7bc444bd-8715-435a-a311-a3cd1fc4bbe9",
      errorCode: "ItemDisplayNameAlreadyInUse",
      message: "Requested 'IoT Device Dashboard' is already in use",
      isRetriable: false,
    },
    "Dashboard",
    "IoT Device Dashboard"
  );

  assert.equal(
    message,
    'A dashboard named "IoT Device Dashboard" already exists in this workspace. Choose a different name and try again.'
  );
});

test("maps nested asynchronous operation errors", () => {
  const message = getFabricCreationErrorMessage(
    200,
    {
      status: "Failed",
      error: {
        code: "Conflict",
        message: "An item with this display name already exists.",
      },
    },
    "Data Agent",
    "Device Health"
  );

  assert.equal(
    message,
    'A data agent named "Device Health" already exists in this workspace. Choose a different name and try again.'
  );
});

test("keeps permission failures distinct from other creation errors", () => {
  assert.match(
    getFabricCreationErrorMessage(403, {}, "Eventstream", "Device Events"),
    /permission to create it/
  );
});

test("maps duplicate connection names explicitly", () => {
  assert.equal(
    getFabricCreationErrorMessage(
      409,
      {
        errorCode: "DuplicateConnectionName",
        message: "The connection DisplayName input is already being used by another connection",
      },
      "Connection",
      "KQL-iot-eh-3fz92"
    ),
    'A connection named "KQL-iot-eh-3fz92" already exists and could not be recovered. Remove the unused connection in Fabric and try again.'
  );
});

test("detects top-level and nested Fabric error codes", () => {
  assert.equal(
    hasFabricErrorCode({ errorCode: "DuplicateConnectionName" }, "DuplicateConnectionName"),
    true
  );
  assert.equal(
    hasFabricErrorCode({ error: { code: "Conflict" } }, "Conflict"),
    true
  );
});

test("preserves the public Fabric message for unmapped errors", () => {
  assert.equal(
    getFabricCreationErrorMessage(
      400,
      {
        errorCode: "InvalidDefinition",
        message: "The dashboard definition is invalid.",
      },
      "Dashboard",
      "Device Health"
    ),
    "Fabric could not create the dashboard: The dashboard definition is invalid."
  );
});

test("parses JSON and plain-text Fabric error responses", () => {
  assert.deepEqual(
    parseFabricErrorPayload('{"errorCode":"Example","message":"Details"}'),
    { errorCode: "Example", message: "Details" }
  );
  assert.deepEqual(parseFabricErrorPayload("Service unavailable"), {
    message: "Service unavailable",
  });
});
