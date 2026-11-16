// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { AccessToken, WorkloadClientAPI } from "@ms-fabric/workload-client";
import { EnvironmentConstants } from "../../constants";
import { callAcquireFrontendAccessToken } from "../../controller/AuthenticationController";
import { Item } from "../../clients/FabricPlatformTypes";
import { FabricPlatformError } from "../../clients/FabricPlatformClient";
import { parseFabricErrorPayload } from "./FabricCreationError";
import { isJsonObject } from "./ModelAuthoring.ts";

const EVENTHOUSE_READ_SCOPE = "https://api.fabric.microsoft.com/Item.Read.All";

/** Connection metadata exposed by an Eventhouse item. */
export interface EventhouseItemProperties {
  queryServiceUri: string;
  ingestionServiceUri: string;
  databasesItemIds: string[];
}

/** An Eventhouse item together with its connection metadata. */
export interface EventhouseItemMetadata extends Item {
  properties: EventhouseItemProperties;
}

/**
 * Fetches an Eventhouse item's metadata (query/ingestion service URIs and the ids of
 * its KQL databases) from the Fabric REST API. Request failures retain their cause.
 */
export async function getEventhouseItem(
  workloadClient: WorkloadClientAPI,
  workspaceId: string,
  eventhouseId: string
): Promise<EventhouseItemMetadata> {
  try {
    const accessToken: AccessToken = await callAcquireFrontendAccessToken(
      workloadClient,
      EVENTHOUSE_READ_SCOPE
    );
    const response = await fetch(
      `${EnvironmentConstants.FabricApiBaseUrl}/v1/workspaces/${workspaceId}/eventhouses/${eventhouseId}`,
      {
        method: "GET",
        headers: {
          Authorization: "Bearer " + accessToken.token,
          "Content-Type": "application/json",
        },
      }
    );

    if (!response.ok) {
      const requestId = response.headers.get("x-ms-request-id") || response.headers.get("request-id") || undefined;
      let payload: unknown;
      try {
        payload = parseFabricErrorPayload(await response.text());
      } catch (error: unknown) {
        throw new FabricPlatformError(response.status, response.statusText, undefined, requestId,
          error instanceof Error ? error : undefined);
      }
      const details = isJsonObject(payload) && isJsonObject(payload.error) ? payload.error : payload;
      const errorResponse = isJsonObject(details) && typeof details.message === "string"
        ? { error: { code: typeof details.errorCode === "string" ? details.errorCode
          : typeof details.code === "string" ? details.code : "EventhouseReadFailed", message: details.message } }
        : undefined;
      throw new FabricPlatformError(response.status, response.statusText, errorResponse, requestId);
    }

    return (await response.json()) as EventhouseItemMetadata;
  } catch (error) {
    console.error("GetEventhouseItem error:", error);
    throw error;
  }
}
