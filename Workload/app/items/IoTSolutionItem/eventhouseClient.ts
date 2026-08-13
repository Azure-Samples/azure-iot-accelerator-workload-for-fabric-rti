import { AccessToken, WorkloadClientAPI } from "@ms-fabric/workload-client";
import { EnvironmentConstants } from "../../constants";
import { callAcquireFrontendAccessToken } from "../../controller/AuthenticationController";
import { Item } from "../../clients/FabricPlatformTypes";

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
 * its KQL databases) from the Fabric REST API. Returns null if the item cannot be read.
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
      const errorMessage = await response.text();
      console.error(`GetEventhouseItem failed (${response.status}): ${errorMessage}`);
      return null;
    }

    return (await response.json()) as EventhouseItemMetadata;
  } catch (error) {
    console.error("GetEventhouseItem error:", error);
    return null;
  }
}
