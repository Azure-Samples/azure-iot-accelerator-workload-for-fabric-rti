import { WorkloadClientAPI } from "@ms-fabric/workload-client";
import { ItemClient } from "./ItemClient";

/** Fabric API clients used by the IoT Solution workload. */
export class FabricPlatformAPIClient {
  public readonly items: ItemClient;

  constructor(workloadClient: WorkloadClientAPI) {
    this.items = new ItemClient(workloadClient);
  }

  static create(workloadClient: WorkloadClientAPI): FabricPlatformAPIClient {
    return new FabricPlatformAPIClient(workloadClient);
  }
}
