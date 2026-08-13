/**
 * Interface representing the definition of an IoT Solution item.
 * This information is stored in Fabric as the Item definition.
 * It tracks the wizard progress and all configured IoT resources.
 */
export interface IoTSolutionItemDefinition {
  /** Current wizard step (persisted so user can resume) */
  wizardStep?: string;

  /** IoT Hub configuration (Step 1) */
  iotHub?: {
    subscriptionId: string;
    resourceGroup: string;
    hubName: string;
    location?: string;
    skuName?: string;
    skuTier?: string;
    hostName?: string;
  };

  /** Eventhouse & database configuration (Step 2) */
  eventhouse?: {
    eventhouseId: string;
    eventhouseName: string;
    databaseId: string;
    databaseName: string;
    queryServiceUri?: string;
    ingestionServiceUri?: string;
  };

  /** KQL tables configuration (Step 3) */
  tables?: {
    telemetryTableName: string;
    propertiesTableName: string;
  };

  /** Eventstream configuration (Step 4) */
  eventstreams?: {
    telemetryStreamId: string;
    telemetryStreamName: string;
    propertiesStreamId: string;
    propertiesStreamName: string;
  };

  /** IoT Hub routing configuration (Step 5) */
  routing?: {
    telemetryRouteName: string;
    propertiesRouteName: string;
    telemetryEndpointName: string;
    propertiesEndpointName: string;
  };

  /** Generated Real-Time Dashboard */
  dashboard?: {
    dashboardId: string;
    dashboardName: string;
  };

  /** Uploaded device model (DTDL) */
  deviceModel?: {
    fileName: string;
    modelName: string;
    modelJson: object | null;
    propsNormalizedName?: string;
    propsLkvViewName?: string;
    modeledDataName?: string;
  };

  /** Activator configuration */
  activator?: {
    activatorId: string;
    activatorName: string;
    eventstreamId: string;
    eventstreamName: string;
    connectionId: string;
  };

  /** Fabric Data Agent – AI-powered natural language query interface */
  dataAgent?: {
    agentId: string;
    agentName: string;
    databaseId: string;
    databaseName: string;
    tableName: string;
  };
}
