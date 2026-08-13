// Main API Client Infrastructure
export { FabricPlatformClient } from './FabricPlatformClient';
export { FabricPlatformAPIClient } from './FabricPlatformAPIClient';
export { FabricAuthenticationService } from './FabricAuthenticationService';
export * from './FabricPlatformTypes';

// API client used by the workload
export { ItemClient as ItemController } from './ItemClient';
export { LongRunningOperationsClient } from './LongRunningOperationsClient';

// Re-export WorkloadClientAPI for convenience
export { WorkloadClientAPI } from '@ms-fabric/workload-client';
