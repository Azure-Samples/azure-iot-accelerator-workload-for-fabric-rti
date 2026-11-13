// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const EnvironmentConstants = {
    FabricApiBaseUrl: "https://api.fabric.microsoft.com",
    OneLakeDFSBaseUrl: "https://onelake.dfs.fabric.microsoft.com",
} as const;
export type EnvironmentConstants = typeof EnvironmentConstants[keyof typeof EnvironmentConstants];





