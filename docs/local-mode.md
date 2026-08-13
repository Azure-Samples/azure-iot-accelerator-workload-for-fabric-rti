# Run the workload in local mode
In this mode, the workload runs from your local machine, connected to a specific Fabric workspace.
You can use this mode to test the workload before deploying it to Azure and across your tenant.

## Prerequisites

- `Fabric Administrator` role, to enable the workload tenant settings and developer mode (first time setup).
- A Fabric [workspace](https://learn.microsoft.com/en-us/fabric/fundamentals/create-workspaces).

## 1. Enable workload tenant settings

In the Fabric admin portal, under `Tenant settings`, [enable workload
options](https://learn.microsoft.com/en-us/fabric/extensibility-toolkit/setup-guide#step-3-enable-developer-features-in-fabric).

## 2. Enable developer settings

In the Fabric portal, enable [workload developer
settings](https://learn.microsoft.com/en-us/fabric/extensibility-toolkit/setup-guide#step-3-enable-developer-features-in-fabric).

## 3. Run the development environment setup script

Run the local development setup script, replacing the placeholder with your [Fabric workspace
Id](https://learn.microsoft.com/en-us/fabric/data-factory/migrate-pipelines-how-to-find-your-fabric-workspace-id).

```powershell
pwsh ./scripts/Setup/SetupDevEnvironment.ps1 `
    -DevWorkspaceId "<Fabric workspace Id>"
```

## 4. Start the dev server and gateway

In two separate terminal windows, start the development server and the development gateway.

```powershell
# In terminal window 1.
pwsh .\scripts\Run\StartDevServer.ps1
```

```powershell
# In terminal window 2 (will prompt for authentication).
pwsh .\scripts\Run\StartDevGateway.ps1
```

Once both server and gateway start, the workload will be available in the workspace specified in the
setup script in the previous step.