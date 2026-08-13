# Create the Microsoft Entra application for the workload

The workload requires a [Microsoft Entra application
registration](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app)
(the workload front end app). The steps below create the application and grant the
delegated permissions needed by the workload.

## Prerequisites

- Permission to register applications in your tenant (`Application Developer` or
  `Application Administrator` role).
- `Global Administrator` or `Privileged Role Administrator` role, to grant tenant-wide admin
  consent.

## 1. Set up the environment variables and connect to Microsoft Graph

```powershell
# Replace this placeholder with the display name for your application.
$applicationName = "<Display name for the workload app registration>"

# Install the Microsoft Graph PowerShell module (if not already installed) and connect.
Install-Module Microsoft.Graph.Applications -Scope CurrentUser -AllowClobber
Connect-MgGraph -Scopes 'Application.ReadWrite.All'

$tenantId        = (Get-MgContext).TenantId
$workload        = "Org.IoTSolutionAccelerator"
```

## 2. Create the application registration

```powershell
# Redirect URIs
$redirectUris = @(
    "http://localhost:60006/close",
    "https://app.powerbi.com/workloadSignIn/$tenantId/$workload",
    "https://app.fabric.microsoft.com/workloadSignIn/$tenantId/$workload"
)

# Application ID URI (the format is specified by Fabric)
$suffix        = -join ((65..90) + (97..122) | Get-Random -Count (Get-Random -Minimum 1 -Maximum 6) |
                    ForEach-Object { [char]$_ })
$identifierUri = "api://localdevinstance/$tenantId/$workload/$suffix"

# Delegated API permissions:
#   Power BI Service          -> Fabric.Extend       (required by every Fabric workload)
#   Azure Data Explorer       -> user_impersonation  (query device data in Kusto/ADX)
#   Azure Service Management  -> user_impersonation  (call Azure Resource Manager)
$requiredResourceAccess = @(
    @{
        ResourceAppId  = "00000009-0000-0000-c000-000000000000"                       # Power BI Service
        ResourceAccess = @( @{ Id = "7ba630b9-8110-4e27-8d17-81e5f2218787"; Type = "Scope" } )  # Fabric.Extend
    },
    @{
        ResourceAppId  = "2746ea77-4702-4b45-80ca-3c97e680e8b7"                       # Azure Data Explorer
        ResourceAccess = @( @{ Id = "00d678f0-da44-4b12-a6d6-c98bcfd1c5fe"; Type = "Scope" } )  # user_impersonation
    },
    @{
        ResourceAppId  = "797f4846-ba00-4fd7-ba43-dac1f8f63013"                       # Azure Service Management
        ResourceAccess = @( @{ Id = "41094075-9dad-400e-a0bd-54e686782033"; Type = "Scope" } )  # user_impersonation
    }
)

$app = New-MgApplication `
    -DisplayName $applicationName `
    -SignInAudience "AzureADMyOrg" `
    -IdentifierUris @($identifierUri) `
    -Spa @{ RedirectUris = $redirectUris } `
    -OptionalClaims @{ AccessToken = @(@{ Name = "idtyp"; Essential = $false }) } `
    -RequiredResourceAccess $requiredResourceAccess

# Create the matching service principal (enterprise application) in your tenant.
New-MgServicePrincipal -AppId $app.AppId | Out-Null

$appId = $app.AppId

# Note this value for the rest of the setup, as it will be needed for subsequent steps.
Write-Host "Application Id (FrontendAppId): $appId"
```

## 3. Grant tenant-wide admin consent
Sign in as a tenant administrator and grant admin consent. All requested permissions are delegated, so admin
consent lets the workload's users sign in without being prompted to consent individually. This can
alternatively be done in the [Azure
Portal](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/grant-admin-consent):
`Microsoft Entra ID -> Enterprise applications -> your app -> Permissions -> Grant admin consent`.

```powershell
# Wait a few seconds for the app to propagate, then open the admin consent page.
Start-Sleep -Seconds 30
Start-Process "https://login.microsoftonline.com/$tenantId/adminconsent?client_id=$appId"
```
