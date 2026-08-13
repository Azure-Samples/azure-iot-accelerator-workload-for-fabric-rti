# Host the workload in Azure and make it available across your tenant

In this deployment option, the workload front end will be hosted in an [Azure Static Web
App](https://learn.microsoft.com/en-us/azure/static-web-apps/) and served from a [custom
domain](https://learn.microsoft.com/en-us/azure/static-web-apps/custom-domain).
The workload manifest will be published through the Fabric admin portal, making it available across
your tenant.

## Prerequisites

- An Azure subscription.
- `Domain Name Administrator` role (or equivalent permissions), to register and verify a custom domain.
- `Application Administrator` role (or application owner permission), to update the Entra app registration with
  the new frontend URL.
- `Fabric Administrator` role, to publish the workload manifest.
- Enable the related [workload
  tenant settings](https://learn.microsoft.com/en-us/fabric/extensibility-toolkit/setup-guide#step-3-enable-developer-features-in-fabric)
  in the Fabric admin portal.
- [Azure CLI](https://learn.microsoft.com/en-us/cli/azure/install-azure-cli) to run the deployment commands below.

## 1. Set up the environment variables and create the resource group

```powershell
# Replace the placeholders below with your own values
$sub        = "<Your Azure subscription ID>"
$rg         = "<Resource group to be created>"
$loc        = "<Deployment region>"
$domain     = "<Desired domain>.com"            # The public domain you want to register and use for the workload front end.
$webappName = "<Name of the Static Web App>"

$frontendHost   = "iot.$domain"
$tenantId       = az account show --query tenantId -o tsv
$workload       = "Org.IoTSolutionAccelerator"
$appId          = (Select-String -Path ".\Workload\.env.prod" -Pattern '^FRONTEND_APPID=(.+)$').Matches[0].Groups[1].Value.Trim()

az account set --subscription $sub
az group create --name $rg --location $loc
```

## 2. Custom domain

Fabric workloads [must use a custom domain that can be registered and verified in your Entra
tenant](https://learn.microsoft.com/en-us/fabric/extensibility-toolkit/publishing-requirements-general#microsoft-entra-requirements).
For this reason, the default Web App URL cannot be used. You can alternativelly use a subdomain of
an existing domain that you already own and have verified in your Entra tenant. If using an existing
domain, you can skip the domain creation and verification steps below.

### 2.1 Create a custom domain

```powershell
# Register the Domain Registration resource provider, if not yet registered in your subscription.
az provider register --namespace Microsoft.DomainRegistration --wait

# Replace your contact information below for the public domain registration.
@'
{
    "address1":    { "value": "<Your address>",         "required": true },
    "city":        { "value": "<Your city>",            "required": true },
    "country":     { "value": "<Your country>",         "required": true, "options": [] },
    "postal_code": { "value": "<Your postal code>",     "required": true },
    "state":       { "value": "<Your state>",           "required": true },
    "email":       { "value": "<Your email>",           "required": true },
    "name_first":  { "value": "<Your first name>",      "required": true },
    "name_last":   { "value": "<Your last name>",       "required": true },
    "phone":       { "value": "<Your phone number>",    "required": true }
}
'@ | Set-Content -Path contact.json -Encoding utf8

# Register the custom domain
az appservice domain create -g $rg --hostname $domain --contact-info "@contact.json" --dryrun
az appservice domain create -g $rg --hostname $domain --contact-info "@contact.json" --accept-term
```

### 2.2 Register and verify the custom domain in your Entra tenant

The steps below can alternatively be done in the [Entra admin
portal](https://learn.microsoft.com/en-us/entra/fundamentals/add-custom-domain).

```powershell
# Install the Microsoft Entra PowerShell module (if not already installed) and connect to your tenant.
Install-Module Microsoft.Entra -Scope CurrentUser -AllowClobber
Connect-Entra -Scopes 'Domain.ReadWrite.All'

# Add the (unverified) domain to the tenant.
New-EntraDomain -Name $domain

# Pull the TXT verification value (MS=msXXXXXXXX) and add it to your DNS zone.
$txt = (Get-EntraDomainVerificationDnsRecord -Name $domain | Where-Object RecordType -eq 'Txt').Text
az network dns record-set txt add-record -g $rg -z $domain -n "@" --value $txt

# Wait a few minutes for the DNS record to propagate, then confirm the domain in your tenant.
Start-Sleep -Seconds 90
Confirm-EntraDomain -Name $domain
Get-EntraDomain -Name $domain | Select-Object Name, IsVerified
```

## 3. Deploy the Static Web App and bind the custom domain

```powershell
cd .\Workload

# Build the front end (using .env.prod).
npm run build:prod

# Create the staticwebapp.config.json file.
@'
{
  "navigationFallback": {
    "rewrite": "/index.html",
    "exclude": ["/assets/*", "*.{js,css,png,jpg,svg,ico,json,txt,woff,woff2,map,config}"]
  }
}
'@ | Set-Content -Path ..\build\Frontend\staticwebapp.config.json -Encoding utf8

# Provision the Static Web App and deploy the front end.
az staticwebapp create -n $webappName -g $rg -l $loc --sku Free
$deployToken = az staticwebapp secrets list -n $webappName -g $rg --query "properties.apiKey" -o tsv
npm install -g @azure/static-web-apps-cli
swa deploy "..\build\Frontend" --deployment-token $deployToken --env production

# Bind the custom domain to the Static Web App and auto-provision TLS.
$webappNameHost = az staticwebapp show -n $webappName -g $rg --query "defaultHostname" -o tsv
az network dns record-set cname set-record -g $rg -z $domain -n "iot" --cname $webappNameHost
az staticwebapp hostname set -n $webappName -g $rg --hostname $frontendHost

# Wait for the TLS certificate to be provisioned and verify that the custom domain is accessible over HTTPS.
Start-Sleep -Seconds 90
Invoke-WebRequest "https://$frontendHost/close" -UseBasicParsing
```

## 4. Reconfigure the Entra app with the new frontend URL
The following steps add the frontend URL to the Redirect URIs of the Entra app. This can
alternatively be done in the [Azure
Portal](https://learn.microsoft.com/en-us/entra/identity-platform/how-to-add-redirect-uri).

```powershell
# Install the Microsoft Graph PowerShell module (if not already installed) and connect to your tenant.
Install-Module Microsoft.Graph.Applications -Scope CurrentUser -AllowClobber
Connect-MgGraph -Scopes 'Application.ReadWrite.All'

# Add the new redirect URI.
$app     = Get-MgApplication -Filter "appId eq '$appId'" | Select-Object -First 1
$spaUris = @($app.Spa.RedirectUris) + "https://$frontendHost/close" | Select-Object -Unique
Update-MgApplication -ApplicationId $app.Id -Spa @{ RedirectUris = $spaUris }

# Verify the updated redirect URIs.
(Get-MgApplication -ApplicationId $app.Id).Spa.RedirectUris
```

## 5. Build the workload manifest package

```powershell
cd ..

# Update the FRONTEND_URL in the .env.prod.
$newProdEnv = (Get-Content ".\Workload\.env.prod" -Raw) -replace '(?m)^FRONTEND_URL=.*', "FRONTEND_URL=https://$frontendHost/"
[System.IO.File]::WriteAllText((Resolve-Path ".\Workload\.env.prod"), $newProdEnv)
Select-String -Path ".\Workload\.env.prod" -Pattern '^FRONTEND_URL=' # Verify the updated FRONTEND_URL

# Build the workload manifest package.
pwsh .\scripts\Build\BuildManifestPackage.ps1 -Environment prod
```

## 6. Upload to Fabric admin portal
Once the workload manifest package is built, you can [upload it to the Fabric admin
portal](https://learn.microsoft.com/en-us/fabric/extensibility-toolkit/tutorial-publish-workload):
- Navigate to `Fabric portal -> gear icon -> Admin portal -> Workloads -> Publish`.
- Upload the workload manifest package under `build\Manifest\Org.IoTSolutionAccelerator.1.0.0.nupkg`.
- In the same page, select the newly uploaded workload. In the right panel, select a version then
  click `Add`.

The workload will now be available across your tenant. You can verify it by navigating to the
Workload Hub page.