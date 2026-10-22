---
name: deploy-iot-workload
description: Build and deploy this Azure IoT Solution Accelerator workload to Azure Static Web Apps and publish it to a Microsoft Fabric tenant. Use for first-time deployment, tenant activation, or resuming an incomplete deployment. Guide the user through the existing repository documentation with state checks and approval gates.
---

# Deploy the IoT workload to Azure and Fabric

Guide an interactive deployment of this repository's existing workload. The
outcome is a working Azure-hosted frontend and an activated Fabric tenant
workload.

## Source of truth

Start with [README setup and deployment](../../../README.md#setup-and-deployment).
Load the relevant guide as its step approaches:

- [Entra app guide](../../../docs/create-entra-app.md): registration, delegated
  permissions, and admin consent.
- [Azure hosting and tenant publishing guide](../../../docs/publish-to-azure-and-tenant.md):
  prerequisites and the numbered deployment, DNS, packaging, and activation steps.
- [Local mode](../../../docs/local-mode.md): only when local testing is requested.

Use those guides for the baseline commands, contact schema, and permission
identifiers. The notes below supplement them. Reuse existing scripts and documented interfaces;
explain a verified blocker before switching approaches.

## Preflight and approvals

- Reuse answers and inspect existing configuration before asking for missing
  inputs, one decision at a time. Confirm the Fabric tenant, Azure hosting
  tenant/subscription, region, resource group, workload name/version, and hostname.
- Verify each CLI, Graph, and browser session before using it. Prefer explicit
  subscription/tenant arguments over mutable global defaults. Select the hosting
  subscription before reading its tenant; do not derive the Fabric tenant from
  that subscription. Pass the intended Fabric tenant to `Connect-MgGraph` and
  `Connect-Entra` using `-TenantId`, and check the connected context.
- Discover existing apps, verified domains, DNS authority, hosting resources,
  and published versions. Confirm reuse rather than creating duplicates. Do
  not assume Azure DNS hosts a domain merely because the guide uses its commands.
- Check the current step's documented tools, permissions, resource providers,
  and tenant settings. Install or enable prerequisites only when needed and
  approved. In a fresh hosting subscription, inspect `Microsoft.Web` and
  `Microsoft.Network`; inspect `Microsoft.DomainRegistration` only if purchasing
  a domain. Register missing providers with approval and verify registration.
  A failed Azure DNS resource query is not proof of a wrong subscription; inspect
  its actual error and provider registration before choosing a corrective action.
- Present the proposed operations and costs before writes. Reuse approval for
  an unchanged scope; ask again when scope, cost, or security impact changes.
  Explicitly confirm domain purchase/renewal, admin consent, tenant activation,
  and destructive changes. Do not assign roles or replace resources implicitly.
- Use an organizationally permitted sign-in method; never request passwords or
  pasted tokens, or relax policy to bypass a blocked flow. Keep deployment tokens
  out of output and persistent credentials. Leave generated configuration and
  build artifacts uncommitted.

### Fresh-tenant API consent

Before admin consent, check that each API in the Entra guide's
`$requiredResourceAccess` has a service principal in the Fabric tenant. Use each
entry's `ResourceAppId` in a `Get-MgServicePrincipal -Filter "appId eq '...'"` query;
with approval, use `New-MgServicePrincipal -AppId` only for missing entries. Then grant and verify
exactly the documented delegated scopes, allowing bounded time for replication.
The missing Azure Data Explorer enterprise application can cause a "not
subscribed to the service" error. This is a directory prerequisite, not a reason
to create an ADX cluster or change the requested permissions.

### Only when purchasing a domain

Prefer an existing suitable domain. If a new purchase is chosen, use the
publishing guide's contact schema. Reuse supplied values or a supplied file;
ask for each missing or invalid required field individually using the available
user-input tool. Fill the JSON file yourself rather than handing the user a
blank template. Ask for optional fields only when needed.

Use a two-letter country code and the CLI's `+countrycode.number` phone format.
Validate required formats, escape values as data, and do not invent registrant
details or substitute example values. Save temporary contact data in
a private, access-restricted local file outside the checkout and remove copies
created for this run when no longer needed. Do not delete a user-supplied source
file. Keep contacts out of the skill, shared logs, checkpoints, and durable memory;
offer a local file if the user prefers not to share them in chat.

Filter dry-run output with a `--query` allowlist: `hostname`, `resource_group`,
`hostname_available`, `price`, `auto_renew`, and `privacy`. The unfiltered result
contains contact details. Confirm the domain, current price, terms, renewal,
privacy, and [subscription eligibility](https://learn.microsoft.com/en-us/azure/app-service/manage-custom-dns-buy-domain#prerequisites)
before purchasing; annual registration and Azure DNS are billed separately.
Availability is not purchase consent.

Use the full `--accept-terms` option after reviewing `az appservice domain
show-terms`. The current CLI's default-enabled `--auto-renew` switch disables
renewal when supplied; `--auto-renew false` is rejected. Check the installed
CLI's help and dry run, apply the same
renewal choice to purchase, then verify the resource's actual `autoRenew` and
registration state. Do not repeat a purchase command without inspecting state.

## Execute and verify

Follow the guides in order. A gate already completed by the user can be skipped
after verifying its current state; do not repeat uploads, purchases, or consent.

| Gate | Reference | Exit evidence |
| --- | --- | --- |
| App and setup | Entra guide; README setup | Correct Fabric tenant/app, documented delegated consent, and generated configuration matching the chosen workload identity. |
| Domain and Azure host | Publishing sections 1-3 | Owned domain verified in the Fabric tenant; production build deployed; valid custom-domain TLS and expected page, assets, and `/close` content. |
| Hosted callback | Publishing section 4 | Exact hosted SPA redirect registered, preserving existing redirects and app settings. |
| Production package | Publishing section 5 | Fresh validated package inspected for matching app ID, HTTPS URL, workload name/version, and no unresolved placeholders. |
| Fabric activation | Publishing section 6 | Exact uploaded version confirmed as `Active in tenant`, not merely listed as uploaded. |
| Smoke check | README usage | User-approved workspace/item; workload discoverable, rendered from the Azure hostname, and no sign-in error. |

## Repository-specific guardrails

- Resolve paths from the repository root and set each command's working directory
  and variables explicitly; separate tool calls may not share shell state.
- Try the documented setup before proposing new flags. Its
  [implementation](../../../scripts/Setup/SetupWorkload.ps1) also prepares
  development configuration and can reuse `FABRIC_DEV_WORKSPACE_GUID`. Confirm
  an inherited workspace before using local mode. Do not invent an ID, force
  an overwrite, or start the dev server/gateway just to publish.
- The frontend build cleans its output. Generate the documented Static Web Apps
  routing configuration afterward. Static Web Apps does not apply IIS
  `web.config`: read the `Content-Security-Policy` header from
  `Workload\app\web.config` and copy its value to
  `globalHeaders.Content-Security-Policy` in
  `build\Frontend\staticwebapp.config.json`, retaining the guide's fallback.
  Derive the policy from the file rather than hardcoding another list of origins.
  Do not replace the required custom domain with an autogenerated hostname or
  an `onmicrosoft.com` domain.
- Capture the documented deployment-token lookup without printing it. Supply
  `SWA_CLI_DEPLOYMENT_TOKEN` only to the deployment process instead of exposing
  the token in command-line arguments. Use `--no-use-keychain` for this token-based
  deployment and clear the temporary environment value in a `finally` block.
- Preserve frontend-only `FERemote`/`AADFEApp` and SPA authentication. Do not add
  a backend, client secret, or cross-tenant marketplace configuration based on
  unrelated toolkit examples.
- Use the existing
  [manifest builder](../../../scripts/Build/BuildManifestPackage.ps1).
  From the repository root in PowerShell 7+, invoke
  `& .\scripts\Build\BuildManifestPackage.ps1 -Environment prod -ValidateFiles $true`
  so its Boolean validation parameter is enabled. Derive identity/version from
  `.env.prod`, not `Workload/package.json` or an example filename. On macOS/Linux,
  the current builder additionally requires Mono. Inspect the package itself.
- Check `$LASTEXITCODE` after native commands and verify artifacts, not success
  banners alone; setup can continue after a failed npm step. Stop on unresolved
  failure rather than allowing later output to mask it.
- Reuse the authenticated Fabric browser and verify its tenant. If the UI
  differs, consult the linked
  [current publishing guide](https://learn.microsoft.com/en-us/fabric/extensibility-toolkit/tutorial-publish-workload)
  and visible controls. Provide the exact package path when a file picker needs
  user input, then verify upload and activation separately. Complete both the
  version selector's Add action and the final Add confirmation. Blank
  version/status fields are not activation evidence. Reuploading a replacement
  package needs a new agreed version; do not delete an active workload to bypass
  a conflict.
- Obtain approval before creating a smoke-test item. Check that no development
  gateway overrides the Azure-hosted iframe. Do not run ingestion, modeling,
  dashboard, Activator, or Data Agent provisioning as part of this smoke check.

### Hostname validation and recovery

Do not treat the guide's fixed sleep as success. Request hostname binding with
`--no-wait` and poll `az staticwebapp hostname list` for the selected hostname's
`status` and `errorMessage`, with a bounded deadline. Verify `Ready` and actual
HTTPS page, assets, and `/close` content with certificate validation enabled.
Report a failed or still-pending state instead of bypassing TLS or looping forever.
An unexpected CLI error, including an HTTP `OK` polling error, requires a live
state check before any retry.

If CNAME validation fails despite the correct publicly resolvable CNAME, use
`--validation-method dns-txt-token --no-wait`. Read the binding's current
`validationToken` and publish it at `_dnsauth.<frontend-label>` in the DNS zone
(`_dnsauth.iot` for the guide's hostname), preserving the routing CNAME. This is
separate from the apex TXT challenge for Entra domain verification. Follow
[Static Web Apps TXT guidance](https://learn.microsoft.com/en-us/azure/static-web-apps/custom-domain#zero-downtime-migration)
if record placement differs.

A persistently failed binding can require removal and recreation. Obtain
approval for removing only that failed hostname binding, never the working
site, purchased domain, or DNS zone. Publish a changed token after recreation
and allow DNS propagation before retrying. An empty token after `Ready` is
normal; do not restart validation. Remove obsolete proof values only with
approval and without deleting current or unrelated records.

## Finish or resume

Record only non-sensitive deployment state in session-local storage: completed
gates/evidence, resource identifiers, package path/version, and the next action.
On resumption, recheck mutable state before continuing. If blocked, state the
actual error, pending gate, and administrator/user action needed; do not silently
substitute defaults or claim a partial deployment is complete. Report the live
HTTPS endpoint and confirmed Fabric version, distinguishing verified results
from user-reported or pending steps.
