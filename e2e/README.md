# Fabric IoT end-to-end tests

This suite provisions Azure IoT test resources, configures the Azure IoT Solution
Accelerator through its Fabric UI, and verifies the resulting data and experiences.
Run it from this `e2e` folder using Node.js, npm, Playwright and the Azure SDKs.

## What the suite covers

| Area | Checks |
| --- | --- |
| Setup | A fresh workload item, five S1 IoT Hubs, managed identities, and three linked Hubs across two ADR namespaces and DPS instances. |
| Ingestion | Nonlinked, combined and linked-only sources; ADR prerequisites, routing, raw telemetry/properties, component values and negative controls. |
| Modeling | DTDL and WoT uploads, support guidance and documentation links, device assignments, typed/enriched KQL values, and model-reference and data-condition filters. |
| Dashboards | Generated dashboards, source bindings, visible device/component values and chart data. |
| Data Agents | Grounding in modeled tables, persisted component/independent latest-value examples, and a new answer to a fixture-specific question. |
| Activator | Generated feed/schema bindings and actual incoming device rows. |
| Cleanup | Removal and absence checks for the run-owned Azure and Fabric resources. |

The full run has 20 stages. Availability checks and local harness self-tests are
separate commands. For stable scenario IDs and implementation details, see the
[scenario catalog](src/scenarios.ts), [runner](src/run.ts), and
[scenario tests](tests/scenarios). Bundled device models are described in the
[fixture guide](fixtures/README.md).

## Costs and test isolation

**Live commands create billable infrastructure. Use a dedicated test subscription
and an empty, dedicated Fabric workspace, not a production environment.**

A full run creates the following resources:

| Location | Resources |
| --- | --- |
| Azure | One resource group, five S1 IoT Hubs with one unit each, two Device Provisioning Service (DPS) instances, two Azure Device Registry (ADR) namespaces, and one user-assigned managed identity. |
| Access configuration | Resource-scoped Azure role assignments and Fabric Contributor access for the test Hub identities. |
| Fabric | A run folder and workload item, Eventhouse/KQL databases, raw tables and six ingestion Eventstreams, modeled tables/views, three dashboards and three Data Agents. |
| Activator | Three scenario-specific Activator items, feed Eventstreams and KQL connections; these are temporary under the default cleanup policy. |
| Local machine | SDK device simulators, a dedicated browser profile, recordings, reports and caches. Simulators run locally rather than in Azure containers. |

Azure service charges and Fabric capacity consumption depend on region, usage and
run duration. Check your subscription pricing and capacity before running.
Resources retained for troubleshooting continue to consume resources and may
incur charges.

Names and tags are randomized per run. Azure resources share the run's resource
group; Fabric items are collected in its folder. Cleanup validates the saved
ownership manifest. The supplied workspace, its capacity and workspace identity
remain in place after cleanup.

## Prerequisites and permissions

| Requirement | Details |
| --- | --- |
| Runtime | Node.js **22.18 or newer**, npm, and Azure CLI (`az`) available in the same terminal. |
| Browser | Installed Microsoft Edge, or an installed Playwright Chromium browser. Live tests require an interactive desktop. |
| Azure access | Permission to create/delete resource groups, IoT Hubs, DPS, ADR namespaces and managed identities, and to create/delete role assignments in the test subscription. |
| Azure roles | Subscription **Owner**, or **Contributor** plus **User Access Administrator** or **Role Based Access Control Administrator** with the required assignment permissions. Contributor alone cannot create role assignments; an access-administrator role alone does not provide resource-provisioning permissions. |
| Azure availability | Sufficient quota for five S1 Hubs, two DPS instances and two ADR namespaces. The selected subscription/region must support the ADR-linked Hub/DPS features used by the suite. Register `Microsoft.Devices`, `Microsoft.DeviceRegistry` and `Microsoft.ManagedIdentity` before running. |
| Fabric | An active capacity and a workspace with the accelerator, Eventhouse/Eventstreams, Data Agent, Activator and connection features available to your account. |
| Fabric role | **Workspace Admin** is recommended for the full suite, which grants Hub identity access and may initialize a workspace identity. Member can manage lower workspace roles, but an Admin must create/configure the workspace identity beforehand if you use Member. |
| Workload | Install and enable the accelerator in the tenant/workspace. For a published run, deploy the intended frontend and workload package and turn Fabric developer mode off. |
| Authentication/network | Azure CLI and Fabric sign-ins targeting the same Microsoft Entra tenant, Fabric MFA, and HTTPS access to Azure/Fabric. Device clients use MQTT over WebSockets on TLS port 443. |

See [Azure role-assignment permissions](https://learn.microsoft.com/en-us/azure/role-based-access-control/role-assignments-portal),
[Fabric workspace roles](https://learn.microsoft.com/en-us/rest/api/fabric/core/workspaces/add-workspace-role-assignment),
and [workspace identity requirements](https://learn.microsoft.com/en-us/fabric/security/workspace-identity).
Tenant administrators may need to approve workload installation, consent and
feature settings before a runner can use them.

For workload installation/deployment, follow the
[repository setup guide](../README.md#setup-and-deployment).

## Setup

The examples use PowerShell. In other shells, set the same environment variables.
Start from the repository root with Node.js 22.18 or newer:

```powershell
Set-Location e2e
node --version
npm ci
$env:PLAYWRIGHT_BROWSERS_PATH = Join-Path $PWD 'artifacts\browsers'
npm run install:video
```

`install:video` downloads the FFmpeg binary used for recordings. Edge is the
default browser. To use Chromium instead:

```powershell
$env:E2E_BROWSER = "chromium"
npx playwright install chromium
```

If you set a custom `PLAYWRIGHT_BROWSERS_PATH`, use the same value when installing
browsers/video support and when running the suite.

Obtain tenant and subscription GUIDs from the Azure portal and the workspace GUID from its Fabric
URL (`/groups/<workspace-guid>`), then configure your test target:

```powershell
$env:E2E_TENANT_ID = "<tenant-guid>"
$env:E2E_SUBSCRIPTION_ID = "<subscription-guid>"
$env:E2E_WORKSPACE_ID = "<workspace-guid>"
$env:E2E_LOCATION = "australiaeast"
$env:E2E_DEPLOYMENT = "published"
$env:E2E_ITEM_LABEL = "Azure IoT Solution Accelerator (preview)"
$env:E2E_RETENTION = "cleanup"
$env:E2E_KEEP_RESOURCES = "0"
```

Use the exact installed workload card label for `E2E_ITEM_LABEL`; include a
preview suffix only if it is displayed in your catalog.

### Settings

The suite reads the current process environment. `.env.example` is a reference;
an `.env` file is not loaded automatically.

| Setting | Required | Default | Meaning |
| --- | --- | --- | --- |
| `E2E_TENANT_ID` | Yes | None | Tenant GUID used for Azure CLI token validation and the Fabric tenant hint. |
| `E2E_SUBSCRIPTION_ID` | Yes | None | Azure subscription GUID where the test resource group is created. |
| `E2E_WORKSPACE_ID` | Yes | None | GUID of the dedicated Fabric workspace. |
| `E2E_LOCATION` | Yes | None | Azure region name, for example `australiaeast`; choose a supported region. |
| `E2E_DEPLOYMENT` | Yes | None | `published` or `development`. |
| `E2E_FABRIC_URL` | No | `https://app.fabric.microsoft.com/groups/<workspace-guid>` | HTTPS `/groups/<workspace-guid>` root on `app.fabric.microsoft.com` or `app.powerbi.com`, without extra page paths/query parameters. The optional `ctid` must match the tenant; the hint is otherwise added automatically. |
| `E2E_BROWSER` | No | `msedge` | `msedge` for installed Edge, or `chromium` for the installed Playwright browser. |
| `E2E_ITEM_LABEL` | No | `Azure IoT Solution Accelerator` | Exact accessible label of the installed workload card. |
| `E2E_RETENTION` | No | `cleanup` | `cleanup` removes run-owned resources automatically; `retain` preserves them. Saved target settings must match when continuing or cleaning a run. |
| `E2E_KEEP_RESOURCES` | No | Unset (`0` effect with `cleanup`) | `1` keeps resources for an invocation; `0` does not override `E2E_RETENTION=retain`. |
| `E2E_RUN_ID` | No | A new UUID | Selects an existing local manifest for staged commands or continuation. Leave unset for a fresh full run. |

Sign in with the account that has the permissions above. If an Azure CLI sign-in
is needed:

```powershell
az login --tenant $env:E2E_TENANT_ID
az account show --subscription $env:E2E_SUBSCRIPTION_ID
```

Fabric authentication uses `artifacts\browser-profile`. Sign-in/MFA happens
before workflow recording. Use one live command at a time with this profile.

### Development runs

For local development, start and bind the workload frontend/gateway using the
repository's setup instructions before testing. Configure the test profile once:

```powershell
$env:E2E_DEPLOYMENT = "development"
npm run auth
npm run configure:development
```

For published testing, use `E2E_DEPLOYMENT=published` and ensure developer mode is
off for the test account/profile. Publishing the package and deploying frontend
assets are separate steps; a Git push alone does not update the hosted UI.

## Running

### Full run

```powershell
Remove-Item Env:E2E_RUN_ID -ErrorAction SilentlyContinue
npm run e2e
```

Allow roughly 90-110 minutes, including two ten-minute RBAC propagation waits
during linked-Hub setup. Complete authentication prompts, then let the suite
operate its browser. To interrupt, press Ctrl+C once and wait for cleanup to
finish. Keep the terminal, machine and desktop session available through cleanup.
Forced termination, sleep, hibernation or a lost desktop session can still leave
cleanup unfinished.

A fresh, unfiltered full pass requires all 20 stages to pass, including cleanup.
Playwright retries are disabled.

### Runner options

Options go after npm's `--` separator:

| Option | Effect |
| --- | --- |
| `--nonlinked-only` | Runs the nonlinked/DTDL path and its generated experiences. Linked and other-path stages are marked skipped. The five-Hub fixture is still provisioned. |
| `--ingestion-only` | Runs ingestion/setup/linking and cleanup, omitting modeling and generated experiences. |
| `--extensions-only` | Requires a retained run with verified nonlinked data and skips that ingestion phase. Other selected phases still execute. Combine with `--nonlinked-only` to restrict it to the nonlinked model/experiences, provided their creation has not already been attempted. |
| `--keep-resources` | Keeps run-owned resources on success/failure for inspection; see [cleanup](#cleanup). |

```powershell
npm run e2e -- --nonlinked-only
npm run e2e -- --ingestion-only
npm run e2e -- --keep-resources
```

A filtered run is subset coverage, not a full 20-stage validation.

### Individual scenarios and continuation

Use the run ID printed in the console. Its manifest is at
`artifacts\runs\<run-id>\manifest.json`.

```powershell
$env:E2E_RUN_ID = "<run-guid>"
npm run test:modeling:combined
```

| Command | Purpose |
| --- | --- |
| `npm run auth` | Fabric sign-in only. |
| `npm run preflight` | Checks the selected deployment mode and workload catalog/creation dialog without submitting item creation. |
| `npm run plan` | Writes a new local manifest without provisioning resources. |
| `npm run test:item` | Creates the run's Fabric folder/workload item; needs an unused plan or creates one if no run ID is set. |
| `npm run provision` | Provisions/verifies the five Hubs and test identity. |
| `npm run provision:linked` | Creates/verifies ADR/DPS links for the provisioned linked Hubs. |
| `npm run test:ingestion` | Nonlinked ingestion and native data checks. |
| `npm run test:combined` / `npm run test:linked-only` | Combined or linked-only ingestion after the prerequisite stages. |
| `npm run test:modeling:nonlinked` / `npm run test:modeling:combined` / `npm run test:modeling:linked-only` | Modeling for a path whose ingestion has passed. |
| `npm run test:dashboards` / `npm run test:data-agents` / `npm run test:activator` | All three scenarios for that experience; requires the corresponding path data/modeling. |
| `npm run cleanup` | Explicit cleanup of the selected run. |

For staged execution, follow the full runner's order: item, Hub provisioning,
nonlinked ingestion, linked provisioning, combined ingestion, linked-only
ingestion, modeling, then generated experiences. Individual scenario commands
leave the overall fixture in place; Activator scenarios remove their temporary
feed resources under the default policy. Finish with `npm run cleanup`.

`E2E_RUN_ID` selects a manifest; it is **not** an automatic resume-from-failure
checkpoint. The full runner skips an already recorded item but still executes its
other selected phases. Repeating creation against retained tables/items can
cause name collisions. Continue only phases whose prerequisites passed and whose
creation has not already been attempted. If a create request has an uncertain
outcome, inspect that run before trying another create.

Keep the original tenant, subscription, workspace URL/ID, region, item label,
deployment and retention settings when using a saved manifest. For a new full
validation, unset `E2E_RUN_ID`.

## Results and artifacts

The console prints numbered START/PASS/FAIL stages and a final results tree.
`SKIPPED` means excluded by scope or retention; `NOT RUN` means blocked by an
earlier failure. A nonzero exit code indicates a test, setup or cleanup failure.

| Path | Contents |
| --- | --- |
| `artifacts\reports\run-<report-id>\summary.json` | Overall run ID, effective options, stage results and failures. |
| `artifacts\reports\run-<report-id>\results.junit.xml` | Overall JUnit results, including setup/cleanup failures. |
| `artifacts\reports\playwright-<timestamp>-<pid>\html\index.html` | Per-phase HTML report with detailed steps and attachments; JSON/JUnit files are alongside `html`. |
| `artifacts\runs\<run-id>\manifest.json` | Target, ownership, observed IDs and verification markers used for continuation/cleanup. |
| `artifacts\runs\<run-id>\` | Native data/assignment/linking evidence and cleanup receipts. |
| `artifacts\runs\<run-id>\videos\` | Workflow recordings and recording manifests. |

Recordings cover workflow creation through native verification, including one
continuous dashboard flow. Descriptive filenames include the scenario and
outcome, such as `2.1-nonlinked-ingestion-passed.webm`; extra tabs receive a suffix.
Temporary assertion captions/outlines use yellow for checking, green for passed
and red for failed. Native API/KQL assertions appear in report steps/evidence.

**Keep artifacts private by default.** The browser profile contains
authentication material, and reports, videos and manifests can expose tenant,
resource, query or device data. Do not commit or publicly upload the entire
`artifacts` folder, browser/storage state, tokens, credentials, connection
strings, HARs, authentication screenshots or traces. Review and sanitize any
report or workflow clip before sharing it; use access-controlled storage.

For Azure Pipelines, use an interactive self-hosted agent when sign-in is required.
Publish the overall JUnit XML with `PublishTestResults@2` and `condition: always()`.
Share only reviewed report/video files as protected artifacts, not the browser
profile or the whole artifacts directory.

## Cleanup

The full runner automatically attempts verified cleanup on success and failure
when `E2E_RETENTION=cleanup` and resources are not explicitly kept. It removes
run-owned Azure resources, Fabric items/folder and temporary Activator connections.
Local simulator clients close when their scenario finishes, including when
resources are retained.

After an interrupted run, use its saved `E2E_RUN_ID` and run cleanup on the same
machine. If cleanup reports that its device process is still running, stop that
run's command first, then retry. An unresolved resource-creation operation can
also require a later cleanup attempt; keep its manifest and check the reported
cleanup result.

To retain resources:

```powershell
npm run e2e -- --keep-resources
```

For individual commands, use `E2E_KEEP_RESOURCES=1`. Alternatively, set
`E2E_RETENTION=retain` for the saved run. Retention can leave Azure services,
Eventstreams and other Fabric items consuming resources until explicitly cleaned.

To delete a selected run later, use its original target settings:

```powershell
$env:E2E_RUN_ID = "<run-guid>"
npm run cleanup
```

Explicit `cleanup` deletes run-owned resources even when keep/retain is set.
`E2E_KEEP_RESOURCES=0` alone does not cancel `E2E_RETENTION=retain`; restore the
normal policy for subsequent fresh runs:

```powershell
Remove-Item Env:E2E_KEEP_RESOURCES -ErrorAction SilentlyContinue
$env:E2E_RETENTION = "cleanup"
Remove-Item Env:E2E_RUN_ID -ErrorAction SilentlyContinue
```

Wait for cleanup to finish before closing the terminal. A cleanup failure makes
the full run fail and leaves its manifest/receipts for investigation; resources
may still incur charges. If ownership or process-state checks block deletion,
preserve the evidence and have a maintainer review the specific run rather than
bypassing the checks.

## Harness self-tests

These checks validate the harness locally without Azure/Fabric provisioning or
live sign-in:

```powershell
npm run check
npm run test:unit
npm run test:pages
npm run test:video
```

The browser checks require Edge; recording checks also require FFmpeg and a
desktop session. Passing self-tests does not replace a live end-to-end run.

## Troubleshooting

| Symptom | What to check |
| --- | --- |
| Missing/invalid configuration or mismatched manifest | Use GUIDs, a supported region and the exact workload label. Restore all saved target settings for continuation/cleanup, or unset the run ID for a new run. |
| Azure CLI not found or authentication fails | Check `az --version`, `az account show --subscription $env:E2E_SUBSCRIPTION_ID`, the CLI sign-in tenant, and Node.js version in the same terminal. |
| Workload unavailable or wrong deployment mode | Confirm tenant/workspace installation, current frontend deployment, exact card label and the appropriate developer switch. Use a stable, already running frontend/gateway for development. |
| Azure quota/provider/role error | Check supported ADR-linked features, registered providers, five S1 Hub capacity and the provisioning/role-assignment permissions. The ten-minute RBAC waits are expected. |
| Ingestion/modeling data checks fail | Check the failing report step, Hub/ADR links, native routes, outbound TLS/443 access and active Fabric capacity. Keep native evidence from the run. |
| Assignment remains unknown | Review the per-device result and assignment timing/status evidence. Unknown does not satisfy the confirmed-assignment check; investigate before starting another fresh attempt. |
| Browser control timeout or unexpected UI | Confirm the suite matches the deployed workload UI and that another process is not using its profile. Keep product changes stable during live execution. |
| Fabric reports that 20 items are already open | Close some open items in Fabric's left navigation pane, then rerun the suite. |
| Missing video binary | Run `npm run install:video` with the intended `PLAYWRIGHT_BROWSERS_PATH`. |
| Interrupted run or cleanup failure | Preserve the manifest/reports, confirm the run's processes have ended, then use explicit cleanup with its saved target. Request a maintainer review if an ownership/state check still blocks it. |
