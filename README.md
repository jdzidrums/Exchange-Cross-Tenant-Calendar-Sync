# Exchange Cross-Tenant Calendar Sync

Production-oriented two-way Microsoft 365 calendar synchronization between:

- `joey@dzidrums.com`
- `jdzidrums@ultrapro.com`

The runtime uses Microsoft Graph, Azure Automation, Azure Key Vault, Azure Blob Storage, and Exchange Online **RBAC for Applications**. GitHub Actions deploys ongoing Exchange RBAC and Azure Automation changes by OIDC federation; no long-lived GitHub cloud secret is required.

## Security model

Each Microsoft 365 tenant has a dedicated **runtime** Entra application. Calendar authorization is granted with Exchange Online `Application Calendars.ReadWrite` and a custom resource scope containing only the intended mailbox.

Do **not** grant the runtime apps tenant-wide Microsoft Graph `Calendars.*` application permissions. Those permissions would be additive to Exchange Application RBAC and would defeat the mailbox scope. The bootstrap and CI/CD deployment both fail closed if such a Graph calendar application role is detected.

Runtime client secrets are stored only in Azure Key Vault. The Azure Automation account uses a system-assigned managed identity to read the two secrets and to read/write the Blob state store.

Mirrored items intentionally omit attendees, organizer state, Teams meeting metadata, body/notes, and attachments. This prevents the mirror from becoming a second meeting that sends invitations or cancellations.

Recommended production privacy mode is `BusyOnly`, which produces `Work - Busy` and `Personal - Busy` mirror appointments while preserving time and free/busy state. Private events are masked when `RespectPrivate` is enabled.

## Repository layout

- `AzureAutomation-CrossTenantCalendarSync.ps1` — production synchronization runbook.
- `Install-AzureAutomationCalendarSync.ps1` — one-time Azure/runtime-app bootstrap.
- `Test-AzureAutomationCalendarSync.ps1` — starts a test job against the deployed runbook.
- `Invoke-AzureAutomationCalendarSync.ps1` — manual Sync/Test/Rebaseline launcher.
- `Rotate-AzureCalendarSyncSecrets.ps1` — runtime app secret rotation helper.
- `scripts/Bootstrap-GitHubOIDC.ps1` — creates the three secretless GitHub deployment identities.
- `scripts/Deploy-ExchangeTenant.ps1` — idempotent mailbox-scoped Exchange RBAC deployment.
- `scripts/Deploy-AzureAutomation.ps1` — publishes the runbook and maintains schedules.
- `scripts/Validate-Repository.ps1` — PowerShell parser/PSScriptAnalyzer validation.
- `.github/workflows/ci-cd.yml` — validation and production CI/CD.

## Runtime behavior

The runbook uses Graph `calendarView/delta` for incremental changes and stores its delta/mapping state in Azure Blob Storage. A Blob lease prevents overlapping jobs from changing state simultaneously.

The initial baseline defaults to 30 days in the past and 365 days in the future. A rolling rebaseline repairs missing mirrors and advances the synchronization window.

The effective five-minute cadence is implemented with 12 Azure Automation schedules, each recurring hourly at minute offsets `00,05,10,...55`.

## 1. One-time infrastructure/runtime bootstrap

Run PowerShell 7.4+ from an administrator workstation:

```powershell
pwsh ./Install-AzureAutomationCalendarSync.ps1 `
  -SubscriptionId '<AZURE-SUBSCRIPTION-GUID>' `
  -Location 'westus2' `
  -DetailMode BusyOnly `
  -SkipSchedules
```

`-SkipSchedules` still publishes the runbook; it only leaves recurring execution disabled for the first verification pass.

The bootstrap creates/reuses:

- Azure resource group
- Key Vault using Azure RBAC
- Storage account for persistent state/locking
- Azure Automation account with system-assigned identity
- PowerShell 7.4 Runtime Environment
- one runtime Entra app/service principal in each M365 tenant
- one exact-mailbox Exchange Application RBAC scope in each tenant
- the two runtime client secrets in Key Vault

The generated `AzureCalendarSync.DeploymentSummary.json` contains IDs/configuration but no plaintext client secret and is ignored by Git.

## 2. Bootstrap GitHub OIDC deployment identities

After the first infrastructure deployment:

```powershell
pwsh ./scripts/Bootstrap-GitHubOIDC.ps1 `
  -DeploymentSummaryPath ./AzureCalendarSync.DeploymentSummary.json
```

This interactive one-time step requires appropriately privileged administrators in both M365 tenants and the Azure tenant.

It creates three federated identities:

1. Dzidrums M365 deployer
2. Ultra PRO M365 deployer
3. Azure Automation deployer

The two M365 deployment identities receive `Exchange.ManageAsApp`, Exchange Administrator, and read-only Microsoft Graph `Application.Read.All`. The Graph read permission is used only to verify on every deployment that the runtime app has not acquired an unscoped `Calendars.*` application permission.

The Azure deployment identity receives `Automation Contributor` only on the target Automation account.

The federated credential is tied to this repository's production environment. The bootstrap writes `GitHubActions.EnvironmentVariables.json`, which contains identifiers only and is ignored by Git.

## 3. Configure the GitHub `production` environment

Create a GitHub environment named `production` and, where your GitHub plan supports it, protect it with required reviewers and restrict deployment to `main`.

Copy the values from `GitHubActions.EnvironmentVariables.json` into GitHub Actions variables. The workflow expects these variables:

```text
CALENDAR_SYNC_CICD_ENABLED
AZURE_TENANT_ID
AZURE_SUBSCRIPTION_ID
AZURE_DEPLOYER_CLIENT_ID
AZURE_RESOURCE_GROUP
AZURE_LOCATION
AUTOMATION_ACCOUNT_NAME
AUTOMATION_RUNTIME_ENVIRONMENT
AUTOMATION_RUNBOOK_NAME
KEY_VAULT_NAME
STATE_STORAGE_ACCOUNT

DZIDRUMS_TENANT_ID
DZIDRUMS_ORGANIZATION
DZIDRUMS_DEPLOYER_CLIENT_ID
DZIDRUMS_RUNTIME_CLIENT_ID
DZIDRUMS_RUNTIME_SP_OBJECT_ID
DZIDRUMS_MAILBOX
DZIDRUMS_SECRET_NAME

ULTRAPRO_TENANT_ID
ULTRAPRO_ORGANIZATION
ULTRAPRO_DEPLOYER_CLIENT_ID
ULTRAPRO_RUNTIME_CLIENT_ID
ULTRAPRO_RUNTIME_SP_OBJECT_ID
ULTRAPRO_MAILBOX
ULTRAPRO_SECRET_NAME

SYNC_DETAIL_MODE
SYNC_RESPECT_PRIVATE
SYNC_COPY_REMINDERS
SYNC_PAST_DAYS
SYNC_FUTURE_DAYS
SYNC_REBASELINE_DAYS
```

No GitHub Actions secret is required for Azure/M365 authentication.

## 4. CI/CD behavior

Pull requests to `main` run validation only.

A push to `main` always validates. Production deployment runs automatically only when `CALENDAR_SYNC_CICD_ENABLED=true`.

The workflow can also be started manually with `workflow_dispatch`; set `deploy=true` to deploy after validation.

Production deployment order is:

1. Validate all PowerShell.
2. OIDC sign in to Dzidrums and deploy/validate its mailbox-scoped Exchange RBAC.
3. OIDC sign in to Ultra PRO and deploy/validate its mailbox-scoped Exchange RBAC.
4. OIDC sign in to Azure, publish the current runbook, and recreate the 12 staggered schedules.

If either tenant fails RBAC validation, Azure publication does not proceed.

## First production activation

After OIDC bootstrap and GitHub variables are configured, use the GitHub Actions **Cross-Tenant Calendar Sync CI/CD** workflow and manually dispatch it with `deploy=true`.

Before enabling automatic main-branch deployments, test the runbook:

```powershell
pwsh ./Test-AzureAutomationCalendarSync.ps1 `
  -DeploymentSummaryPath ./AzureCalendarSync.DeploymentSummary.json
```

Then set `CALENDAR_SYNC_CICD_ENABLED=true`. Future approved changes merged to `main` will update both Exchange tenants first and publish Azure Automation only after both tenant deployments succeed.

## Secret rotation

Runtime app credentials are intentionally separate from GitHub OIDC. Rotate them with:

```powershell
pwsh ./Rotate-AzureCalendarSyncSecrets.ps1 `
  -DeploymentSummaryPath ./AzureCalendarSync.DeploymentSummary.json
```

The new secrets are written to Key Vault rather than GitHub.

## Rollback

Azure runbook code is source-controlled. To roll back code, revert the offending Git commit and deploy the reverted `main` commit.

Exchange RBAC deployment is idempotent and maintains the exact-mailbox scope. If CI/CD detects a tenant-wide runtime `Calendars.*` Graph application role, remove that role assignment in Entra before rerunning deployment.

Do not delete the Blob state file during a normal rollback; preserving delta/mapping state avoids duplicate mirrors.
