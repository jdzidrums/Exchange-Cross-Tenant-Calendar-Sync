# Exchange Cross-Tenant Calendar Sync

Production-oriented two-way Microsoft 365 calendar synchronization between:

- `joey@dzidrums.com`
- `jdzidrums@ultrapro.com`

The runtime uses Microsoft Graph, Azure Automation, Azure Key Vault, Azure Blob Storage, and Exchange Online **RBAC for Applications**. GitHub Actions handles ongoing deployment through OIDC federation; no long-lived Azure/M365 deployment secret is stored in GitHub.

## Security model

Each Microsoft 365 tenant has a dedicated **runtime** Entra application. Calendar authorization is granted with Exchange Online `Application Calendars.ReadWrite` and a custom resource scope containing only the intended mailbox.

Do **not** grant either runtime app tenant-wide Microsoft Graph `Calendars.*` application permissions. Those permissions would be additive to Exchange Application RBAC and would defeat the mailbox scope. Both bootstrap and CI/CD fail closed if an unscoped Graph calendar application role is detected.

Runtime client secrets are stored only in Azure Key Vault. The Azure Automation account uses a system-assigned managed identity to read those secrets and the Azure Blob state/lock store.

Mirrored items intentionally omit attendees, organizer state, Teams meeting metadata, body/notes, and attachments. Recommended production mode is `BusyOnly`, producing `Work - Busy` and `Personal - Busy` mirrors while preserving time/free-busy state. Private events remain masked when `RespectPrivate` is enabled.

## Repository layout

- `AzureAutomation-CrossTenantCalendarSync.ps1` — production synchronization runbook.
- `Install-AzureAutomationCalendarSync.ps1` — one-time Azure/runtime-app bootstrap.
- `Test-AzureAutomationCalendarSync.ps1` — runbook authorization test launcher.
- `Invoke-AzureAutomationCalendarSync.ps1` — manual Sync/Test/Rebaseline launcher.
- `Rotate-AzureCalendarSyncSecrets.ps1` — runtime credential rotation.
- `scripts/Bootstrap-GitHubOIDC.ps1` — creates three secretless GitHub deployment identities.
- `scripts/Configure-GitHubRepositoryVariables.ps1` — creates the production environment and loads non-secret GitHub repository variables.
- `scripts/Deploy-ExchangeTenant.ps1` — idempotent mailbox-scoped Exchange RBAC deployment.
- `scripts/Deploy-AzureAutomation.ps1` — publishes the runbook and maintains schedules.
- `scripts/Validate-Repository.ps1` — parser/PSScriptAnalyzer validation.
- `.github/workflows/ci-cd.yml` — validation and protected production deployment.

## Runtime behavior

The runbook uses Graph `calendarView/delta` for incremental changes and stores its delta/mapping state in Azure Blob Storage. A Blob lease prevents concurrent jobs from modifying state simultaneously.

The default window is 30 days in the past and 365 days in the future. A rolling rebaseline advances the window and repairs missing mirrors.

The effective five-minute cadence uses 12 Azure Automation schedules, each recurring hourly at minute offsets `00,05,10,...55`.

## 1. One-time infrastructure/runtime bootstrap

Run PowerShell 7.4+ from an administrator workstation:

```powershell
pwsh ./Install-AzureAutomationCalendarSync.ps1 `
  -SubscriptionId '<AZURE-SUBSCRIPTION-GUID>' `
  -Location 'westus2' `
  -DetailMode BusyOnly `
  -SkipSchedules
```

`-SkipSchedules` still publishes the runbook; it only suppresses recurring schedule creation for the first verification pass.

The bootstrap creates/reuses the Azure resource group, Key Vault, state Storage account, Automation account/system identity, PowerShell 7.4 Runtime Environment, one runtime app/service principal per M365 tenant, exact-mailbox Exchange Application RBAC scopes, and the two runtime credentials in Key Vault.

It writes `AzureCalendarSync.DeploymentSummary.json`. The file contains IDs/configuration but no plaintext runtime secret and is ignored by Git.

## 2. Bootstrap GitHub OIDC deployment identities

```powershell
pwsh ./scripts/Bootstrap-GitHubOIDC.ps1 `
  -DeploymentSummaryPath ./AzureCalendarSync.DeploymentSummary.json
```

This one-time interactive step requires appropriately privileged administrators in both M365 tenants and the Azure tenant.

It creates:

1. a Dzidrums M365 deployer;
2. an Ultra PRO M365 deployer;
3. an Azure Automation deployer.

The M365 deployment identities receive `Exchange.ManageAsApp`, Exchange Administrator, and read-only Microsoft Graph `Application.Read.All`. The Graph permission is used only for fail-closed verification that the runtime app has not acquired tenant-wide `Calendars.*` access.

The Azure deployment identity receives `Automation Contributor` only on the target Automation account.

The OIDC credential is restricted to this repository and the `production` GitHub environment. The script writes non-secret configuration to `GitHubActions.EnvironmentVariables.json` (ignored by Git).

## 3. Configure GitHub for production

GitHub's environment-level configuration variables are not available early enough for the workflow's pre-job conditions/OIDC inputs, so this repository stores all non-sensitive deployment IDs/configuration as **repository-level Actions variables**. The `production` environment is used for deployment approval/protection and as the OIDC subject context.

If GitHub CLI is installed and authenticated, configure the repository automatically:

```powershell
pwsh ./scripts/Configure-GitHubRepositoryVariables.ps1 `
  -ConfigurationPath ./GitHubActions.EnvironmentVariables.json
```

The helper creates/reuses the `production` environment, loads the values as repository Actions variables, and deliberately forces:

```text
CALENDAR_SYNC_CICD_ENABLED=false
```

Then, in GitHub, add appropriate protection/reviewer rules to the `production` environment. Restrict production deployment to `main` where supported.

The workflow uses these repository Actions variables:

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

A push to `main` always validates. Automatic production deployment only occurs when repository variable `CALENDAR_SYNC_CICD_ENABLED=true`.

Manual production deployment is available with `workflow_dispatch`; set `deploy=true`.

Deployment order is:

1. validate PowerShell;
2. OIDC sign in to Dzidrums and enforce/verify its mailbox-scoped Exchange RBAC;
3. OIDC sign in to Ultra PRO and enforce/verify its mailbox-scoped Exchange RBAC;
4. only after both tenant jobs succeed, OIDC sign in to Azure, publish the runbook, and recreate the 12 schedules.

If either M365 tenant fails validation, Azure publication does not proceed.

## First production activation

After OIDC bootstrap and repository variables are configured, manually run **Cross-Tenant Calendar Sync CI/CD** with `deploy=true`.

Before enabling automatic main-branch deployment, test the runbook:

```powershell
pwsh ./Test-AzureAutomationCalendarSync.ps1 `
  -DeploymentSummaryPath ./AzureCalendarSync.DeploymentSummary.json
```

After a successful manual deployment and calendar verification, change repository variable `CALENDAR_SYNC_CICD_ENABLED` to `true`. Subsequent approved `main` changes will deploy automatically.

## Secret rotation

Runtime credentials are deliberately separate from GitHub OIDC. Rotate them with:

```powershell
pwsh ./Rotate-AzureCalendarSyncSecrets.ps1 `
  -DeploymentSummaryPath ./AzureCalendarSync.DeploymentSummary.json
```

The new secrets are written to Key Vault, not GitHub.

## Rollback

To roll back runbook code, revert the Git commit and deploy the reverted `main` commit. Exchange deployment is idempotent and maintains the exact-mailbox resource scope.

If CI/CD reports an unscoped `Calendars.*` Graph application role on either runtime app, remove that Entra app-role assignment before rerunning deployment.

Do not delete the Blob state file during a normal code rollback; preserving delta/mapping state avoids duplicate mirror creation.
