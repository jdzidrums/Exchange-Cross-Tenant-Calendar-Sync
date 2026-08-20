#Requires -Version 7.2
<#
.SYNOPSIS
  One-time bootstrap for secretless GitHub Actions deployment into both
  Exchange Online tenants and the Azure Automation account.

.DESCRIPTION
  Creates one GitHub OIDC workload identity in each M365 tenant and one in the
  Azure tenant. The two M365 deployers receive Exchange.ManageAsApp and the
  Exchange Administrator directory role so they can maintain the narrowly
  scoped runtime Application RBAC assignment. The Azure deployer receives
  Automation Contributor only on the target Automation account.

  This bootstrap is intentionally interactive and must be run by appropriately
  privileged administrators. The resulting GitHubActions.EnvironmentVariables.json
  contains identifiers/configuration only; it contains no client secret.
#>

[CmdletBinding()]
param(
    [string]$DeploymentSummaryPath = './AzureCalendarSync.DeploymentSummary.json',
    [string]$AzureTenantId,
    [string]$GitHubOwner = 'jdzidrums',
    [string]$GitHubOwnerId = '36529445',
    [string]$GitHubRepository = 'Exchange-Cross-Tenant-Calendar-Sync',
    [string]$GitHubRepositoryId = '1341092190',
    [string]$GitHubEnvironment = 'production'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
}

foreach ($module in @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications',
    'Az.Accounts',
    'Az.Resources'
)) {
    Ensure-Module -Name $module
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications
Import-Module Az.Accounts
Import-Module Az.Resources

$summaryPath = [System.IO.Path]::GetFullPath($DeploymentSummaryPath)
if (-not (Test-Path $summaryPath)) { throw "Deployment summary not found: $summaryPath" }
$summary = Get-Content -Path $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 50

$subject = "repo:$GitHubOwner@$GitHubOwnerId/$GitHubRepository@$GitHubRepositoryId:environment:$GitHubEnvironment"
Write-Host "GitHub OIDC subject: $subject"

function Get-OrCreateApplication {
    param([Parameter(Mandatory)][string]$DisplayName)
    $escaped = $DisplayName.Replace("'", "''")
    $apps = @(Get-MgApplication -Filter "displayName eq '$escaped'" -All)
    if ($apps.Count -gt 1) { throw "Multiple app registrations named '$DisplayName' exist." }
    if ($apps.Count -eq 1) { return $apps[0] }
    New-MgApplication -DisplayName $DisplayName -SignInAudience 'AzureADMyOrg'
}

function Get-OrCreateServicePrincipal {
    param([Parameter(Mandatory)][string]$AppId)
    $sps = @(Get-MgServicePrincipal -Filter "appId eq '$AppId'" -All)
    if ($sps.Count -gt 1) { throw "Multiple service principals were returned for AppId $AppId." }
    if ($sps.Count -eq 1) { return $sps[0] }
    New-MgServicePrincipal -AppId $AppId
}

function Set-GitHubFederatedCredential {
    param(
        [Parameter(Mandatory)][string]$ApplicationObjectId,
        [Parameter(Mandatory)][string]$Name
    )
    $collectionUri = "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectId/federatedIdentityCredentials"
    $existingResponse = Invoke-MgGraphRequest -Method GET -Uri $collectionUri
    $existing = @($existingResponse.value | Where-Object { $_.name -eq $Name }) | Select-Object -First 1
    $body = @{
        name        = $Name
        issuer      = 'https://token.actions.githubusercontent.com'
        subject     = $subject
        audiences   = @('api://AzureADTokenExchange')
        description = "GitHub Actions $GitHubEnvironment deployment for $GitHubOwner/$GitHubRepository"
    }
    if ($existing) {
        Invoke-MgGraphRequest -Method PATCH -Uri "$collectionUri/$($existing.id)" -Body $body | Out-Null
        Write-Host "Updated federated credential: $Name"
    }
    else {
        Invoke-MgGraphRequest -Method POST -Uri $collectionUri -Body $body | Out-Null
        Write-Host "Created federated credential: $Name"
    }
}

function Grant-GraphApplicationReadAll {
    param([Parameter(Mandatory)]$ServicePrincipal)

    $graphResource = @(Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -All) | Select-Object -First 1
    if (-not $graphResource) { throw 'Microsoft Graph enterprise application was not found.' }
    $appRole = @($graphResource.AppRoles | Where-Object {
        $_.Value -eq 'Application.Read.All' -and $_.AllowedMemberTypes -contains 'Application'
    }) | Select-Object -First 1
    if (-not $appRole) { throw 'Microsoft Graph Application.Read.All application role was not found.' }

    $assignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -All -ErrorAction SilentlyContinue)
    $exists = $assignments | Where-Object {
        [string]$_.ResourceId -eq [string]$graphResource.Id -and [string]$_.AppRoleId -eq [string]$appRole.Id
    }
    if (-not $exists) {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $ServicePrincipal.Id `
            -PrincipalId $ServicePrincipal.Id `
            -ResourceId $graphResource.Id `
            -AppRoleId $appRole.Id | Out-Null
        Write-Host 'Granted Microsoft Graph Application.Read.All for fail-closed runtime-app validation.'
    }
}

function Grant-ExchangeManageAsApp {
    param([Parameter(Mandatory)]$ServicePrincipal)

    $exchangeResource = @(Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'" -All) | Select-Object -First 1
    if (-not $exchangeResource) { throw 'Office 365 Exchange Online enterprise application was not found.' }

    $appRole = @($exchangeResource.AppRoles | Where-Object {
        $_.Value -in @('Exchange.ManageAsApp','Exchange.ManageAsAppV2') -and $_.AllowedMemberTypes -contains 'Application'
    }) | Sort-Object @{ Expression = { if ($_.Value -eq 'Exchange.ManageAsApp') { 0 } else { 1 } } } | Select-Object -First 1

    if (-not $appRole) { throw 'Exchange.ManageAsApp application role was not found on the Exchange Online service principal.' }

    $assignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -All -ErrorAction SilentlyContinue)
    $exists = $assignments | Where-Object {
        [string]$_.ResourceId -eq [string]$exchangeResource.Id -and [string]$_.AppRoleId -eq [string]$appRole.Id
    }
    if (-not $exists) {
        New-MgServicePrincipalAppRoleAssignment `
            -ServicePrincipalId $ServicePrincipal.Id `
            -PrincipalId $ServicePrincipal.Id `
            -ResourceId $exchangeResource.Id `
            -AppRoleId $appRole.Id | Out-Null
        Write-Host "Granted $($appRole.Value)."
    }
}

function Grant-ExchangeAdministratorDirectoryRole {
    param([Parameter(Mandatory)]$ServicePrincipal)

    $roleUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$filter=displayName%20eq%20%27Exchange%20Administrator%27"
    $roleResponse = Invoke-MgGraphRequest -Method GET -Uri $roleUri
    $role = @($roleResponse.value) | Select-Object -First 1
    if (-not $role) { throw 'Exchange Administrator directory role definition was not found.' }

    $assignUri = "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=principalId%20eq%20%27$($ServicePrincipal.Id)%27"
    $assignmentResponse = Invoke-MgGraphRequest -Method GET -Uri $assignUri
    $existing = @($assignmentResponse.value | Where-Object {
        [string]$_.roleDefinitionId -eq [string]$role.id -and [string]$_.directoryScopeId -eq '/'
    })
    if ($existing.Count -eq 0) {
        Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' -Body @{
            principalId      = $ServicePrincipal.Id
            roleDefinitionId = $role.id
            directoryScopeId = '/'
        } | Out-Null
        Write-Host 'Assigned Exchange Administrator directory role to the GitHub deployer.'
    }
}

function Get-InitialDomain {
    $domains = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/domains'
    $initial = @($domains.value | Where-Object { $_.isInitial -eq $true }) | Select-Object -First 1
    if (-not $initial) { throw 'Unable to determine the tenant initial .onmicrosoft.com domain.' }
    [string]$initial.id
}

function New-M365GitHubDeployer {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$FriendlyName
    )

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Host "Sign in to bootstrap the $FriendlyName tenant ($TenantId)."
    Connect-MgGraph -TenantId $TenantId -Scopes @(
        'Application.ReadWrite.All',
        'AppRoleAssignment.ReadWrite.All',
        'RoleManagement.ReadWrite.Directory',
        'Domain.Read.All'
    ) -NoWelcome

    try {
        $app = Get-OrCreateApplication -DisplayName "GitHub Actions Calendar Sync Deployer - $FriendlyName"
        $sp = Get-OrCreateServicePrincipal -AppId $app.AppId
        Set-GitHubFederatedCredential -ApplicationObjectId $app.Id -Name 'github-production'
        Grant-ExchangeManageAsApp -ServicePrincipal $sp
        Grant-GraphApplicationReadAll -ServicePrincipal $sp
        Grant-ExchangeAdministratorDirectoryRole -ServicePrincipal $sp
        $organization = Get-InitialDomain

        [pscustomobject]@{
            TenantId      = $TenantId
            ClientId      = [string]$app.AppId
            ServiceObject = [string]$sp.Id
            Organization  = $organization
        }
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

$dz = New-M365GitHubDeployer -TenantId ([string]$summary.Dzidrums.TenantId) -FriendlyName 'Dzidrums'
$up = New-M365GitHubDeployer -TenantId ([string]$summary.UltraPro.TenantId) -FriendlyName 'Ultra PRO'

Write-Host 'Sign in to the Azure subscription that hosts the Automation account.'
if ($AzureTenantId) {
    Connect-AzAccount -Tenant $AzureTenantId -Subscription ([string]$summary.Azure.SubscriptionId) | Out-Null
}
else {
    Connect-AzAccount -Subscription ([string]$summary.Azure.SubscriptionId) | Out-Null
}
$azContext = Get-AzContext
$AzureTenantId = [string]$azContext.Tenant.Id

try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
Write-Host "Sign in to Microsoft Graph for Azure tenant $AzureTenantId to create the Azure GitHub deployer."
Connect-MgGraph -TenantId $AzureTenantId -Scopes 'Application.ReadWrite.All' -NoWelcome
try {
    $azureApp = Get-OrCreateApplication -DisplayName 'GitHub Actions Calendar Sync Deployer - Azure Automation'
    $azureSp = Get-OrCreateServicePrincipal -AppId $azureApp.AppId
    Set-GitHubFederatedCredential -ApplicationObjectId $azureApp.Id -Name 'github-production'
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

$automationScope = "/subscriptions/$($summary.Azure.SubscriptionId)/resourceGroups/$($summary.Azure.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($summary.Azure.AutomationAccountName)"
$role = Get-AzRoleAssignment -ObjectId $azureSp.Id -RoleDefinitionName 'Automation Contributor' -Scope $automationScope -ErrorAction SilentlyContinue
if (-not $role) {
    New-AzRoleAssignment -ObjectId $azureSp.Id -RoleDefinitionName 'Automation Contributor' -Scope $automationScope | Out-Null
    Write-Host "Granted Automation Contributor at $automationScope"
}

function Resolve-RuntimeServicePrincipalId {
    param([Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$ClientId)
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Connect-MgGraph -TenantId $TenantId -Scopes 'Application.Read.All' -NoWelcome
    try {
        $sp = @(Get-MgServicePrincipal -Filter "appId eq '$ClientId'" -All) | Select-Object -First 1
        if (-not $sp) { throw "Runtime service principal $ClientId was not found in tenant $TenantId." }
        [string]$sp.Id
    }
    finally { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
}

$dzRuntimeSp = if ($summary.Dzidrums.PSObject.Properties['ServicePrincipalObjectId']) {
    [string]$summary.Dzidrums.ServicePrincipalObjectId
} else {
    Resolve-RuntimeServicePrincipalId -TenantId ([string]$summary.Dzidrums.TenantId) -ClientId ([string]$summary.Dzidrums.ClientId)
}
$upRuntimeSp = if ($summary.UltraPro.PSObject.Properties['ServicePrincipalObjectId']) {
    [string]$summary.UltraPro.ServicePrincipalObjectId
} else {
    Resolve-RuntimeServicePrincipalId -TenantId ([string]$summary.UltraPro.TenantId) -ClientId ([string]$summary.UltraPro.ClientId)
}

$variables = [ordered]@{
    CALENDAR_SYNC_CICD_ENABLED       = 'true'
    AZURE_TENANT_ID                  = $AzureTenantId
    AZURE_SUBSCRIPTION_ID            = [string]$summary.Azure.SubscriptionId
    AZURE_DEPLOYER_CLIENT_ID         = [string]$azureApp.AppId
    AZURE_RESOURCE_GROUP             = [string]$summary.Azure.ResourceGroupName
    AZURE_LOCATION                   = [string]$summary.Azure.Location
    AUTOMATION_ACCOUNT_NAME          = [string]$summary.Azure.AutomationAccountName
    AUTOMATION_RUNTIME_ENVIRONMENT   = [string]$summary.Azure.RuntimeEnvironmentName
    AUTOMATION_RUNBOOK_NAME          = [string]$summary.Azure.RunbookName
    KEY_VAULT_NAME                   = [string]$summary.Azure.KeyVaultName
    STATE_STORAGE_ACCOUNT            = [string]$summary.Azure.StorageAccountName

    DZIDRUMS_TENANT_ID               = [string]$summary.Dzidrums.TenantId
    DZIDRUMS_ORGANIZATION            = [string]$dz.Organization
    DZIDRUMS_DEPLOYER_CLIENT_ID      = [string]$dz.ClientId
    DZIDRUMS_RUNTIME_CLIENT_ID       = [string]$summary.Dzidrums.ClientId
    DZIDRUMS_RUNTIME_SP_OBJECT_ID    = $dzRuntimeSp
    DZIDRUMS_MAILBOX                 = [string]$summary.Dzidrums.Mailbox
    DZIDRUMS_SECRET_NAME             = [string]$summary.Dzidrums.KeyVaultSecretName

    ULTRAPRO_TENANT_ID               = [string]$summary.UltraPro.TenantId
    ULTRAPRO_ORGANIZATION            = [string]$up.Organization
    ULTRAPRO_DEPLOYER_CLIENT_ID      = [string]$up.ClientId
    ULTRAPRO_RUNTIME_CLIENT_ID       = [string]$summary.UltraPro.ClientId
    ULTRAPRO_RUNTIME_SP_OBJECT_ID    = $upRuntimeSp
    ULTRAPRO_MAILBOX                 = [string]$summary.UltraPro.Mailbox
    ULTRAPRO_SECRET_NAME             = [string]$summary.UltraPro.KeyVaultSecretName

    SYNC_DETAIL_MODE                 = if ($summary.Sync.DetailMode) { [string]$summary.Sync.DetailMode } else { 'BusyOnly' }
    SYNC_RESPECT_PRIVATE             = ([bool]$summary.Sync.RespectPrivate).ToString().ToLowerInvariant()
    SYNC_COPY_REMINDERS              = ([bool]$summary.Sync.CopyReminders).ToString().ToLowerInvariant()
    SYNC_PAST_DAYS                   = [string]$summary.Sync.PastDays
    SYNC_FUTURE_DAYS                 = [string]$summary.Sync.FutureDays
    SYNC_REBASELINE_DAYS             = [string]$summary.Sync.RebaselineDays
}

$output = [ordered]@{
    GeneratedUtc = [datetime]::UtcNow.ToString('o')
    GitHub = [ordered]@{
        Owner        = $GitHubOwner
        OwnerId      = $GitHubOwnerId
        Repository   = $GitHubRepository
        RepositoryId = $GitHubRepositoryId
        Environment  = $GitHubEnvironment
        OidcSubject  = $subject
    }
    Variables = $variables
}

$outputPath = Join-Path (Split-Path -Parent $summaryPath) 'GitHubActions.EnvironmentVariables.json'
$output | ConvertTo-Json -Depth 20 | Set-Content -Path $outputPath -Encoding UTF8

Write-Host ''
Write-Host 'OIDC BOOTSTRAP COMPLETE'
Write-Host "Configuration written to: $outputPath"
Write-Host 'No client secrets were written to this file.'
Write-Host 'Create GitHub environment production and copy these Variables into GitHub Actions variables.'
