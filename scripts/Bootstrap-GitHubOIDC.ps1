#Requires -Version 7.2
<#
.SYNOPSIS
  One-time bootstrap for secretless GitHub Actions deployment into both
  Exchange Online tenants and the Azure Automation account.

.DESCRIPTION
  Creates one GitHub OIDC workload identity in each M365 tenant and one in the
  Azure tenant. The two M365 deployers receive Exchange.ManageAsApp,
  Microsoft Graph Application.Read.All, and the Exchange Administrator
  directory role so CI/CD can maintain the narrowly scoped runtime Application
  RBAC assignment. The Azure deployer receives Automation Contributor only on
  the target Automation account.

  Interactive Microsoft Graph authentication uses device-code flow with an
  extended timeout, and Graph changes are made through Invoke-MgGraphRequest.
  No GitHub client secret is created.
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
$ProgressPreference = 'SilentlyContinue'

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name | Select-Object -First 1)) {
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
}

foreach ($module in @(
    'Microsoft.Graph.Authentication',
    'Az.Accounts',
    'Az.Resources'
)) {
    Ensure-Module -Name $module
}

$summaryPath = [System.IO.Path]::GetFullPath($DeploymentSummaryPath)
if (-not (Test-Path $summaryPath)) { throw "Deployment summary not found: $summaryPath" }
$summary = Get-Content -Path $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 50

$subject = "repo:$GitHubOwner@$GitHubOwnerId/$GitHubRepository@$GitHubRepositoryId`:environment:$GitHubEnvironment"
$issuer = 'https://token.actions.githubusercontent.com'
$audience = 'api://AzureADTokenExchange'
Write-Host "GitHub OIDC subject: $subject"

function Invoke-IsolatedAzureBootstrap {
    param(
        [Parameter(Mandatory)][string]$SummaryPath,
        [string]$TenantId,
        [string]$ServicePrincipalObjectId
    )

    # Microsoft.Graph.Authentication and Az.Accounts currently ship
    # incompatible Azure.Identity versions. Run Az commands in a child pwsh
    # process so either module can use the dependency version it was built for.
    $helperPath = Join-Path $PSScriptRoot 'Bootstrap-GitHubOIDC.Azure.ps1'
    if (-not (Test-Path $helperPath)) { throw "Azure OIDC helper not found: $helperPath" }

    $contextPath = Join-Path ([System.IO.Path]::GetTempPath()) "calendar-sync-azure-context-$([guid]::NewGuid().ToString('N')).json"
    $processArgs = @(
        '-NoLogo'
        '-NoProfile'
        '-File'
        $helperPath
        '-DeploymentSummaryPath'
        $SummaryPath
        '-ContextOutputPath'
        $contextPath
    )
    if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
        $processArgs += @('-AzureTenantId', $TenantId)
    }
    if (-not [string]::IsNullOrWhiteSpace($ServicePrincipalObjectId)) {
        $processArgs += @('-AzureServicePrincipalObjectId', $ServicePrincipalObjectId)
    }

    try {
        $pwshPath = (Get-Command pwsh -ErrorAction Stop).Source
        & $pwshPath @processArgs | Out-Host
        if ($LASTEXITCODE -ne 0) {
            throw "Isolated Azure OIDC bootstrap failed with exit code $LASTEXITCODE."
        }
        if (-not (Test-Path $contextPath)) {
            throw 'Isolated Azure OIDC bootstrap did not return its Azure context.'
        }
        return Get-Content -Path $contextPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    finally {
        Remove-Item -Path $contextPath -Force -ErrorAction SilentlyContinue
    }
}

function Connect-CalendarSyncGraph {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string[]]$Scopes,
        [Parameter(Mandatory)][string]$FriendlyName
    )

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    Write-Host "`nMicrosoft Graph sign-in required for $FriendlyName ($TenantId)." -ForegroundColor Yellow
    Write-Host 'When the device code appears, open https://microsoft.com/devicelogin immediately and complete sign-in.' -ForegroundColor Yellow

    $command = Get-Command Connect-MgGraph -ErrorAction Stop
    $params = @{
        TenantId     = $TenantId
        Scopes       = $Scopes
        ContextScope = 'Process'
        NoWelcome    = $true
        ErrorAction  = 'Stop'
    }
    if ($command.Parameters.ContainsKey('UseDeviceCode')) {
        $params.UseDeviceCode = $true
    }
    elseif ($command.Parameters.ContainsKey('UseDeviceAuthentication')) {
        $params.UseDeviceAuthentication = $true
    }
    else {
        throw 'Installed Microsoft.Graph.Authentication does not support device-code authentication. Update the module and retry.'
    }
    if ($command.Parameters.ContainsKey('ClientTimeout')) {
        $params.ClientTimeout = 600
    }

    # Device-code instructions are emitted to the success stream. Send them to
    # the host so callers can capture only the Graph context returned below.
    Connect-MgGraph @params | Out-Host
    $ctx = Get-MgContext
    if (-not $ctx -or [string]::IsNullOrWhiteSpace([string]$ctx.TenantId)) {
        throw "Microsoft Graph authentication did not establish a usable context for $FriendlyName."
    }
    Write-Host "Microsoft Graph authenticated: $($ctx.Account) / tenant $($ctx.TenantId)"
    return $ctx
}

function Get-GraphCollection {
    param(
        [Parameter(Mandatory)][string]$Resource,
        [Parameter(Mandatory)][string]$Filter,
        [string]$Select
    )

    $encodedFilter = [Uri]::EscapeDataString($Filter)
    $uri = "https://graph.microsoft.com/v1.0/${Resource}?`$filter=$encodedFilter"
    if ($Select) { $uri += "&`$select=$([Uri]::EscapeDataString($Select))" }
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    return @($response.value)
}

function Get-OrCreateApplication {
    param([Parameter(Mandatory)][string]$DisplayName)

    $escaped = $DisplayName.Replace("'", "''")
    $apps = @(Get-GraphCollection -Resource 'applications' -Filter "displayName eq '$escaped'")
    if ($apps.Count -gt 1) { throw "Multiple app registrations named '$DisplayName' exist." }
    if ($apps.Count -eq 1) { return $apps[0] }

    return Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' -Body @{
        displayName    = $DisplayName
        signInAudience = 'AzureADMyOrg'
    } -ErrorAction Stop
}

function Get-OrCreateServicePrincipal {
    param([Parameter(Mandatory)][string]$AppId)

    $sps = @(Get-GraphCollection -Resource 'servicePrincipals' -Filter "appId eq '$AppId'")
    if ($sps.Count -gt 1) { throw "Multiple service principals were returned for appId $AppId." }
    if ($sps.Count -eq 1) { return $sps[0] }

    return Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body @{
        appId = $AppId
    } -ErrorAction Stop
}

function Set-GitHubFederatedCredential {
    param(
        [Parameter(Mandatory)][string]$ApplicationObjectId,
        [Parameter(Mandatory)][string]$Name
    )

    $collectionUri = "https://graph.microsoft.com/v1.0/applications/$ApplicationObjectId/federatedIdentityCredentials"
    $response = Invoke-MgGraphRequest -Method GET -Uri $collectionUri -ErrorAction Stop
    $existing = @($response.value | Where-Object { $_.name -eq $Name }) | Select-Object -First 1

    $desired = @{
        name        = $Name
        issuer      = $issuer
        subject     = $subject
        audiences   = @($audience)
        description = "GitHub Actions $GitHubEnvironment deployment for $GitHubOwner/$GitHubRepository"
    }

    if ($existing) {
        $same = ([string]$existing.issuer -eq $issuer) -and
                ([string]$existing.subject -eq $subject) -and
                (@($existing.audiences) -contains $audience)
        if ($same) {
            Write-Host "Federated credential already correct: $Name"
            return
        }

        Invoke-MgGraphRequest -Method DELETE -Uri "$collectionUri/$($existing.id)" -ErrorAction Stop | Out-Null
    }

    Invoke-MgGraphRequest -Method POST -Uri $collectionUri -Body $desired -ErrorAction Stop | Out-Null
    Write-Host "Created federated credential: $Name"
}

function Grant-ServicePrincipalAppRole {
    param(
        [Parameter(Mandatory)]$ServicePrincipal,
        [Parameter(Mandatory)][string]$ResourceAppId,
        [Parameter(Mandatory)][string[]]$RoleValues
    )

    $resources = @(Get-GraphCollection -Resource 'servicePrincipals' -Filter "appId eq '$ResourceAppId'" -Select 'id,appRoles,appId,displayName')
    $resource = $resources | Select-Object -First 1
    if (-not $resource) { throw "Resource service principal $ResourceAppId was not found." }

    $role = $null
    foreach ($value in $RoleValues) {
        $role = @($resource.appRoles | Where-Object {
            [string]$_.value -eq $value -and @($_.allowedMemberTypes) -contains 'Application'
        }) | Select-Object -First 1
        if ($role) { break }
    }
    if (-not $role) { throw "None of the requested application roles were found: $($RoleValues -join ', ')." }

    $assignmentUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$($ServicePrincipal.id)/appRoleAssignments"
    $assignmentResponse = Invoke-MgGraphRequest -Method GET -Uri $assignmentUri -ErrorAction Stop
    $exists = @($assignmentResponse.value | Where-Object {
        [string]$_.resourceId -eq [string]$resource.id -and [string]$_.appRoleId -eq [string]$role.id
    }).Count -gt 0

    if (-not $exists) {
        Invoke-MgGraphRequest -Method POST -Uri $assignmentUri -Body @{
            principalId = [string]$ServicePrincipal.id
            resourceId  = [string]$resource.id
            appRoleId   = [string]$role.id
        } -ErrorAction Stop | Out-Null
        Write-Host "Granted application role $($role.value)."
    }
}

function Grant-ExchangeAdministratorDirectoryRole {
    param([Parameter(Mandatory)]$ServicePrincipal)

    $filter = [Uri]::EscapeDataString("displayName eq 'Exchange Administrator'")
    $roleResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleDefinitions?`$filter=$filter" -ErrorAction Stop
    $role = @($roleResponse.value) | Select-Object -First 1
    if (-not $role) { throw 'Exchange Administrator directory role definition was not found.' }

    $assignmentFilter = [Uri]::EscapeDataString("principalId eq '$($ServicePrincipal.id)'")
    $assignmentResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=$assignmentFilter" -ErrorAction Stop
    $exists = @($assignmentResponse.value | Where-Object {
        [string]$_.roleDefinitionId -eq [string]$role.id -and [string]$_.directoryScopeId -eq '/'
    }).Count -gt 0

    if (-not $exists) {
        Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments' -Body @{
            principalId      = [string]$ServicePrincipal.id
            roleDefinitionId = [string]$role.id
            directoryScopeId = '/'
        } -ErrorAction Stop | Out-Null
        Write-Host 'Assigned Exchange Administrator directory role to the GitHub deployer.'
    }
}

function Get-InitialDomain {
    $domains = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/domains' -ErrorAction Stop
    $initial = @($domains.value | Where-Object { $_.isInitial -eq $true }) | Select-Object -First 1
    if (-not $initial) { throw 'Unable to determine the tenant initial .onmicrosoft.com domain.' }
    return [string]$initial.id
}

function New-M365GitHubDeployer {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$FriendlyName
    )

    Connect-CalendarSyncGraph `
        -TenantId $TenantId `
        -Scopes @('Application.ReadWrite.All','AppRoleAssignment.ReadWrite.All','RoleManagement.ReadWrite.Directory','Domain.Read.All') `
        -FriendlyName "$FriendlyName GitHub deployer" | Out-Null

    try {
        $app = Get-OrCreateApplication -DisplayName "GitHub Actions Calendar Sync Deployer - $FriendlyName"
        $sp = Get-OrCreateServicePrincipal -AppId ([string]$app.appId)

        Set-GitHubFederatedCredential -ApplicationObjectId ([string]$app.id) -Name 'github-production'
        Grant-ServicePrincipalAppRole `
            -ServicePrincipal $sp `
            -ResourceAppId '00000002-0000-0ff1-ce00-000000000000' `
            -RoleValues @('Exchange.ManageAsApp','Exchange.ManageAsAppV2')
        Grant-ServicePrincipalAppRole `
            -ServicePrincipal $sp `
            -ResourceAppId '00000003-0000-0000-c000-000000000000' `
            -RoleValues @('Application.Read.All')
        Grant-ExchangeAdministratorDirectoryRole -ServicePrincipal $sp

        return [pscustomobject]@{
            TenantId      = $TenantId
            ClientId      = [string]$app.appId
            ServiceObject = [string]$sp.id
            Organization  = Get-InitialDomain
        }
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

$azureContext = Invoke-IsolatedAzureBootstrap `
    -SummaryPath $summaryPath `
    -TenantId $AzureTenantId
$AzureTenantId = [string]$azureContext.AzureTenantId

Import-Module Microsoft.Graph.Authentication

$dz = New-M365GitHubDeployer -TenantId ([string]$summary.Dzidrums.TenantId) -FriendlyName 'Dzidrums'
$up = New-M365GitHubDeployer -TenantId ([string]$summary.UltraPro.TenantId) -FriendlyName 'Ultra PRO'

Connect-CalendarSyncGraph `
    -TenantId $AzureTenantId `
    -Scopes @('Application.ReadWrite.All') `
    -FriendlyName 'Azure Automation GitHub deployer' | Out-Null
try {
    $azureApp = Get-OrCreateApplication -DisplayName 'GitHub Actions Calendar Sync Deployer - Azure Automation'
    $azureSp = Get-OrCreateServicePrincipal -AppId ([string]$azureApp.appId)
    Set-GitHubFederatedCredential -ApplicationObjectId ([string]$azureApp.id) -Name 'github-production'
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

Invoke-IsolatedAzureBootstrap `
    -SummaryPath $summaryPath `
    -TenantId $AzureTenantId `
    -ServicePrincipalObjectId ([string]$azureSp.id) | Out-Null

function Resolve-RuntimeServicePrincipalId {
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string]$FriendlyName
    )

    Connect-CalendarSyncGraph -TenantId $TenantId -Scopes @('Application.Read.All') -FriendlyName $FriendlyName | Out-Null
    try {
        $sp = @(Get-GraphCollection -Resource 'servicePrincipals' -Filter "appId eq '$ClientId'") | Select-Object -First 1
        if (-not $sp) { throw "Runtime service principal $ClientId was not found in tenant $TenantId." }
        return [string]$sp.id
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

$dzRuntimeSp = if ($summary.Dzidrums.PSObject.Properties['ServicePrincipalObjectId']) {
    [string]$summary.Dzidrums.ServicePrincipalObjectId
}
else {
    Resolve-RuntimeServicePrincipalId `
        -TenantId ([string]$summary.Dzidrums.TenantId) `
        -ClientId ([string]$summary.Dzidrums.ClientId) `
        -FriendlyName 'Dzidrums runtime service principal lookup'
}

$upRuntimeSp = if ($summary.UltraPro.PSObject.Properties['ServicePrincipalObjectId']) {
    [string]$summary.UltraPro.ServicePrincipalObjectId
}
else {
    Resolve-RuntimeServicePrincipalId `
        -TenantId ([string]$summary.UltraPro.TenantId) `
        -ClientId ([string]$summary.UltraPro.ClientId) `
        -FriendlyName 'Ultra PRO runtime service principal lookup'
}

$variables = [ordered]@{
    CALENDAR_SYNC_CICD_ENABLED       = 'true'
    AZURE_TENANT_ID                  = $AzureTenantId
    AZURE_SUBSCRIPTION_ID            = [string]$summary.Azure.SubscriptionId
    AZURE_DEPLOYER_CLIENT_ID         = [string]$azureApp.appId
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
        Owner          = $GitHubOwner
        OwnerId        = $GitHubOwnerId
        Repository     = $GitHubRepository
        RepositoryId   = $GitHubRepositoryId
        Environment    = $GitHubEnvironment
        ImmutableSubject = $subject
    }
    Variables = $variables
}

$outputPath = Join-Path (Split-Path -Parent $summaryPath) 'GitHubActions.EnvironmentVariables.json'
$output | ConvertTo-Json -Depth 20 | Set-Content -Path $outputPath -Encoding UTF8

Write-Host ''
Write-Host 'OIDC BOOTSTRAP COMPLETE'
Write-Host "Configuration written to: $outputPath"
Write-Host 'No client secrets were written to this file.'
Write-Host "OIDC subject: $subject"
Write-Host 'Next: run scripts/Configure-GitHubRepositoryVariables.ps1, then perform one manual deploy=true workflow run.'
