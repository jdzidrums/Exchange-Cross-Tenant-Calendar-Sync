#Requires -Version 7.4
<#
.SYNOPSIS
  One-time bootstrap for the Azure Automation cross-tenant calendar sync.

.DESCRIPTION
  Creates the Azure resource group, Key Vault, Storage account, Automation
  account managed identity, and one mailbox-scoped runtime app in each M365
  tenant. Client secrets are written directly to Key Vault. The script then
  publishes the runbook by calling scripts/Deploy-AzureAutomation.ps1.

  Interactive Microsoft Graph authentication uses device-code flow with an
  extended timeout. Graph bootstrap operations use Invoke-MgGraphRequest rather
  than generated Microsoft.Graph.Applications cmdlets to avoid cross-version
  authentication-provider issues on macOS.

  After this one-time bootstrap, run scripts/Bootstrap-GitHubOIDC.ps1 and let
  GitHub Actions handle ongoing Exchange RBAC and runbook deployments.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SubscriptionId,
    [string]$ResourceGroupName = 'rg-crosscalendar-prod',
    [string]$Location = 'westus2',
    [string]$AutomationAccountName = 'aa-crosscalendar-prod',
    [string]$RuntimeEnvironmentName = 'CalendarSync_PS74',
    [string]$RunbookName = 'CrossTenantCalendarSync',
    [string]$KeyVaultName,
    [string]$StorageAccountName,
    [string]$DzidrumsMailbox = 'joey@dzidrums.com',
    [string]$UltraProMailbox = 'jdzidrums@ultrapro.com',
    [string]$DzidrumsSecretName = 'dzidrums-calendar-sync-client-secret',
    [string]$UltraProSecretName = 'ultrapro-calendar-sync-client-secret',
    [ValidateSet('BusyOnly','SubjectLocation')][string]$DetailMode = 'BusyOnly',
    [bool]$RespectPrivate = $true,
    [bool]$CopyReminders = $false,
    [int]$PastDays = 30,
    [int]$FutureDays = 365,
    [int]$RebaselineDays = 1,
    [ValidateRange(1,36)][int]$SecretLifetimeMonths = 12,
    [switch]$SkipSchedules
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

function Write-Step {
    param([Parameter(Mandatory)][string]$Text)
    Write-Host "`n=== $Text ===" -ForegroundColor Cyan
}

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name | Select-Object -First 1)) {
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
}

foreach ($module in @(
    'Az.Accounts','Az.Resources','Az.Automation','Az.KeyVault','Az.Storage',
    'Microsoft.Graph.Authentication','ExchangeOnlineManagement'
)) {
    Ensure-Module $module
}

Import-Module Az.Accounts
Import-Module Az.Resources
Import-Module Az.Automation
Import-Module Az.KeyVault
Import-Module Az.Storage
Import-Module Microsoft.Graph.Authentication
Import-Module ExchangeOnlineManagement

function Get-StableSuffix {
    param([Parameter(Mandatory)][string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text.ToLowerInvariant())) }
    finally { $sha.Dispose() }
    ([Convert]::ToHexString($hash).ToLowerInvariant()).Substring(0,10)
}

function Ensure-ResourceProvider {
    param([Parameter(Mandatory)][string]$ProviderNamespace)

    $providers = @(Get-AzResourceProvider -ProviderNamespace $ProviderNamespace -ErrorAction SilentlyContinue)
    if (-not ($providers | Where-Object { $_.RegistrationState -eq 'Registered' })) {
        Write-Host "Registering Azure resource provider $ProviderNamespace..."
        Register-AzResourceProvider -ProviderNamespace $ProviderNamespace | Out-Null
        for ($attempt = 1; $attempt -le 60; $attempt++) {
            Start-Sleep -Seconds 3
            $providers = @(Get-AzResourceProvider -ProviderNamespace $ProviderNamespace -ErrorAction SilentlyContinue)
            if ($providers | Where-Object { $_.RegistrationState -eq 'Registered' }) { return }
        }
        throw "Azure resource provider $ProviderNamespace did not reach Registered state."
    }
}

function Ensure-RoleAssignment {
    param(
        [Parameter(Mandatory)][string]$ObjectId,
        [Parameter(Mandatory)][string]$Role,
        [Parameter(Mandatory)][string]$Scope
    )

    $existing = Get-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName $Role -Scope $Scope -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName $Role -Scope $Scope | Out-Null
        Write-Host "Assigned '$Role' at $Scope"
    }
}

function Set-KeyVaultSecretWithRetry {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingConvertToSecureStringWithPlainText',
        '',
        Justification = 'Microsoft Graph returns newly generated client secrets as plaintext, while Set-AzKeyVaultSecret requires a SecureString.'
    )]
    param(
        [Parameter(Mandatory)][string]$VaultName,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Value
    )

    $secureValue = ConvertTo-SecureString -String $Value -AsPlainText -Force
    for ($attempt = 1; $attempt -le 30; $attempt++) {
        try {
            Set-AzKeyVaultSecret -VaultName $VaultName -Name $Name -SecretValue $secureValue -ErrorAction Stop | Out-Null
            return
        }
        catch {
            if ($attempt -eq 30) { throw }
            Write-Host 'Waiting for Key Vault RBAC propagation before writing secrets...'
            Start-Sleep -Seconds 10
        }
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

    $connectCommand = Get-Command Connect-MgGraph -ErrorAction Stop
    $connectParams = @{
        TenantId     = $TenantId
        Scopes       = $Scopes
        ContextScope = 'Process'
        NoWelcome    = $true
        ErrorAction  = 'Stop'
    }
    if ($connectCommand.Parameters.ContainsKey('UseDeviceCode')) {
        $connectParams.UseDeviceCode = $true
    }
    elseif ($connectCommand.Parameters.ContainsKey('UseDeviceAuthentication')) {
        $connectParams.UseDeviceAuthentication = $true
    }
    else {
        throw 'Installed Microsoft.Graph.Authentication does not support device-code authentication. Update the module and retry.'
    }
    if ($connectCommand.Parameters.ContainsKey('ClientTimeout')) {
        $connectParams.ClientTimeout = 600
    }

    # Device-code instructions are emitted to the success stream. Send them to
    # the host so callers can capture only the Graph context returned below.
    Connect-MgGraph @connectParams | Out-Host
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
    if ($Select) {
        $uri += "&`$select=$([Uri]::EscapeDataString($Select))"
    }
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
    return @($response.value)
}

function Connect-CalendarSyncExchange {
    param(
        [Parameter(Mandatory)][string]$AdminUpn,
        [Parameter(Mandatory)][string]$TenantHint
    )

    $connectCommand = Get-Command Connect-ExchangeOnline -ErrorAction Stop
    $params = @{
        UserPrincipalName = $AdminUpn
        ShowBanner        = $false
        ErrorAction       = 'Stop'
    }
    if ($connectCommand.Parameters.ContainsKey('DisableWAM')) {
        $params.DisableWAM = $true
    }

    Write-Host "Connecting to Exchange Online for $TenantHint as $AdminUpn..."
    Connect-ExchangeOnline @params
}

function New-RuntimeApp {
    param(
        [Parameter(Mandatory)][string]$TenantHint,
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$FriendlyName
    )

    $ctx = Connect-CalendarSyncGraph `
        -TenantId $TenantHint `
        -Scopes @('Application.ReadWrite.All','Application.Read.All') `
        -FriendlyName $FriendlyName

    $exchangeConnected = $false
    try {
        $displayName = "CrossTenant Calendar Sync Runtime - $FriendlyName - $Mailbox"
        $escapedDisplayName = $displayName.Replace("'", "''")
        $apps = @(Get-GraphCollection -Resource 'applications' -Filter "displayName eq '$escapedDisplayName'")
        if ($apps.Count -gt 1) {
            throw "Multiple app registrations named '$displayName' exist. Remove duplicates before continuing."
        }

        if ($apps.Count -eq 1) {
            $app = $apps[0]
            Write-Host "Using existing runtime app: $($app.appId)"
        }
        else {
            $app = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' -Body @{
                displayName    = $displayName
                signInAudience = 'AzureADMyOrg'
            } -ErrorAction Stop
            Write-Host "Created runtime app: $($app.appId)"
        }

        $sps = @(Get-GraphCollection -Resource 'servicePrincipals' -Filter "appId eq '$($app.appId)'")
        if ($sps.Count -gt 1) { throw "Multiple service principals exist for appId $($app.appId)." }
        if ($sps.Count -eq 1) {
            $sp = $sps[0]
        }
        else {
            $sp = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/servicePrincipals' -Body @{
                appId = [string]$app.appId
            } -ErrorAction Stop
            Write-Host "Created runtime service principal: $($sp.id)"
        }

        # Fail closed if this runtime app has any tenant-wide Microsoft Graph
        # Calendars.* application role. Such a grant would override the intended
        # mailbox-only Exchange Application RBAC boundary.
        $graphSps = @(Get-GraphCollection -Resource 'servicePrincipals' -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -Select 'id,appRoles')
        $graphSp = $graphSps | Select-Object -First 1
        if (-not $graphSp) { throw 'Microsoft Graph service principal was not found in the tenant.' }

        $assignmentResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" -ErrorAction Stop
        $badPermissions = foreach ($assignment in @($assignmentResponse.value)) {
            if ([string]$assignment.resourceId -ne [string]$graphSp.id) { continue }
            $role = @($graphSp.appRoles | Where-Object { [string]$_.id -eq [string]$assignment.appRoleId }) | Select-Object -First 1
            if ($role -and [string]$role.value -like 'Calendars.*') { [string]$role.value }
        }
        if (@($badPermissions).Count -gt 0) {
            throw "Runtime app $($app.appId) has unscoped Microsoft Graph calendar permission(s): $(@($badPermissions) -join ', '). Remove them before continuing."
        }

        $exoAdmin = Read-Host "Exchange administrator UPN for $TenantHint"
        if ([string]::IsNullOrWhiteSpace($exoAdmin)) { throw 'Exchange administrator UPN is required.' }

        Connect-CalendarSyncExchange -AdminUpn $exoAdmin -TenantHint $TenantHint
        $exchangeConnected = $true

        $mailboxObject = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
        $scopeName = "CalendarSync-$($Mailbox.Replace('@','-at-').Replace('.','-'))"
        $assignmentName = "$scopeName-CalendarsRW"
        $escapedDn = $mailboxObject.DistinguishedName.Replace("'", "''")
        $recipientFilter = "DistinguishedName -eq '$escapedDn'"

        $scope = Get-ManagementScope -Identity $scopeName -ErrorAction SilentlyContinue
        if ($scope) {
            Set-ManagementScope -Identity $scopeName -RecipientRestrictionFilter $recipientFilter -ErrorAction Stop
        }
        else {
            New-ManagementScope -Name $scopeName -RecipientRestrictionFilter $recipientFilter -ErrorAction Stop | Out-Null
        }

        $exoSp = Get-ServicePrincipal -Identity ([string]$sp.id) -ErrorAction SilentlyContinue
        if (-not $exoSp) {
            $created = $false
            for ($attempt = 1; $attempt -le 12 -and -not $created; $attempt++) {
                try {
                    New-ServicePrincipal `
                        -AppId ([string]$app.appId) `
                        -ObjectId ([string]$sp.id) `
                        -DisplayName $displayName `
                        -ErrorAction Stop | Out-Null
                    $created = $true
                }
                catch {
                    if ($attempt -eq 12) { throw }
                    Write-Host 'Waiting for the Entra service principal to propagate to Exchange Online...'
                    Start-Sleep -Seconds 5
                }
            }
        }

        $roleAssignment = Get-ManagementRoleAssignment -Identity $assignmentName -ErrorAction SilentlyContinue
        if ($roleAssignment) {
            Set-ManagementRoleAssignment -Identity $assignmentName -CustomResourceScope $scopeName -ErrorAction Stop
        }
        else {
            New-ManagementRoleAssignment `
                -Name $assignmentName `
                -Role 'Application Calendars.ReadWrite' `
                -App ([string]$sp.id) `
                -CustomResourceScope $scopeName `
                -ErrorAction Stop | Out-Null
        }

        $authorization = @(Test-ServicePrincipalAuthorization -Identity ([string]$sp.id) -Resource $Mailbox -ErrorAction Stop)
        if (-not ($authorization | Where-Object { $_.RoleName -eq 'Application Calendars.ReadWrite' -and $_.InScope -eq $true })) {
            throw "Mailbox-scoped Exchange authorization failed for $Mailbox."
        }
        Write-Host "Exchange mailbox-scoped authorization PASSED for $Mailbox."

        # Create the runtime credential only after Exchange scoping succeeds.
        $secretExpiry = [datetime]::UtcNow.AddMonths($SecretLifetimeMonths)
        $passwordResult = Invoke-MgGraphRequest `
            -Method POST `
            -Uri "https://graph.microsoft.com/v1.0/applications/$($app.id)/addPassword" `
            -Body @{
                passwordCredential = @{
                    displayName = "Calendar sync $(Get-Date -Format yyyy-MM-dd)"
                    endDateTime = $secretExpiry.ToString('o')
                }
            } `
            -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace([string]$passwordResult.secretText)) {
            throw "Runtime client-secret creation failed for $FriendlyName."
        }

        return [pscustomobject]@{
            TenantId                 = [string]$ctx.TenantId
            ClientId                 = [string]$app.appId
            ApplicationObjectId      = [string]$app.id
            ServicePrincipalObjectId = [string]$sp.id
            Mailbox                  = $Mailbox
            Secret                   = [string]$passwordResult.secretText
            SecretKeyId              = [string]$passwordResult.keyId
            SecretExpiresUtc         = $secretExpiry.ToString('o')
            ExchangeScopeName        = $scopeName
            ExchangeAssignmentName   = $assignmentName
        }
    }
    finally {
        if ($exchangeConnected) {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

Write-Step 'Connecting to Azure'
if (-not (Get-AzContext -ErrorAction SilentlyContinue)) {
    Connect-AzAccount -UseDeviceAuthentication | Out-Null
}
Set-AzContext -SubscriptionId $SubscriptionId | Out-Null

foreach ($provider in @('Microsoft.Automation','Microsoft.KeyVault','Microsoft.Storage')) {
    Ensure-ResourceProvider -ProviderNamespace $provider
}

$suffix = Get-StableSuffix "$SubscriptionId|$ResourceGroupName|$AutomationAccountName"
if (-not $KeyVaultName) { $KeyVaultName = "kv-calsync-$suffix" }
if (-not $StorageAccountName) { $StorageAccountName = "stcalsync$suffix" }

Write-Step 'Creating or reusing Azure resources'
$rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $rg) {
    $rg = New-AzResourceGroup -Name $ResourceGroupName -Location $Location
}

$vault = Get-AzKeyVault -VaultName $KeyVaultName -ErrorAction SilentlyContinue
if (-not $vault) {
    $newVaultParams = @{
        Name               = $KeyVaultName
        ResourceGroupName  = $ResourceGroupName
        Location           = $Location
        EnablePurgeProtection = $true
    }
    $newVaultCommand = Get-Command New-AzKeyVault -ErrorAction Stop
    if ($newVaultCommand.Parameters.ContainsKey('EnableRbacAuthorization')) {
        # Az.KeyVault < 6.0.0: RBAC was opt-in.
        $newVaultParams.EnableRbacAuthorization = $true
    }
    # Az.KeyVault 6.0.0+: RBAC is the default, so no enable switch is used.
    $vault = New-AzKeyVault @newVaultParams
}
elseif ($vault.PSObject.Properties['EnableRbacAuthorization'] -and -not [bool]$vault.EnableRbacAuthorization) {
    $updateVaultCommand = Get-Command Update-AzKeyVault -ErrorAction Stop
    if ($updateVaultCommand.Parameters.ContainsKey('DisableRbacAuthorization')) {
        Update-AzKeyVault -VaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -DisableRbacAuthorization $false | Out-Null
    }
    elseif ($updateVaultCommand.Parameters.ContainsKey('EnableRbacAuthorization')) {
        Update-AzKeyVault -VaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -EnableRbacAuthorization $true | Out-Null
    }
    else {
        throw 'The existing Key Vault is not RBAC-enabled and the installed Az.KeyVault module cannot switch its authorization model.'
    }
    $vault = Get-AzKeyVault -VaultName $KeyVaultName
}

$storage = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction SilentlyContinue
if (-not $storage) {
    $storage = New-AzStorageAccount `
        -ResourceGroupName $ResourceGroupName `
        -Name $StorageAccountName `
        -Location $Location `
        -SkuName Standard_LRS `
        -Kind StorageV2 `
        -MinimumTlsVersion TLS1_2 `
        -AllowBlobPublicAccess $false `
        -AllowSharedKeyAccess $false
}

try {
    Update-AzStorageBlobServiceProperty `
        -ResourceGroupName $ResourceGroupName `
        -StorageAccountName $StorageAccountName `
        -IsVersioningEnabled $true `
        -ErrorAction Stop | Out-Null
}
catch {
    Write-Warning "Could not enable blob versioning automatically: $($_.Exception.Message)"
}

$automation = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction SilentlyContinue
if (-not $automation) {
    $automation = New-AzAutomationAccount `
        -ResourceGroupName $ResourceGroupName `
        -Name $AutomationAccountName `
        -Location $Location `
        -AssignSystemIdentity
}
else {
    $automation = Set-AzAutomationAccount `
        -ResourceGroupName $ResourceGroupName `
        -Name $AutomationAccountName `
        -AssignSystemIdentity
}

$principalId = $null
for ($attempt = 1; $attempt -le 30 -and -not $principalId; $attempt++) {
    $automation = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName
    $principalId = [string]$automation.Identity.PrincipalId
    if (-not $principalId) { Start-Sleep -Seconds 5 }
}
if (-not $principalId) { throw 'Automation managed identity was not available.' }

Ensure-RoleAssignment -ObjectId $principalId -Role 'Key Vault Secrets User' -Scope $vault.ResourceId
Ensure-RoleAssignment -ObjectId $principalId -Role 'Storage Blob Data Contributor' -Scope $storage.Id

Write-Step 'Granting temporary deployer access to write Key Vault secrets'
$me = Get-AzADUser -SignedIn -ErrorAction SilentlyContinue
if (-not $me) { throw 'Unable to resolve the signed-in Azure user for temporary Key Vault access.' }
$hadVaultRole = [bool](Get-AzRoleAssignment -ObjectId $me.Id -RoleDefinitionName 'Key Vault Secrets Officer' -Scope $vault.ResourceId -ErrorAction SilentlyContinue)
if (-not $hadVaultRole) {
    New-AzRoleAssignment -ObjectId $me.Id -RoleDefinitionName 'Key Vault Secrets Officer' -Scope $vault.ResourceId | Out-Null
}

try {
    Write-Step 'Configuring Dzidrums runtime identity and Exchange scope'
    $dz = New-RuntimeApp -TenantHint 'dzidrums.com' -Mailbox $DzidrumsMailbox -FriendlyName 'Dzidrums'

    Write-Step 'Configuring Ultra PRO runtime identity and Exchange scope'
    $up = New-RuntimeApp -TenantHint 'ultrapro.com' -Mailbox $UltraProMailbox -FriendlyName 'Ultra PRO'

    Write-Step 'Writing runtime credentials to Azure Key Vault'
    Set-KeyVaultSecretWithRetry -VaultName $KeyVaultName -Name $DzidrumsSecretName -Value $dz.Secret
    Set-KeyVaultSecretWithRetry -VaultName $KeyVaultName -Name $UltraProSecretName -Value $up.Secret
    $dz.Secret = $null
    $up.Secret = $null

    Write-Step 'Publishing Azure Automation runbook'
    $deployScript = Join-Path $PSScriptRoot 'scripts/Deploy-AzureAutomation.ps1'
    if (-not (Test-Path $deployScript)) { throw "Missing deployment helper: $deployScript" }

    & $deployScript `
        -SubscriptionId $SubscriptionId `
        -ResourceGroupName $ResourceGroupName `
        -Location $Location `
        -AutomationAccountName $AutomationAccountName `
        -RuntimeEnvironmentName $RuntimeEnvironmentName `
        -RunbookName $RunbookName `
        -RunbookFile (Join-Path $PSScriptRoot 'AzureAutomation-CrossTenantCalendarSync.ps1') `
        -KeyVaultName $KeyVaultName `
        -StateStorageAccount $StorageAccountName `
        -DzidrumsSecretName $DzidrumsSecretName `
        -UltraProSecretName $UltraProSecretName `
        -DzidrumsTenantId $dz.TenantId `
        -DzidrumsClientId $dz.ClientId `
        -DzidrumsMailbox $DzidrumsMailbox `
        -UltraProTenantId $up.TenantId `
        -UltraProClientId $up.ClientId `
        -UltraProMailbox $UltraProMailbox `
        -DetailMode $DetailMode `
        -RespectPrivate $RespectPrivate `
        -CopyReminders $CopyReminders `
        -PastDays $PastDays `
        -FutureDays $FutureDays `
        -RebaselineDays $RebaselineDays `
        -SkipSchedules:$SkipSchedules

    $summary = [ordered]@{
        GeneratedUtc = [datetime]::UtcNow.ToString('o')
        Azure = [ordered]@{
            SubscriptionId         = $SubscriptionId
            ResourceGroupName      = $ResourceGroupName
            Location               = $Location
            AutomationAccountName  = $AutomationAccountName
            AutomationPrincipalId  = $principalId
            RuntimeEnvironmentName = $RuntimeEnvironmentName
            RunbookName            = $RunbookName
            KeyVaultName           = $KeyVaultName
            StorageAccountName     = $StorageAccountName
            EffectiveSchedule      = if ($SkipSchedules) { 'Not created' } else { 'Every 5 minutes via 12 staggered hourly schedules' }
        }
        Dzidrums = [ordered]@{
            TenantId                 = $dz.TenantId
            ClientId                 = $dz.ClientId
            ApplicationObjectId      = $dz.ApplicationObjectId
            ServicePrincipalObjectId = $dz.ServicePrincipalObjectId
            Mailbox                  = $DzidrumsMailbox
            KeyVaultSecretName       = $DzidrumsSecretName
            SecretKeyId              = $dz.SecretKeyId
            SecretExpiresUtc         = $dz.SecretExpiresUtc
            ExchangeScopeName        = $dz.ExchangeScopeName
            ExchangeAssignmentName   = $dz.ExchangeAssignmentName
        }
        UltraPro = [ordered]@{
            TenantId                 = $up.TenantId
            ClientId                 = $up.ClientId
            ApplicationObjectId      = $up.ApplicationObjectId
            ServicePrincipalObjectId = $up.ServicePrincipalObjectId
            Mailbox                  = $UltraProMailbox
            KeyVaultSecretName       = $UltraProSecretName
            SecretKeyId              = $up.SecretKeyId
            SecretExpiresUtc         = $up.SecretExpiresUtc
            ExchangeScopeName        = $up.ExchangeScopeName
            ExchangeAssignmentName   = $up.ExchangeAssignmentName
        }
        Sync = [ordered]@{
            DetailMode     = $DetailMode
            RespectPrivate = $RespectPrivate
            CopyReminders  = $CopyReminders
            PastDays       = $PastDays
            FutureDays     = $FutureDays
            RebaselineDays = $RebaselineDays
        }
    }

    $summaryPath = Join-Path $PSScriptRoot 'AzureCalendarSync.DeploymentSummary.json'
    $summary | ConvertTo-Json -Depth 20 | Set-Content -Path $summaryPath -Encoding UTF8

    Write-Step 'Bootstrap complete'
    Write-Host "Deployment summary: $summaryPath"
    Write-Host "Key Vault: $KeyVaultName"
    Write-Host "Storage account: $StorageAccountName"
    Write-Host "Automation account: $AutomationAccountName"
    Write-Host "Runbook: $RunbookName"
    if ($SkipSchedules) { Write-Host 'Recurring schedules were intentionally not created.' }
}
finally {
    if (-not $hadVaultRole) {
        Remove-AzRoleAssignment `
            -ObjectId $me.Id `
            -RoleDefinitionName 'Key Vault Secrets Officer' `
            -Scope $vault.ResourceId `
            -ErrorAction SilentlyContinue | Out-Null
    }
}
