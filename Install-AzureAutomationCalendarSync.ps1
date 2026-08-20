#Requires -Version 7.4
<#
.SYNOPSIS
  One-time bootstrap for the Azure Automation cross-tenant calendar sync.

.DESCRIPTION
  Creates the Azure resource group, Key Vault, Storage account, Automation
  account managed identity, and one mailbox-scoped runtime app in each M365
  tenant. Client secrets are written directly to Key Vault. The script then
  publishes the runbook by calling scripts/Deploy-AzureAutomation.ps1.

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

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name | Select-Object -First 1)) {
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
}

foreach ($module in @(
    'Az.Accounts','Az.Resources','Az.Automation','Az.KeyVault','Az.Storage',
    'Microsoft.Graph.Authentication','Microsoft.Graph.Applications',
    'ExchangeOnlineManagement'
)) { Ensure-Module $module }

Import-Module Az.Accounts
Import-Module Az.Resources
Import-Module Az.Automation
Import-Module Az.KeyVault
Import-Module Az.Storage
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Applications
Import-Module ExchangeOnlineManagement

function Get-StableSuffix {
    param([Parameter(Mandatory)][string]$Text)
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = $sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text.ToLowerInvariant())) }
    finally { $sha.Dispose() }
    ([Convert]::ToHexString($hash).ToLowerInvariant()).Substring(0,10)
}

function Ensure-RoleAssignment {
    param([string]$ObjectId,[string]$Role,[string]$Scope)
    $existing = Get-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName $Role -Scope $Scope -ErrorAction SilentlyContinue
    if (-not $existing) {
        New-AzRoleAssignment -ObjectId $ObjectId -RoleDefinitionName $Role -Scope $Scope | Out-Null
    }
}

function Set-KeyVaultSecretWithRetry {
    param(
        [Parameter(Mandatory)][string]$VaultName,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][string]$Value
    )

    $secureValue = ConvertTo-SecureString -String $Value -AsPlainText -Force
    for ($attempt = 1; $attempt -le 30; $attempt++) {
        try {
            Set-AzKeyVaultSecret -VaultName $VaultName -Name $Name -SecretValue $secureValue | Out-Null
            return
        }
        catch {
            if ($attempt -eq 30) { throw }
            Write-Host 'Waiting for Key Vault RBAC propagation before writing secrets...'
            Start-Sleep -Seconds 10
        }
    }
}

function New-RuntimeApp {
    param(
        [Parameter(Mandatory)][string]$TenantHint,
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$FriendlyName
    )

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Connect-MgGraph -TenantId $TenantHint -Scopes 'Application.ReadWrite.All','Application.Read.All' -NoWelcome
    $ctx = Get-MgContext

    try {
        $displayName = "CrossTenant Calendar Sync Runtime - $FriendlyName - $Mailbox"
        $escaped = $displayName.Replace("'","''")
        $apps = @(Get-MgApplication -Filter "displayName eq '$escaped'" -All)
        if ($apps.Count -gt 1) { throw "Multiple app registrations named '$displayName' exist." }
        $app = if ($apps.Count -eq 1) { $apps[0] } else { New-MgApplication -DisplayName $displayName -SignInAudience 'AzureADMyOrg' }

        $sps = @(Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'" -All)
        if ($sps.Count -gt 1) { throw "Multiple service principals exist for $($app.AppId)." }
        $sp = if ($sps.Count -eq 1) { $sps[0] } else { New-MgServicePrincipal -AppId $app.AppId }

        # Fail closed if a tenant-wide Graph Calendars.* application role exists.
        $graph = @(Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" -All) | Select-Object -First 1
        $assignments = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All -ErrorAction SilentlyContinue)
        $bad = foreach ($assignment in $assignments) {
            if ([string]$assignment.ResourceId -ne [string]$graph.Id) { continue }
            $role = @($graph.AppRoles | Where-Object { [string]$_.Id -eq [string]$assignment.AppRoleId }) | Select-Object -First 1
            if ($role -and [string]$role.Value -like 'Calendars.*') { [string]$role.Value }
        }
        if (@($bad).Count -gt 0) {
            throw "Runtime app $($app.AppId) has unscoped Graph calendar permission(s): $(@($bad) -join ', ')."
        }

        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
            displayName = "Calendar sync $(Get-Date -Format yyyy-MM-dd)"
            endDateTime = (Get-Date).ToUniversalTime().AddMonths($SecretLifetimeMonths)
        }
        if (-not $secret.SecretText) { throw "Failed to create runtime credential for $FriendlyName." }

        $exoAdmin = Read-Host "Exchange administrator UPN for $TenantHint"
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Connect-ExchangeOnline -UserPrincipalName $exoAdmin -ShowBanner:$false
        try {
            $mailboxObject = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
            $scopeName = "CalendarSync-$($Mailbox.Replace('@','-at-').Replace('.','-'))"
            $assignmentName = "$scopeName-CalendarsRW"
            $escapedDn = $mailboxObject.DistinguishedName.Replace("'","''")
            $filter = "DistinguishedName -eq '$escapedDn'"

            if (Get-ManagementScope -Identity $scopeName -ErrorAction SilentlyContinue) {
                Set-ManagementScope -Identity $scopeName -RecipientRestrictionFilter $filter
            } else {
                New-ManagementScope -Name $scopeName -RecipientRestrictionFilter $filter | Out-Null
            }

            if (-not (Get-ServicePrincipal -Identity $sp.Id -ErrorAction SilentlyContinue)) {
                New-ServicePrincipal -AppId $app.AppId -ObjectId $sp.Id -DisplayName $displayName | Out-Null
            }

            if (Get-ManagementRoleAssignment -Identity $assignmentName -ErrorAction SilentlyContinue) {
                Set-ManagementRoleAssignment -Identity $assignmentName -CustomResourceScope $scopeName
            } else {
                New-ManagementRoleAssignment -Name $assignmentName -Role 'Application Calendars.ReadWrite' -App $sp.Id -CustomResourceScope $scopeName | Out-Null
            }

            $test = @(Test-ServicePrincipalAuthorization -Identity $sp.Id -Resource $Mailbox)
            if (-not ($test | Where-Object { $_.RoleName -eq 'Application Calendars.ReadWrite' -and $_.InScope -eq $true })) {
                throw "Mailbox-scoped Exchange authorization failed for $Mailbox."
            }
        }
        finally { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null }

        [pscustomobject]@{
            TenantId = [string]$ctx.TenantId
            ClientId = [string]$app.AppId
            ApplicationObjectId = [string]$app.Id
            ServicePrincipalObjectId = [string]$sp.Id
            Mailbox = $Mailbox
            Secret = [string]$secret.SecretText
            SecretKeyId = [string]$secret.KeyId
            SecretExpiresUtc = $secret.EndDateTime.ToUniversalTime().ToString('o')
            ExchangeScopeName = $scopeName
            ExchangeAssignmentName = $assignmentName
        }
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) { Connect-AzAccount | Out-Null }
Set-AzContext -SubscriptionId $SubscriptionId | Out-Null

$suffix = Get-StableSuffix "$SubscriptionId|$ResourceGroupName|$AutomationAccountName"
if (-not $KeyVaultName) { $KeyVaultName = "kv-calsync-$suffix" }
if (-not $StorageAccountName) { $StorageAccountName = "stcalsync$suffix" }

$rg = Get-AzResourceGroup -Name $ResourceGroupName -ErrorAction SilentlyContinue
if (-not $rg) { $rg = New-AzResourceGroup -Name $ResourceGroupName -Location $Location }

$vault = Get-AzKeyVault -VaultName $KeyVaultName -ErrorAction SilentlyContinue
if (-not $vault) {
    $vault = New-AzKeyVault -Name $KeyVaultName -ResourceGroupName $ResourceGroupName -Location $Location -EnableRbacAuthorization $true -EnablePurgeProtection
}

$storage = Get-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -ErrorAction SilentlyContinue
if (-not $storage) {
    $storage = New-AzStorageAccount -ResourceGroupName $ResourceGroupName -Name $StorageAccountName -Location $Location -SkuName Standard_LRS -Kind StorageV2 -MinimumTlsVersion TLS1_2 -AllowBlobPublicAccess $false -AllowSharedKeyAccess $false
}

$automation = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -ErrorAction SilentlyContinue
if (-not $automation) {
    $automation = New-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -Location $Location -AssignSystemIdentity
} else {
    $automation = Set-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName -AssignSystemIdentity
}
$automation = Get-AzAutomationAccount -ResourceGroupName $ResourceGroupName -Name $AutomationAccountName
$principalId = [string]$automation.Identity.PrincipalId
if (-not $principalId) { throw 'Automation managed identity was not available.' }

Ensure-RoleAssignment -ObjectId $principalId -Role 'Key Vault Secrets User' -Scope $vault.ResourceId
Ensure-RoleAssignment -ObjectId $principalId -Role 'Storage Blob Data Contributor' -Scope $storage.Id

$me = Get-AzADUser -SignedIn -ErrorAction SilentlyContinue
if (-not $me) { throw 'Unable to resolve the signed-in Azure user for temporary Key Vault access.' }
$hadVaultRole = [bool](Get-AzRoleAssignment -ObjectId $me.Id -RoleDefinitionName 'Key Vault Secrets Officer' -Scope $vault.ResourceId -ErrorAction SilentlyContinue)
if (-not $hadVaultRole) {
    New-AzRoleAssignment -ObjectId $me.Id -RoleDefinitionName 'Key Vault Secrets Officer' -Scope $vault.ResourceId | Out-Null
}

try {
    $dz = New-RuntimeApp -TenantHint 'dzidrums.com' -Mailbox $DzidrumsMailbox -FriendlyName 'Dzidrums'
    $up = New-RuntimeApp -TenantHint 'ultrapro.com' -Mailbox $UltraProMailbox -FriendlyName 'Ultra PRO'

    Set-KeyVaultSecretWithRetry -VaultName $KeyVaultName -Name $DzidrumsSecretName -Value $dz.Secret
    Set-KeyVaultSecretWithRetry -VaultName $KeyVaultName -Name $UltraProSecretName -Value $up.Secret
    $dz.Secret = $null
    $up.Secret = $null

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
            SubscriptionId = $SubscriptionId
            ResourceGroupName = $ResourceGroupName
            Location = $Location
            AutomationAccountName = $AutomationAccountName
            AutomationPrincipalId = $principalId
            RuntimeEnvironmentName = $RuntimeEnvironmentName
            RunbookName = $RunbookName
            KeyVaultName = $KeyVaultName
            StorageAccountName = $StorageAccountName
            EffectiveSchedule = if ($SkipSchedules) { 'Not created' } else { 'Every 5 minutes via 12 staggered hourly schedules' }
        }
        Dzidrums = [ordered]@{
            TenantId = $dz.TenantId
            ClientId = $dz.ClientId
            ApplicationObjectId = $dz.ApplicationObjectId
            ServicePrincipalObjectId = $dz.ServicePrincipalObjectId
            Mailbox = $DzidrumsMailbox
            KeyVaultSecretName = $DzidrumsSecretName
            SecretKeyId = $dz.SecretKeyId
            SecretExpiresUtc = $dz.SecretExpiresUtc
            ExchangeScopeName = $dz.ExchangeScopeName
            ExchangeAssignmentName = $dz.ExchangeAssignmentName
        }
        UltraPro = [ordered]@{
            TenantId = $up.TenantId
            ClientId = $up.ClientId
            ApplicationObjectId = $up.ApplicationObjectId
            ServicePrincipalObjectId = $up.ServicePrincipalObjectId
            Mailbox = $UltraProMailbox
            KeyVaultSecretName = $UltraProSecretName
            SecretKeyId = $up.SecretKeyId
            SecretExpiresUtc = $up.SecretExpiresUtc
            ExchangeScopeName = $up.ExchangeScopeName
            ExchangeAssignmentName = $up.ExchangeAssignmentName
        }
        Sync = [ordered]@{
            DetailMode = $DetailMode
            RespectPrivate = $RespectPrivate
            CopyReminders = $CopyReminders
            PastDays = $PastDays
            FutureDays = $FutureDays
            RebaselineDays = $RebaselineDays
        }
    }

    $summaryPath = Join-Path $PSScriptRoot 'AzureCalendarSync.DeploymentSummary.json'
    $summary | ConvertTo-Json -Depth 20 | Set-Content -Path $summaryPath -Encoding UTF8
    Write-Host "Bootstrap complete. Deployment summary: $summaryPath"
}
finally {
    if (-not $hadVaultRole) {
        Remove-AzRoleAssignment -ObjectId $me.Id -RoleDefinitionName 'Key Vault Secrets Officer' -Scope $vault.ResourceId -ErrorAction SilentlyContinue | Out-Null
    }
}
