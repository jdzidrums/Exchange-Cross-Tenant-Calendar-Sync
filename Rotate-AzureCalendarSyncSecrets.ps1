#Requires -Version 7.4
<#
.SYNOPSIS
    Rotates both cross-tenant app client secrets and immediately writes the new
    values to Azure Key Vault without changing runbook configuration.

.DESCRIPTION
    Creates a new password credential on each existing Entra app registration,
    then replaces the corresponding Key Vault secret value. Existing app password
    credentials are not automatically removed, which provides rollback safety.
    Remove stale credentials after validating the sync with the test script.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DeploymentSummaryPath,

    [ValidateRange(1,36)]
    [int]$SecretLifetimeMonths = 12,

    [string]$DeployerObjectId,
    [switch]$KeepDeployerKeyVaultAccess
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

foreach ($module in @('Az.Accounts','Az.Resources','Az.KeyVault','Microsoft.Graph.Authentication','Microsoft.Graph.Applications')) {
    Import-Module $module -ErrorAction Stop
}

$path = [IO.Path]::GetFullPath($DeploymentSummaryPath)
if (-not (Test-Path $path)) { throw "Deployment summary not found: $path" }
$cfg = Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 30

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) { Connect-AzAccount | Out-Null }
Set-AzContext -SubscriptionId $cfg.Azure.SubscriptionId | Out-Null

if ([string]::IsNullOrWhiteSpace($DeployerObjectId)) {
    $signedIn = Get-AzADUser -SignedIn -ErrorAction SilentlyContinue
    if ($signedIn) { $DeployerObjectId = [string]$signedIn.Id }
}
if ([string]::IsNullOrWhiteSpace($DeployerObjectId)) {
    throw 'Could not determine your Azure-tenant object ID. Supply -DeployerObjectId <guid>.'
}

$vault = Get-AzKeyVault -VaultName $cfg.Azure.KeyVaultName -ErrorAction Stop
$temporaryRole = $false
$existing = Get-AzRoleAssignment `
    -ObjectId $DeployerObjectId `
    -RoleDefinitionName 'Key Vault Secrets Officer' `
    -Scope $vault.ResourceId `
    -ErrorAction SilentlyContinue
if (-not $existing) {
    New-AzRoleAssignment `
        -ObjectId $DeployerObjectId `
        -RoleDefinitionName 'Key Vault Secrets Officer' `
        -Scope $vault.ResourceId | Out-Null
    $temporaryRole = $true
}

function Set-VaultSecretWithRetry {
    [System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSAvoidUsingConvertToSecureStringWithPlainText',
        '',
        Justification = 'Microsoft Graph returns newly generated client secrets as plaintext, while Set-AzKeyVaultSecret requires a SecureString.'
    )]
    param([string]$Name, [string]$Value)
    $secure = ConvertTo-SecureString -String $Value -AsPlainText -Force
    for ($i = 1; $i -le 30; $i++) {
        try {
            Set-AzKeyVaultSecret -VaultName $cfg.Azure.KeyVaultName -Name $Name -SecretValue $secure | Out-Null
            return
        }
        catch {
            if ($i -eq 30) { throw }
            Start-Sleep -Seconds 10
        }
    }
}

function Rotate-One {
    param(
        [Parameter(Mandatory)]$Tenant,
        [Parameter(Mandatory)][int]$Months
    )

    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Connect-MgGraph -TenantId $Tenant.TenantId -Scopes 'Application.ReadWrite.All' -NoWelcome
    try {
        $credential = Add-MgApplicationPassword `
            -ApplicationId $Tenant.ApplicationObjectId `
            -PasswordCredential @{
                displayName = "Azure Automation calendar sync rotation $(Get-Date -Format yyyy-MM-dd)"
                endDateTime = (Get-Date).ToUniversalTime().AddMonths($Months)
            }

        if (-not $credential.SecretText) { throw "Secret rotation failed for $($Tenant.Mailbox)." }
        Set-VaultSecretWithRetry -Name $Tenant.KeyVaultSecretName -Value ([string]$credential.SecretText)

        $Tenant.SecretKeyId = [string]$credential.KeyId
        $Tenant.SecretExpiresUtc = $credential.EndDateTime.ToUniversalTime().ToString('o')
        Write-Host "Rotated secret for $($Tenant.Mailbox); expires $($Tenant.SecretExpiresUtc)."
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

try {
    Rotate-One -Tenant $cfg.Dzidrums -Months $SecretLifetimeMonths
    Rotate-One -Tenant $cfg.UltraPro -Months $SecretLifetimeMonths
    $cfg.GeneratedUtc = [datetime]::UtcNow.ToString('o')
    $cfg | ConvertTo-Json -Depth 30 | Set-Content -Path $path -Encoding UTF8

    Write-Host "`nRotation complete. Run Test-AzureAutomationCalendarSync.ps1 now, then remove obsolete app password credentials after validation." -ForegroundColor Green
}
finally {
    if ($temporaryRole -and -not $KeepDeployerKeyVaultAccess) {
        try {
            Remove-AzRoleAssignment `
                -ObjectId $DeployerObjectId `
                -RoleDefinitionName 'Key Vault Secrets Officer' `
                -Scope $vault.ResourceId | Out-Null
        }
        catch {
            Write-Warning "Could not remove temporary Key Vault Secrets Officer role: $($_.Exception.Message)"
        }
    }
}
