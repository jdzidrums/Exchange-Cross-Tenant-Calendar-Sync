#Requires -Version 7.2
<#
.SYNOPSIS
  Runs the Azure-specific portion of Bootstrap-GitHubOIDC.ps1 in an isolated
  PowerShell process.

.DESCRIPTION
  Microsoft.Graph.Authentication and Az.Accounts can carry incompatible
  Azure.Identity assemblies. This helper keeps Az authentication and role
  assignment out of the Microsoft Graph process while preserving the original
  interactive bootstrap behavior.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DeploymentSummaryPath,
    [string]$AzureTenantId,
    [string]$AzureServicePrincipalObjectId,
    [Parameter(Mandatory)][string]$ContextOutputPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

Import-Module Az.Accounts
Import-Module Az.Resources

$summaryPath = [System.IO.Path]::GetFullPath($DeploymentSummaryPath)
if (-not (Test-Path $summaryPath)) { throw "Deployment summary not found: $summaryPath" }
$summary = Get-Content -Path $summaryPath -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 50

Write-Host "`nAzure sign-in required for the subscription that hosts the Automation account." -ForegroundColor Yellow
$currentAz = Get-AzContext -ErrorAction SilentlyContinue
$wrongSubscription = -not $currentAz -or
    [string]$currentAz.Subscription.Id -ne [string]$summary.Azure.SubscriptionId
$wrongTenant = -not [string]::IsNullOrWhiteSpace($AzureTenantId) -and
    $currentAz -and [string]$currentAz.Tenant.Id -ne $AzureTenantId

if ($wrongSubscription -or $wrongTenant) {
    $azParams = @{
        Subscription            = [string]$summary.Azure.SubscriptionId
        UseDeviceAuthentication = $true
        ErrorAction             = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($AzureTenantId)) {
        $azParams.Tenant = $AzureTenantId
    }
    Connect-AzAccount @azParams | Out-Host
}

Set-AzContext -SubscriptionId ([string]$summary.Azure.SubscriptionId) | Out-Null
$azureContext = Get-AzContext -ErrorAction Stop
$resolvedTenantId = [string]$azureContext.Tenant.Id

if (-not [string]::IsNullOrWhiteSpace($AzureTenantId) -and $resolvedTenantId -ne $AzureTenantId) {
    throw "Azure context tenant $resolvedTenantId does not match requested tenant $AzureTenantId."
}

if (-not [string]::IsNullOrWhiteSpace($AzureServicePrincipalObjectId)) {
    $automationScope = "/subscriptions/$($summary.Azure.SubscriptionId)/resourceGroups/$($summary.Azure.ResourceGroupName)/providers/Microsoft.Automation/automationAccounts/$($summary.Azure.AutomationAccountName)"
    $role = Get-AzRoleAssignment `
        -ObjectId $AzureServicePrincipalObjectId `
        -RoleDefinitionName 'Automation Contributor' `
        -Scope $automationScope `
        -ErrorAction SilentlyContinue
    if (-not $role) {
        New-AzRoleAssignment `
            -ObjectId $AzureServicePrincipalObjectId `
            -ObjectType 'ServicePrincipal' `
            -RoleDefinitionName 'Automation Contributor' `
            -Scope $automationScope | Out-Null
        Write-Host "Granted Automation Contributor at $automationScope"
    }
}

[ordered]@{
    AzureTenantId = $resolvedTenantId
} | ConvertTo-Json | Set-Content -Path $ContextOutputPath -Encoding UTF8
