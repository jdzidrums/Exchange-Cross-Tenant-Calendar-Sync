#Requires -Version 7.2
<#
.SYNOPSIS
  Idempotently deploys the mailbox-scoped Exchange Online Application RBAC
  configuration for one calendar-sync runtime service principal.

.DESCRIPTION
  This script is designed for GitHub Actions. The job must already be logged in
  with azure/login using an OIDC/federated Entra application in the target M365
  tenant. That deployer app needs Exchange.ManageAsApp plus sufficient Exchange
  administrative authorization to manage scopes/service-principal pointers/role
  assignments. The bootstrap script assigns the Exchange Administrator directory
  role to that dedicated, secretless deployment identity.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$TenantName,
    [Parameter(Mandatory)][string]$Organization,
    [Parameter(Mandatory)][string]$Mailbox,
    [Parameter(Mandatory)][string]$RuntimeClientId,
    [Parameter(Mandatory)][string]$RuntimeServicePrincipalObjectId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
}

Ensure-Module -Name ExchangeOnlineManagement
Import-Module ExchangeOnlineManagement

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw 'Azure CLI is required. Run this through the GitHub Actions workflow after azure/login.'
}

Write-Host "Validating that the runtime app has no tenant-wide Microsoft Graph calendar application permission ..."
$graphToken = (& az account get-access-token --resource 'https://graph.microsoft.com' --query accessToken -o tsv)
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($graphToken)) {
    throw "Unable to acquire a Microsoft Graph token for $TenantName."
}
$graphHeaders = @{ Authorization = "Bearer $graphToken"; Accept = 'application/json' }
$runtimeSpUri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId%20eq%20%27$RuntimeClientId%27&`$select=id,appId"
$runtimeSpResponse = Invoke-RestMethod -Method GET -Uri $runtimeSpUri -Headers $graphHeaders
$runtimeSp = @($runtimeSpResponse.value) | Select-Object -First 1
if (-not $runtimeSp) { throw "Runtime service principal $RuntimeClientId was not found in $TenantName." }
if ([string]$runtimeSp.id -ne $RuntimeServicePrincipalObjectId) {
    throw "Runtime service-principal object ID mismatch. Expected $RuntimeServicePrincipalObjectId, got $($runtimeSp.id)."
}

$graphResourceUri = "https://graph.microsoft.com/v1.0/servicePrincipals?`$filter=appId%20eq%20%2700000003-0000-0000-c000-000000000000%27&`$select=id,appRoles"
$graphResourceResponse = Invoke-RestMethod -Method GET -Uri $graphResourceUri -Headers $graphHeaders
$graphResource = @($graphResourceResponse.value) | Select-Object -First 1
if (-not $graphResource) { throw 'Microsoft Graph enterprise application was not found.' }

$assignmentUri = "https://graph.microsoft.com/v1.0/servicePrincipals/$RuntimeServicePrincipalObjectId/appRoleAssignments"
$assignmentResponse = Invoke-RestMethod -Method GET -Uri $assignmentUri -Headers $graphHeaders
$unscopedCalendarRoles = foreach ($assignment in @($assignmentResponse.value)) {
    if ([string]$assignment.resourceId -ne [string]$graphResource.id) { continue }
    $role = @($graphResource.appRoles | Where-Object { [string]$_.id -eq [string]$assignment.appRoleId }) | Select-Object -First 1
    if ($role -and [string]$role.value -like 'Calendars.*') { [string]$role.value }
}
if (@($unscopedCalendarRoles).Count -gt 0) {
    throw "FAIL-CLOSED: runtime app $RuntimeClientId has unscoped Microsoft Graph calendar application permission(s): $(@($unscopedCalendarRoles) -join ', '). Remove those Entra app-role assignments before deployment."
}
Write-Host 'Runtime app Graph calendar permission check passed.'

Write-Host "Acquiring Exchange Online access token for $TenantName ..."
$accessToken = (& az account get-access-token --resource 'https://outlook.office365.com' --query accessToken -o tsv)
if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($accessToken)) {
    throw "Unable to acquire an Exchange Online token for $TenantName."
}

Connect-ExchangeOnline -AccessToken $accessToken -Organization $Organization -ShowBanner:$false
try {
    $mailboxObject = Get-Mailbox -Identity $Mailbox -ErrorAction Stop
    $scopeName = "CalendarSync-$($Mailbox.Replace('@','-at-').Replace('.','-'))"
    $assignmentName = "$scopeName-CalendarsRW"
    $pointerName = "CrossTenant Calendar Sync - $TenantName - $Mailbox"

    $escapedDn = $mailboxObject.DistinguishedName.Replace("'", "''")
    $recipientFilter = "DistinguishedName -eq '$escapedDn'"

    $scope = Get-ManagementScope -Identity $scopeName -ErrorAction SilentlyContinue
    if ($scope) {
        Set-ManagementScope -Identity $scopeName -RecipientRestrictionFilter $recipientFilter
        Write-Host "Updated management scope: $scopeName"
    }
    else {
        New-ManagementScope -Name $scopeName -RecipientRestrictionFilter $recipientFilter | Out-Null
        Write-Host "Created management scope: $scopeName"
    }

    $exoSp = Get-ServicePrincipal -Identity $RuntimeServicePrincipalObjectId -ErrorAction SilentlyContinue
    if (-not $exoSp) {
        New-ServicePrincipal `
            -AppId $RuntimeClientId `
            -ObjectId $RuntimeServicePrincipalObjectId `
            -DisplayName $pointerName | Out-Null
        Write-Host 'Created Exchange service-principal pointer.'
    }

    $assignment = Get-ManagementRoleAssignment -Identity $assignmentName -ErrorAction SilentlyContinue
    if ($assignment) {
        if ([string]$assignment.Role -ne 'Application Calendars.ReadWrite') {
            throw "Existing role assignment '$assignmentName' has unexpected role '$($assignment.Role)'."
        }
        Set-ManagementRoleAssignment -Identity $assignmentName -CustomResourceScope $scopeName
        Write-Host "Updated application role assignment: $assignmentName"
    }
    else {
        New-ManagementRoleAssignment `
            -Name $assignmentName `
            -Role 'Application Calendars.ReadWrite' `
            -App $RuntimeServicePrincipalObjectId `
            -CustomResourceScope $scopeName | Out-Null
        Write-Host "Created Application Calendars.ReadWrite assignment for $Mailbox only."
    }

    $test = @(Test-ServicePrincipalAuthorization -Identity $RuntimeServicePrincipalObjectId -Resource $Mailbox)
    $calendarGrant = @($test | Where-Object {
        $_.RoleName -eq 'Application Calendars.ReadWrite' -and $_.InScope -eq $true
    })

    if ($calendarGrant.Count -eq 0) {
        $test | Format-Table RoleName, GrantedPermissions, AllowedResourceScope, ScopeType, InScope -AutoSize
        throw "Exchange authorization validation failed for $Mailbox."
    }

    Write-Host "Exchange Application RBAC PASSED for $Mailbox."
}
finally {
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
}
