#Requires -Version 7.2
<#
.SYNOPSIS
  Publishes the calendar-sync runbook into an existing Azure Automation account
  and maintains twelve staggered hourly schedules for an effective 5-minute cadence.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$SubscriptionId,
    [Parameter(Mandatory)][string]$ResourceGroupName,
    [Parameter(Mandatory)][string]$Location,
    [Parameter(Mandatory)][string]$AutomationAccountName,
    [Parameter(Mandatory)][string]$RuntimeEnvironmentName,
    [Parameter(Mandatory)][string]$RunbookName,
    [Parameter(Mandatory)][string]$RunbookFile,
    [Parameter(Mandatory)][string]$KeyVaultName,
    [Parameter(Mandatory)][string]$StateStorageAccount,
    [Parameter(Mandatory)][string]$DzidrumsSecretName,
    [Parameter(Mandatory)][string]$UltraProSecretName,
    [Parameter(Mandatory)][string]$DzidrumsTenantId,
    [Parameter(Mandatory)][string]$DzidrumsClientId,
    [Parameter(Mandatory)][string]$DzidrumsMailbox,
    [Parameter(Mandatory)][string]$UltraProTenantId,
    [Parameter(Mandatory)][string]$UltraProClientId,
    [Parameter(Mandatory)][string]$UltraProMailbox,
    [ValidateSet('BusyOnly','SubjectLocation')][string]$DetailMode = 'BusyOnly',
    [bool]$RespectPrivate = $true,
    [bool]$CopyReminders = $false,
    [int]$PastDays = 30,
    [int]$FutureDays = 365,
    [int]$RebaselineDays = 1,
    [switch]$SkipSchedules
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }
}

Ensure-Module -Name Az.Accounts
Ensure-Module -Name Az.Automation
Import-Module Az.Accounts
Import-Module Az.Automation

if (-not (Test-Path $RunbookFile)) {
    throw "Runbook file not found: $RunbookFile"
}

$context = Get-AzContext
if (-not $context -or [string]$context.Subscription.Id -ne $SubscriptionId) {
    Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
}

function Convert-TokenToPlainText {
    param([Parameter(Mandatory)]$TokenValue)
    if ($TokenValue -is [securestring]) {
        return [System.Net.NetworkCredential]::new('', $TokenValue).Password
    }
    return [string]$TokenValue
}

function Get-ArmAccessToken {
    $token = Get-AzAccessToken -ResourceUrl 'https://management.azure.com/'
    Convert-TokenToPlainText $token.Token
}

function Invoke-ArmRest {
    param(
        [Parameter(Mandatory)][ValidateSet('GET','PUT','POST','DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [object]$Body,
        [string]$ContentType = 'application/json'
    )

    $params = @{
        Method      = $Method
        Uri         = $Uri
        Headers     = @{ Authorization = "Bearer $(Get-ArmAccessToken)" }
        ErrorAction = 'Stop'
    }
    if ($null -ne $Body) {
        $params.ContentType = $ContentType
        $params.Body = if ($ContentType -eq 'application/json') {
            $Body | ConvertTo-Json -Depth 30 -Compress
        }
        else {
            [string]$Body
        }
    }
    Invoke-RestMethod @params
}

function Wait-ForArmProvisioning {
    param([Parameter(Mandatory)][string]$Uri, [int]$MaxAttempts = 60)
    for ($i = 1; $i -le $MaxAttempts; $i++) {
        Start-Sleep -Seconds 3
        $item = Invoke-ArmRest -Method GET -Uri $Uri
        $stateProperty = if ($item.PSObject.Properties['properties']) {
            $item.properties.PSObject.Properties['provisioningState']
        }
        else {
            $null
        }
        $state = if ($stateProperty) { [string]$stateProperty.Value } else { '' }
        if ($state -eq 'Succeeded' -or [string]::IsNullOrWhiteSpace($state)) { return $item }
        if ($state -in @('Failed','Canceled')) { throw "ARM provisioning ended in state '$state'." }
    }
    throw "Timed out waiting for ARM provisioning: $Uri"
}

$api = '2024-10-23'
$base = "https://management.azure.com/subscriptions/$SubscriptionId/resourceGroups/$ResourceGroupName/providers/Microsoft.Automation/automationAccounts/$AutomationAccountName"

$runtimeUri = "$base/runtimeEnvironments/${RuntimeEnvironmentName}?api-version=$api"
$runtimeBody = @{
    name = $RuntimeEnvironmentName
    properties = @{
        runtime = @{ language = 'PowerShell'; version = '7.4' }
        defaultPackages = @{ Az = '12.3.0' }
    }
}
Invoke-ArmRest -Method PUT -Uri $runtimeUri -Body $runtimeBody | Out-Null
Wait-ForArmProvisioning -Uri $runtimeUri | Out-Null
Write-Host "Runtime Environment ready: $RuntimeEnvironmentName"

$runbookUri = "$base/runbooks/${RunbookName}?api-version=$api"
$runbookBody = @{
    name = $RunbookName
    location = $Location
    properties = @{
        description        = 'Cross-tenant calendar synchronization for Dzidrums and Ultra PRO.'
        draft              = @{}
        logProgress        = $false
        logVerbose         = $true
        logActivityTrace   = 1
        runbookType        = 'PowerShell'
        runtimeEnvironment = $RuntimeEnvironmentName
    }
    tags = @{ workload = 'calendar-sync'; managedBy = 'github-actions' }
}
Invoke-ArmRest -Method PUT -Uri $runbookUri -Body $runbookBody | Out-Null
Wait-ForArmProvisioning -Uri $runbookUri | Out-Null

$contentUri = "$base/runbooks/${RunbookName}/draft/content?api-version=$api"
$content = Get-Content -Path $RunbookFile -Raw -Encoding UTF8
Invoke-ArmRest -Method PUT -Uri $contentUri -Body $content -ContentType 'text/plain' | Out-Null

$publishUri = "$base/runbooks/${RunbookName}/publish?api-version=$api"
$accepted = $false
for ($attempt = 1; $attempt -le 12 -and -not $accepted; $attempt++) {
    Start-Sleep -Seconds 5
    try {
        Invoke-ArmRest -Method POST -Uri $publishUri | Out-Null
        $accepted = $true
    }
    catch {
        if ($attempt -eq 12) { throw }
    }
}

for ($i = 1; $i -le 60; $i++) {
    Start-Sleep -Seconds 3
    $rb = Invoke-ArmRest -Method GET -Uri $runbookUri
    if ([string]$rb.properties.state -eq 'Published') { break }
    if ($i -eq 60) { throw "Timed out waiting for runbook '$RunbookName' to publish." }
}
Write-Host "Runbook published from Git commit: $env:GITHUB_SHA"

$runbookParameters = @{
    KeyVaultName        = $KeyVaultName
    DzidrumsSecretName  = $DzidrumsSecretName
    UltraProSecretName  = $UltraProSecretName
    StateStorageAccount = $StateStorageAccount
    StateContainer      = 'calendar-sync'
    StateBlob           = 'calendar-sync-state.json'
    LockBlob            = 'calendar-sync.lock'
    DzidrumsTenantId    = $DzidrumsTenantId
    DzidrumsClientId    = $DzidrumsClientId
    DzidrumsMailbox     = $DzidrumsMailbox
    UltraProTenantId    = $UltraProTenantId
    UltraProClientId    = $UltraProClientId
    UltraProMailbox     = $UltraProMailbox
    DetailMode          = $DetailMode
    RespectPrivate      = $RespectPrivate
    CopyReminders       = $CopyReminders
    PastDays            = $PastDays
    FutureDays          = $FutureDays
    RebaselineDays      = $RebaselineDays
    MaxDeltaPages       = 1000
    Mode                = 'Sync'
}

if (-not $SkipSchedules) {
    $minimumStart = [DateTimeOffset]::UtcNow.AddMinutes(10)
    $hourBase = [DateTimeOffset]::new(
        $minimumStart.Year, $minimumStart.Month, $minimumStart.Day,
        $minimumStart.Hour, 0, 0, [TimeSpan]::Zero
    )

    foreach ($minute in 0,5,10,15,20,25,30,35,40,45,50,55) {
        $name = 'CalendarSync-05m-{0:D2}' -f $minute
        $start = $hourBase.AddMinutes($minute)
        while ($start -lt $minimumStart) { $start = $start.AddHours(1) }

        $association = Get-AzAutomationScheduledRunbook `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -RunbookName $RunbookName `
            -ScheduleName $name `
            -ErrorAction SilentlyContinue
        if ($association) {
            Unregister-AzAutomationScheduledRunbook `
                -ResourceGroupName $ResourceGroupName `
                -AutomationAccountName $AutomationAccountName `
                -RunbookName $RunbookName `
                -ScheduleName $name `
                -Force -ErrorAction SilentlyContinue
        }

        $existingSchedule = Get-AzAutomationSchedule `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -Name $name -ErrorAction SilentlyContinue
        if ($existingSchedule) {
            Remove-AzAutomationSchedule `
                -ResourceGroupName $ResourceGroupName `
                -AutomationAccountName $AutomationAccountName `
                -Name $name -Force
        }

        New-AzAutomationSchedule `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -Name $name `
            -StartTime $start `
            -HourInterval 1 `
            -TimeZone 'UTC' `
            -Description "Calendar synchronization minute offset $minute" | Out-Null

        Register-AzAutomationScheduledRunbook `
            -ResourceGroupName $ResourceGroupName `
            -AutomationAccountName $AutomationAccountName `
            -RunbookName $RunbookName `
            -ScheduleName $name `
            -Parameters $runbookParameters | Out-Null

        Write-Host "Scheduled $name hourly starting $($start.ToString('u'))"
    }
}
else {
    Write-Host 'Schedule creation skipped; runbook remains published.'
}

Write-Host 'Azure Automation deployment completed.'
