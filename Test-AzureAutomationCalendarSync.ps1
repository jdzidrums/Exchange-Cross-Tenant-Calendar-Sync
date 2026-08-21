#Requires -Version 7.4
<#
.SYNOPSIS
    Starts the deployed calendar-sync runbook in Test mode and waits for the
    Azure Automation job result.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DeploymentSummaryPath,

    [ValidateRange(1,30)]
    [int]$PollSeconds = 5
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Automation -ErrorAction Stop

$path = [IO.Path]::GetFullPath($DeploymentSummaryPath)
if (-not (Test-Path $path)) { throw "Deployment summary not found: $path" }
$cfg = Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 30

if (-not (Get-AzContext -ErrorAction SilentlyContinue)) { Connect-AzAccount | Out-Null }
Set-AzContext -SubscriptionId $cfg.Azure.SubscriptionId | Out-Null

$params = @{
    KeyVaultName        = [string]$cfg.Azure.KeyVaultName
    DzidrumsSecretName  = [string]$cfg.Dzidrums.KeyVaultSecretName
    UltraProSecretName  = [string]$cfg.UltraPro.KeyVaultSecretName
    StateStorageAccount = [string]$cfg.Azure.StorageAccountName
    StateContainer      = 'calendar-sync'
    StateBlob           = 'calendar-sync-state.json'
    LockBlob            = 'calendar-sync.lock'
    DzidrumsTenantId    = [string]$cfg.Dzidrums.TenantId
    DzidrumsClientId    = [string]$cfg.Dzidrums.ClientId
    DzidrumsMailbox     = [string]$cfg.Dzidrums.Mailbox
    UltraProTenantId    = [string]$cfg.UltraPro.TenantId
    UltraProClientId    = [string]$cfg.UltraPro.ClientId
    UltraProMailbox     = [string]$cfg.UltraPro.Mailbox
    DetailMode          = [string]$cfg.Sync.DetailMode
    RespectPrivate      = [bool]$cfg.Sync.RespectPrivate
    CopyReminders       = [bool]$cfg.Sync.CopyReminders
    PastDays            = [int]$cfg.Sync.PastDays
    FutureDays          = [int]$cfg.Sync.FutureDays
    RebaselineDays      = [int]$cfg.Sync.RebaselineDays
    MaxDeltaPages       = 1000
    Mode                = 'Test'
}

Write-Host "Starting $($cfg.Azure.RunbookName) in Test mode..."
$job = Start-AzAutomationRunbook `
    -ResourceGroupName $cfg.Azure.ResourceGroupName `
    -AutomationAccountName $cfg.Azure.AutomationAccountName `
    -Name $cfg.Azure.RunbookName `
    -Parameters $params

$terminalStates = @('Completed','Failed','Stopped','Suspended')
do {
    Start-Sleep -Seconds $PollSeconds
    $current = Get-AzAutomationJob `
        -ResourceGroupName $cfg.Azure.ResourceGroupName `
        -AutomationAccountName $cfg.Azure.AutomationAccountName `
        -Id $job.JobId
    Write-Host "Job $($job.JobId): $($current.Status)"
} while ($current.Status -notin $terminalStates)

Write-Host "`n--- Job output ---"
$output = Get-AzAutomationJobOutput `
    -ResourceGroupName $cfg.Azure.ResourceGroupName `
    -AutomationAccountName $cfg.Azure.AutomationAccountName `
    -Id $job.JobId `
    -Stream Any | Get-AzAutomationJobOutputRecord

foreach ($record in @($output)) {
    if ($record.Type -eq 'Error' -and $record.Value.Exception) {
        Write-Host $record.Value.Exception
    }
    elseif ($record.Type -eq 'Warning') {
        if ($record.Value -is [System.Collections.IDictionary] -and $record.Value.Contains('Message')) {
            Write-Host $record.Value['Message']
        }
        elseif ($null -ne $record.Value -and $record.Value.PSObject.Properties['Message']) {
            Write-Host $record.Value.Message
        }
        else {
            Write-Host ([string]$record.Value)
        }
    }
    elseif ($record.Value -is [System.Collections.IDictionary] -and $record.Value.Contains('value')) {
        Write-Host $record.Value['value']
    }
    elseif ($null -ne $record.Value -and $record.Value.PSObject.Properties['value']) {
        Write-Host $record.Value.value
    }
    elseif ($null -ne $record.Value) {
        Write-Host $record.Value
    }
}

if ($current.Status -ne 'Completed') {
    if (-not [string]::IsNullOrWhiteSpace([string]$current.Exception)) {
        Write-Host "`n--- Job exception ---"
        Write-Host $current.Exception
    }
    throw "Test runbook job ended in state '$($current.Status)'."
}

Write-Host "`nTEST JOB COMPLETED. Look for 'TEST PASSED' in the output above." -ForegroundColor Green
