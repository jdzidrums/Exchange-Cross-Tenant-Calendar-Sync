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

$params = @{}
foreach ($p in $cfg.Sync.RunbookParameters.PSObject.Properties) {
    $params[$p.Name] = $p.Value
}
$params.Mode = 'Test'

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
    elseif ($null -ne $record.Value) {
        Write-Host $record.Value
    }
}

if ($current.Status -ne 'Completed') {
    throw "Test runbook job ended in state '$($current.Status)'."
}

Write-Host "`nTEST JOB COMPLETED. Look for 'TEST PASSED' in the output above." -ForegroundColor Green
