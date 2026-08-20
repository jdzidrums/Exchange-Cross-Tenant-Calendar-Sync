#Requires -Version 7.4
<#
.SYNOPSIS
    Starts the deployed Azure Automation calendar sync on demand in Sync,
    Test, or Rebaseline mode and optionally waits for completion.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$DeploymentSummaryPath,

    [ValidateSet('Sync','Test','Rebaseline')]
    [string]$Mode = 'Sync',

    [switch]$NoWait,

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
$params.Mode = $Mode

$job = Start-AzAutomationRunbook `
    -ResourceGroupName $cfg.Azure.ResourceGroupName `
    -AutomationAccountName $cfg.Azure.AutomationAccountName `
    -Name $cfg.Azure.RunbookName `
    -Parameters $params

Write-Host "Started $Mode job: $($job.JobId)"
if ($NoWait) { return $job }

$terminalStates = @('Completed','Failed','Stopped','Suspended')
do {
    Start-Sleep -Seconds $PollSeconds
    $current = Get-AzAutomationJob `
        -ResourceGroupName $cfg.Azure.ResourceGroupName `
        -AutomationAccountName $cfg.Azure.AutomationAccountName `
        -Id $job.JobId
    Write-Host "Status: $($current.Status)"
} while ($current.Status -notin $terminalStates)

Get-AzAutomationJobOutput `
    -ResourceGroupName $cfg.Azure.ResourceGroupName `
    -AutomationAccountName $cfg.Azure.AutomationAccountName `
    -Id $job.JobId `
    -Stream Any | Get-AzAutomationJobOutputRecord | ForEach-Object {
        if ($_.Type -eq 'Error' -and $_.Value.Exception) { Write-Host $_.Value.Exception }
        elseif ($null -ne $_.Value) { Write-Host $_.Value }
    }

if ($current.Status -ne 'Completed') {
    throw "$Mode run ended in state '$($current.Status)'."
}
