#Requires -Version 7.4
[CmdletBinding()]
param([string]$Path = '.')

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$root = [System.IO.Path]::GetFullPath($Path)
$files = Get-ChildItem -Path $root -Recurse -File -Filter '*.ps1'
$errors = New-Object System.Collections.Generic.List[object]

foreach ($file in $files) {
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
    foreach ($parseError in @($parseErrors)) {
        $errors.Add([pscustomobject]@{ File = $file.FullName; Message = $parseError.Message; Extent = $parseError.Extent.Text })
    }

    if ($file.Name -eq 'AzureAutomation-CrossTenantCalendarSync.ps1') {
        $writeLog = $ast.Find({
            param($node)
            $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Write-Log'
        }, $true)
        $writesToOutput = $writeLog -and $writeLog.Body.Find({
            param($node)
            $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Write-Output'
        }, $true)
        if (-not $writeLog -or -not $writesToOutput) {
            $errors.Add([pscustomobject]@{
                File = $file.FullName
                Message = 'Write-Log must emit concise informational records with Write-Output.'
                Extent = 'Write-Log'
            })
        }

        foreach ($functionName in @('Acquire-StorageLease','Get-State','Invoke-Graph')) {
            $dataFunction = $ast.Find({
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $functionName
            }, $true)
            $nestedLogs = if ($dataFunction) {
                @($dataFunction.Body.FindAll({
                    param($node)
                    $node -is [System.Management.Automation.Language.CommandAst] -and $node.GetCommandName() -eq 'Write-Log'
                }, $true))
            }
            else {
                @()
            }
            if (-not $dataFunction -or @($nestedLogs).Count -gt 0) {
                $errors.Add([pscustomobject]@{
                    File = $file.FullName
                    Message = "$functionName must not log to the success pipeline while returning data."
                    Extent = $functionName
                })
            }
        }
    }
}

if ($errors.Count -gt 0) {
    $errors | Format-Table -AutoSize | Out-String | Write-Error
    throw "$($errors.Count) PowerShell parser error(s) found."
}

if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer | Select-Object -First 1)) {
    Install-Module PSScriptAnalyzer -Scope CurrentUser -Force -Repository PSGallery
}
Import-Module PSScriptAnalyzer
$analysis = Invoke-ScriptAnalyzer -Path $root -Recurse -Severity Error
if ($analysis) {
    $analysis | Format-Table RuleName,Severity,ScriptName,Line,Message -AutoSize
    throw "$(@($analysis).Count) PSScriptAnalyzer error(s) found."
}

Write-Host "Validation passed for $($files.Count) PowerShell files."
