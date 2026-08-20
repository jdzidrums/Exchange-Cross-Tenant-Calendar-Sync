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
    [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors) | Out-Null
    foreach ($parseError in @($parseErrors)) {
        $errors.Add([pscustomobject]@{ File = $file.FullName; Message = $parseError.Message; Extent = $parseError.Extent.Text })
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
