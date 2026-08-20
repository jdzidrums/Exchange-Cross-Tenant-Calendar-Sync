#Requires -Version 7.2
<#
.SYNOPSIS
  Creates the GitHub production environment and loads the bootstrap output as
  repository-level Actions variables.

.DESCRIPTION
  GitHub environment-level configuration variables are not available early
  enough for job-level conditions and OIDC action inputs. This helper therefore
  stores all non-secret deployment configuration as repository Actions variables.
  The production environment is still used for deployment protection and as the
  OIDC subject context.

  CALENDAR_SYNC_CICD_ENABLED is forced to false during bootstrap. Enable it only
  after a successful manual workflow_dispatch deployment.
#>

[CmdletBinding()]
param(
    [string]$ConfigurationPath = './GitHubActions.EnvironmentVariables.json',
    [string]$Repository = 'jdzidrums/Exchange-Cross-Tenant-Calendar-Sync',
    [string]$Environment = 'production'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Get-Command gh -ErrorAction SilentlyContinue)) {
    throw 'GitHub CLI (gh) is required. Install it and run gh auth login first.'
}

& gh auth status
if ($LASTEXITCODE -ne 0) { throw 'GitHub CLI is not authenticated.' }

$path = [System.IO.Path]::GetFullPath($ConfigurationPath)
if (-not (Test-Path $path)) { throw "Configuration file not found: $path" }
$config = Get-Content -Path $path -Raw -Encoding UTF8 | ConvertFrom-Json -Depth 30
if (-not $config.Variables) { throw 'Configuration does not contain a Variables object.' }

# Create/reuse the deployment environment. Protection rules can then be added in GitHub UI.
'{}' | & gh api --method PUT "repos/$Repository/environments/$Environment" --input - | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Unable to create/update GitHub environment '$Environment'." }

foreach ($property in $config.Variables.PSObject.Properties) {
    $name = [string]$property.Name
    $value = [string]$property.Value
    if ($name -eq 'CALENDAR_SYNC_CICD_ENABLED') { $value = 'false' }

    & gh variable set $name --body $value --repo $Repository
    if ($LASTEXITCODE -ne 0) { throw "Unable to set repository Actions variable '$name'." }
}

Write-Host "Configured repository Actions variables for $Repository."
Write-Host "Created/reused deployment environment '$Environment'."
Write-Host 'CALENDAR_SYNC_CICD_ENABLED=false was enforced.'
Write-Host 'Next: add production environment protection/review rules, then manually dispatch the CI/CD workflow with deploy=true.'
