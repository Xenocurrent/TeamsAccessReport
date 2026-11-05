#Requires -Version 7.4.6
<#
.SYNOPSIS
    Process boundary for phase 2. Started by Start-TarSharePointScan - not intended to be run manually.
    Loads only the required private files + PnP.PowerShell (never Microsoft.Graph in the same process).
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$ConfigPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$InformationPreference = 'Continue'

$private = Join-Path $PSScriptRoot '..' 'Private'
foreach ($f in 'Constants', 'Logging', 'Util', 'Store', 'SharePointScan') { . (Join-Path $private "$f.ps1") }

$cfg = Get-Content $ConfigPath -Raw | ConvertFrom-Json
$ctx = $cfg.Context
Initialize-TarLog -Path (Join-Path $ctx.OutputPath $script:TarFile.Log)

$pnp = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1
if (-not $pnp) { throw 'PnP.PowerShell is missing. Install: Install-Module PnP.PowerShell -Scope CurrentUser' }
if ($pnp.Version -lt [version]'3.0.0') { throw "PnP.PowerShell $($pnp.Version) found - 3.0+ is required." }
Import-Module PnP.PowerShell -RequiredVersion $pnp.Version -ErrorAction Stop
Write-TarLog "Worker started. PnP $($pnp.Version)."

Invoke-TarSharePointScan -Context $ctx -Sites ([string[]]$cfg.Sites) -ClientId $cfg.ClientId -Thumbprint $cfg.Thumbprint -Tenant $cfg.Tenant
exit 0
