#Requires -Version 7.4.6
<#
.SYNOPSIS
    Processgräns för fas 2. Startas av Start-TarSharePointScan – inte avsedd att köras manuellt.
    Laddar endast de privata filer som behövs + PnP.PowerShell (aldrig Microsoft.Graph i samma process).
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
if (-not $pnp) { throw 'PnP.PowerShell saknas. Installera: Install-Module PnP.PowerShell -Scope CurrentUser' }
if ($pnp.Version -lt [version]'3.0.0') { throw "PnP.PowerShell $($pnp.Version) hittad – 3.0+ krävs." }
Import-Module PnP.PowerShell -RequiredVersion $pnp.Version -ErrorAction Stop
Write-TarLog "Worker startad. PnP $($pnp.Version)."

Invoke-TarSharePointScan -Context $ctx -Sites ([string[]]$cfg.Sites) -ClientId $cfg.ClientId -Thumbprint $cfg.Thumbprint -Tenant $cfg.Tenant
exit 0
