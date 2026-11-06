#Requires -Version 7.2
<#
.SYNOPSIS
    Runs New-TeamsAccessReportApp for one tenant in its own process (used by the start window's 'Create app...'
    button) so the admin sign-in never mixes with the report's Graph session. Writes the result row as JSON.
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)][string]$ResultPath)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try {
    Import-Module (Join-Path $PSScriptRoot '..' 'TeamsAccessReport.psd1') -Force
    # The start window saves the profile itself (with the options chosen there).
    $r = New-TeamsAccessReportApp -Tenant $Tenant -NoProfile -Confirm:$false
    $r | ConvertTo-Json | Set-Content -Path $ResultPath -Encoding utf8
    Write-Information "Status: $($r.Status). This window can be closed." -InformationAction Continue
} catch {
    [ordered]@{ Tenant = $Tenant; Status = 'Failed'; Error = $_.Exception.Message } | ConvertTo-Json | Set-Content -Path $ResultPath -Encoding utf8
    Write-Warning $_.Exception.Message
}
