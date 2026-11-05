#Requires -Version 7.2
<#
.SYNOPSIS
    Backward-compatible entry point. All logic lives in the ./TeamsAccessReport module (Invoke-TeamsAccessReport).
    Arguments are forwarded unchanged (@args preserves parameter names and switches).
.EXAMPLE
    ./Get-TeamsAccessReport.ps1 -AllTeams -Force
#>
Import-Module (Join-Path $PSScriptRoot 'TeamsAccessReport' 'TeamsAccessReport.psd1') -Force
Invoke-TeamsAccessReport @args
