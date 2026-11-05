#Requires -Version 7.2
<#
.SYNOPSIS
    Bakåtkompatibel ingång. All logik finns i modulen ./TeamsAccessReport (Invoke-TeamsAccessReport).
    Argument vidarebefordras oförändrade (@args bevarar parameternamn och switchar).
.EXAMPLE
    ./Get-TeamsAccessReport.ps1 -AllTeams -Force
#>
Import-Module (Join-Path $PSScriptRoot 'TeamsAccessReport' 'TeamsAccessReport.psd1') -Force
Invoke-TeamsAccessReport @args
