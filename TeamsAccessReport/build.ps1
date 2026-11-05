#Requires -Version 7.2
<#
.SYNOPSIS
    Kvalitetsgrind: PSScriptAnalyzer (0 tillåtna anmärkningar) + Pester med kodtäckning.
.EXAMPLE
    ./build.ps1            # analys + tester
    ./build.ps1 -Format    # formatera alla filer enligt PSScriptAnalyzerSettings.psd1 först
#>
[CmdletBinding()]
param([switch]$Format, [ValidateRange(0, 100)][int]$MinimumCoverage = 80)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$settings = Join-Path $root 'PSScriptAnalyzerSettings.psd1'
foreach ($m in @{ Name = 'PSScriptAnalyzer'; Min = '1.22.0' }, @{ Name = 'Pester'; Min = '5.5.0' }) {
    if (-not (Get-Module -ListAvailable $m.Name | Where-Object Version -ge $m.Min)) { throw "$($m.Name) >= $($m.Min) krävs: Install-Module $($m.Name) -Scope CurrentUser" }
}
$sources = Get-ChildItem $root -Recurse -Include *.ps1, *.psm1, *.psd1 | Where-Object FullName -NotMatch '[\\/]Tests[\\/]'

if ($Format) {
    foreach ($f in $sources) {
        $src = [IO.File]::ReadAllText($f.FullName)
        $fmt = (Invoke-Formatter -ScriptDefinition $src -Settings $settings) -replace '[ \t]+(\r?\n)', '$1'
        if ($fmt -ne $src) { [IO.File]::WriteAllText($f.FullName, $fmt, [Text.UTF8Encoding]::new($true)); Write-Information "Formaterad: $($f.Name)" -InformationAction Continue }
    }
}

$issues = @($sources | ForEach-Object { Invoke-ScriptAnalyzer -Path $_.FullName -Settings $settings })
$issues | Format-Table -AutoSize ScriptName, Line, Severity, RuleName, Message | Out-String | Write-Information -InformationAction Continue
if ($issues.Count) { throw "PSScriptAnalyzer: $($issues.Count) anmärkningar." }
Write-Information 'PSScriptAnalyzer: 0 anmärkningar.' -InformationAction Continue

$cfg = New-PesterConfiguration
$cfg.Run.Path = Join-Path $root 'Tests'
$cfg.Run.Exit = $false
$cfg.Run.PassThru = $true
$cfg.Output.Verbosity = 'Detailed'
$cfg.CodeCoverage.Enabled = $true
$cfg.CodeCoverage.Path = @((Join-Path $root 'Private'), (Join-Path $root 'Public'))
$cfg.CodeCoverage.OutputPath = Join-Path $root 'coverage.xml'
$cfg.TestResult.Enabled = $true
$cfg.TestResult.OutputPath = Join-Path $root 'testResults.xml'
$result = Invoke-Pester -Configuration $cfg
$coverage = [math]::Round($result.CodeCoverage.CoveragePercent, 1)
Write-Information "Pester: $($result.PassedCount) ok, $($result.FailedCount) fel, täckning $coverage %" -InformationAction Continue
if ($result.FailedCount -gt 0) { throw "Pester: $($result.FailedCount) tester misslyckades." }
if ($coverage -lt $MinimumCoverage) { throw "Kodtäckning $coverage % < $MinimumCoverage %." }
