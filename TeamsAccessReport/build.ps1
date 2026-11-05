#Requires -Version 7.2
<#
.SYNOPSIS
    Quality gate: PSScriptAnalyzer (0 findings allowed) + Pester with code coverage.
.EXAMPLE
    ./build.ps1            # analysis + tests
    ./build.ps1 -Format    # format all files according to PSScriptAnalyzerSettings.psd1 first
#>
[CmdletBinding()]
param([switch]$Format, [ValidateRange(0, 100)][int]$MinimumCoverage = 80)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$settings = Join-Path $root 'PSScriptAnalyzerSettings.psd1'
foreach ($m in @{ Name = 'PSScriptAnalyzer'; Min = '1.22.0' }, @{ Name = 'Pester'; Min = '5.5.0' }) {
    if (-not (Get-Module -ListAvailable $m.Name | Where-Object Version -ge $m.Min)) { throw "$($m.Name) >= $($m.Min) is required: Install-Module $($m.Name) -Scope CurrentUser" }
}
$sources = Get-ChildItem $root -Recurse -Include *.ps1, *.psm1, *.psd1 | Where-Object FullName -NotMatch '[\\/]Tests[\\/]'

if ($Format) {
    foreach ($f in $sources) {
        $src = [IO.File]::ReadAllText($f.FullName)
        $fmt = (Invoke-Formatter -ScriptDefinition $src -Settings $settings) -replace '[ \t]+(\r?\n)', '$1'
        if ($fmt -ne $src) { [IO.File]::WriteAllText($f.FullName, $fmt, [Text.UTF8Encoding]::new($true)); Write-Information "Formatted: $($f.Name)" -InformationAction Continue }
    }
}

$issues = @($sources | ForEach-Object { Invoke-ScriptAnalyzer -Path $_.FullName -Settings $settings })
$issues | Format-Table -AutoSize ScriptName, Line, Severity, RuleName, Message | Out-String | Write-Information -InformationAction Continue
if ($issues.Count) { throw "PSScriptAnalyzer: $($issues.Count) findings." }
Write-Information 'PSScriptAnalyzer: 0 findings.' -InformationAction Continue

$cfg = New-PesterConfiguration
$cfg.Run.Path = Join-Path $root 'Tests'
$cfg.Run.Exit = $false
$cfg.Run.PassThru = $true
$cfg.Output.Verbosity = 'Detailed'
$cfg.CodeCoverage.Enabled = $true
# Gui.ps1 (Windows Forms) can only execute on Windows; it is measured there and excluded elsewhere.
$coverageFiles = @(Get-ChildItem (Join-Path $root 'Private'), (Join-Path $root 'Public') -Filter *.ps1 | Where-Object { $IsWindows -or $_.Name -ne 'Gui.ps1' })
$cfg.CodeCoverage.Path = @($coverageFiles.FullName)
$cfg.CodeCoverage.OutputPath = Join-Path $root 'coverage.xml'
$cfg.TestResult.Enabled = $true
$cfg.TestResult.OutputPath = Join-Path $root 'testResults.xml'
$result = Invoke-Pester -Configuration $cfg
$coverage = [math]::Round($result.CodeCoverage.CoveragePercent, 1)
Write-Information "Pester: $($result.PassedCount) passed, $($result.FailedCount) failed, coverage $coverage %" -InformationAction Continue
if ($result.FailedCount -gt 0) { throw "Pester: $($result.FailedCount) tests failed." }
if ($coverage -lt $MinimumCoverage) { throw "Code coverage $coverage % < $MinimumCoverage %." }
