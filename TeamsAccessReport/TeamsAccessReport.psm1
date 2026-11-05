Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load order is explicit: Constants before anything that references $script:TarFile etc.
$privateFiles = 'Constants', 'Logging', 'Util', 'Store', 'Graph', 'Connection', 'Selection', 'Gui', 'Inventory', 'SharePointScan', 'Resolution', 'Analysis'
foreach ($f in $privateFiles) { . (Join-Path $PSScriptRoot 'Private' "$f.ps1") }
foreach ($f in Get-ChildItem (Join-Path $PSScriptRoot 'Public') -Filter *.ps1) { . $f.FullName }

Export-ModuleMember -Function 'Invoke-TeamsAccessReport', 'New-TeamsAccessReportProfile', 'Get-TeamsAccessReportProfile', 'Remove-TeamsAccessReportProfile'

# Tab completion of saved profile names (runs in module scope, so private functions are available).
$profileNameCompleter = {
    param($commandName, $parameterName, $wordToComplete)
    $null = $commandName, $parameterName
    try { @(Read-TarProfileStore) | Where-Object { $_.Name -like "$wordToComplete*" } | ForEach-Object { [System.Management.Automation.CompletionResult]::new($_.Name, $_.Name, 'ParameterValue', "$($_.AuthMode) $($_.TenantQuery ?? $_.TenantId)") } }
    catch { $null = $_ }
}
Register-ArgumentCompleter -CommandName 'Invoke-TeamsAccessReport' -ParameterName 'ProfileName' -ScriptBlock $profileNameCompleter
Register-ArgumentCompleter -CommandName 'Get-TeamsAccessReportProfile', 'Remove-TeamsAccessReportProfile', 'New-TeamsAccessReportProfile' -ParameterName 'Name' -ScriptBlock $profileNameCompleter
