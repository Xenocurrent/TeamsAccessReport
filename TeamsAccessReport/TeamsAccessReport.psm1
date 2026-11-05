Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Load order is explicit: Constants before anything that references $script:TarFile etc.
$privateFiles = 'Constants', 'Logging', 'Util', 'Store', 'Graph', 'Selection', 'Inventory', 'SharePointScan', 'Resolution', 'Analysis'
foreach ($f in $privateFiles) { . (Join-Path $PSScriptRoot 'Private' "$f.ps1") }
foreach ($f in Get-ChildItem (Join-Path $PSScriptRoot 'Public') -Filter *.ps1) { . $f.FullName }

Export-ModuleMember -Function 'Invoke-TeamsAccessReport'
