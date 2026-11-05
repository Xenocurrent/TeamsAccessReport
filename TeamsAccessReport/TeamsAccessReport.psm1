Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ordningen är explicit: Constants före allt som refererar $script:TarFile m.fl.
$privateFiles = 'Constants', 'Logging', 'Util', 'Store', 'Graph', 'Selection', 'Inventory', 'SharePointScan', 'Resolution', 'Analysis'
foreach ($f in $privateFiles) { . (Join-Path $PSScriptRoot 'Private' "$f.ps1") }
foreach ($f in Get-ChildItem (Join-Path $PSScriptRoot 'Public') -Filter *.ps1) { . $f.FullName }

Export-ModuleMember -Function 'Invoke-TeamsAccessReport'
