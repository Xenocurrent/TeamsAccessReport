# Strukturerad loggning: JSON Lines till fil + ström-baserad konsolutskrift.
# Ingen Write-Host: INFO -> Information-strömmen, WARN/ERROR -> Warning, DEBUG -> Verbose.

$script:TarLogPath = $null

function Initialize-TarLog {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $script:TarLogPath = $Path
}

function Write-TarLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')][string]$Level = 'INFO',
        [hashtable]$Data
    )
    $entry = [ordered]@{ ts = (Get-Date).ToString('o'); level = $Level; msg = $Message }
    if ($Data) { $entry.data = $Data }
    if ($script:TarLogPath) {
        Add-Content -Path $script:TarLogPath -Value ($entry | ConvertTo-Json -Compress -Depth 5) -Encoding utf8
    }
    switch ($Level) {
        'ERROR' { Write-Warning "ERROR: $Message" }
        'WARN' { Write-Warning $Message }
        'DEBUG' { Write-Verbose $Message }
        default { Write-Information $Message }
    }
}
