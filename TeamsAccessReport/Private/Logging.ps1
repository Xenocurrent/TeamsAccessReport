# Structured logging: JSON Lines to file + stream-based console output.
# No Write-Host: INFO -> Information stream, WARN/ERROR -> Warning, DEBUG -> Verbose.

$script:TarLogPath = $null

function Initialize-TarLog {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $script:TarLogPath = $Path
}

function Clear-TarLogTarget {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Resets an in-memory module variable only.')]
    # Detaches the file sink so entries from a later command never land in a previous run's run.jsonl.
    [CmdletBinding()]
    param()
    $script:TarLogPath = $null
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
