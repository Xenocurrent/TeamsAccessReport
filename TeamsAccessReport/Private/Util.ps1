# Pure helper functions without side effects (unit tested).

$script:TarRandom = [System.Random]::new()

function Get-TarProperty {
    # StrictMode-safe property read for hashtable, PSCustomObject and $null.
    [CmdletBinding()]
    param([AllowNull()]$InputObject, [Parameter(Mandatory)][string]$Name)
    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($Name)) { return $InputObject[$Name] }
        return $null
    }
    $prop = $InputObject.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertTo-TarArray {
    # Normalizes a collection: array, {results:[]} (odata=verbose) or a single object.
    # Emits elements to the pipeline - call as @(ConvertTo-TarArray $x).
    [CmdletBinding()]
    param([AllowNull()]$InputObject)
    if ($null -eq $InputObject) { return }
    $results = Get-TarProperty $InputObject 'results'
    if ($null -ne $results) { return $results }
    return $InputObject
}

function Get-TarHttpStatusFromError {
    # Only for sources that do not expose a status code (PnP). Parses status phrases, never arbitrary digits.
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $resp = Get-TarProperty $ErrorRecord.Exception 'Response'
    $code = Get-TarProperty $resp 'StatusCode'
    if ($code) { return [int]$code }
    $msg = [string]$ErrorRecord.Exception.Message
    if ($msg -match 'status code[^0-9]{0,40}\b([1-5]\d{2})\b') { return [int]$Matches[1] }
    if ($msg -match '\(([1-5]\d{2})\)') { return [int]$Matches[1] }
    $map = [ordered]@{ TooManyRequests = 429; InternalServerError = 500; BadGateway = 502; ServiceUnavailable = 503; GatewayTimeout = 504; Forbidden = 403; Unauthorized = 401; NotFound = 404 }
    foreach ($k in $map.Keys) { if ($msg -match "\b$k\b") { return $map[$k] } }
    return 0
}

function Get-TarBackoffDelay {
    # Exponential backoff with jitter, capped at 120 s. Retry-After takes precedence.
    [CmdletBinding()]
    [OutputType([int])]
    param([Parameter(Mandatory)][int]$Attempt, [Parameter(Mandatory)][int]$BaseDelaySeconds, [int]$RetryAfterSeconds = 0)
    $base = if ($RetryAfterSeconds -gt 0) { $RetryAfterSeconds } else { [math]::Min(120, $BaseDelaySeconds * [math]::Pow(2, $Attempt - 1)) }
    return [int]($base * 1000) + $script:TarRandom.Next(0, 1000)
}

function Protect-TarCsvValue {
    # Protects against formula injection in Excel (external users control their own display names).
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)][AllowNull()]$InputObject)
    process {
        $copy = [ordered]@{}
        foreach ($p in $InputObject.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [string] -and $v.Length -eq 0) { $v = $null }   # canonical form: empty string and $null are written identically
            elseif ($v -is [string] -and $v -match '^[=+\-@\t\r]') { $v = "'" + $v }
            $copy[$p.Name] = $v
        }
        [pscustomobject]$copy
    }
}

function Test-TarGuestPrincipal {
    [CmdletBinding()]
    [OutputType([bool])]
    param([AllowNull()][string]$UserType, [AllowNull()][string]$Upn)
    return ($UserType -eq 'Guest' -or [string]$Upn -match '#EXT#')
}
