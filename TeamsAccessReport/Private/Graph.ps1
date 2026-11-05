# Microsoft Graph client. Invoke-TarGraphRaw is the ONLY place that touches Invoke-MgGraphRequest
# (test seam). The status code is read deterministically via -StatusCodeVariable, not by parsing error messages.
# The SDK's built-in retry handler (429/503/504, Retry-After) is configured in Connect-TarGraph.

function Invoke-TarGraphRaw {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)
    $statusCode = 0; $headers = $null
    $body = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -SkipHttpErrorCheck `
        -StatusCodeVariable 'statusCode' -ResponseHeadersVariable 'headers' -ErrorAction Stop
    [pscustomobject]@{ StatusCode = [int]$statusCode; Body = $body; RetryAfterSeconds = (ConvertFrom-TarRetryAfterHeader -Headers $headers) }
}

function Invoke-TarGraphRequest {
    # Outer retry (max 3) on top of SDK retry: handles 500/502 and exhausted SDK retry. Returns the body, or $null (404 with -AllowNotFound).
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Uri, [switch]$AllowNotFound)
    $outerMax = 3
    for ($attempt = 1; $attempt -le $outerMax; $attempt++) {
        Write-TarLog "GET $Uri ($attempt/$outerMax)" -Level DEBUG
        $r = Invoke-TarGraphRaw -Uri $Uri
        if ($r.StatusCode -ge 200 -and $r.StatusCode -lt 300) { return $r.Body }
        if ($r.StatusCode -eq 404 -and $AllowNotFound) { return $null }
        $graphError = Get-TarProperty (Get-TarProperty $r.Body 'error') 'message'
        if ($r.StatusCode -notin $script:TarRetryableStatus -or $attempt -eq $outerMax) {
            $ex = [System.Net.Http.HttpRequestException]::new("Graph HTTP $($r.StatusCode) for $Uri :: $graphError")
            # Callers branch on these instead of parsing the message.
            $ex.Data['StatusCode'] = $r.StatusCode
            $ex.Data['GraphMessage'] = [string]$graphError
            throw $ex
        }
        $wait = Get-TarBackoffDelay -Attempt ($attempt + 2) -BaseDelaySeconds $Context.BaseDelaySeconds -RetryAfterSeconds $r.RetryAfterSeconds
        Write-TarLog "Graph HTTP $($r.StatusCode) after SDK retry; outer attempt $attempt/$outerMax, waiting $wait ms" -Level WARN -Data @{ uri = $Uri }
        Start-Sleep -Milliseconds $wait
    }
}

function Get-TarGraphCollection {
    # Follows @odata.nextLink. Emits elements to the pipeline - call as @(Get-TarGraphCollection ...).
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Uri)
    $next = $Uri
    while ($next) {
        $page = Invoke-TarGraphRequest -Context $Context -Uri $next
        foreach ($v in @(Get-TarProperty $page 'value')) { if ($null -ne $v) { $v } }
        $next = Get-TarProperty $page '@odata.nextLink'
    }
}

function Get-TarRequiredGraphScope {
    [CmdletBinding()]
    [OutputType([string[]])]
    param([switch]$AppOnly)
    $common = 'Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All', 'ChannelMember.Read.All', 'Files.Read.All', 'GroupMember.Read.All'
    if ($AppOnly) { return [string[]]($common + 'User.Read.All') }
    return [string[]](@('User.Read') + $common + 'User.ReadBasic.All')
}

function Connect-TarGraph {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [string]$TenantId)
    $scopes = Get-TarRequiredGraphScope -AppOnly:$Context.AppOnly
    if ($Context.AppOnly) {
        Write-TarLog "Connecting app-only. Tenant=$TenantId ClientId=$($Context.ClientId)"
        Connect-MgGraph -TenantId $TenantId -ClientId $Context.ClientId -CertificateThumbprint $Context.CertificateThumbprint -NoWelcome -ErrorAction Stop
    } else {
        $connectArgs = @{ Scopes = $scopes; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($TenantId) { $connectArgs.TenantId = $TenantId }
        Write-TarLog "Connecting delegated. Tenant=$(if ($TenantId) { $TenantId } else { '<resolved at sign-in>' })"
        Write-TarLog 'Delegated mode is limited by your own role - app-only gives the complete picture.' -Level WARN
        Connect-MgGraph @connectArgs
    }
    $mg = Get-MgContext
    if (-not $mg -or -not $mg.TenantId) { throw 'No Graph context after sign-in.' }
    $null = Set-MgRequestContext -MaxRetry $Context.MaxRetries -RetryDelay $Context.BaseDelaySeconds
    $Context.TenantId = $mg.TenantId
    $Context.AuthType = [string]$mg.AuthType

    try {
        $org = @(Get-TarProperty (Invoke-TarGraphRequest -Context $Context -Uri "$($Context.GraphBase)/organization?`$select=id,displayName,verifiedDomains") 'value')[0]
        $domains = @(Get-TarProperty $org 'verifiedDomains')
        $Context.TenantName = Get-TarProperty $org 'displayName'
        $Context.TenantDomain = Get-TarProperty (@($domains | Where-Object { Get-TarProperty $_ 'isDefault' })[0]) 'name'
        $Context.InitialDomain = Get-TarProperty (@($domains | Where-Object { Get-TarProperty $_ 'isInitial' })[0]) 'name'
    } catch { Write-TarLog "Could not read tenant name: $($_.Exception.Message). Continuing with ID." -Level WARN }

    Write-TarLog ('Tenant: {0} | {1} | {2} | {3} | {4}' -f $Context.TenantId, ($Context.TenantName ?? '-'), ($Context.TenantDomain ?? '-'), ($mg.Account ?? $mg.AppName), $Context.AuthType)
    $missing = @($scopes | Where-Object { $_ -ne 'User.Read' -and $_ -notin @($mg.Scopes) })
    if ($missing) { Write-TarLog "Token is missing: $($missing -join ', ') - the corresponding steps will return 403 (see Errors.csv/Teams.csv)." -Level WARN }
}
