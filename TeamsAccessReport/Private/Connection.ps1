# Connection resolution (data logic only): tenant discovery, saved connection profiles, certificate lookup.
# Invoke-TarHttpRaw is the ONLY place that calls the public Entra discovery endpoint (test seam).
# Prompts live in Selection.ps1 (UI rule); this file never reads from the console.

function Invoke-TarHttpRaw {
    # Unauthenticated GET. Returns @{ StatusCode; Body }. Network failures (DNS, proxy, timeout) throw.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Uri)
    $status = 0
    $body = Invoke-RestMethod -Uri $Uri -Method Get -SkipHttpErrorCheck -StatusCodeVariable status -TimeoutSec 20 -ErrorAction Stop
    [pscustomobject]@{ StatusCode = [int]$status; Body = $body }
}

function ConvertTo-TarTenantQuery {
    # Pure: normalizes what a person types to an authority segment, or $null when it cannot be a tenant.
    #   GUID                 -> GUID (lower case)
    #   contoso              -> contoso.onmicrosoft.com
    #   contoso.com          -> contoso.com            (any verified domain of the tenant)
    #   admin@contoso.com    -> contoso.com            (an e-mail/UPN -> its domain)
    #   https://login.../x/  -> scheme and path are stripped
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$InputText)
    $t = $InputText.Trim().Trim('"', "'") -replace '^(?i)https?://', ''
    $t = ($t -split '[/?#]')[0]
    if ($t.Contains('@')) { $t = ($t -split '@')[-1] }
    $t = $t.Trim().TrimEnd('.').ToLowerInvariant()
    if (-not $t) { return $null }
    $g = [guid]::Empty
    if ([guid]::TryParse($t, [ref]$g)) { return $g.ToString() }
    if ($t -match '^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$') { return "$t.onmicrosoft.com" }
    if ($t.Length -le 253 -and $t -match '^(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+[a-z]{2,63}$') { return $t }
    return $null
}

function Resolve-TarTenant {
    # Resolves a tenant name/domain/ID to its tenant ID via the public OpenID discovery document (no sign-in).
    # Throws with Exception.Data['TarReason'] = 'Invalid' | 'NotFound' | 'Unavailable' so callers can branch.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Tenant, [ValidateRange(1, 10)][int]$MaxRetries = 3, [ValidateRange(1, 60)][int]$BaseDelaySeconds = 1)
    $fail = { param($Reason, $Message) $e = [System.InvalidOperationException]::new($Message); $e.Data['TarReason'] = $Reason; throw $e }
    $query = ConvertTo-TarTenantQuery -InputText $Tenant
    if (-not $query) { & $fail 'Invalid' "'$Tenant' is not a tenant ID, domain, e-mail address or tenant name." }
    $uri = "$script:TarAuthority/$query/v2.0/.well-known/openid-configuration"
    $r = $null
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        $why = $null
        try { $r = Invoke-TarHttpRaw -Uri $uri } catch { $r = $null; $why = $_.Exception.Message }
        if ($r -and $r.StatusCode -eq 200) { break }
        if ($r -and $r.StatusCode -in 400, 404) { & $fail 'NotFound' "Tenant '$query' was not found in Microsoft Entra ID. Check the spelling or use the tenant ID." }
        if (-not $why) { $why = "HTTP $($r.StatusCode)" }
        if ($attempt -eq $MaxRetries) { & $fail 'Unavailable' "Tenant lookup for '$query' failed after $MaxRetries attempts: $why" }
        $wait = Get-TarBackoffDelay -Attempt $attempt -BaseDelaySeconds $BaseDelaySeconds
        Write-TarLog "Tenant lookup for '$query' failed ($why), attempt $attempt/$MaxRetries; waiting $wait ms" -Level WARN
        Start-Sleep -Milliseconds $wait
    }
    $issuer = [string](Get-TarProperty $r.Body 'issuer')
    if ($issuer -notmatch '/([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})/v2\.0/?$') {
        & $fail 'Unavailable' "Unexpected discovery response for '$query' (issuer '$issuer')."
    }
    $tenantId = $Matches[1].ToLowerInvariant()
    $cloud = [string](Get-TarProperty $r.Body 'cloud_instance_name')
    if ($cloud -and $cloud -ne 'microsoftonline.com') {
        Write-TarLog "Tenant $tenantId is in the '$cloud' cloud; this module connects to the global Microsoft cloud only." -Level WARN
    }
    [pscustomobject]@{ TenantId = $tenantId; Query = $query; Region = [string](Get-TarProperty $r.Body 'tenant_region_scope'); Cloud = $cloud }
}

function Resolve-TarTenantOrPassThrough {
    # Like Resolve-TarTenant, but when the lookup service is unreachable the normalized input is passed on
    # (Connect-MgGraph accepts domains) so an offline/proxied machine is not blocked. Invalid/NotFound still throw.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Tenant)
    try { Resolve-TarTenant -Tenant $Tenant }
    catch {
        if ($_.Exception.Data['TarReason'] -ne 'Unavailable') { throw }
        $q = ConvertTo-TarTenantQuery -InputText $Tenant
        Write-TarLog "$($_.Exception.Message) Continuing with '$q'; the sign-in will validate it." -Level WARN
        [pscustomobject]@{ TenantId = $q; Query = $q; Region = $null; Cloud = $null }
    }
}

function ConvertTo-TarThumbprint {
    # Pure: removes spaces, colons and invisible characters (certmgr copies a leading U+200E) and upper-cases.
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Thumbprint)
    $t = ($Thumbprint -replace '[^0-9A-Fa-f]', '').ToUpperInvariant()
    if ($t.Length -ne 40) { throw "Certificate thumbprint must be 40 hexadecimal characters (SHA-1); got '$Thumbprint'." }
    $t
}

function ConvertTo-TarClientId {
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$ClientId)
    $g = [guid]::Empty
    if (-not [guid]::TryParse($ClientId.Trim(), [ref]$g)) { throw "Client ID must be a GUID (Application (client) ID of the Entra app); got '$ClientId'." }
    $g.ToString()
}

function Get-TarCertificateCandidate {
    # Certificates usable for app-only sign-in: private key present and valid now. CurrentUser\My, then LocalMachine\My.
    # Emits rows to the pipeline (Id = thumbprint), newest expiry first.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$Thumbprint)
    $now = Get-Date
    $rows = foreach ($location in 'CurrentUser', 'LocalMachine') {
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', $location)
        try { $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]'ReadOnly, OpenExistingOnly') } catch { continue }
        try {
            foreach ($c in $store.Certificates) {
                if ($Thumbprint -and $c.Thumbprint -ne $Thumbprint) { continue }
                [pscustomobject]@{
                    Id = $c.Thumbprint; Subject = $c.Subject; Expires = $c.NotAfter.ToString('yyyy-MM-dd'); Store = "$location\My"
                    HasPrivateKey = $c.HasPrivateKey; Valid = ($c.NotBefore -le $now -and $c.NotAfter -ge $now); NotAfter = $c.NotAfter
                }
            }
        } finally { $store.Close() }
    }
    $rows | Sort-Object NotAfter -Descending
}

function Assert-TarCertificate {
    # Fails early with a precise reason instead of an opaque sign-in error.
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Thumbprint)
    $found = @(Get-TarCertificateCandidate -Thumbprint $Thumbprint)
    if ($found.Count -eq 0) { throw "Certificate $Thumbprint was not found in CurrentUser\My or LocalMachine\My on this machine." }
    $usable = @($found | Where-Object { $_.HasPrivateKey -and $_.Valid })
    if ($usable.Count -eq 0) {
        $c = $found[0]
        $reason = if (-not $c.HasPrivateKey) { 'has no private key' } else { "is not valid now (expires $($c.Expires))" }
        throw "Certificate $Thumbprint ($($c.Subject)) in $($c.Store) $reason."
    }
}

# --- Connection profiles: non-secret settings per tenant, stored per user -------------------------------

function Get-TarProfileStorePath {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    if ($env:TEAMSACCESSREPORT_PROFILE_PATH) { return $env:TEAMSACCESSREPORT_PROFILE_PATH }
    $root = [Environment]::GetFolderPath('LocalApplicationData')
    if (-not $root) { $root = Join-Path $HOME '.local' 'share' }
    Join-Path $root 'TeamsAccessReport' 'profiles.json'
}

function Read-TarProfileStore {
    # Emits profile objects to the pipeline - call as @(Read-TarProfileStore).
    [CmdletBinding()]
    param()
    $p = Get-TarProfileStorePath
    if (-not (Test-Path $p)) { return }
    $doc = Get-Content -Path $p -Raw -Encoding utf8 | ConvertFrom-Json
    if ([int](Get-TarProperty $doc 'SchemaVersion') -gt $script:TarProfileSchemaVersion) { throw "$p uses profile schema $($doc.SchemaVersion); this module supports $script:TarProfileSchemaVersion. Update the module." }
    foreach ($x in @(Get-TarProperty $doc 'Profiles')) { if ($x) { $x } }
}

function Write-TarProfileStore {
    # Atomic replace: write a temp file next to the store, then move it over the old one.
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Profiles)
    $p = Get-TarProfileStorePath
    if (-not $PSCmdlet.ShouldProcess($p, 'Write connection profiles')) { return }
    $dir = Split-Path $p -Parent
    if (-not (Test-Path $dir)) { $null = New-Item -Path $dir -ItemType Directory -Force }
    $tmp = "$p.$PID.tmp"
    [pscustomobject]@{ SchemaVersion = $script:TarProfileSchemaVersion; Profiles = @($Profiles | Sort-Object Name) } |
        ConvertTo-Json -Depth 5 | Set-Content -Path $tmp -Encoding utf8
    Move-Item -Path $tmp -Destination $p -Force
}

function New-TarProfileRecord {
    # Pure constructor with validation. Stores identifiers only - never secrets (the certificate stays in the store).
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory constructor; does not change system state.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$TenantId, [string]$TenantQuery,
        [Parameter(Mandatory)][ValidateSet('AppOnly', 'Delegated')][string]$AuthMode, [string]$ClientId, [string]$CertificateThumbprint, [string]$PnPClientId)
    if ($Name -notmatch '^[A-Za-z0-9][A-Za-z0-9._-]{0,63}$') { throw "Profile name '$Name' is invalid: use 1-64 letters, digits, '.', '_' or '-'." }
    if ($AuthMode -eq 'AppOnly') {
        if (-not $ClientId -or -not $CertificateThumbprint) { throw 'An app-only profile needs both a client ID and a certificate thumbprint.' }
        $ClientId = ConvertTo-TarClientId $ClientId
        $CertificateThumbprint = ConvertTo-TarThumbprint $CertificateThumbprint
    } else { $ClientId = $null; $CertificateThumbprint = $null }
    if ($PnPClientId) { $PnPClientId = ConvertTo-TarClientId $PnPClientId }
    [pscustomobject]@{
        Name = $Name; TenantId = $TenantId; TenantQuery = $TenantQuery; AuthMode = $AuthMode
        ClientId = $ClientId; CertificateThumbprint = $CertificateThumbprint; PnPClientId = ($PnPClientId ? $PnPClientId : $null)
        UpdatedAt = (Get-Date).ToString('o')
    }
}

function Save-TarProfile {
    # Insert or replace by name (case-insensitive).
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][pscustomobject]$ProfileRecord)
    $others = @(Read-TarProfileStore | Where-Object { $_.Name -ne $ProfileRecord.Name })
    if ($PSCmdlet.ShouldProcess($ProfileRecord.Name, 'Save connection profile')) {
        Write-TarProfileStore -Profiles (@($others) + $ProfileRecord)
        Write-TarLog "Saved connection profile '$($ProfileRecord.Name)' ($($ProfileRecord.AuthMode), tenant $($ProfileRecord.TenantId)) to $(Get-TarProfileStorePath)"
    }
}

function Get-TarProfile {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Name)
    $all = @(Read-TarProfileStore)
    $p = @($all | Where-Object { $_.Name -eq $Name })
    if ($p.Count -eq 0) {
        $known = if ($all.Count) { ($all.Name | Sort-Object) -join ', ' } else { '<none>' }
        throw "Connection profile '$Name' not found. Saved profiles: $known. Create one with New-TeamsAccessReportProfile."
    }
    $p[0]
}

function ConvertFrom-TarProfile {
    # Profile -> connection object used by Invoke-TeamsAccessReport.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)]$ProfileRecord)
    [pscustomobject]@{
        AuthMode = $ProfileRecord.AuthMode; TenantId = $ProfileRecord.TenantId; ClientId = $ProfileRecord.ClientId
        CertificateThumbprint = $ProfileRecord.CertificateThumbprint; PnPClientId = $ProfileRecord.PnPClientId; ProfileName = $ProfileRecord.Name
    }
}

function Resolve-TarConnection {
    # Decides how to sign in. Order:
    #   1. -ProfileName                         -> that profile
    #   2. -ClientId + -CertificateThumbprint   -> app-only; tenant from -TenantId, Selection.json (-Resume) or a prompt
    #   3. -Resume with a profile in Selection.json, or exactly one saved profile for the resumed tenant
    #   4. -TenantId only                       -> delegated sign-in to that tenant
    #   5. nothing + console                    -> pick a saved profile / create one / delegated (prompts in Selection.ps1)
    #   6. nothing + no console                 -> delegated (the sign-in itself decides the tenant)
    # Returns $null when the user cancels.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$TenantId, [string]$ClientId, [string]$CertificateThumbprint, [string]$ProfileName,
        $ResumeSelection, [switch]$CanPrompt, [string]$SelectionMode = 'Auto')
    $conn = $null
    if ($ProfileName) {
        $conn = ConvertFrom-TarProfile (Get-TarProfile -Name $ProfileName)
    } elseif ($ClientId -or $CertificateThumbprint) {
        if (-not ($ClientId -and $CertificateThumbprint)) { throw 'App-only sign-in needs both -ClientId and -CertificateThumbprint.' }
        $tenant = if ($TenantId) { (Resolve-TarTenantOrPassThrough -Tenant $TenantId).TenantId }
        elseif ($ResumeSelection) { [string]$ResumeSelection.TenantId }
        elseif ($CanPrompt) { $t = Read-TarTenantInput; if (-not $t) { return }; $t.TenantId }
        else { throw 'App-only sign-in needs a tenant: pass -TenantId (name, domain or ID) or -ProfileName.' }
        $conn = [pscustomobject]@{ AuthMode = 'AppOnly'; TenantId = $tenant; ClientId = (ConvertTo-TarClientId $ClientId)
            CertificateThumbprint = (ConvertTo-TarThumbprint $CertificateThumbprint); PnPClientId = $null; ProfileName = $null
        }
    } elseif ($ResumeSelection -and -not $TenantId) {
        $named = [string](Get-TarProperty $ResumeSelection 'ProfileName')
        $match = @(Read-TarProfileStore | Where-Object { $_.TenantId -eq [string]$ResumeSelection.TenantId })
        $conn = if ($named -and @($match | Where-Object Name -eq $named).Count) { ConvertFrom-TarProfile (@($match | Where-Object Name -eq $named)[0]) }
        elseif ($match.Count -eq 1) { ConvertFrom-TarProfile $match[0] }
        else { [pscustomobject]@{ AuthMode = 'Delegated'; TenantId = [string]$ResumeSelection.TenantId; ClientId = $null; CertificateThumbprint = $null; PnPClientId = $null; ProfileName = $null } }
    } elseif ($TenantId) {
        $conn = [pscustomobject]@{ AuthMode = 'Delegated'; TenantId = (Resolve-TarTenantOrPassThrough -Tenant $TenantId).TenantId; ClientId = $null; CertificateThumbprint = $null; PnPClientId = $null; ProfileName = $null }
    } elseif ($CanPrompt) {
        $conn = Select-TarConnection -SelectionMode $SelectionMode
        if (-not $conn) { return }
    } else {
        $conn = [pscustomobject]@{ AuthMode = 'Delegated'; TenantId = $null; ClientId = $null; CertificateThumbprint = $null; PnPClientId = $null; ProfileName = $null }
    }
    if ($conn.AuthMode -eq 'AppOnly') { Assert-TarCertificate -Thumbprint $conn.CertificateThumbprint }
    $conn
}
