# Provisioning of the report app in a tenant (New-TeamsAccessReportApp). This is the ONLY code that writes to a
# tenant; the report itself stays read-only. Graph writes go through Invoke-TarGraphWrite (test seam).

function Invoke-TarGraphWriteRaw {
    # Seam: the only Graph write call. Returns @{ StatusCode; Body }.
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('POST', 'PATCH')][string]$Method, [Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$Json)
    $status = 0
    $body = Invoke-MgGraphRequest -Method $Method -Uri $Uri -Body $Json -ContentType 'application/json' -OutputType Hashtable -SkipHttpErrorCheck -StatusCodeVariable 'status' -ErrorAction Stop
    [pscustomobject]@{ StatusCode = [int]$status; Body = $body }
}

function Invoke-TarGraphWrite {
    # POST/PATCH with retry on throttling and on directory replication delay (a new object is not yet visible).
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('POST', 'PATCH')][string]$Method, [Parameter(Mandatory)][string]$Uri,
        [Parameter(Mandatory)][hashtable]$Body, [ValidateRange(1, 10)][int]$MaxRetries = 6, [ValidateRange(1, 30)][int]$BaseDelaySeconds = 2)
    $json = $Body | ConvertTo-Json -Depth 10 -Compress
    for ($attempt = 1; ; $attempt++) {
        $raw = Invoke-TarGraphWriteRaw -Method $Method -Uri $Uri -Json $json
        $status = $raw.StatusCode; $r = $raw.Body
        if ($status -ge 200 -and $status -lt 300) { return $r }
        $message = [string](Get-TarProperty (Get-TarProperty $r 'error') 'message')
        # 400/404 right after creating an object is usually replication ("does not exist"); 429/5xx are transient.
        $transient = $status -in 429, 500, 502, 503, 504 -or ($status -in 400, 404 -and $message -match 'does not exist|not found|Invalid object identifier|ResourceNotFound')
        if (-not $transient -or $attempt -ge $MaxRetries) {
            $e = [System.InvalidOperationException]::new("Graph $Method $Uri failed (HTTP $status): $message")
            $e.Data['StatusCode'] = $status; throw $e
        }
        $wait = Get-TarBackoffDelay -Attempt $attempt -BaseDelaySeconds $BaseDelaySeconds
        Write-TarLog "Graph $Method HTTP $status ($message), attempt $attempt/$MaxRetries; waiting $wait ms" -Level WARN
        Start-Sleep -Milliseconds $wait
    }
}

function Get-TarGraphOne {
    # GET with the same retry as the report; returns the body.
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)
    Invoke-TarGraphRequest -Context ([pscustomobject]@{ BaseDelaySeconds = 2 }) -Uri $Uri
}

function New-TarAppCertificate {
    # Self-signed RSA 2048 / SHA-256 in CurrentUser\My with a NON-EXPORTABLE private key: the key cannot be copied
    # off this machine as a PFX. Pure .NET (no PKI module). Returns the stored certificate.
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param([Parameter(Mandatory)][string]$Subject, [ValidateRange(1, 3)][int]$Years = 1)
    if (-not $PSCmdlet.ShouldProcess("CurrentUser\My: $Subject", 'Create certificate')) { return }
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $temp = $null; $stored = $null
    try {
        $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new($Subject, $rsa,
            [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $req.CertificateExtensions.Add([System.Security.Cryptography.X509Certificates.X509KeyUsageExtension]::new(
                [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]::DigitalSignature, $true))
        $temp = $req.CreateSelfSigned([DateTimeOffset]::UtcNow.AddMinutes(-5), [DateTimeOffset]::UtcNow.AddYears($Years))
        # Round-trip through an in-memory PFX so the key is persisted WITHOUT the Exportable flag.
        $password = [guid]::NewGuid().ToString('N')
        $pfx = $temp.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $password)
        $flags = [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]'PersistKeySet, UserKeySet'
        $stored = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($pfx, $password, $flags)
        [Array]::Clear($pfx, 0, $pfx.Length)
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
        $store.Open('ReadWrite')
        try { $store.Add($stored) } finally { $store.Close() }
        Write-TarLog "Created certificate $Subject ($($stored.Thumbprint)), expires $($stored.NotAfter.ToString('yyyy-MM-dd')), CurrentUser\My, key not exportable"
        $stored
    } finally {
        if ($temp) { $temp.Dispose() }
        $rsa.Dispose()
    }
}

function Remove-TarStoreCertificate {
    # Cleanup of a certificate this command created when the app could not be created.
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Thumbprint)
    if (-not $PSCmdlet.ShouldProcess("CurrentUser\My\$Thumbprint", 'Remove certificate')) { return }
    $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', 'CurrentUser')
    $store.Open('ReadWrite')
    try { foreach ($c in @($store.Certificates.Find('FindByThumbprint', $Thumbprint, $false))) { $store.Remove($c) } } finally { $store.Close() }
}

function Get-TarAppRoleId {
    # Permission names -> app role IDs, read from the resource's service principal in this tenant (no hard-coded IDs).
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$GraphBase, [Parameter(Mandatory)][string]$ResourceAppId, [Parameter(Mandatory)][string[]]$Roles)
    $sp = @(Get-TarProperty (Get-TarGraphOne -Uri "$GraphBase/servicePrincipals?`$filter=appId eq '$ResourceAppId'&`$select=id,displayName,appRoles") 'value')[0]
    if (-not $sp) { throw "Service principal for resource $ResourceAppId was not found in the tenant." }
    $byName = @{}
    foreach ($r in @(Get-TarProperty $sp 'appRoles')) { if ((Get-TarProperty $r 'allowedMemberTypes') -contains 'Application') { $byName[[string](Get-TarProperty $r 'value')] = [string](Get-TarProperty $r 'id') } }
    $missing = @($Roles | Where-Object { -not $byName.ContainsKey($_) })
    if ($missing) { throw "$(Get-TarProperty $sp 'displayName') has no application permission named: $($missing -join ', ')." }
    [pscustomobject]@{ ResourceSpId = [string](Get-TarProperty $sp 'id'); ResourceAppId = $ResourceAppId; Name = [string](Get-TarProperty $sp 'displayName'); Roles = @($Roles | ForEach-Object { [pscustomobject]@{ Name = $_; Id = $byName[$_] } }) }
}

function New-TarProvisionedApp {
    # Signed in (delegated) to the target tenant already. Creates app + certificate + service principal and grants
    # admin consent for exactly $script:TarAppPermissions. Never deletes anything it did not create in this call.
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$ApplicationName, [string]$CertificateThumbprint,
        [ValidateRange(1, 3)][int]$CertificateYears = 1, [string]$GraphBase = 'https://graph.microsoft.com/v1.0')
    $escaped = $ApplicationName.Replace("'", "''")
    $existing = @(Get-TarProperty (Get-TarGraphOne -Uri "$GraphBase/applications?`$filter=displayName eq '$escaped'&`$select=id,appId") 'value')
    if ($existing.Count) {
        throw "An app named '$ApplicationName' already exists in this tenant (client ID $(Get-TarProperty $existing[0] 'appId')). Use it with New-TeamsAccessReportProfile, or pass -ApplicationName for a new one."
    }
    $resources = @(foreach ($k in $script:TarAppPermissions.Keys) { Get-TarAppRoleId -GraphBase $GraphBase -ResourceAppId $k -Roles $script:TarAppPermissions[$k].Roles })
    if (-not $PSCmdlet.ShouldProcess($TenantId, "Create app '$ApplicationName' with $(@($resources.Roles).Count) application permissions and grant admin consent")) { return }

    $createdCert = $false
    if ($CertificateThumbprint) {
        Assert-TarCertificate -Thumbprint $CertificateThumbprint
        $cert = @(Get-TarCertificateCandidate -Thumbprint $CertificateThumbprint)[0]
        $raw = $null
        $store = [System.Security.Cryptography.X509Certificates.X509Store]::new('My', $cert.Store.Split('\')[0]); $store.Open('ReadOnly')
        try { $raw = @($store.Certificates.Find('FindByThumbprint', $CertificateThumbprint, $false))[0].RawData } finally { $store.Close() }
        $thumb = $CertificateThumbprint; $expires = [datetime]$cert.NotAfter
    } else {
        $c = New-TarAppCertificate -Subject "CN=$ApplicationName-$((Get-Date).ToString('yyyy-MM'))" -Years $CertificateYears -Confirm:$false
        $raw = $c.RawData; $thumb = $c.Thumbprint; $expires = $c.NotAfter; $createdCert = $true
    }

    $app = $null
    try {
        $body = @{
            displayName = $ApplicationName; signInAudience = 'AzureADMyOrg'
            notes = 'Read-only Teams/SharePoint access report (TeamsAccessReport). Created by New-TeamsAccessReportApp.'
            requiredResourceAccess = @($resources | ForEach-Object { @{ resourceAppId = $_.ResourceAppId; resourceAccess = @($_.Roles | ForEach-Object { @{ id = $_.Id; type = 'Role' } }) } })
            keyCredentials = @(@{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [Convert]::ToBase64String($raw); displayName = "CN=$ApplicationName" })
        }
        $app = Invoke-TarGraphWrite -Method POST -Uri "$GraphBase/applications" -Body $body
        Write-TarLog "Created app '$ApplicationName' (client ID $($app['appId']))"
    } catch {
        if ($createdCert) { Remove-TarStoreCertificate -Thumbprint $thumb -Confirm:$false }
        throw
    }

    $result = [pscustomobject]@{
        TenantId = $TenantId; ApplicationName = $ApplicationName; ClientId = [string]$app['appId']; ApplicationObjectId = [string]$app['id']
        ServicePrincipalId = $null; CertificateThumbprint = $thumb; CertificateExpires = $expires.ToString('yyyy-MM-dd')
        Granted = [System.Collections.Generic.List[string]]::new(); ConsentComplete = $false
        ConsentUrl = "$script:TarAuthority/$TenantId/adminconsent?client_id=$($app['appId'])"
    }
    try {
        $sp = Invoke-TarGraphWrite -Method POST -Uri "$GraphBase/servicePrincipals" -Body @{ appId = $result.ClientId }
        $result.ServicePrincipalId = [string]$sp['id']
        foreach ($res in $resources) {
            foreach ($role in $res.Roles) {
                $null = Invoke-TarGraphWrite -Method POST -Uri "$GraphBase/servicePrincipals/$($res.ResourceSpId)/appRoleAssignedTo" `
                    -Body @{ principalId = $result.ServicePrincipalId; resourceId = $res.ResourceSpId; appRoleId = $role.Id }
                $result.Granted.Add("$($res.Name): $($role.Name)")
            }
        }
        $result.ConsentComplete = $true
        Write-TarLog "Admin consent granted for $($result.Granted.Count) permissions."
    } catch {
        # The app exists; do not delete it. Report what is missing and how to finish.
        Write-TarLog "App created but consent is incomplete ($($_.Exception.Message)). Finish with: $($result.ConsentUrl)" -Level WARN
        $result | Add-Member -NotePropertyName ConsentError -NotePropertyValue $_.Exception.Message
    }
    $result
}
