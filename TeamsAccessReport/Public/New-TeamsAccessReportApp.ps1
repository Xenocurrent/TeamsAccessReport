function New-TeamsAccessReportApp {
    <#
    .SYNOPSIS
        Sets up the report app in one or many tenants: Entra app, certificate, admin consent and connection profile.
    .DESCRIPTION
        Per tenant: you sign in once as an admin (browser), then the command
          1. creates a self-signed certificate in CurrentUser\My with a NON-EXPORTABLE private key,
          2. creates the app registration (single tenant) with exactly the application permissions the report needs
             (Microsoft Graph: Team.ReadBasic.All, Channel.ReadBasic.All, TeamMember.Read.All, ChannelMember.Read.All,
             Files.Read.All, GroupMember.Read.All, User.Read.All, Organization.Read.All; SharePoint: Sites.FullControl.All),
          3. creates the service principal and grants admin consent,
          4. saves a connection profile, so the next step is: Invoke-TeamsAccessReport -ProfileName <name>.
        One app and one certificate per tenant: a lost certificate affects one tenant only.
        Required role of the signing-in admin: Privileged Role Administrator (to consent to Microsoft Graph application
        permissions) plus Application Administrator or Cloud Application Administrator - or Global Administrator.
        Nothing is deleted except a certificate this command created when the app itself could not be created.
        If consent fails the app is kept and the result contains a ConsentUrl to finish it in a browser.
    .PARAMETER Tenant
        Tenant name, domain or ID. Accepts several values and pipeline input (e.g. Import-Csv with a 'Tenant' column).
        Omit it to take the tenant from the admin sign-in: you sign in, the command shows the tenant it found
        (name, domain, ID, account) and creates nothing until you confirm. Use -Confirm:$false to skip that question.
    .PARAMETER ProfileName
        Profile name (single tenant, or a 'ProfileName' column in piped CSV). Default: first label of the tenant domain
        (without -Tenant: first label of the tenant's initial onmicrosoft.com domain, lower case).
    .PARAMETER ApplicationName
        Display name of the app registration. Default: TeamsAccessReport.
    .PARAMETER CertificateYears
        Certificate lifetime in years (1-3). Default: 1.
    .PARAMETER CertificateThumbprint
        Use an existing certificate instead of creating one.
    .PARAMETER NoProfile
        Do not save a connection profile.
    .PARAMETER Force
        Replace an existing profile with the same name.
    .EXAMPLE
        New-TeamsAccessReportApp
        Sign in as an admin of the customer tenant; the tenant and profile name come from that sign-in.
    .EXAMPLE
        New-TeamsAccessReportApp -Tenant contoso
    .EXAMPLE
        Import-Csv .\tenants.csv | New-TeamsAccessReportApp | Export-Csv .\setup-results.csv -NoTypeInformation
        tenants.csv: Tenant[,ProfileName]. One browser sign-in per tenant; one result row per tenant.
    .OUTPUTS
        PSCustomObject per tenant: Tenant, Status (Ready | ConsentPending | Failed), ProfileName, ClientId,
        CertificateThumbprint, CertificateExpires, ConsentUrl, Error. Status Cancelled when you answer No.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)][Alias('TenantId', 'Domain')][string[]]$Tenant,
        [Parameter(ValueFromPipelineByPropertyName)][string]$ProfileName,
        [ValidateNotNullOrEmpty()][string]$ApplicationName = 'TeamsAccessReport',
        [ValidateRange(1, 3)][int]$CertificateYears = 1,
        [ValidateNotNullOrEmpty()][string]$CertificateThumbprint,
        [switch]$NoProfile,
        [switch]$Force
    )
    begin {
        if (-not $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = 'Continue' }
        Clear-TarLogTarget
        $tally = [ordered]@{ Ready = 0; ConsentPending = 0; Failed = 0; Cancelled = 0 }
        $askTenant = -not ($PSBoundParameters.ContainsKey('Confirm') -and -not $PSBoundParameters['Confirm'])
    }
    process {
        $targets = @($Tenant | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($ProfileName -and $targets.Count -gt 1) { throw '-ProfileName applies to one tenant; for many tenants use a ProfileName column in the piped input.' }
        if ($targets.Count -eq 0) { $targets = @($null) }   # no tenant given: the admin sign-in decides
        foreach ($t in $targets) {
            $row = [ordered]@{ Tenant = $t; Status = 'Failed'; ProfileName = $null; ClientId = $null; CertificateThumbprint = $null; CertificateExpires = $null; ConsentUrl = $null; Error = $null }
            $connected = $false
            try {
                if ($t) {
                    $resolved = Resolve-TarTenant -Tenant $t
                    $tenantId = $resolved.TenantId; $query = $resolved.Query
                    $name = if ($ProfileName) { $ProfileName } else { ($query -split '\.')[0] }
                    Assert-TarAppProfileName -Name $name -TenantId $tenantId -NoProfile:$NoProfile -Force:$Force
                    if (-not $PSCmdlet.ShouldProcess("$query ($tenantId)", "Create app '$ApplicationName', certificate and admin consent")) { $row.Status = 'Skipped'; [pscustomobject]$row; continue }
                    Write-TarLog "[$query] Sign in as an admin of this tenant in the browser window ..."
                    Connect-MgGraph -TenantId $tenantId -Scopes $script:TarProvisioningScopes -ContextScope Process -NoWelcome -ErrorAction Stop
                    $connected = $true
                    $mg = Get-MgContext
                    if (-not $mg -or $mg.TenantId -ne $tenantId) { throw "Signed in to tenant $($mg.TenantId), expected $tenantId. Sign in with an account of the target tenant." }
                } else {
                    if (-not (Test-TarCanPrompt)) { throw 'No -Tenant and no interactive session. Pass -Tenant (name, domain or ID) or pipe a CSV with a Tenant column.' }
                    if (-not $PSCmdlet.ShouldProcess('the tenant of the admin who signs in', "Create app '$ApplicationName', certificate and admin consent")) { $row.Status = 'Skipped'; [pscustomobject]$row; continue }
                    Write-TarLog 'Sign in as an admin of the tenant to set up, in the browser window. The tenant is taken from that sign-in ...'
                    Connect-MgGraph -Scopes $script:TarProvisioningScopes -ContextScope Process -NoWelcome -ErrorAction Stop
                    $connected = $true
                    $mg = Get-MgContext
                    if (-not $mg -or -not $mg.TenantId) { throw 'The sign-in returned no tenant. Run the command again.' }
                    $org = Get-TarSignedInOrganization
                    $tenantId = ([string]$mg.TenantId).ToLowerInvariant()
                    $query = if ($org.InitialDomain) { $org.InitialDomain.ToLowerInvariant() } else { $tenantId }
                    $row.Tenant = $query
                    $name = if ($ProfileName) { $ProfileName } else { ($query -split '\.')[0] }
                    Assert-TarAppProfileName -Name $name -TenantId $tenantId -NoProfile:$NoProfile -Force:$Force
                    if ($askTenant) {
                        $nl = [Environment]::NewLine
                        $profilePart = if ($NoProfile) { '?' } else { " and save profile '$name'?" }
                        $text = "Signed in as: $($mg.Account)$nl" + "Tenant: $($org.DisplayName)$nl" + "Domain: $query$nl" + "Tenant ID: $tenantId$nl$nl" + "Create app '$ApplicationName' and a certificate in THIS tenant$profilePart"
                        Write-TarLog "Signed in to '$($org.DisplayName)' ($query, $tenantId) as $($mg.Account); waiting for confirmation."
                        if (-not (Read-TarYesNo -Prompt $text -Default $false -Mode Auto)) {
                            $row.Status = 'Cancelled'
                            $row.Error = 'Cancelled: wrong tenant? Run the command again and choose another account in the sign-in window.'
                            Write-TarLog "[$query] $($row.Error)" -Level WARN
                            $tally.Cancelled++
                            [pscustomobject]$row
                            continue
                        }
                    }
                }
                $r = New-TarProvisionedApp -TenantId $tenantId -ApplicationName $ApplicationName -CertificateThumbprint $CertificateThumbprint -CertificateYears $CertificateYears -Confirm:$false
                $row.ClientId = $r.ClientId; $row.CertificateThumbprint = $r.CertificateThumbprint; $row.CertificateExpires = $r.CertificateExpires
                $row.ConsentUrl = $r.ConsentUrl
                $row.Status = if ($r.ConsentComplete) { 'Ready' } else { 'ConsentPending' }
                if (-not $r.ConsentComplete) { $row.Error = [string](Get-TarProperty $r 'ConsentError') }
                if (-not $NoProfile) {
                    Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name $name -TenantId $tenantId -TenantQuery $query -AuthMode AppOnly `
                            -ClientId $r.ClientId -CertificateThumbprint $r.CertificateThumbprint -Options @{ IncludeSharePoint = $true; IncludeUserAssociations = $true }) -Confirm:$false
                    $row.ProfileName = $name
                }
                $next = if ($row.Status -eq 'Ready') { "Ready. New permissions can take a few minutes to apply; then run Invoke-TeamsAccessReport and pick profile '$name' in the start window." } else { "Consent incomplete - open as admin: $($r.ConsentUrl)" }
                Write-TarLog "[$query] $next" -Level $(if ($row.Status -eq 'Ready') { 'INFO' } else { 'WARN' })
            } catch {
                $row.Error = $_.Exception.Message
                Write-TarLog "[$($row.Tenant)] Failed: $($row.Error)" -Level ERROR
            } finally {
                if ($connected) { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue -WarningAction SilentlyContinue }
            }
            if ($row.Status -ne 'Cancelled' -and $tally.Contains($row.Status)) { $tally[$row.Status]++ }
            [pscustomobject]$row
        }
    }
    end {
        Write-TarLog "Setup finished: $($tally.Ready) ready, $($tally.ConsentPending) consent pending, $($tally.Failed) failed, $($tally.Cancelled) cancelled."
    }
}
