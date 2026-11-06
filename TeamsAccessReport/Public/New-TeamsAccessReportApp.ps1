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
    .PARAMETER ProfileName
        Profile name (single tenant, or a 'ProfileName' column in piped CSV). Default: first label of the tenant domain.
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
        New-TeamsAccessReportApp -Tenant contoso
    .EXAMPLE
        Import-Csv .\tenants.csv | New-TeamsAccessReportApp | Export-Csv .\setup-results.csv -NoTypeInformation
        tenants.csv: Tenant[,ProfileName]. One browser sign-in per tenant; one result row per tenant.
    .OUTPUTS
        PSCustomObject per tenant: Tenant, Status (Ready | ConsentPending | Failed), ProfileName, ClientId,
        CertificateThumbprint, CertificateExpires, ConsentUrl, Error
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory, Position = 0, ValueFromPipeline, ValueFromPipelineByPropertyName)][Alias('TenantId', 'Domain')][ValidateNotNullOrEmpty()][string[]]$Tenant,
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
        $tally = @{ Ready = 0; ConsentPending = 0; Failed = 0 }
    }
    process {
        if ($ProfileName -and $Tenant.Count -gt 1) { throw '-ProfileName applies to one tenant; for many tenants use a ProfileName column in the piped input.' }
        foreach ($t in $Tenant) {
            $row = [ordered]@{ Tenant = $t; Status = 'Failed'; ProfileName = $null; ClientId = $null; CertificateThumbprint = $null; CertificateExpires = $null; ConsentUrl = $null; Error = $null }
            try {
                $resolved = Resolve-TarTenant -Tenant $t
                $name = if ($ProfileName) { $ProfileName } else { ($resolved.Query -split '\.')[0] }
                if (-not $NoProfile) {
                    $null = New-TarProfileRecord -Name $name -TenantId $resolved.TenantId -AuthMode Delegated   # validates the name before anything is created
                    if (-not $Force -and @(Read-TarProfileStore | Where-Object Name -eq $name).Count) { throw "Profile '$name' already exists. Use -Force to replace it or -ProfileName for another name." }
                }
                if (-not $PSCmdlet.ShouldProcess("$($resolved.Query) ($($resolved.TenantId))", "Create app '$ApplicationName', certificate and admin consent")) { $row.Status = 'Skipped'; [pscustomobject]$row; continue }
                Write-TarLog "[$($resolved.Query)] Sign in as an admin of this tenant in the browser window ..."
                Connect-MgGraph -TenantId $resolved.TenantId -Scopes $script:TarProvisioningScopes -ContextScope Process -NoWelcome -ErrorAction Stop
                try {
                    $mg = Get-MgContext
                    if (-not $mg -or $mg.TenantId -ne $resolved.TenantId) { throw "Signed in to tenant $($mg.TenantId), expected $($resolved.TenantId). Sign in with an account of the target tenant." }
                    $r = New-TarProvisionedApp -TenantId $resolved.TenantId -ApplicationName $ApplicationName -CertificateThumbprint $CertificateThumbprint -CertificateYears $CertificateYears -Confirm:$false
                } finally { $null = Disconnect-MgGraph -ErrorAction SilentlyContinue }
                $row.ClientId = $r.ClientId; $row.CertificateThumbprint = $r.CertificateThumbprint; $row.CertificateExpires = $r.CertificateExpires
                $row.ConsentUrl = $r.ConsentUrl
                $row.Status = if ($r.ConsentComplete) { 'Ready' } else { 'ConsentPending' }
                if (-not $r.ConsentComplete) { $row.Error = [string](Get-TarProperty $r 'ConsentError') }
                if (-not $NoProfile) {
                    Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name $name -TenantId $resolved.TenantId -TenantQuery $resolved.Query -AuthMode AppOnly `
                            -ClientId $r.ClientId -CertificateThumbprint $r.CertificateThumbprint -Options @{ IncludeSharePoint = $true; IncludeUserAssociations = $true }) -Confirm:$false
                    $row.ProfileName = $name
                }
                $next = if ($row.Status -eq 'Ready') { "Ready. New permissions can take a few minutes to apply; then: Invoke-TeamsAccessReport -ProfileName $name" } else { "Consent incomplete - open as admin: $($r.ConsentUrl)" }
                Write-TarLog "[$($resolved.Query)] $next" -Level $(if ($row.Status -eq 'Ready') { 'INFO' } else { 'WARN' })
            } catch {
                $row.Error = $_.Exception.Message
                Write-TarLog "[$t] Failed: $($row.Error)" -Level ERROR
            }
            if ($tally.ContainsKey($row.Status)) { $tally[$row.Status]++ }
            [pscustomobject]$row
        }
    }
    end {
        Write-TarLog "Setup finished: $($tally.Ready) ready, $($tally.ConsentPending) consent pending, $($tally.Failed) failed."
    }
}
