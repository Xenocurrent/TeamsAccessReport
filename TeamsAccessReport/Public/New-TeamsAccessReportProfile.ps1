function New-TeamsAccessReportProfile {
    <#
    .SYNOPSIS
        Saves a connection profile (tenant, app ID, certificate thumbprint) so reports run without typing IDs.
    .DESCRIPTION
        Asks only for what is missing: tenant (name, domain, e-mail or ID - resolved to the tenant ID), sign-in method,
        application (client) ID and certificate (picked from the certificate store). Stores identifiers only, never
        secrets; the certificate and its private key stay in the Windows/.NET certificate store.
        Store: %LOCALAPPDATA%\TeamsAccessReport\profiles.json (Windows) or ~/.local/share/TeamsAccessReport/profiles.json.
        Override with the environment variable TEAMSACCESSREPORT_PROFILE_PATH.
    .PARAMETER Name
        Profile name (letters, digits, '.', '_', '-'). Defaults to the first label of the tenant domain.
    .PARAMETER Tenant
        Tenant name ('contoso'), domain ('contoso.com'), e-mail address or tenant ID.
    .PARAMETER ClientId
        Application (client) ID for app-only sign-in.
    .PARAMETER CertificateThumbprint
        Certificate thumbprint for app-only sign-in. Omit to pick from the certificate store.
    .PARAMETER PnPClientId
        Optional separate app ID for PnP.PowerShell (phase 2). Defaults to ClientId in app-only mode.
    .PARAMETER Delegated
        Create a delegated profile (sign in as yourself; only the tenant is stored).
    .PARAMETER SkipCertificateCheck
        Do not require the certificate to be present on this machine (profile prepared for another machine).
    .PARAMETER Force
        Replace an existing profile with the same name.
    .PARAMETER SelectionMode
        Auto, Gui (pop-up), ConsoleGrid or Menu - used for the certificate and sign-in method pickers.
    .EXAMPLE
        New-TeamsAccessReportProfile
        Interactive: one window on Windows (tenant, sign-in method, app ID, certificate picker), console prompts elsewhere.
    .EXAMPLE
        New-TeamsAccessReportProfile -Name contoso -Tenant contoso -ClientId '<APP_ID>' -CertificateThumbprint '<THUMBPRINT>'
        Unattended.
    .OUTPUTS
        PSCustomObject (the saved profile)
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Position = 0)][ValidateNotNullOrEmpty()][string]$Name,
        [ValidateNotNullOrEmpty()][string]$Tenant,
        [ValidateNotNullOrEmpty()][string]$ClientId,
        [ValidateNotNullOrEmpty()][string]$CertificateThumbprint,
        [ValidateNotNullOrEmpty()][string]$PnPClientId,
        [switch]$Delegated,
        [switch]$SkipCertificateCheck,
        [switch]$Force,
        [ValidateSet('Auto', 'Gui', 'ConsoleGrid', 'Menu')][string]$SelectionMode = 'Auto'
    )
    if (-not $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = 'Continue' }
    Clear-TarLogTarget
    if ($Delegated -and ($ClientId -or $CertificateThumbprint)) { throw '-Delegated cannot be combined with -ClientId/-CertificateThumbprint.' }
    $mode = if ($Delegated) { 'Delegated' } elseif ($ClientId -or $CertificateThumbprint) { 'AppOnly' } else { '' }
    $complete = $Tenant -and ($mode -eq 'Delegated' -or ($ClientId -and $CertificateThumbprint))
    $canPrompt = Test-TarCanPrompt

    if ($complete) {
        $t = Resolve-TarTenantOrPassThrough -Tenant $Tenant
        if (-not $Name) { $Name = ($t.Query -split '\.')[0] }
        $record = New-TarProfileRecord -Name $Name -TenantId $t.TenantId -TenantQuery $t.Query -AuthMode $mode -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -PnPClientId $PnPClientId
    } elseif ($canPrompt -and (Resolve-TarSelectionMode -Mode $SelectionMode) -eq 'Gui' -and (Test-TarGuiAvailable)) {
        $start = Show-TarStartForm -ConnectionOnly
        if (-not $start) { Write-TarLog 'Cancelled; no profile saved.' -Level WARN; return }
        $record = $start.ProfileRecord
        $Force = $true   # the window already refuses an existing name unless that profile was selected for editing
    } elseif ($canPrompt) {
        $record = Read-TarConnectionWizard -Name $Name -Tenant $Tenant -AuthMode $mode -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -PnPClientId $PnPClientId -SelectionMode $SelectionMode
        if (-not $record) { Write-TarLog 'Cancelled; no profile saved.' -Level WARN; return }
    } else {
        throw 'No interactive console: pass -Tenant and either -Delegated or -ClientId with -CertificateThumbprint.'
    }

    if ($record.AuthMode -eq 'AppOnly' -and -not $SkipCertificateCheck) { Assert-TarCertificate -Thumbprint $record.CertificateThumbprint }
    if (-not $Force -and @(Read-TarProfileStore | Where-Object Name -eq $record.Name).Count) {
        throw "Profile '$($record.Name)' already exists. Use -Force to replace it or -Name to choose another name."
    }
    if ($PSCmdlet.ShouldProcess($record.Name, 'Save connection profile')) {
        Save-TarProfile -ProfileRecord $record -Confirm:$false
        Write-TarLog "Use it with: Invoke-TeamsAccessReport -ProfileName $($record.Name)"
        $record
    }
}
