function New-TeamsAccessReportPnPApp {
    <#
    .SYNOPSIS
        Creates the Entra app PnP PowerShell needs for delegated SharePoint scanning and stores it in a profile.
    .DESCRIPTION
        Runs Register-PnPEntraIDAppForInteractiveLogin in a separate PowerShell window (PnP and the Graph SDK cannot
        share a session). One delegated permission: SharePoint AllSites.FullControl; redirect URI http://localhost.
        A browser opens for sign-in and then for admin consent. Not needed for app-only sign-in.
        Requires PnP.PowerShell and an account allowed to register apps (for example Application Administrator).
    .PARAMETER ProfileName
        Profile to update with the new PnP app ID (its tenant is used).
    .PARAMETER Tenant
        Tenant name, domain or ID, when no profile is used.
    .PARAMETER ApplicationName
        Display name of the app registration. Default: TeamsAccessReport-PnP.
    .EXAMPLE
        New-TeamsAccessReportPnPApp -ProfileName contoso
    .OUTPUTS
        System.String (the client ID)
    #>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Profile')]
    [OutputType([string])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Profile', Position = 0)][ValidateNotNullOrEmpty()][string]$ProfileName,
        [Parameter(Mandatory, ParameterSetName = 'Tenant')][ValidateNotNullOrEmpty()][string]$Tenant,
        [ValidateNotNullOrEmpty()][string]$ApplicationName = 'TeamsAccessReport-PnP'
    )
    if (-not $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = 'Continue' }
    Clear-TarLogTarget
    $profileRecord = if ($ProfileName) { Get-TarProfile -Name $ProfileName } else { $null }
    $target = if ($profileRecord) { [string]$(if ($profileRecord.TenantQuery) { $profileRecord.TenantQuery } else { $profileRecord.TenantId }) } else { (Resolve-TarTenantOrPassThrough -Tenant $Tenant).Query }
    if (-not $PSCmdlet.ShouldProcess($target, "Register Entra app '$ApplicationName'")) { return }
    $id = New-TarPnPApp -Tenant $target -ApplicationName $ApplicationName -Confirm:$false
    if ($profileRecord) {
        $opts = Get-TarProperty $profileRecord 'Options'
        $updated = New-TarProfileRecord -Name $profileRecord.Name -TenantId $profileRecord.TenantId -TenantQuery $profileRecord.TenantQuery -AuthMode $profileRecord.AuthMode `
            -ClientId $profileRecord.ClientId -CertificateThumbprint $profileRecord.CertificateThumbprint -PnPClientId $id `
            -Options $(if ($opts) { @{ IncludeSharePoint = [bool]$opts.IncludeSharePoint; IncludeUserAssociations = [bool]$opts.IncludeUserAssociations; SelectChannels = [bool]$opts.SelectChannels } } else { $null })
        Save-TarProfile -ProfileRecord $updated -Confirm:$false
    }
    $id
}
