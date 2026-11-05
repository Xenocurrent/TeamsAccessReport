function Invoke-TeamsAccessReport {
    <#
    .SYNOPSIS
        Complete access report for Microsoft Teams and the associated SharePoint sites.
    .DESCRIPTION
        Phase 1  Teams via Graph (teams, channels, direct/indirect membership, channel sites)  - checkpoint per team
        Phase 2  SharePoint via PnP in a separate process (admins, groups, unique permissions) - checkpoint per site
        Phase 3  Resolution of SharePoint claims to Entra users
        Phase 4  Drift between Teams and SharePoint, external associations, summary
        Returns the summary object; all details are written as CSV/JSON to -OutputPath.
    .PARAMETER TenantId
        Optional for delegated sign-in (resolved at sign-in). Required for app-only.
    .PARAMETER ClientId
        App registration (client) ID for app-only authentication.
    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate used for app-only authentication (CurrentUser or LocalMachine store).
    .PARAMETER TeamId
        Limit the report to these team (group) IDs. Skips interactive team selection.
    .PARAMETER ChannelId
        Limit the report to these channel IDs. Requires -TeamId.
    .PARAMETER AllTeams
        Include every team in the tenant without interactive selection.
    .PARAMETER Resume
        Resume an interrupted run in -OutputPath. Incomplete teams and failed sites are re-run.
    .PARAMETER SelectChannels
        Prompt for channel selection per team.
    .PARAMETER SelectionMode
        Auto (default), Gui (Out-GridView), ConsoleGrid (Out-ConsoleGridView) or Menu (numbered console menu).
    .PARAMETER Force
        Skip confirmation prompts (tenant confirmation and channel-selection question).
    .PARAMETER OutputPath
        Report directory. Defaults to ./TeamsAccess_<timestamp>.
    .PARAMETER IncludeSharePoint
        Run phases 2-4: SharePoint permissions, claim resolution and drift analysis.
    .PARAMETER PnPClientId
        Entra app ID used by PnP.PowerShell. Defaults to -ClientId in app-only mode.
    .PARAMETER MaxItemsPerList
        Maximum number of list items to scan per list (0 = unlimited).
    .PARAMETER MaxSubwebDepth
        Maximum subweb recursion depth.
    .PARAMETER IncludeUserAssociations
        Find internal users' memberships in shared channels hosted by other tenants (app-only only).
    .PARAMETER MaxRetries
        Maximum retries in the Graph SDK retry handler and the SharePoint retry loop.
    .PARAMETER BaseDelaySeconds
        Base delay for exponential backoff.
    .EXAMPLE
        Invoke-TeamsAccessReport
        Interactive: sign in, confirm the tenant, select teams and channels.
    .EXAMPLE
        Invoke-TeamsAccessReport -TenantId '<TENANT_ID>' -ClientId '<APP_ID>' -CertificateThumbprint '<THUMBPRINT>' -AllTeams -Force -IncludeSharePoint -IncludeUserAssociations
        Unattended full report.
    .OUTPUTS
        PSCustomObject (summary, same content as Summary.json)
    #>
    [CmdletBinding(DefaultParameterSetName = 'Delegated', SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(ParameterSetName = 'Delegated')]
        [Parameter(ParameterSetName = 'AppOnly', Mandatory)]
        [ValidateNotNullOrEmpty()][string]$TenantId,

        [Parameter(ParameterSetName = 'AppOnly', Mandatory)][ValidateNotNullOrEmpty()][string]$ClientId,
        [Parameter(ParameterSetName = 'AppOnly', Mandatory)][ValidateNotNullOrEmpty()][string]$CertificateThumbprint,

        [string[]]$TeamId,
        [string[]]$ChannelId,
        [switch]$AllTeams,
        [switch]$Resume,
        [switch]$SelectChannels,
        [ValidateSet('Auto', 'Gui', 'ConsoleGrid', 'Menu')][string]$SelectionMode = 'Auto',
        [switch]$Force,
        [ValidateNotNullOrEmpty()][string]$OutputPath = (Join-Path (Get-Location) ('TeamsAccess_{0:yyyyMMdd_HHmmss}' -f (Get-Date))),

        [switch]$IncludeSharePoint,
        [string]$PnPClientId,
        [ValidateRange(0, [int]::MaxValue)][int]$MaxItemsPerList = 0,
        [ValidateRange(0, 10)][int]$MaxSubwebDepth = 5,
        [switch]$IncludeUserAssociations,

        [ValidateRange(1, 10)][int]$MaxRetries = 6,
        [ValidateRange(1, 60)][int]$BaseDelaySeconds = 3
    )
    if (-not $PSBoundParameters.ContainsKey('InformationAction')) { $InformationPreference = 'Continue' }
    $appOnly = $PSCmdlet.ParameterSetName -eq 'AppOnly'

    # --- Validate combinations that parameter sets cannot express ---
    if ($ChannelId -and -not $TeamId) { throw '-ChannelId requires -TeamId.' }
    if ($TeamId -and $AllTeams) { throw 'Specify either -TeamId or -AllTeams, not both.' }
    if ($Resume -and ($TeamId -or $AllTeams -or $ChannelId)) { throw '-Resume reads the selection from Selection.json; do not combine it with -TeamId/-AllTeams/-ChannelId.' }
    $spClientId = if ($PnPClientId) { $PnPClientId } elseif ($appOnly) { $ClientId } else { $null }
    if ($IncludeSharePoint -and -not $spClientId) {
        throw "-IncludeSharePoint requires a dedicated Entra app for PnP (-PnPClientId). Create one: Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP-TeamsAccess' -Tenant <tenant>.onmicrosoft.com -SharePointDelegatePermissions AllSites.FullControl"
    }
    if ($IncludeUserAssociations -and -not $appOnly) { Write-Warning '-IncludeUserAssociations requires app-only (delegated only supports "me"). The step is skipped.' }
    $interactive = -not $Resume -and -not $TeamId -and -not $AllTeams
    if ($interactive -and ([Console]::IsInputRedirected -or -not [Environment]::UserInteractive)) { throw 'No interactive console. Use -TeamId, -AllTeams or -Resume.' }
    if (Get-Module -Name PnP.PowerShell) { throw 'PnP.PowerShell is loaded in this session - start a new pwsh (assembly conflict with the Graph SDK).' }
    if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) { throw 'Microsoft.Graph.Authentication is missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }

    # --- Context, store, logging ---
    $ctx = New-TarContext -OutputPath $OutputPath -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds -MaxItemsPerList $MaxItemsPerList -MaxSubwebDepth $MaxSubwebDepth
    $ctx.AppOnly = $appOnly
    if ($appOnly) { $ctx.ClientId = $ClientId; $ctx.CertificateThumbprint = $CertificateThumbprint }
    if (-not $PSCmdlet.ShouldProcess($ctx.OutputPath, 'Create access report')) { return }
    Initialize-TarStore -Context $ctx -Resume:$Resume
    Initialize-TarLog -Path (Get-TarStorePath $ctx $script:TarFile.Log)

    try {
        Connect-TarGraph -Context $ctx -TenantId $TenantId
        if ($interactive -and -not $Force -and -not (Read-TarYesNo "Continue with tenant '$($ctx.TenantName ?? $ctx.TenantId)' ($($ctx.TenantDomain ?? $ctx.TenantId))?")) {
            Write-TarLog 'Cancelled at tenant confirmation.' -Level WARN
            return
        }

        if ($Resume) {
            $selection = Import-TarSelection -Context $ctx
            Write-TarLog "Resuming: $($selection.Teams.Count) teams from Selection.json"
            Reset-TarIncompleteWork -Context $ctx
        } else {
            $selection = Get-TarTeamSelection -Context $ctx -TeamId $TeamId -ChannelId $ChannelId -AllTeams:$AllTeams -Interactive:$interactive `
                -SelectChannels:$SelectChannels -Force:$Force -SelectionMode $SelectionMode
            if (-not $selection) { return }
            Save-TarSelection -Context $ctx -Selection $selection
        }
        Write-TarLog "Teams to process: $($selection.Teams.Count)"

        Invoke-TarTeamInventory -Context $ctx -Teams $selection.Teams -ChannelFilter $selection.ChannelFilter
        if ($IncludeSharePoint) {
            Start-TarSharePointScan -Context $ctx -ClientId $spClientId
            Invoke-TarPrincipalResolution -Context $ctx
            Invoke-TarDriftAnalysis -Context $ctx
        }
        if ($IncludeUserAssociations -and $appOnly) { Invoke-TarExternalAssociationScan -Context $ctx }

        $summary = Export-TarSummary -Context $ctx -SharePointIncluded:$IncludeSharePoint
        Write-TarLog ("Done: Status={0} Teams={1} Rows={2} Drift={3} (High={4}) Errors={5} -> {6}" -f $summary.ResultStatus, $summary.Teams, $summary.AccessRows, $summary.DriftFindings, $summary.DriftHigh, $summary.Errors, $ctx.OutputPath)
        if ($summary.Errors) { Write-TarLog "See Errors.csv - $($summary.Errors) steps failed (HTTP 403 = missing permission/consent; other errors include the reason)." -Level WARN }
        $summary
    } finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
