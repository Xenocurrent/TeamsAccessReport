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
    .PARAMETER ProfileName
        Saved connection profile (tenant, app ID, certificate thumbprint). Create one with New-TeamsAccessReportProfile,
        or let the start screen create it when the command is run without connection parameters. Tab-completes.
    .PARAMETER TenantId
        Tenant name, domain, e-mail address or tenant ID - 'contoso', 'contoso.com', 'admin@contoso.com' and
        'contoso.onmicrosoft.com' all work; it is resolved to the tenant ID before sign-in. Alone: delegated sign-in.
    .PARAMETER ClientId
        App registration (client) ID for app-only authentication. Use together with -CertificateThumbprint.
    .PARAMETER CertificateThumbprint
        Thumbprint of the certificate used for app-only authentication (CurrentUser or LocalMachine store).
        Spaces and invisible characters copied from certmgr are removed.
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
        Interactive. On Windows one window asks for everything: saved connection or tenant (name, domain or ID), sign-in
        method, app ID, certificate (picked from the store), report options and output folder. Tenant confirmation and
        team/channel selection are pop-ups too. Elsewhere (or with -SelectionMode Menu) the same questions are asked in
        the console.
    .EXAMPLE
        Invoke-TeamsAccessReport -ProfileName contoso -IncludeSharePoint
        Saved connection, no IDs to type.
    .EXAMPLE
        Invoke-TeamsAccessReport -TenantId '<TENANT_ID>' -ClientId '<APP_ID>' -CertificateThumbprint '<THUMBPRINT>' -AllTeams -Force -IncludeSharePoint -IncludeUserAssociations
        Unattended full report.
    .OUTPUTS
        PSCustomObject (summary, same content as Summary.json)
    #>
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([pscustomobject])]
    param(
        [ValidateNotNullOrEmpty()][string]$ProfileName,
        [ValidateNotNullOrEmpty()][string]$TenantId,
        [ValidateNotNullOrEmpty()][string]$ClientId,
        [ValidateNotNullOrEmpty()][string]$CertificateThumbprint,

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
    Clear-TarLogTarget

    # --- Validate parameter combinations ---
    if ($ChannelId -and -not $TeamId) { throw '-ChannelId requires -TeamId.' }
    if ($TeamId -and $AllTeams) { throw 'Specify either -TeamId or -AllTeams, not both.' }
    if ($Resume -and ($TeamId -or $AllTeams -or $ChannelId)) { throw '-Resume reads the selection from Selection.json; do not combine it with -TeamId/-AllTeams/-ChannelId.' }
    if ($ProfileName -and ($TenantId -or $ClientId -or $CertificateThumbprint)) { throw 'Use either -ProfileName or -TenantId/-ClientId/-CertificateThumbprint, not both.' }
    $canPrompt = Test-TarCanPrompt

    # --- Connection: profile, explicit parameters, resumed tenant or start screen ---
    $resumeSelection = $null
    $selectionPath = Join-Path ([System.IO.Path]::GetFullPath($OutputPath)) $script:TarFile.Selection
    if ($Resume -and (Test-Path $selectionPath)) { $resumeSelection = Get-Content -Path $selectionPath -Raw -Encoding utf8 | ConvertFrom-Json }
    $noConnectionParameters = -not ($ProfileName -or $TenantId -or $ClientId -or $CertificateThumbprint -or $Resume)
    $useStartForm = $noConnectionParameters -and $canPrompt -and (Resolve-TarSelectionMode -Mode $SelectionMode) -eq 'Gui' -and (Test-TarGuiAvailable)
    if ($useStartForm) {
        # One window for connection + report options; switches given on the command line pre-fill and win.
        $defaults = @{ OutputPath = $OutputPath }
        foreach ($k in 'IncludeSharePoint', 'IncludeUserAssociations', 'SelectChannels', 'PnPClientId') { if ($PSBoundParameters.ContainsKey($k)) { $defaults[$k] = $PSBoundParameters[$k] } }
        $start = Show-TarStartForm -Defaults $defaults
        if (-not $start) { Write-TarLog 'Cancelled on the start screen.' -Level WARN; return }
        if ($start.ProfileRecord) { Save-TarProfile -ProfileRecord $start.ProfileRecord -Confirm:$false }
        $conn = $start.Connection
        $IncludeSharePoint = [switch]$start.Options.IncludeSharePoint
        $IncludeUserAssociations = [switch]$start.Options.IncludeUserAssociations
        $SelectChannels = [switch]$start.Options.SelectChannels
        $OutputPath = $start.Options.OutputPath
    } else {
        $conn = Resolve-TarConnection -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -ProfileName $ProfileName `
            -ResumeSelection $resumeSelection -CanPrompt:$canPrompt -SelectionMode $SelectionMode
    }
    if (-not $conn) { Write-TarLog 'Cancelled at connection selection.' -Level WARN; return }
    if ($ProfileName) {
        # Report options saved in the profile apply unless the switch is given on the command line.
        $saved = Get-TarProperty (Get-TarProfile -Name $ProfileName) 'Options'
        $applied = @(foreach ($k in 'IncludeSharePoint', 'IncludeUserAssociations', 'SelectChannels') {
                if (-not $PSBoundParameters.ContainsKey($k) -and $saved -and [bool](Get-TarProperty $saved $k)) { Set-Variable -Name $k -Value ([switch]$true); $k }
            })
        if ($applied.Count) { Write-TarLog "Options from profile '$ProfileName': $($applied -join ', ')" }
    }
    $appOnly = $conn.AuthMode -eq 'AppOnly'
    $spClientId = if ($PnPClientId) { $PnPClientId } elseif ($conn.PnPClientId) { $conn.PnPClientId } elseif ($appOnly) { $conn.ClientId } else { $null }
    if ($IncludeSharePoint -and -not $spClientId) {
        throw '-IncludeSharePoint with delegated sign-in needs a PnP app (-PnPClientId). Create it once with: New-TeamsAccessReportPnPApp -ProfileName <profile> (or -Tenant <tenant>), or use app-only.'
    }
    if ($IncludeUserAssociations -and -not $appOnly) { Write-Warning '-IncludeUserAssociations requires app-only (delegated only supports "me"). The step is skipped.' }
    $interactive = -not $Resume -and -not $TeamId -and -not $AllTeams
    if ($interactive -and -not $canPrompt) { throw 'No interactive console. Use -TeamId, -AllTeams or -Resume.' }
    if (Get-Module -Name PnP.PowerShell) { throw 'PnP.PowerShell is loaded in this session - start a new pwsh (assembly conflict with the Graph SDK).' }
    if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) { throw 'Microsoft.Graph.Authentication is missing: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }

    # --- Context, store, logging ---
    $ctx = New-TarContext -OutputPath $OutputPath -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds -MaxItemsPerList $MaxItemsPerList -MaxSubwebDepth $MaxSubwebDepth
    $ctx.AppOnly = $appOnly
    $ctx.ProfileName = $conn.ProfileName
    if ($appOnly) { $ctx.ClientId = $conn.ClientId; $ctx.CertificateThumbprint = $conn.CertificateThumbprint }
    if (-not $PSCmdlet.ShouldProcess($ctx.OutputPath, 'Create access report')) { return }
    Initialize-TarStore -Context $ctx -Resume:$Resume
    Initialize-TarLog -Path (Get-TarStorePath $ctx $script:TarFile.Log)
    $via = if ($conn.ProfileName) { "profile '$($conn.ProfileName)'" } else { 'parameters' }
    Write-TarLog "Connection: $($conn.AuthMode), tenant $(if ($conn.TenantId) { $conn.TenantId } else { '<chosen at sign-in>' }) via $via"

    try {
        Connect-TarGraph -Context $ctx -TenantId $conn.TenantId
        if ($interactive -and -not $Force -and -not (Read-TarYesNo "Continue with tenant '$($ctx.TenantName ?? $ctx.TenantId)' ($($ctx.TenantDomain ?? $ctx.TenantId))?" -Mode $SelectionMode)) {
            Write-TarLog 'Cancelled at tenant confirmation.' -Level WARN
            return
        }

        if ($Resume) {
            $selection = Import-TarSelection -Context $ctx
            Write-TarLog "Resuming: $($selection.Teams.Count) teams from Selection.json"
            Reset-TarIncompleteWork -Context $ctx
        } else {
            $selection = Get-TarTeamSelection -Context $ctx -TeamId $TeamId -ChannelId $ChannelId -AllTeams:$AllTeams -Interactive:$interactive `
                -SelectChannels:$SelectChannels -Force:$Force -SelectionMode $SelectionMode -NoChannelPrompt:$useStartForm
            if (-not $selection) { return }
            Save-TarSelection -Context $ctx -Selection $selection
        }
        Write-TarLog "Teams to process: $($selection.Teams.Count)"

        Invoke-TarTeamInventory -Context $ctx -Teams $selection.Teams -ChannelFilter $selection.ChannelFilter
        # Header-only files mark "phase ran, 0 rows"; files of phases that did not run stay absent.
        Initialize-TarRecordFile -Context $ctx -Name $script:TarFile.Teams, $script:TarFile.TeamsAccess, $script:TarFile.Channels, $script:TarFile.IncomingChannels
        if ($IncludeSharePoint) {
            Start-TarSharePointScan -Context $ctx -ClientId $spClientId
            Initialize-TarRecordFile -Context $ctx -Name $script:TarFile.SpAssignments, $script:TarFile.SpGroupMembers, $script:TarFile.SharingLinks
            Invoke-TarPrincipalResolution -Context $ctx
            Invoke-TarDriftAnalysis -Context $ctx
        }
        if ($IncludeUserAssociations -and $appOnly) { Invoke-TarExternalAssociationScan -Context $ctx }

        Initialize-TarRecordFile -Context $ctx -Name $script:TarFile.Errors
        $summary = Export-TarSummary -Context $ctx -SharePointIncluded:$IncludeSharePoint
        Write-TarLog ("Done: Status={0} Teams={1} Rows={2} Drift={3} (High={4}) Errors={5} -> {6}" -f $summary.ResultStatus, $summary.Teams, $summary.AccessRows, $summary.DriftFindings, $summary.DriftHigh, $summary.Errors, $ctx.OutputPath)
        if ($summary.Errors) { Write-TarLog "See Errors.csv - $($summary.Errors) steps failed (HTTP 403 = missing permission/consent; other errors include the reason)." -Level WARN }
        $summary
    } finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
