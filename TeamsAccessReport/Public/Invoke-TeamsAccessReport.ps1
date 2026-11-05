function Invoke-TeamsAccessReport {
    <#
    .SYNOPSIS
        Fullständig åtkomstrapport för Microsoft Teams och tillhörande SharePoint-webbplatser.
    .DESCRIPTION
        Fas 1  Teams via Graph (team, kanaler, direkt/indirekt medlemskap, kanalsajter)     – checkpoint per team
        Fas 2  SharePoint via PnP i separat process (admins, grupper, unika behörigheter)    – checkpoint per webbplats
        Fas 3  Upplösning av SharePoint-claims till Entra-användare
        Fas 4  Avvikelser Teams <-> SharePoint, externa kopplingar, sammanfattning
        Returnerar sammanfattningsobjektet; alla detaljer skrivs som CSV/JSON i -OutputPath.
    .PARAMETER TenantId
        Valfritt delegerat (identifieras vid inloggning). Krävs app-only.
    .PARAMETER Resume
        Återupptar en avbruten körning i -OutputPath. Ofullständiga team och misslyckade webbplatser körs om.
    .EXAMPLE
        Invoke-TeamsAccessReport
        Interaktivt: logga in, bekräfta tenant, välj team och kanaler.
    .EXAMPLE
        Invoke-TeamsAccessReport -TenantId '<TENANT_ID>' -ClientId '<APP_ID>' -CertificateThumbprint '<THUMBPRINT>' -AllTeams -Force -IncludeSharePoint -IncludeUserAssociations
    .OUTPUTS
        PSCustomObject (sammanfattning, samma innehåll som Summary.json)
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

    # --- Validering av kombinationer som parameter sets inte kan uttrycka ---
    if ($ChannelId -and -not $TeamId) { throw '-ChannelId kräver -TeamId.' }
    if ($TeamId -and $AllTeams) { throw 'Ange antingen -TeamId eller -AllTeams, inte båda.' }
    if ($Resume -and ($TeamId -or $AllTeams -or $ChannelId)) { throw '-Resume läser urvalet från Selection.json; kombinera inte med -TeamId/-AllTeams/-ChannelId.' }
    $spClientId = if ($PnPClientId) { $PnPClientId } elseif ($appOnly) { $ClientId } else { $null }
    if ($IncludeSharePoint -and -not $spClientId) {
        throw "-IncludeSharePoint kräver en egen Entra-app för PnP (-PnPClientId). Skapa: Register-PnPEntraIDAppForInteractiveLogin -ApplicationName 'PnP-TeamsAccess' -Tenant <tenant>.onmicrosoft.com -SharePointDelegatePermissions AllSites.FullControl"
    }
    if ($IncludeUserAssociations -and -not $appOnly) { Write-Warning '-IncludeUserAssociations kräver app-only (delegerat stöder endast "me"). Steget hoppas över.' }
    $interactive = -not $Resume -and -not $TeamId -and -not $AllTeams
    if ($interactive -and ([Console]::IsInputRedirected -or -not [Environment]::UserInteractive)) { throw 'Ingen interaktiv konsol. Använd -TeamId, -AllTeams eller -Resume.' }
    if (Get-Module -Name PnP.PowerShell) { throw 'PnP.PowerShell är laddad i sessionen – starta en ny pwsh (assemblykonflikt med Graph SDK).' }
    if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) { throw 'Microsoft.Graph.Authentication saknas: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser' }
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }

    # --- Kontext, lager, loggning ---
    $ctx = New-TarContext -OutputPath $OutputPath -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds -MaxItemsPerList $MaxItemsPerList -MaxSubwebDepth $MaxSubwebDepth
    $ctx.AppOnly = $appOnly
    if ($appOnly) { $ctx.ClientId = $ClientId; $ctx.CertificateThumbprint = $CertificateThumbprint }
    if (-not $PSCmdlet.ShouldProcess($ctx.OutputPath, 'Skapa åtkomstrapport')) { return }
    Initialize-TarStore -Context $ctx -Resume:$Resume
    Initialize-TarLog -Path (Get-TarStorePath $ctx $script:TarFile.Log)

    try {
        Connect-TarGraph -Context $ctx -TenantId $TenantId
        if ($interactive -and -not $Force -and -not (Read-TarYesNo "Fortsätt mot tenant '$($ctx.TenantName ?? $ctx.TenantId)' ($($ctx.TenantDomain ?? $ctx.TenantId))?")) {
            Write-TarLog 'Avbrutet vid tenantbekräftelse.' -Level WARN
            return
        }

        if ($Resume) {
            $selection = Import-TarSelection -Context $ctx
            Write-TarLog "Återupptar: $($selection.Teams.Count) team från Selection.json"
            Reset-TarIncompleteWork -Context $ctx
        } else {
            $selection = Get-TarTeamSelection -Context $ctx -TeamId $TeamId -ChannelId $ChannelId -AllTeams:$AllTeams -Interactive:$interactive `
                -SelectChannels:$SelectChannels -Force:$Force -SelectionMode $SelectionMode
            if (-not $selection) { return }
            Save-TarSelection -Context $ctx -Selection $selection
        }
        Write-TarLog "Antal team att behandla: $($selection.Teams.Count)"

        Invoke-TarTeamInventory -Context $ctx -Teams $selection.Teams -ChannelFilter $selection.ChannelFilter
        if ($IncludeSharePoint) {
            Start-TarSharePointScan -Context $ctx -ClientId $spClientId
            Invoke-TarPrincipalResolution -Context $ctx
            Invoke-TarDriftAnalysis -Context $ctx
        }
        if ($IncludeUserAssociations -and $appOnly) { Invoke-TarExternalAssociationScan -Context $ctx }

        $summary = Export-TarSummary -Context $ctx -SharePointIncluded:$IncludeSharePoint
        Write-TarLog ("Klart: Status={0} Team={1} Rader={2} Avvikelser={3} (High={4}) Fel={5} -> {6}" -f $summary.ResultStatus, $summary.Teams, $summary.AccessRows, $summary.DriftFindings, $summary.DriftHigh, $summary.Errors, $ctx.OutputPath)
        if ($summary.Errors) { Write-TarLog "Se Errors.csv – $($summary.Errors) steg misslyckades (403 = saknad behörighet/consent)." -Level WARN }
        $summary
    } finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
