# Phase 4 - drift between Teams and SharePoint, external associations, summary.

function New-TarDriftRecord {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory constructor; does not change system state.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][hashtable]$Site, [Parameter(Mandatory)][string]$Finding, [Parameter(Mandatory)][ValidateSet('High', 'Medium', 'Low', 'Info')][string]$Severity,
        $UserId, $Upn, $DisplayName, $Scope, $ObjectPath, $Roles, $ExpansionPath, [string]$Cause)
    [pscustomobject]@{ SiteUrl = $Site.Url; Site = $Site.Label; SiteKind = $Site.Kind; Finding = $Finding; Severity = $Severity; Cause = $Cause; UserId = $UserId; Upn = $Upn
        DisplayName = $DisplayName; Scope = $Scope; ObjectPath = $ObjectPath; Roles = $Roles; ExpansionPath = $ExpansionPath
    }
}

function Get-TarSiteMap {
    # SiteUrl -> which Teams population is the "expected" access (team members or channel members).
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Channels, [AllowEmptyCollection()][object[]]$Teams = @())
    $visibility = @{}
    foreach ($t in $Teams) { $visibility[[string]$t.TeamId] = [string]$t.Visibility }
    $map = @{}
    foreach ($c in ($Channels | Where-Object { $_.SiteUrl })) {
        $url = $c.SiteUrl.TrimEnd('/')
        $vis = $visibility[[string]$c.TeamId]
        if ($c.ChannelType -in 'private', 'shared') { $map[$url] = @{ Url = $url; Kind = 'ChannelSite'; TeamId = $c.TeamId; ChannelId = $c.ChannelId; Label = "$($c.TeamName) / $($c.ChannelName)"; Visibility = $vis } }
        elseif (-not $map.ContainsKey($url)) { $map[$url] = @{ Url = $url; Kind = 'TeamSite'; TeamId = $c.TeamId; ChannelId = $null; Label = $c.TeamName; Visibility = $vis } }
    }
    $map
}

function Get-TarDriftCause {
    # Pure function: returns a known by-design explanation for a broad grant, or '' when none applies.
    # PublicTeamDefault requires every condition, so a manually added grant is never explained away.
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][hashtable]$Site, [Parameter(Mandatory)]$Row)
    $siteIsPublicTeam = $Site.Kind -eq 'TeamSite' -and $Site.Visibility -eq 'public'
    $rowIsEeeu = $Row.Kind -eq 'BroadGrant' -and $Row.Upn -eq 'EveryoneExceptExternalUsers'
    # Get-TarProperty: rows from older or partial producers may lack AssociatedGroup/WebUrl (strict mode).
    $viaRootMembers = $Row.Scope -eq 'Web' -and (Get-TarProperty $Row 'AssociatedGroup') -eq 'Member' -and ([string](Get-TarProperty $Row 'WebUrl')).TrimEnd('/') -eq $Site.Url
    $isDefault = $siteIsPublicTeam -and $rowIsEeeu -and $viaRootMembers
    if ($isDefault) { $script:TarCause.PublicTeamDefault } else { '' }
}

function Get-TarSiteDrift {
    # Pure function: compares expected (Teams) with actual (SharePoint) access for ONE site.
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Site, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TeamsUsers,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SiteRows)
    $teamsSet = [System.Collections.Generic.HashSet[string]]::new([string[]]@($TeamsUsers | ForEach-Object { $_.UserId }), [System.StringComparer]::OrdinalIgnoreCase)
    $real = @($SiteRows | Where-Object { $_.LimitedAccessOnly -ne 'True' })

    foreach ($b in ($real | Where-Object { $_.Kind -in 'BroadGrant', 'AnonymousLink', 'OrganizationLink' })) {
        New-TarDriftRecord -Site $Site -Finding $b.Kind -Severity High -UserId $b.UserId -Upn $b.Upn -DisplayName $b.DisplayName -Scope $b.Scope -ObjectPath $b.ObjectPath -Roles $b.Roles -ExpansionPath $b.ExpansionPath `
            -Cause (Get-TarDriftCause -Site $Site -Row $b)
    }
    $spUsers = @($real | Where-Object { $_.Kind -eq 'User' })
    foreach ($g in ($spUsers | Group-Object { if ($_.UserId) { $_.UserId } else { $_.Upn } })) {
        $r = $g.Group[0]
        if ($r.UserId -and $teamsSet.Contains($r.UserId)) { continue }
        $scopes = @($g.Group.Scope | Sort-Object -Unique)
        $finding = if ($scopes -contains 'SiteCollectionAdmin') { $script:TarFinding.SiteAdminNotInTeams } elseif ($scopes -contains 'Web') { $script:TarFinding.SiteNotInTeams } else { $script:TarFinding.ItemNotInTeams }
        $sev = if (Test-TarGuestPrincipal -UserType $r.UserType -Upn $r.Upn) { 'High' } elseif ($finding -eq $script:TarFinding.ItemNotInTeams) { 'Low' } else { 'Medium' }
        New-TarDriftRecord -Site $Site -Finding $finding -Severity $sev -UserId $r.UserId -Upn $r.Upn -DisplayName $r.DisplayName -Scope ($scopes -join ';') `
            -ObjectPath ((@($g.Group.ObjectPath | Sort-Object -Unique | Select-Object -First 5)) -join ' | ') `
            -Roles ((@($g.Group.Roles | Sort-Object -Unique)) -join ';') -ExpansionPath ((@($g.Group.ExpansionPath | Sort-Object -Unique)) -join ' | ')
    }
    if (@($SiteRows | Where-Object Scope -eq 'Web').Count -gt 0) {
        $spSet = [System.Collections.Generic.HashSet[string]]::new([string[]]@($spUsers | Where-Object { $_.Scope -in 'Web', 'SiteCollectionAdmin' -and $_.UserId } | ForEach-Object { $_.UserId }), [System.StringComparer]::OrdinalIgnoreCase)
        foreach ($t in ($TeamsUsers | Group-Object UserId)) {
            if ($spSet.Contains($t.Name)) { continue }
            $tr = $t.Group[0]
            New-TarDriftRecord -Site $Site -Finding $script:TarFinding.TeamsNotInSp -Severity Info -UserId $tr.UserId -Upn $tr.Email -DisplayName $tr.DisplayName -Scope $tr.Scope -ObjectPath $Site.Url -Roles $tr.Role -ExpansionPath "Teams:$($tr.AccessPath)"
        }
    }
    foreach ($u in ($real | Where-Object Kind -eq 'Unresolved')) {
        New-TarDriftRecord -Site $Site -Finding $script:TarFinding.Unresolved -Severity Medium -Upn $u.Upn -DisplayName $u.DisplayName -Scope $u.Scope -ObjectPath $u.ObjectPath -Roles $u.Roles -ExpansionPath $u.ExpansionPath
    }
}

function Invoke-TarDriftAnalysis {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)
    $effective = @(Read-TarRecord -Context $Context -Name $script:TarFile.EffectiveSpAccess)
    if ($effective.Count -eq 0) { Set-TarRecord -Context $Context -Name $script:TarFile.Drift -InputObject @(); return }
    $teamsAccess = @(Read-TarRecord -Context $Context -Name $script:TarFile.TeamsAccess)
    $siteMap = Get-TarSiteMap -Channels @(Read-TarRecord -Context $Context -Name $script:TarFile.Channels) -Teams @(Read-TarRecord -Context $Context -Name $script:TarFile.Teams)
    $drift = foreach ($site in $siteMap.Values) {
        # @(if ...) is required: an empty array returned from an if expression otherwise becomes $null.
        $teamsUsers = @(if ($site.Kind -eq 'ChannelSite') { $teamsAccess | Where-Object { $_.ChannelId -eq $site.ChannelId -and $_.UserId } }
            else { $teamsAccess | Where-Object { $_.TeamId -eq $site.TeamId -and $_.Scope -eq 'Team' -and $_.UserId } })
        Get-TarSiteDrift -Site $site -TeamsUsers $teamsUsers -SiteRows @($effective | Where-Object { $_.SiteUrl.TrimEnd('/') -eq $site.Url })
    }
    $drift = @($drift | Sort-Object SiteUrl, Severity, Finding, Upn)
    Set-TarRecord -Context $Context -Name $script:TarFile.Drift -InputObject $drift
    Write-TarLog "Phase 4: $($drift.Count) drift findings (High=$(@($drift | Where-Object Severity -eq 'High').Count))"
}

function Invoke-TarExternalAssociationScan {
    # Requires app-only (delegated only supports 'me').
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)
    $internal = @(Read-TarRecord -Context $Context -Name $script:TarFile.TeamsAccess | Where-Object { $_.UserId -and $_.UserTenantId -eq $Context.TenantId } | ForEach-Object UserId | Sort-Object -Unique)
    Write-TarLog "Phase 4b: external associations for $($internal.Count) internal users"
    $assoc = foreach ($uid in $internal) {
        try {
            foreach ($at in @(Get-TarGraphCollection -Context $Context -Uri "$($Context.GraphBase)/users/$uid/teamwork/associatedTeams")) {
                $tid = Get-TarProperty $at 'tenantId'
                if ($tid -and $tid -ne $Context.TenantId) { [pscustomobject]@{ UserId = $uid; HostTenantId = $tid; HostTeamId = Get-TarProperty $at 'id'; HostTeamName = Get-TarProperty $at 'displayName' } }
            }
        } catch { Add-TarErrorRecord -Context $Context -Stage 'AssociatedTeams' -Target $uid -Message $_.Exception.Message }
    }
    Set-TarRecord -Context $Context -Name $script:TarFile.ExternalAssoc -InputObject @($assoc)
    Write-TarLog "    External associations: $(@($assoc).Count)"
}

function Export-TarSummary {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [switch]$SharePointIncluded)
    $teams = @(Read-TarRecord -Context $Context -Name $script:TarFile.Teams)
    $access = @(Read-TarRecord -Context $Context -Name $script:TarFile.TeamsAccess)
    $channels = @(Read-TarRecord -Context $Context -Name $script:TarFile.Channels)
    $errors = @(Read-TarRecord -Context $Context -Name $script:TarFile.Errors)
    $drift = @(Read-TarRecord -Context $Context -Name $script:TarFile.Drift)
    $incoming = @(Read-TarRecord -Context $Context -Name $script:TarFile.IncomingChannels)
    $partial = @($teams | Where-Object Completeness -eq 'Partial').Count
    $summary = [pscustomobject][ordered]@{
        SchemaVersion = 1; TenantId = $Context.TenantId; TenantName = $Context.TenantName; AuthType = $Context.AuthType; OutputPath = $Context.OutputPath
        Teams = $teams.Count; TeamsPartial = $partial; ArchivedTeams = @($teams | Where-Object IsArchived -eq 'True').Count
        Channels = $channels.Count; PrivateChannels = @($channels | Where-Object ChannelType -eq 'private').Count; SharedChannels = @($channels | Where-Object ChannelType -eq 'shared').Count
        IncomingCrossTenantChannels = @($incoming | Where-Object CrossTenant -eq 'True').Count
        AccessRows = $access.Count; IndirectRows = @($access | Where-Object AccessPath -eq 'IndirectViaSharedTeam').Count
        GuestRows = @($access | Where-Object Role -eq 'Guest').Count; ExternalTenantRows = @($access | Where-Object { $_.UserTenantId -and $_.UserTenantId -ne $Context.TenantId }).Count
        SharePointIncluded = [bool]$SharePointIncluded; DriftFindings = $drift.Count; DriftHigh = @($drift | Where-Object Severity -eq 'High').Count
        Errors = $errors.Count; ResultStatus = if ($errors.Count -eq 0 -and $partial -eq 0) { 'Complete' } else { 'Incomplete' }
    }
    $summary | ConvertTo-Json | Set-Content (Get-TarStorePath $Context $script:TarFile.Summary) -Encoding utf8
    $summary
}
