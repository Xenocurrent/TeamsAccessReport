# Phase 1 - Teams inventory via Graph. Checkpoint per team; rows are written only when the team is complete.

function Get-TarChannelAccessPath {
    # Pure: how a channel member got access. Graph sets originalSourceMembershipUrl on allMembers rows; for private
    # channels it can point at the host team itself (seen in a real tenant), which is NOT indirect access.
    #   private                       -> DirectChannelMember
    #   shared, no source             -> DirectChannelMember
    #   shared, source = host team    -> IndirectViaHostTeam   (channel shared with its own team)
    #   shared, source = other team   -> IndirectViaSharedTeam
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$ChannelType, [string]$SourceTeamId, [Parameter(Mandatory)][string]$HostTeamId)
    if ($ChannelType -eq 'private' -or -not $SourceTeamId) { return 'DirectChannelMember' }
    if ($SourceTeamId -eq $HostTeamId) { 'IndirectViaHostTeam' } else { 'IndirectViaSharedTeam' }
}

function ConvertTo-TarAccessRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Member, [Parameter(Mandatory)]$Team, $Channel,
        [Parameter(Mandatory)][string]$Scope, [Parameter(Mandatory)][string]$AccessPath, [string]$SourceTeamId)
    $roles = @(Get-TarProperty $Member 'roles')
    $role = if ($roles -contains 'owner') { 'Owner' } elseif ($roles -contains 'guest') { 'Guest' } else { 'Member' }
    [pscustomobject]@{
        TeamId = $Team.id; TeamName = $Team.displayName; Scope = $Scope
        ChannelId = Get-TarProperty $Channel 'id'; ChannelName = Get-TarProperty $Channel 'displayName'; ChannelType = Get-TarProperty $Channel 'membershipType'
        UserId = Get-TarProperty $Member 'userId'; DisplayName = Get-TarProperty $Member 'displayName'; Email = Get-TarProperty $Member 'email'
        UserTenantId = Get-TarProperty $Member 'tenantId'; Role = $role; RawRoles = ($roles -join ';')
        AccessPath = $AccessPath; SourceTeamId = $SourceTeamId; MemberType = Get-TarProperty $Member '@odata.type'
    }
}

function Test-TarFolderNotProvisioned {
    # Graph returns 404 "Folder location for this channel is not ready yet" until the channel's files folder
    # has been provisioned (typically when someone opens the channel's Files tab for the first time).
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    $data = $ErrorRecord.Exception.Data
    return ($data -and $data['StatusCode'] -eq 404 -and [string]$data['GraphMessage'] -match 'not ready')
}

function Get-TarChannelSite {
    # filesFolder -> driveItem -> sharepointIds.siteUrl (works for all channel types and managed paths).
    # Standard channels fall back to the team's own document library when the folder is not provisioned yet,
    # because standard channels always live on the team site.
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$TeamId, [Parameter(Mandatory)][string]$ChannelId,
        [Parameter(Mandatory)][string]$ChannelName, [Parameter(Mandatory)][string]$MembershipType)
    $base = $Context.GraphBase
    try {
        $ff = Invoke-TarGraphRequest -Context $Context -Uri "$base/teams/$TeamId/channels/$ChannelId/filesFolder"
    } catch {
        if (-not (Test-TarFolderNotProvisioned $_)) { throw }
        if ($MembershipType -ne 'standard') {
            throw "Files folder for $MembershipType channel '$ChannelName' is not provisioned yet. Open the channel's Files tab in Teams once, then re-run with -Resume."
        }
        Write-TarLog "    Files folder for '$ChannelName' not provisioned yet; using the team's document library (standard channel)." -Level WARN
        $root = Invoke-TarGraphRequest -Context $Context -Uri "$base/groups/$TeamId/drive/root`?`$select=id,webUrl,parentReference,sharepointIds"
        return [pscustomobject]@{
            SiteUrl = Get-TarProperty (Get-TarProperty $root 'sharepointIds') 'siteUrl'; FilesFolderUrl = Get-TarProperty $root 'webUrl'
            DriveId = Get-TarProperty (Get-TarProperty $root 'parentReference') 'driveId'; FolderItemId = Get-TarProperty $root 'id'; SiteSource = 'GroupDriveFallback'
        }
    }
    $folderId = Get-TarProperty $ff 'id'
    $driveId = Get-TarProperty (Get-TarProperty $ff 'parentReference') 'driveId'
    $siteUrl = $null
    if ($driveId -and $folderId) {
        $item = Invoke-TarGraphRequest -Context $Context -Uri "$base/drives/$driveId/items/$folderId`?`$select=id,sharepointIds"
        $siteUrl = Get-TarProperty (Get-TarProperty $item 'sharepointIds') 'siteUrl'
    }
    [pscustomobject]@{ SiteUrl = $siteUrl; FilesFolderUrl = Get-TarProperty $ff 'webUrl'; DriveId = $driveId; FolderItemId = $folderId; SiteSource = 'FilesFolder' }
}

function Invoke-TarTeamInventory {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][object[]]$Teams, [Parameter(Mandatory)][hashtable]$ChannelFilter)
    $base = $Context.GraphBase
    $done = Get-TarCheckpoint -Context $Context -Name $script:TarCheckpoint.Graph
    $i = 0
    foreach ($team in $Teams) {
        $i++
        if ($done.Contains($team.id)) { Write-TarLog "[$i] Skipping (done): $($team.displayName)"; continue }
        Write-Progress -Activity 'Phase 1: Teams' -Status "$($team.displayName) ($i/$($Teams.Count))" -PercentComplete (100 * $i / [math]::Max(1, $Teams.Count))
        Write-TarLog "[$i/$($Teams.Count)] Team '$($team.displayName)' ($($team.id))"
        $failed = [System.Collections.Generic.List[string]]::new()
        $access = [System.Collections.Generic.List[object]]::new()
        $channelRecords = [System.Collections.Generic.List[object]]::new()
        $incomingRecords = [System.Collections.Generic.List[object]]::new()
        $fail = { param($stage, $channelId, $err) $failed.Add($stage); Add-TarErrorRecord -Context $Context -Stage ($stage -replace ':.*$') -TeamId $team.id -ChannelId $channelId -Message $err.Exception.Message }

        $info = $null
        try { $info = Invoke-TarGraphRequest -Context $Context -Uri "$base/teams/$($team.id)?`$select=id,displayName,visibility,isArchived,webUrl,createdDateTime" }
        catch { & $fail 'TeamInfo' $null $_ }

        try {
            foreach ($m in @(Get-TarGraphCollection -Context $Context -Uri "$base/teams/$($team.id)/members")) {
                $access.Add((ConvertTo-TarAccessRecord -Member $m -Team $team -Scope 'Team' -AccessPath 'TeamMembership' -SourceTeamId $team.id))
            }
        } catch { & $fail 'TeamMembers' $null $_ }

        try {
            foreach ($ic in @(Get-TarGraphCollection -Context $Context -Uri "$base/teams/$($team.id)/incomingChannels")) {
                $hostTenant = Get-TarProperty $ic 'tenantId'
                $incomingRecords.Add([pscustomobject]@{ TeamId = $team.id; TeamName = $team.displayName; ChannelId = Get-TarProperty $ic 'id'
                        ChannelName = Get-TarProperty $ic 'displayName'; ChannelType = Get-TarProperty $ic 'membershipType'; HostTenantId = $hostTenant
                        CrossTenant = [bool]($hostTenant -and $hostTenant -ne $Context.TenantId)
                    })
            }
        } catch { & $fail 'IncomingChannels' $null $_ }

        $channels = @()
        try { $channels = @(Get-TarGraphCollection -Context $Context -Uri "$base/teams/$($team.id)/channels?`$select=id,displayName,membershipType") }
        catch { & $fail 'Channels' $null $_ }
        if ($ChannelFilter.ContainsKey($team.id)) {
            $sel = $ChannelFilter[$team.id]
            $channels = @($channels | Where-Object { $sel.Contains([string]$_.id) })
            Write-TarLog "    Channel filter applied: $($channels.Count) channels"
        }

        foreach ($u in @($channels | Where-Object { $_.membershipType -notin 'standard', 'private', 'shared' })) {
            Write-TarLog "    Channel '$($u.displayName)' has membershipType '$($u.membershipType)'; treated as a channel with its own site and member list." -Level WARN

        }

        foreach ($ch in $channels) {
            $site = [pscustomobject]@{ SiteUrl = $null; FilesFolderUrl = $null; DriveId = $null; FolderItemId = $null; SiteSource = $null }
            try { $site = Get-TarChannelSite -Context $Context -TeamId $team.id -ChannelId $ch.id -ChannelName $ch.displayName -MembershipType $ch.membershipType }
            catch { & $fail "FilesFolder:$($ch.displayName)" $ch.id $_ }
            $channelRecords.Add([pscustomobject]@{ TeamId = $team.id; TeamName = $team.displayName; ChannelId = $ch.id; ChannelName = $ch.displayName
                    ChannelType = $ch.membershipType; SiteUrl = $site.SiteUrl; FilesFolderUrl = $site.FilesFolderUrl; DriveId = $site.DriveId
                    FolderItemId = $site.FolderItemId; SiteSource = $site.SiteSource; OwnSharePointSite = ($ch.membershipType -ne 'standard')
                })

            # Anything but 'standard' has its own site and member list (defensive for future enum values).
            if ($ch.membershipType -ne 'standard') {
                try {
                    $members = @(Get-TarGraphCollection -Context $Context -Uri "$base/teams/$($team.id)/channels/$($ch.id)/allMembers")
                    foreach ($cm in $members) {
                        $src = Get-TarProperty $cm '@microsoft.graph.originalSourceMembershipUrl'
                        $srcTeam = if ($src -and $src -match '/teams/([^/]+)') { $Matches[1] } else { $null }
                        $path = Get-TarChannelAccessPath -ChannelType $ch.membershipType -SourceTeamId $srcTeam -HostTeamId $team.id
                        $access.Add((ConvertTo-TarAccessRecord -Member $cm -Team $team -Channel $ch -Scope 'Channel' -AccessPath $path -SourceTeamId $srcTeam))
                    }
                    Write-TarLog "    Channel '$($ch.displayName)' [$($ch.membershipType)]: $($members.Count) membership rows"
                } catch { & $fail "AllMembers:$($ch.displayName)" $ch.id $_ }
            }
        }

        Write-TarRecord -Context $Context -Name $script:TarFile.TeamsAccess -InputObject $access
        Write-TarRecord -Context $Context -Name $script:TarFile.Channels -InputObject $channelRecords
        Write-TarRecord -Context $Context -Name $script:TarFile.IncomingChannels -InputObject $incomingRecords
        Write-TarRecord -Context $Context -Name $script:TarFile.Teams -InputObject @([pscustomobject]@{
                TeamId = $team.id; TeamName = $team.displayName; Visibility = Get-TarProperty $info 'visibility'; IsArchived = Get-TarProperty $info 'isArchived'
                WebUrl = Get-TarProperty $info 'webUrl'; CreatedDateTime = Get-TarProperty $info 'createdDateTime'; Channels = $channels.Count
                Completeness = if ($failed.Count) { 'Partial' } else { 'Complete' }; FailedStages = ($failed -join ';')
            })
        Add-TarCheckpoint -Context $Context -Name $script:TarCheckpoint.Graph -Key $team.id
    }
    Write-Progress -Activity 'Phase 1: Teams' -Completed
}

function Reset-TarIncompleteWork {
    # On -Resume: archive old errors and remove teams with Completeness=Partial so they run again.
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context)
    $errFile = Get-TarStorePath $Context $script:TarFile.Errors
    if ((Test-Path $errFile) -and $PSCmdlet.ShouldProcess($errFile, 'Archive')) {
        $archived = Join-Path $Context.StateDir ('Errors_{0:yyyyMMdd_HHmmss}.csv' -f (Get-Date))
        Move-Item $errFile $archived -Force
        Write-TarLog "Previous errors archived: $archived"
    }
    $partial = [string[]]@(Read-TarRecord -Context $Context -Name $script:TarFile.Teams | Where-Object Completeness -eq 'Partial' | ForEach-Object { $_.TeamId })
    if ($partial.Count -eq 0) { return }
    foreach ($name in $script:TarFile.TeamsAccess, $script:TarFile.Channels, $script:TarFile.IncomingChannels, $script:TarFile.Teams) {
        $keep = @(Read-TarRecord -Context $Context -Name $name | Where-Object { $_.TeamId -notin $partial })
        Set-TarRecord -Context $Context -Name $name -InputObject $keep
    }
    Remove-TarCheckpoint -Context $Context -Name $script:TarCheckpoint.Graph -Key $partial
    Write-TarLog "Re-running $($partial.Count) incomplete teams: $($partial -join ', ')" -Level WARN
}
