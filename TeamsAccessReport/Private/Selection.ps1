# Interactive selection (UI). Isolated from data logic so everything else can run unattended and be tested.

function Resolve-TarSelectionMode {
    [CmdletBinding()]
    [OutputType([string])]
    param([ValidateSet('Auto', 'Gui', 'ConsoleGrid', 'Menu')][string]$Mode = 'Auto')
    if ($Mode -ne 'Auto') { return $Mode }
    if ($IsWindows -and (Get-Command Out-GridView -ErrorAction SilentlyContinue)) { return 'Gui' }
    if (Get-Command Out-ConsoleGridView -ErrorAction SilentlyContinue) { return 'ConsoleGrid' }
    return 'Menu'
}

function ConvertFrom-TarSelectionInput {
    # "1,3,5-7" -> zero-based indexes. $null on invalid input.
    [CmdletBinding()]
    [OutputType([int[]])]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text, [Parameter(Mandatory)][int]$Count)
    $idx = [System.Collections.Generic.SortedSet[int]]::new()
    foreach ($part in ($Text -split '[,\s]+' | Where-Object { $_ })) {
        if ($part -match '^(\d+)-(\d+)$') {
            $a = [int]$Matches[1]; $b = [int]$Matches[2]
            if ($a -lt 1 -or $b -gt $Count -or $a -gt $b) { return $null }
            foreach ($n in $a..$b) { [void]$idx.Add($n - 1) }
        } elseif ($part -match '^\d+$') {
            $n = [int]$part
            if ($n -lt 1 -or $n -gt $Count) { return $null }
            [void]$idx.Add($n - 1)
        } else { return $null }
    }
    Write-Output -NoEnumerate -InputObject ([int[]]@($idx))
}

function Select-TarFromMenu {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive menu rendering requires direct console output.')]
    [CmdletBinding()]
    [OutputType([psobject], [object[]])]
    param([Parameter(Mandatory)][object[]]$Rows, [Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string[]]$Columns)
    $filter = $null
    while ($true) {
        $view = @(if ($filter) { $Rows | Where-Object { ($_.PSObject.Properties.Value -join ' ') -like "*$filter*" } } else { $Rows })
        Write-Host ''
        Write-Host "=== $Title ===  ($($view.Count) of $($Rows.Count) shown$(if ($filter) { ", filter: '$filter'" }))" -ForegroundColor Cyan
        for ($n = 0; $n -lt $view.Count; $n++) {
            $cells = foreach ($c in $Columns) { [string]$view[$n].$c }
            Write-Host ('{0,4}) {1}' -f ($n + 1), ($cells -join '  |  '))
        }
        Write-Host 'Select: numbers/ranges (1,3,5-7) | a = all shown | /text = filter | c = clear filter | q = cancel' -ForegroundColor DarkGray
        $answer = (Read-Host 'Your choice').Trim()
        if ($answer -eq 'q') { return }
        if ($answer -eq 'c') { $filter = $null; continue }
        if ($answer -match '^/(.+)$') { $filter = $Matches[1]; continue }
        if ($answer -eq 'a') { if ($view.Count) { $view; return }; continue }
        $sel = ConvertFrom-TarSelectionInput -Text $answer -Count $view.Count
        if ($null -eq $sel -or $sel.Count -eq 0) { Write-Host "Invalid choice: '$answer'" -ForegroundColor Yellow; continue }
        foreach ($ix in $sel) { $view[$ix] }
        return
    }
}

function Select-TarItem {
    # Rows must have an Id property. Emits the selected rows to the pipeline.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseCompatibleCmdlets', '', Justification = 'Out-GridView is only used when available (Resolve-TarSelectionMode checks $IsWindows and Get-Command).')]
    [CmdletBinding()]
    [OutputType([psobject])]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string[]]$Columns, [string]$Mode = 'Auto')
    if ($Rows.Count -eq 0) { return }
    $resolved = Resolve-TarSelectionMode -Mode $Mode
    Write-TarLog "Selection '$Title' via $resolved ($($Rows.Count) options)" -Level DEBUG
    $picked = switch ($resolved) {
        'Gui' { $Rows | Select-Object -Property (@($Columns) + 'Id') | Out-GridView -Title "$Title (Ctrl/Shift for multi-select)" -OutputMode Multiple }
        'ConsoleGrid' { $Rows | Select-Object -Property (@($Columns) + 'Id') | Out-ConsoleGridView -Title $Title -OutputMode Multiple }
        default { Select-TarFromMenu -Rows $Rows -Title $Title -Columns $Columns }
    }
    $ids = @($picked | ForEach-Object { $_.Id })
    $Rows | Where-Object { $_.Id -in $ids }
}

function Read-TarYesNo {
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Prompt, [bool]$Default = $true)
    $suffix = if ($Default) { '[Y/n]' } else { '[y/N]' }
    while ($true) {
        $a = (Read-Host "$Prompt $suffix").Trim().ToLowerInvariant()
        if (-not $a) { return $Default }
        if ($a -in 'y', 'yes') { return $true }
        if ($a -in 'n', 'no') { return $false }
    }
}

function Get-TarTeamSelection {
    # Returns @{ Teams = [hashtable[]]; ChannelFilter = @{ TeamId = HashSet[string] } } or $null if the user cancels.
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)]$Context, [string[]]$TeamId, [string[]]$ChannelId, [switch]$AllTeams,
        [switch]$Interactive, [switch]$SelectChannels, [switch]$Force, [string]$SelectionMode = 'Auto')
    $base = $Context.GraphBase
    $filter = @{}
    if ($TeamId) {
        $teams = @(foreach ($id in $TeamId) { Invoke-TarGraphRequest -Context $Context -Uri "$base/teams/$id`?`$select=id,displayName" })
    } else {
        Write-TarLog "Retrieving all teams in tenant $($Context.TenantId) ..."
        $discovered = @(Get-TarGraphCollection -Context $Context -Uri "$base/teams?`$select=id,displayName,description,visibility")
        Write-TarLog "Found $($discovered.Count) teams."
        if ($discovered.Count -eq 0) { Write-TarLog 'No teams found. Check Team.ReadBasic.All.' -Level WARN; return $null }
        if ($AllTeams) { $teams = $discovered }
        else {
            $rows = @($discovered | Sort-Object { $_.displayName } | ForEach-Object {
                    $d = [string]$_.description
                    [pscustomobject]@{ Name = $_.displayName; Visibility = $_.visibility; Description = if ($d.Length -gt 60) { $d.Substring(0, 57) + '...' } else { $d }; Id = $_.id } })
            $ids = @(Select-TarItem -Rows $rows -Title "Select teams ($($Context.TenantName ?? $Context.TenantId))" -Columns 'Name', 'Visibility', 'Description' -Mode $SelectionMode | ForEach-Object { $_.Id })
            if ($ids.Count -eq 0) { Write-TarLog 'No teams selected - exiting.' -Level WARN; return $null }
            $teams = @($discovered | Where-Object { $_.id -in $ids })
        }
    }
    if ($ChannelId) { foreach ($t in $teams) { $filter[$t.id] = [System.Collections.Generic.HashSet[string]]::new([string[]]$ChannelId) } }

    $pick = $SelectChannels -or ($Interactive -and -not $Force -and (Read-TarYesNo 'Do you want to select specific channels per team? (No = all channels)' $false))
    if ($pick) {
        foreach ($t in $teams) {
            try {
                $chs = @(Get-TarGraphCollection -Context $Context -Uri "$base/teams/$($t.id)/channels?`$select=id,displayName,membershipType")
                $rows = @($chs | Sort-Object { $_.membershipType }, { $_.displayName } | ForEach-Object { [pscustomobject]@{ Channel = $_.displayName; Type = $_.membershipType; Id = $_.id } })
                $ids = [string[]]@(Select-TarItem -Rows $rows -Title "Channels in '$($t.displayName)'" -Columns 'Channel', 'Type' -Mode $SelectionMode | ForEach-Object { $_.Id })
                $filter[$t.id] = [System.Collections.Generic.HashSet[string]]::new($ids)
                Write-TarLog "    '$($t.displayName)': $($ids.Count) of $($chs.Count) channels selected"
            } catch { Write-TarLog "Channel selection failed for '$($t.displayName)': $($_.Exception.Message) - all channels are included." -Level WARN }
        }
    }
    @{ Teams = $teams; ChannelFilter = $filter }
}

function Save-TarSelection {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][hashtable]$Selection)
    $teams = $Selection.Teams; $filter = $Selection.ChannelFilter
    $rerun = "Invoke-TeamsAccessReport -TenantId '$($Context.TenantId)' -TeamId " + ((@($teams | ForEach-Object { "'$($_.id)'" })) -join ',')
    if ($filter.Count) { $rerun += ' -ChannelId ' + ((@($filter.Values | ForEach-Object { $_ } | Select-Object -Unique | ForEach-Object { "'$_'" })) -join ',') }
    if ($Context.AppOnly) { $rerun += " -ClientId '$($Context.ClientId)' -CertificateThumbprint '$($Context.CertificateThumbprint)'" }
    [pscustomobject]@{
        SchemaVersion = 1; TenantId = $Context.TenantId; TenantName = $Context.TenantName; TenantDomain = $Context.TenantDomain; RunAt = (Get-Date).ToString('o')
        Teams = @($teams | ForEach-Object { [pscustomobject]@{ Id = $_.id; Name = $_.displayName; ChannelIds = if ($filter.ContainsKey($_.id)) { @($filter[$_.id]) } else { 'ALL' } } })
        RerunCommand = $rerun
    } | ConvertTo-Json -Depth 5 | Set-Content (Get-TarStorePath $Context $script:TarFile.Selection) -Encoding utf8
}

function Import-TarSelection {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)]$Context)
    $sel = Get-Content (Get-TarStorePath $Context $script:TarFile.Selection) -Raw | ConvertFrom-Json
    if ($sel.TenantId -ne $Context.TenantId) { throw "Selection.json belongs to tenant $($sel.TenantId); the signed-in tenant is $($Context.TenantId)." }
    $filter = @{}
    foreach ($t in $sel.Teams) { if ($t.ChannelIds -ne 'ALL') { $filter[$t.Id] = [System.Collections.Generic.HashSet[string]]::new([string[]]@($t.ChannelIds)) } }
    @{ Teams = @($sel.Teams | ForEach-Object { @{ id = $_.Id; displayName = $_.Name } }); ChannelFilter = $filter }
}
