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
        [Parameter(Mandatory)][string[]]$Columns, [string]$Mode = 'Auto', [switch]$Single, [string[]]$CheckedId = @())
    if ($Rows.Count -eq 0) { return }
    $resolved = Resolve-TarSelectionMode -Mode $Mode
    Write-TarLog "Selection '$Title' via $resolved ($($Rows.Count) options)" -Level DEBUG
    $picked = switch ($resolved) {
        'Gui' {
            # Multi-select: own check-box list (discoverable, filter, select all). Single: Out-GridView.
            if (-not $Single -and (Test-TarGuiAvailable)) { return (Show-TarCheckList -Rows $Rows -Title $Title -Columns $Columns -CheckedId $CheckedId) }
            $hint = if ($Single) { '' } else { ' (Ctrl/Shift for multi-select)' }
            $Rows | Select-Object -Property (@($Columns) + 'Id') | Out-GridView -Title "$Title$hint" -OutputMode ($Single ? 'Single' : 'Multiple')
        }
        'ConsoleGrid' { $Rows | Select-Object -Property (@($Columns) + 'Id') | Out-ConsoleGridView -Title $Title -OutputMode ($Single ? 'Single' : 'Multiple') }
        default { Select-TarFromMenu -Rows $Rows -Title $Title -Columns $Columns }
    }
    $ids = @($picked | ForEach-Object { $_.Id })
    $Rows | Where-Object { $_.Id -in $ids }
}

function Read-TarYesNo {
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory)][string]$Prompt, [bool]$Default = $true, [string]$Mode = 'Menu')
    # Gui mode: Yes/No pop-up instead of console input.
    if ((Resolve-TarSelectionMode -Mode $Mode) -eq 'Gui' -and (Test-TarGuiAvailable)) { return (Show-TarMessageBox -Text $Prompt -Default $Default) }
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
        [switch]$Interactive, [switch]$SelectChannels, [switch]$Force, [string]$SelectionMode = 'Auto', [switch]$NoChannelPrompt)
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

    # -NoChannelPrompt: the start screen already answered this question.
    $pick = $SelectChannels -or ($Interactive -and -not $Force -and -not $NoChannelPrompt -and (Read-TarYesNo 'Do you want to select specific channels per team? (No = all channels)' $false -Mode $SelectionMode))
    if ($pick) {
        foreach ($t in $teams) {
            try {
                $chs = @(Get-TarGraphCollection -Context $Context -Uri "$base/teams/$($t.id)/channels?`$select=id,displayName,membershipType")
                $rows = @($chs | Sort-Object { $_.membershipType }, { $_.displayName } | ForEach-Object { [pscustomobject]@{ Channel = $_.displayName; Type = $_.membershipType; Id = $_.id } })
                $ids = [string[]]@(Select-TarItem -Rows $rows -Title "Channels in '$($t.displayName)'" -Columns 'Channel', 'Type' -Mode $SelectionMode -CheckedId @($rows.Id) | ForEach-Object { $_.Id })
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
    $connection = if ($Context.ProfileName) { "-ProfileName '$($Context.ProfileName)'" } else { "-TenantId '$($Context.TenantId)'" }
    $rerun = "Invoke-TeamsAccessReport $connection -TeamId " + ((@($teams | ForEach-Object { "'$($_.id)'" })) -join ',')
    if ($filter.Count) { $rerun += ' -ChannelId ' + ((@($filter.Values | ForEach-Object { $_ } | Select-Object -Unique | ForEach-Object { "'$_'" })) -join ',') }
    if ($Context.AppOnly -and -not $Context.ProfileName) { $rerun += " -ClientId '$($Context.ClientId)' -CertificateThumbprint '$($Context.CertificateThumbprint)'" }
    [pscustomobject]@{
        SchemaVersion = $script:TarSelectionSchemaVersion; TenantId = $Context.TenantId; TenantName = $Context.TenantName; TenantDomain = $Context.TenantDomain; ProfileName = $Context.ProfileName; RunAt = (Get-Date).ToString('o')
        Teams = @($teams | ForEach-Object {
                $all = -not $filter.ContainsKey($_.id)
                # Assigned via a typed variable: an if expression would unroll a one-element array into a scalar.
                $ids = [string[]]@(if (-not $all) { $filter[$_.id] | Sort-Object })
                [pscustomobject]@{ Id = $_.id; Name = $_.displayName; AllChannels = $all; ChannelIds = $ids }
            })
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
    if ([int]$sel.SchemaVersion -gt $script:TarSelectionSchemaVersion) { throw "Selection.json schema $($sel.SchemaVersion) is newer than this module supports ($script:TarSelectionSchemaVersion). Update the module." }
    foreach ($t in $sel.Teams) {
        # Schema 2: AllChannels + ChannelIds array. Schema 1: ChannelIds = 'ALL', a single id or an array of ids.
        $all = if ($t.PSObject.Properties['AllChannels']) { [bool]$t.AllChannels } else { [string]$t.ChannelIds -eq 'ALL' }
        if (-not $all) { $filter[$t.Id] = [System.Collections.Generic.HashSet[string]]::new([string[]]@($t.ChannelIds | Where-Object { $_ })) }
    }
    @{ Teams = @($sel.Teams | ForEach-Object { @{ id = $_.Id; displayName = $_.Name } }); ChannelFilter = $filter }
}


# --- Connection prompts (data logic in Connection.ps1) ----------------------------------------------------

function Test-TarCanPrompt {
    # True when a person can answer prompts. Single seam so tests never block on Read-Host.
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    -not [Console]::IsInputRedirected -and [Environment]::UserInteractive
}

function Read-TarTenantInput {
    # Asks for a tenant until it resolves. Accepts a name (contoso), a domain, an e-mail address or a tenant ID.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive prompt feedback requires direct console output.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([switch]$AllowEmpty)
    while ($true) {
        $hint = if ($AllowEmpty) { 'Enter = choose at sign-in, ' } else { '' }
        $answer = (Read-Host "Tenant (e.g. contoso, contoso.com, admin@contoso.com or tenant ID; ${hint}q = cancel)").Trim()
        if ($answer -eq 'q') { return }
        if (-not $answer) { if ($AllowEmpty) { return [pscustomobject]@{ TenantId = $null; Query = $null; Region = $null; Cloud = $null } }; continue }
        try {
            $t = Resolve-TarTenantOrPassThrough -Tenant $answer
            Write-Host "  -> $($t.Query) = tenant $($t.TenantId)$(if ($t.Region) { " (region $($t.Region))" })" -ForegroundColor Green
            return $t
        } catch { Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}

function Read-TarGuidInput {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive prompt feedback requires direct console output.')]
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Prompt)
    while ($true) {
        $answer = (Read-Host "$Prompt (q = cancel)").Trim()
        if ($answer -eq 'q') { return }
        try { return (ConvertTo-TarClientId $answer) } catch { Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}

function Select-TarCertificate {
    # Pop-up/menu over usable certificates; falls back to typing a thumbprint when none are found.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive prompt feedback requires direct console output.')]
    [CmdletBinding()]
    [OutputType([string])]
    param([string]$SelectionMode = 'Auto')
    $rows = @(Get-TarCertificateCandidate | Where-Object { $_.HasPrivateKey -and $_.Valid })
    while ($rows.Count) {
        $picked = @(Select-TarItem -Rows $rows -Title 'Certificate for app-only sign-in (pick one)' -Columns 'Subject', 'Expires', 'Store' -Mode $SelectionMode -Single)
        if ($picked.Count -eq 1) { Write-Host "  -> $($picked[0].Subject) ($($picked[0].Id))" -ForegroundColor Green; return $picked[0].Id }
        if ($picked.Count -eq 0) { break }
        Write-Host '  Pick exactly one certificate.' -ForegroundColor Yellow
    }
    if (-not $rows.Count) { Write-Host '  No certificate with a private key was found in CurrentUser\My or LocalMachine\My.' -ForegroundColor Yellow }
    while ($true) {
        $answer = (Read-Host 'Certificate thumbprint (q = cancel)').Trim()
        if ($answer -eq 'q') { return }
        try { return (ConvertTo-TarThumbprint $answer) } catch { Write-Host "  $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}

function Read-TarConnectionWizard {
    # Creates a connection (and optionally saves it as a profile). Returns a profile record, or $null when cancelled.
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Interactive prompt feedback requires direct console output.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$Name, [string]$Tenant, [ValidateSet('', 'AppOnly', 'Delegated')][string]$AuthMode = '', [string]$ClientId,
        [string]$CertificateThumbprint, [string]$PnPClientId, [string]$SelectionMode = 'Auto', [switch]$OfferSave)
    Write-Host ''
    Write-Host '=== New connection ===' -ForegroundColor Cyan
    $t = if ($Tenant) { Resolve-TarTenantOrPassThrough -Tenant $Tenant } else { Read-TarTenantInput }
    if (-not $t) { return }
    if (-not $AuthMode) {
        $modes = @(
            [pscustomobject]@{ Id = 'AppOnly'; Mode = 'App-only (certificate)'; Note = 'Complete picture; needs an Entra app with application permissions' }
            [pscustomobject]@{ Id = 'Delegated'; Mode = 'Delegated (sign in as yourself)'; Note = 'Limited by your own role' }
        )
        $m = @(Select-TarItem -Rows $modes -Title 'Sign-in method (pick one)' -Columns 'Mode', 'Note' -Mode $SelectionMode -Single)
        if ($m.Count -ne 1) { return }
        $AuthMode = $m[0].Id
    }
    if ($AuthMode -eq 'AppOnly') {
        if (-not $ClientId) { $ClientId = Read-TarGuidInput -Prompt 'Application (client) ID of the Entra app'; if (-not $ClientId) { return } }
        if (-not $CertificateThumbprint) { $CertificateThumbprint = Select-TarCertificate -SelectionMode $SelectionMode; if (-not $CertificateThumbprint) { return } }
    }
    if (-not $Name) { $Name = (($t.Query ?? 'tenant') -split '\.')[0] }
    $record = New-TarProfileRecord -Name $Name -TenantId $t.TenantId -TenantQuery $t.Query -AuthMode $AuthMode -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -PnPClientId $PnPClientId
    if ($OfferSave -and (Read-TarYesNo "Save as profile '$Name' for next time (no secrets are stored)?")) {
        $custom = (Read-Host "Profile name [Enter = $Name]").Trim()
        if ($custom) { $record = New-TarProfileRecord -Name $custom -TenantId $t.TenantId -TenantQuery $t.Query -AuthMode $AuthMode -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -PnPClientId $PnPClientId }
        Save-TarProfile -ProfileRecord $record
        Write-Host "  Next time: Invoke-TeamsAccessReport -ProfileName $($record.Name)" -ForegroundColor Green
    }
    $record
}

function Select-TarConnection {
    # Start screen when no connection parameters were given: saved profiles, new connection, or delegated.
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string]$SelectionMode = 'Auto')
    $saved = @(Read-TarProfileStore)
    $rows = @($saved | ForEach-Object { [pscustomobject]@{ Id = "profile:$($_.Name)"; Connection = $_.Name; Tenant = ($_.TenantQuery ?? $_.TenantId); Method = $_.AuthMode } })
    $rows += [pscustomobject]@{ Id = 'new'; Connection = '<New connection...>'; Tenant = 'enter tenant, app and certificate once'; Method = '' }
    $rows += [pscustomobject]@{ Id = 'delegated'; Connection = '<Sign in as myself>'; Tenant = 'delegated, nothing saved'; Method = 'Delegated' }
    # @() around the whole if: an if expression unrolls an empty array to $null (strict mode then fails on .Count).
    $pick = @(if ($saved.Count -eq 0) { $rows | Where-Object Id -eq 'new' } else { Select-TarItem -Rows $rows -Title 'Connection (pick one)' -Columns 'Connection', 'Tenant', 'Method' -Mode $SelectionMode -Single })
    if ($pick.Count -ne 1) { Write-TarLog 'No connection selected.' -Level WARN; return }
    switch -Regex ($pick[0].Id) {
        '^profile:(.+)$' { return ConvertFrom-TarProfile (Get-TarProfile -Name $Matches[1]) }
        '^new$' {
            $r = Read-TarConnectionWizard -SelectionMode $SelectionMode -OfferSave
            if (-not $r) { return }
            $c = ConvertFrom-TarProfile $r
            # Not saved -> no profile name (so Selection.json does not point at a profile that does not exist).
            if (@(Read-TarProfileStore | Where-Object Name -eq $r.Name).Count -eq 0) { $c.ProfileName = $null }
            return $c
        }
        default {
            $t = Read-TarTenantInput -AllowEmpty
            if (-not $t) { return }
            return [pscustomobject]@{ AuthMode = 'Delegated'; TenantId = $t.TenantId; ClientId = $null; CertificateThumbprint = $null; PnPClientId = $null; ProfileName = $null }
        }
    }
}
