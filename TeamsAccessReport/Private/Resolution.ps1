# Phase 3 - resolves SharePoint principals (claims) to Entra users. Cache per run.

function New-TarResolverCache {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory constructor; does not change system state.')]
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    @{ Users = @{}; Groups = @{} }
}

function Resolve-TarUserByUpn {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][hashtable]$Cache, [string]$Upn)
    if (-not $Upn) { return $null }
    $key = $Upn.ToLowerInvariant()
    if ($Cache.Users.ContainsKey($key)) { return $Cache.Users[$key] }
    $u = $null
    try { $u = Invoke-TarGraphRequest -Context $Context -AllowNotFound -Uri ("$($Context.GraphBase)/users/{0}?`$select=id,displayName,userPrincipalName,userType" -f [uri]::EscapeDataString($Upn)) }
    catch { Add-TarErrorRecord -Context $Context -Stage 'ResolveUser' -Target $Upn -Message $_.Exception.Message }
    $Cache.Users[$key] = $u
    $u
}

function Get-TarGroupUser {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][hashtable]$Cache, [Parameter(Mandatory)][string]$GroupId,
        [Parameter(Mandatory)][ValidateSet('members', 'owners')][string]$Relation)
    $key = "$GroupId|$Relation"
    if (-not $Cache.Groups.ContainsKey($key)) {
        $rel = if ($Relation -eq 'owners') { 'owners' } else { 'transitiveMembers' }
        $users = @()
        try {
            $users = @(Get-TarGraphCollection -Context $Context -Uri "$($Context.GraphBase)/groups/$GroupId/$rel`?`$select=id,displayName,userPrincipalName,userType" |
                    Where-Object { (Get-TarProperty $_ '@odata.type') -eq '#microsoft.graph.user' })
        } catch { Add-TarErrorRecord -Context $Context -Stage 'ResolveGroup' -Target $key -Message $_.Exception.Message }
        $Cache.Groups[$key] = $users
    }
    $Cache.Groups[$key]
}

function New-TarPrincipal {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Pure in-memory constructor; does not change system state.')]
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([Parameter(Mandatory)][string]$Kind, $UserId, $Upn, $DisplayName, $UserType, [string]$Path)
    [pscustomobject]@{ Kind = $Kind; UserId = $UserId; Upn = $Upn; DisplayName = $DisplayName; UserType = $UserType; Path = $Path }
}

function Resolve-TarSpPrincipal {
    # Claim formats: i:0#.f = user, c:0o.c = Microsoft 365 group (_o suffix = owners), c:0t.c = security group.
    # Emits one or more principal records to the pipeline.
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][hashtable]$Cache,
        [AllowEmptyString()][string]$Login, [string]$Title, [string]$Email, [string]$PathPrefix = 'Direct')
    switch -Regex ($Login) {
        '^i:0#\.f\|membership\|(.+)$' {
            $upn = $Matches[1]; $u = Resolve-TarUserByUpn -Context $Context -Cache $Cache -Upn $upn
            New-TarPrincipal -Kind User -UserId (Get-TarProperty $u 'id') -Upn ((Get-TarProperty $u 'userPrincipalName') ?? $upn) `
                -DisplayName ((Get-TarProperty $u 'displayName') ?? $Title) -UserType (Get-TarProperty $u 'userType') -Path $PathPrefix
            return
        }
        'federateddirectoryclaimprovider\|([0-9a-fA-F-]{36})(_o)?$' {
            $gid = $Matches[1]; $rel = if ($Matches[2]) { 'owners' } else { 'members' }
            foreach ($u in @(Get-TarGroupUser -Context $Context -Cache $Cache -GroupId $gid -Relation $rel)) {
                New-TarPrincipal -Kind User -UserId $u.id -Upn $u.userPrincipalName -DisplayName $u.displayName -UserType (Get-TarProperty $u 'userType') -Path "$PathPrefix > M365Group:$gid/$rel"
            }
            return
        }
        '^c:0t\.c\|tenant\|([0-9a-fA-F-]{36})$' {
            $gid = $Matches[1]
            foreach ($u in @(Get-TarGroupUser -Context $Context -Cache $Cache -GroupId $gid -Relation members)) {
                New-TarPrincipal -Kind User -UserId $u.id -Upn $u.userPrincipalName -DisplayName $u.displayName -UserType (Get-TarProperty $u 'userType') -Path "$PathPrefix > SecurityGroup:$gid"
            }
            return
        }
        'spo-grid-all-users' { New-TarPrincipal -Kind BroadGrant -UserId '*' -Upn 'EveryoneExceptExternalUsers' -DisplayName $Title -Path $PathPrefix; return }
        '^c:0\(\.s\|true$' { New-TarPrincipal -Kind BroadGrant -UserId '*' -Upn 'Everyone' -DisplayName $Title -Path $PathPrefix; return }
        # Token principal of a sharing link ('SLinkClaim.*'); the access itself is reported via its SharingLinks group.
        '^c:0u\.c\|tenant\|' { New-TarPrincipal -Kind SharingLinkClaim -Upn $Login -DisplayName $Title -Path $PathPrefix; return }
        '^SHAREPOINT\\system$|^i:0i\.t\|ms\.sp\.ext\|' { New-TarPrincipal -Kind SystemOrApp -Upn $Login -DisplayName $Title -Path $PathPrefix; return }
    }
    $kind = if ($Title -match '^SharingLinks\.[^.]+\.Anonymous\w*\.') { 'AnonymousLink' } elseif ($Title -match '^SharingLinks\.[^.]+\.Organization\w*\.') { 'OrganizationLink' } else { 'Unresolved' }
    New-TarPrincipal -Kind $kind -Upn ($Email ? $Email : $Login) -DisplayName $Title -Path $PathPrefix
}

function Invoke-TarPrincipalResolution {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context)
    $assignments = @(Read-TarRecord -Context $Context -Name $script:TarFile.SpAssignments)
    if ($assignments.Count -eq 0) { Set-TarRecord -Context $Context -Name $script:TarFile.EffectiveSpAccess -InputObject @(); return }
    Write-TarLog "Phase 3: resolving $($assignments.Count) SharePoint role assignments"
    $cache = New-TarResolverCache
    $groupIndex = @{}
    foreach ($gm in @(Read-TarRecord -Context $Context -Name $script:TarFile.SpGroupMembers)) {
        $k = "$($gm.SiteUrl)|$($gm.GroupTitle)"
        if (-not $groupIndex.ContainsKey($k)) { $groupIndex[$k] = [System.Collections.Generic.List[object]]::new() }
        $groupIndex[$k].Add($gm)
    }
    $effective = [System.Collections.Generic.List[object]]::new()
    $n = 0
    foreach ($a in $assignments) {
        $n++
        if ($n % 500 -eq 0) { Write-TarLog "    $n/$($assignments.Count) resolved" }
        $principals = if ([string]$a.PrincipalType -eq '8') {
            $prefix = "SPGroup:$($a.PrincipalTitle)"
            $isLink = $a.PrincipalTitle -match '^SharingLinks\.'
            if ($isLink -and $a.PrincipalTitle -match '\.(Anonymous|Organization)\w*\.') {
                Resolve-TarSpPrincipal -Context $Context -Cache $cache -Login $a.PrincipalLogin -Title $a.PrincipalTitle -PathPrefix $prefix
            }
            $k = "$($a.SiteUrl)|$($a.PrincipalTitle)"
            if ($groupIndex.ContainsKey($k)) {
                foreach ($gm in $groupIndex[$k]) { Resolve-TarSpPrincipal -Context $Context -Cache $cache -Login $gm.MemberLogin -Title $gm.MemberTitle -Email $gm.MemberEmail -PathPrefix $prefix }
            } elseif (-not $isLink) { New-TarPrincipal -Kind EmptySPGroup -DisplayName $a.PrincipalTitle -Path $prefix }
        } else {
            Resolve-TarSpPrincipal -Context $Context -Cache $cache -Login $a.PrincipalLogin -Title $a.PrincipalTitle -Email $a.PrincipalEmail
        }
        foreach ($p in @($principals)) {
            $effective.Add([pscustomobject]@{
                    SiteUrl = $a.SiteUrl; WebUrl = $a.WebUrl; Scope = $a.Scope; ObjectPath = $a.ObjectPath; ListTitle = $a.ListTitle; ItemId = $a.ItemId
                    Roles = $a.Roles; LimitedAccessOnly = $a.LimitedAccessOnly; AssociatedGroup = $a.AssociatedGroup; Kind = $p.Kind; UserId = $p.UserId; Upn = $p.Upn
                    DisplayName = $p.DisplayName; UserType = $p.UserType; ExpansionPath = $p.Path
                })
        }
    }
    Set-TarRecord -Context $Context -Name $script:TarFile.EffectiveSpAccess -InputObject $effective
    Write-TarLog "    Effective SharePoint rows: $($effective.Count)"
}
