# Phase 2 - SharePoint scan via PnP. Runs in a separate pwsh process (Worker/Invoke-SharePointScan.ps1)
# because PnP.PowerShell and the Microsoft Graph SDK load incompatible assemblies in the same session.
# Invoke-TarSpRaw is the ONLY place that touches Invoke-PnPSPRestMethod (test seam).

$script:TarSpRoleExpand = 'roleassignments?$expand=Member,RoleDefinitionBindings&$select=Member/Title,Member/LoginName,Member/Email,Member/PrincipalType,RoleDefinitionBindings/Name'

function Invoke-TarSpRaw {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)]$Connection)
    Invoke-PnPSPRestMethod -Url $Url -Connection $Connection -ErrorAction Stop
}

function Invoke-TarSpRequest {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)]$Connection)
    for ($attempt = 1; $attempt -le $Context.MaxRetries; $attempt++) {
        try {
            Write-TarLog "SP GET $Url ($attempt/$($Context.MaxRetries))" -Level DEBUG
            return Invoke-TarSpRaw -Url $Url -Connection $Connection
        } catch {
            $status = Get-TarHttpStatusFromError $_
            if ($status -notin $script:TarRetryableStatus -or $attempt -eq $Context.MaxRetries) { throw }
            $wait = Get-TarBackoffDelay -Attempt $attempt -BaseDelaySeconds $Context.BaseDelaySeconds
            Write-TarLog "SP HTTP $status ($attempt/$($Context.MaxRetries)); waiting $wait ms" -Level WARN -Data @{ url = $Url }
            Start-Sleep -Milliseconds $wait
        }
    }
}

function Get-TarSpCollection {
    # Follows odata.nextLink (nometadata) and d.__next (verbose). Emits elements to the pipeline.
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)]$Connection, [int]$Max = 0)
    $next = $Url; $count = 0
    while ($next) {
        $r = Invoke-TarSpRequest -Context $Context -Url $next -Connection $Connection
        $d = Get-TarProperty $r 'd'
        $values = Get-TarProperty $r 'value'
        if ($null -eq $values -and $d) { $values = Get-TarProperty $d 'results' }
        foreach ($v in @($values)) {
            if ($null -eq $v) { continue }
            $v; $count++
            if ($Max -gt 0 -and $count -ge $Max) { Write-TarLog "Max $Max items reached for $Url (see -MaxItemsPerList)." -Level WARN; return }
        }
        $next = Get-TarProperty $r 'odata.nextLink'
        if (-not $next -and $d) { $next = Get-TarProperty $d '__next' }
    }
}

function ConvertTo-TarSpAssignmentRecord {
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)]$RoleAssignment, [Parameter(Mandatory)][string]$SiteUrl, [Parameter(Mandatory)][string]$WebUrl,
        [Parameter(Mandatory)][string]$Scope, [string]$ObjectPath, [string]$ListTitle, [string]$ItemId)
    process {
        $m = Get-TarProperty $RoleAssignment 'Member'
        $roles = @(ConvertTo-TarArray (Get-TarProperty $RoleAssignment 'RoleDefinitionBindings') | ForEach-Object { Get-TarProperty $_ 'Name' } | Where-Object { $_ })
        $effective = @($roles | Where-Object { $_ -notin $script:TarLimitedAccessRoles })
        [pscustomobject]@{
            SiteUrl = $SiteUrl; WebUrl = $WebUrl; Scope = $Scope; ObjectPath = $ObjectPath; ListTitle = $ListTitle; ItemId = $ItemId
            PrincipalType = Get-TarProperty $m 'PrincipalType'; PrincipalTitle = Get-TarProperty $m 'Title'
            PrincipalLogin = Get-TarProperty $m 'LoginName'; PrincipalEmail = Get-TarProperty $m 'Email'
            Roles = ($roles -join ';'); LimitedAccessOnly = ($roles.Count -gt 0 -and $effective.Count -eq 0); AssociatedGroup = $null
        }
    }
}

function Get-TarSpRoleAssignment {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$SecurableUrl, [Parameter(Mandatory)]$Connection)
    Get-TarSpCollection -Context $Context -Url "$SecurableUrl/$script:TarSpRoleExpand" -Connection $Connection
}

function Invoke-TarSpWebScan {
    # Web -> lists -> items (5000/page) -> subwebs (recursive). Emits assignment records to the pipeline.
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$SiteUrl, [Parameter(Mandatory)][string]$WebUrl,
        [Parameter(Mandatory)]$Connection, [int]$Depth = 0)
    $web = Invoke-TarSpRequest -Context $Context -Url "$WebUrl/_api/web?`$select=Url,HasUniqueRoleAssignments" -Connection $Connection
    if ($Depth -eq 0 -or (Get-TarProperty $web 'HasUniqueRoleAssignments')) {
        Get-TarSpRoleAssignment -Context $Context -SecurableUrl "$WebUrl/_api/web" -Connection $Connection |
            ConvertTo-TarSpAssignmentRecord -SiteUrl $SiteUrl -WebUrl $WebUrl -Scope 'Web' -ObjectPath $WebUrl
    }
    $lists = @(Get-TarSpCollection -Context $Context -Connection $Connection `
            -Url "$WebUrl/_api/web/lists?`$select=Id,Title,Hidden,ItemCount,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl&`$expand=RootFolder&`$filter=Hidden eq false")
    foreach ($l in $lists) {
        $listId = Get-TarProperty $l 'Id'; $title = Get-TarProperty $l 'Title'
        $listUrl = "$WebUrl/_api/web/lists(guid'$listId')"
        $path = Get-TarProperty (Get-TarProperty $l 'RootFolder') 'ServerRelativeUrl'
        try {
            if (Get-TarProperty $l 'HasUniqueRoleAssignments') {
                Get-TarSpRoleAssignment -Context $Context -SecurableUrl $listUrl -Connection $Connection |
                    ConvertTo-TarSpAssignmentRecord -SiteUrl $SiteUrl -WebUrl $WebUrl -Scope 'List' -ObjectPath $path -ListTitle $title
            }
            if ([int](Get-TarProperty $l 'ItemCount') -le 0) { continue }
            $items = @(Get-TarSpCollection -Context $Context -Connection $Connection -Max $Context.MaxItemsPerList `
                    -Url "$listUrl/items?`$select=Id,FileRef,HasUniqueRoleAssignments&`$top=5000")
            $unique = @($items | Where-Object { Get-TarProperty $_ 'HasUniqueRoleAssignments' })
            Write-TarLog "    List '$title': $($items.Count) items, $($unique.Count) with unique permissions"
            foreach ($it in $unique) {
                $itemId = [string](Get-TarProperty $it 'Id')
                Get-TarSpRoleAssignment -Context $Context -SecurableUrl "$listUrl/items($itemId)" -Connection $Connection |
                    ConvertTo-TarSpAssignmentRecord -SiteUrl $SiteUrl -WebUrl $WebUrl -Scope 'Item' -ObjectPath (Get-TarProperty $it 'FileRef') -ListTitle $title -ItemId $itemId
            }
        } catch {
            Write-TarLog "    List '$title' failed: $($_.Exception.Message)" -Level ERROR
            Add-TarErrorRecord -Context $Context -Stage 'SPList' -Target "$WebUrl :: $title" -Message $_.Exception.Message
        }
    }
    if ($Depth -ge $Context.MaxSubwebDepth) { Write-TarLog "    Max depth $($Context.MaxSubwebDepth) reached at $WebUrl" -Level WARN; return }
    foreach ($s in @(Get-TarSpCollection -Context $Context -Url "$WebUrl/_api/web/webs?`$select=Url" -Connection $Connection)) {
        $sub = Get-TarProperty $s 'Url'
        Write-TarLog "    Subweb (depth $($Depth + 1)): $sub"
        Invoke-TarSpWebScan -Context $Context -SiteUrl $SiteUrl -WebUrl $sub -Connection $Connection -Depth ($Depth + 1)
    }
}

function Get-TarSpAssociatedGroup {
    # Group title -> 'Owner' | 'Member' | 'Visitor' for the site's associated (default) groups.
    # Best effort: a failure only means drift findings cannot be attributed to a default, so it is logged, not fatal.
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$SiteUrl, [Parameter(Mandatory)]$Connection)
    $map = @{}
    $nav = 'AssociatedOwnerGroup', 'AssociatedMemberGroup', 'AssociatedVisitorGroup'
    $url = "$SiteUrl/_api/web?`$select=" + ((@($nav | ForEach-Object { "$_/Title" })) -join ',') + '&$expand=' + ($nav -join ',')
    try {
        $web = Invoke-TarSpRequest -Context $Context -Url $url -Connection $Connection
        foreach ($n in $nav) {
            $title = Get-TarProperty (Get-TarProperty $web $n) 'Title'
            if ($title) { $map[[string]$title] = $n -replace '^Associated(\w+)Group$', '$1' }
        }
    } catch {
        Write-TarLog "    Associated groups unavailable for $SiteUrl ($($_.Exception.Message)); drift Cause will be empty for this site." -Level WARN
    }
    $map
}

function Invoke-TarSpSiteScan {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$SiteUrl, [Parameter(Mandatory)]$Connection)
    $records = [System.Collections.Generic.List[object]]::new()
    foreach ($a in @(Get-TarSpCollection -Context $Context -Connection $Connection -Url "$SiteUrl/_api/web/siteusers?`$filter=IsSiteAdmin eq true&`$select=Title,LoginName,Email,PrincipalType")) {
        $records.Add([pscustomobject]@{ SiteUrl = $SiteUrl; WebUrl = $SiteUrl; Scope = 'SiteCollectionAdmin'; ObjectPath = $SiteUrl; ListTitle = $null; ItemId = $null
                PrincipalType = Get-TarProperty $a 'PrincipalType'; PrincipalTitle = Get-TarProperty $a 'Title'; PrincipalLogin = Get-TarProperty $a 'LoginName'
                PrincipalEmail = Get-TarProperty $a 'Email'; Roles = 'SiteCollectionAdmin'; LimitedAccessOnly = $false; AssociatedGroup = $null
            })
    }
    $memberships = [System.Collections.Generic.List[object]]::new()
    $groups = @(Get-TarSpCollection -Context $Context -Connection $Connection -Url "$SiteUrl/_api/web/sitegroups?`$select=Id,Title,LoginName")
    foreach ($g in $groups) {
        $gid = Get-TarProperty $g 'Id'
        foreach ($u in @(Get-TarSpCollection -Context $Context -Connection $Connection -Url "$SiteUrl/_api/web/sitegroups/getbyid($gid)/users?`$select=Title,LoginName,Email,PrincipalType")) {
            $memberships.Add([pscustomobject]@{ SiteUrl = $SiteUrl; GroupTitle = Get-TarProperty $g 'Title'; MemberTitle = Get-TarProperty $u 'Title'
                    MemberLogin = Get-TarProperty $u 'LoginName'; MemberEmail = Get-TarProperty $u 'Email'; MemberPrincipalType = Get-TarProperty $u 'PrincipalType'
                })
        }
    }
    $associated = Get-TarSpAssociatedGroup -Context $Context -SiteUrl $SiteUrl -Connection $Connection
    foreach ($r in @(Invoke-TarSpWebScan -Context $Context -SiteUrl $SiteUrl -WebUrl $SiteUrl -Connection $Connection)) {
        # Associated groups are defined per web; only tag root-web assignments to SharePoint groups.
        if ($r.Scope -eq 'Web' -and [string]$r.PrincipalType -eq '8' -and $r.WebUrl -eq $SiteUrl -and $associated.ContainsKey([string]$r.PrincipalTitle)) {
            $r.AssociatedGroup = $associated[[string]$r.PrincipalTitle]
        }
        $records.Add($r)
    }
    [pscustomobject]@{ Assignments = $records; GroupMembers = $memberships; GroupCount = $groups.Count }
}

function Invoke-TarSharePointScan {
    # Worker entry point. Assumes PnP.PowerShell is loaded (or stubbed in tests).
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string[]]$Sites, [Parameter(Mandatory)][string]$ClientId,
        [string]$Thumbprint, [string]$Tenant)
    $done = Get-TarCheckpoint -Context $Context -Name $script:TarCheckpoint.SharePoint
    Write-TarLog "SharePoint scan: $($Sites.Count) sites, already done: $($done.Count)"
    $n = 0
    foreach ($site in $Sites) {
        $n++
        if ($done.Contains($site)) { Write-TarLog "[$n] Skipping (done): $site"; continue }
        Write-TarLog "[$n/$($Sites.Count)] Site $site"
        try {
            $connArgs = @{ Url = $site; ClientId = $ClientId; ReturnConnection = $true; ErrorAction = 'Stop' }
            if ($Thumbprint) { $connArgs.Thumbprint = $Thumbprint; $connArgs.Tenant = $Tenant }
            $conn = Connect-PnPOnline @connArgs
            $result = Invoke-TarSpSiteScan -Context $Context -SiteUrl $site -Connection $conn
            Write-TarRecord -Context $Context -Name $script:TarFile.SpAssignments -InputObject $result.Assignments
            Write-TarRecord -Context $Context -Name $script:TarFile.SpGroupMembers -InputObject $result.GroupMembers
            Add-TarCheckpoint -Context $Context -Name $script:TarCheckpoint.SharePoint -Key $site
            Write-TarLog "    Done: $($result.Assignments.Count) role assignments, $($result.GroupCount) groups, $($result.GroupMembers.Count) group memberships"
        } catch {
            Write-TarLog "    Site failed: $($_.Exception.Message)" -Level ERROR
            Add-TarErrorRecord -Context $Context -Stage 'SPSite' -Target $site -Message $_.Exception.Message
        }
    }
}

function Start-TarSharePointScan {
    # Starts phase 2 in a separate process (production) or in the same process (tests).
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$ClientId)
    $sites = [string[]]@(Read-TarRecord -Context $Context -Name $script:TarFile.Channels | Where-Object { $_.SiteUrl } | ForEach-Object { $_.SiteUrl.TrimEnd('/') } | Sort-Object -Unique)
    Write-TarLog "Phase 2: $($sites.Count) unique SharePoint sites"
    if ($sites.Count -eq 0 -or -not $PSCmdlet.ShouldProcess("$($sites.Count) sites", 'SharePoint scan')) { return }
    $tenant = if ($Context.InitialDomain) { $Context.InitialDomain } else { $Context.TenantId }
    $thumb = if ($Context.AppOnly) { $Context.CertificateThumbprint } else { $null }

    if ($script:TarSharePointExecutionMode -eq 'InProcess') {
        Invoke-TarSharePointScan -Context $Context -Sites $sites -ClientId $ClientId -Thumbprint $thumb -Tenant $tenant
        return
    }
    $cfgPath = Join-Path $Context.StateDir 'sp-worker.json'
    [pscustomobject]@{ Context = $Context; Sites = $sites; ClientId = $ClientId; Thumbprint = $thumb; Tenant = $tenant } |
        ConvertTo-Json -Depth 5 | Set-Content $cfgPath -Encoding utf8
    $pwsh = (Get-Process -Id $PID).Path
    $worker = Join-Path $PSScriptRoot '..' 'Worker' 'Invoke-SharePointScan.ps1'
    Write-TarLog "Starting SharePoint worker in a separate process: $pwsh"
    # The child process stdout/stderr must NEVER reach the output stream (it would pollute the function's return value).
    & $pwsh -NoProfile -File $worker -ConfigPath $cfgPath 2>&1 | ForEach-Object { Write-Information "[worker] $_" }
    if ($LASTEXITCODE -ne 0) {
        Write-TarLog "SharePoint worker exited with code $LASTEXITCODE - see run.jsonl. Re-run with -Resume." -Level ERROR
        Add-TarErrorRecord -Context $Context -Stage 'SPWorker' -Message "Exit code $LASTEXITCODE"
    }
}
