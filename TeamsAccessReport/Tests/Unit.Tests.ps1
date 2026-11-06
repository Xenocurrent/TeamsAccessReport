BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'TeamsAccessReport.psd1') -Force
}

Describe 'Module surface' {
    It 'exports the report command and the profile commands' {
        (Get-Command -Module TeamsAccessReport).Name | Sort-Object | Should -Be @('Get-TeamsAccessReportProfile', 'Invoke-TeamsAccessReport', 'New-TeamsAccessReportPnPApp', 'New-TeamsAccessReportProfile', 'Remove-TeamsAccessReportProfile')
    }
    It 'has no mandatory parameters (everything missing is asked for or resolved)' {
        @((Get-Command Invoke-TeamsAccessReport).Parameters.Values | Where-Object { $_.Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] -and $_.Mandatory } }).Count | Should -Be 0
    }
    It 'the worker process never references the Graph SDK (assembly conflict)' {
        foreach ($f in 'Worker/Invoke-SharePointScan.ps1', 'Private/SharePointScan.ps1', 'Private/Store.ps1', 'Private/Util.ps1', 'Private/Logging.ps1') {
            $tokens = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..' $f), [ref]$tokens, [ref]$null)
            @($tokens | Where-Object { $_.Kind -in 'Generic', 'Identifier' -and $_.Text -match 'MgGraph|-Mg[A-Z]' }).Count | Should -Be 0 -Because $f
        }
    }
}

Describe 'Get-TarProperty' {
    It 'reads <Case>' -ForEach @(
        @{ Case = 'Hashtable'; Factory = { @{ 'Retry-After' = '5' } } }
        @{ Case = 'generic Dictionary (Graph SDK response headers)'; Factory = { $d = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.IEnumerable[string]]]::new(); $d['Retry-After'] = [string[]]@('5'); $d } }
        @{ Case = 'case-insensitive header name'; Factory = { $d = [System.Collections.Generic.Dictionary[string, string]]::new(); $d['retry-after'] = '5'; $d } }
        @{ Case = 'PSCustomObject'; Factory = { [pscustomobject]@{ 'Retry-After' = '5' } } }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ Factory = $Factory } {
            [string](@(Get-TarProperty -InputObject (& $Factory) -Name 'Retry-After')[0]) | Should -Be '5'
        }
    }
    It 'returns $null for missing keys and $null input' {
        InModuleScope TeamsAccessReport {
            $d = [System.Collections.Generic.Dictionary[string, string]]::new()
            Get-TarProperty -InputObject $d -Name 'x' | Should -BeNullOrEmpty
            Get-TarProperty -InputObject $null -Name 'x' | Should -BeNullOrEmpty
        }
    }
}

Describe 'ConvertFrom-TarRetryAfterHeader (adapter contract with the real Graph SDK header type)' {
    It '<Case> -> <Expected>' -ForEach @(
        @{ Case = 'generic Dictionary[string, IEnumerable[string]]'; Expected = 7; Factory = { $h = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.IEnumerable[string]]]::new(); $h['Retry-After'] = [string[]]@('7'); $h } }
        @{ Case = 'Hashtable'; Expected = 3; Factory = { @{ 'Retry-After' = '3' } } }
        @{ Case = 'header missing'; Expected = 0; Factory = { [System.Collections.Generic.Dictionary[string, System.Collections.Generic.IEnumerable[string]]]::new() } }
        @{ Case = 'non-numeric (HTTP date)'; Expected = 0; Factory = { @{ 'Retry-After' = 'Wed, 21 Oct 2026 07:28:00 GMT' } } }
        @{ Case = '$null headers'; Expected = 0; Factory = { $null } }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ Factory = $Factory; Expected = $Expected } {
            ConvertFrom-TarRetryAfterHeader -Headers (& $Factory) | Should -Be $Expected
        }
    }
}

Describe 'Get-TarChannelSite when the files folder is not provisioned' {
    BeforeAll {
        Mock -ModuleName TeamsAccessReport Write-TarLog { }
        Mock -ModuleName TeamsAccessReport Invoke-TarGraphRequest -ParameterFilter { $Uri -like '*/filesFolder' } {
            $ex = [System.Net.Http.HttpRequestException]::new('Graph HTTP 404 for x :: Folder location for this channel is not ready yet')
            $ex.Data['StatusCode'] = 404; $ex.Data['GraphMessage'] = 'Folder location for this channel is not ready yet, please try again later.'
            throw $ex
        }
        Mock -ModuleName TeamsAccessReport Invoke-TarGraphRequest -ParameterFilter { $Uri -like '*/drive/root*' } {
            @{ id = 'root'; webUrl = 'https://t.sharepoint.com/sites/T/Shared Documents'; parentReference = @{ driveId = 'd1' }; sharepointIds = @{ siteUrl = 'https://t.sharepoint.com/sites/T' } }
        }
    }
    It 'standard channel falls back to the group drive' {
        InModuleScope TeamsAccessReport {
            $r = Get-TarChannelSite -Context ([pscustomobject]@{ GraphBase = 'g' }) -TeamId 't' -ChannelId 'c' -ChannelName 'General' -MembershipType 'standard'
            $r.SiteUrl | Should -Be 'https://t.sharepoint.com/sites/T'; $r.SiteSource | Should -Be 'GroupDriveFallback'; $r.DriveId | Should -Be 'd1'
        }
    }
    It '<_> channel fails with an actionable message' -ForEach @('private', 'shared') {
        InModuleScope TeamsAccessReport -Parameters @{ Type = $_ } {
            { Get-TarChannelSite -Context ([pscustomobject]@{ GraphBase = 'g' }) -TeamId 't' -ChannelId 'c' -ChannelName 'Payroll' -MembershipType $Type } |
                Should -Throw -ExpectedMessage "*Open the channel's Files tab*-Resume*"
        }
    }
    It 'other 404 errors are not swallowed' {
        InModuleScope TeamsAccessReport {
            Mock Invoke-TarGraphRequest -ParameterFilter { $Uri -like '*/filesFolder' } { $ex = [System.Net.Http.HttpRequestException]::new('Graph HTTP 404 for x :: Not found'); $ex.Data['StatusCode'] = 404; $ex.Data['GraphMessage'] = 'Not found'; throw $ex }
            { Get-TarChannelSite -Context ([pscustomobject]@{ GraphBase = 'g' }) -TeamId 't' -ChannelId 'c' -ChannelName 'General' -MembershipType 'standard' } | Should -Throw -ExpectedMessage '*Not found*'
        }
    }
}

Describe 'Get-TarHttpStatusFromError' {
    It 'parses <Message> as <Expected>' -ForEach @(
        @{ Message = 'Response status code does not indicate success: 429 (Too Many Requests).'; Expected = 429 }
        @{ Message = 'The remote server returned an error: (503) Server Unavailable.'; Expected = 503 }
        @{ Message = 'Operation failed: Forbidden'; Expected = 403 }
        @{ Message = 'GET https://x/teams/5030-aaaa-503b failed'; Expected = 0 }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ Message = $Message; Expected = $Expected } {
            try { throw [System.Exception]::new($Message) } catch { Get-TarHttpStatusFromError -ErrorRecord $_ | Should -Be $Expected }
        }
    }
}

Describe 'Get-TarBackoffDelay' {
    It 'Retry-After takes precedence' {
        InModuleScope TeamsAccessReport { $d = Get-TarBackoffDelay -Attempt 5 -BaseDelaySeconds 3 -RetryAfterSeconds 7; $d | Should -BeGreaterOrEqual 7000; $d | Should -BeLessThan 8000 }
    }
    It 'is capped at 120 s + jitter' {
        InModuleScope TeamsAccessReport { Get-TarBackoffDelay -Attempt 20 -BaseDelaySeconds 60 | Should -BeLessThan 121000 }
    }
}

Describe 'ConvertFrom-TarSelectionInput' {
    It '<Text> -> <Expected>' -ForEach @(
        @{ Text = '1,3,5-7'; Count = 8; Expected = '0,2,4,5,6' }
        @{ Text = '2 1 2'; Count = 3; Expected = '0,1' }
        @{ Text = '9'; Count = 3; Expected = 'null' }
        @{ Text = '3-1'; Count = 3; Expected = 'null' }
        @{ Text = 'x'; Count = 3; Expected = 'null' }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ Text = $Text; Count = $Count; Expected = $Expected } {
            $r = ConvertFrom-TarSelectionInput -Text $Text -Count $Count
            $(if ($null -eq $r) { 'null' } else { $r -join ',' }) | Should -Be $Expected
        }
    }
}

Describe 'Protect-TarCsvValue' {
    It 'neutralizes formula prefixes and leaves everything else untouched' {
        InModuleScope TeamsAccessReport {
            $o = [pscustomobject]@{ A = '=SUM(1)'; B = '+1'; C = '-x'; D = '@y'; E = 'Anna'; F = 5 } | Protect-TarCsvValue
            $o.A | Should -Be "'=SUM(1)"; $o.B | Should -Be "'+1"; $o.C | Should -Be "'-x"; $o.D | Should -Be "'@y"; $o.E | Should -Be 'Anna'; $o.F | Should -Be 5
        }
    }
}

Describe 'ConvertTo-TarSpAssignmentRecord' {
    It 'reads roles in nometadata and verbose format and flags Limited Access' -ForEach @(
        @{ Bindings = @([pscustomobject]@{ Name = 'Edit' }); Roles = 'Edit'; Limited = $false }
        @{ Bindings = [pscustomobject]@{ results = @([pscustomobject]@{ Name = 'Read' }, [pscustomobject]@{ Name = 'Limited Access' }) }; Roles = 'Read;Limited Access'; Limited = $false }
        @{ Bindings = @([pscustomobject]@{ Name = 'Limited Access' }); Roles = 'Limited Access'; Limited = $true }
        @{ Bindings = @([pscustomobject]@{ Name = 'Web-Only Limited Access' }); Roles = 'Web-Only Limited Access'; Limited = $true }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ Bindings = $Bindings; Roles = $Roles; Limited = $Limited } {
            $ra = [pscustomobject]@{ Member = [pscustomobject]@{ Title = 'X'; LoginName = 'x'; Email = ''; PrincipalType = 1 }; RoleDefinitionBindings = $Bindings }
            $r = $ra | ConvertTo-TarSpAssignmentRecord -SiteUrl 's' -WebUrl 's' -Scope 'Web'
            $r.Roles | Should -Be $Roles; $r.LimitedAccessOnly | Should -Be $Limited
        }
    }
}

Describe 'Resolve-TarSpPrincipal' {
    BeforeAll {
        Mock -ModuleName TeamsAccessReport Invoke-TarGraphRequest { @{ id = 'u1'; displayName = 'Anna'; userPrincipalName = 'anna@contoso.se'; userType = 'Member' } }
        Mock -ModuleName TeamsAccessReport Get-TarGraphCollection { @{ '@odata.type' = '#microsoft.graph.user'; id = 'u1'; displayName = 'Anna'; userPrincipalName = 'anna@contoso.se' }; @{ '@odata.type' = '#microsoft.graph.group'; id = 'g' } }
    }
    It '<Login> -> <Kind>' -ForEach @(
        @{ Login = 'i:0#.f|membership|anna@contoso.se'; Title = 'Anna'; Kind = 'User'; Path = 'Direct' }
        @{ Login = 'c:0o.c|federateddirectoryclaimprovider|aaaaaaaa-0000-0000-0000-000000000001'; Title = 'M'; Kind = 'User'; Path = 'Direct > M365Group:aaaaaaaa-0000-0000-0000-000000000001/members' }
        @{ Login = 'c:0o.c|federateddirectoryclaimprovider|aaaaaaaa-0000-0000-0000-000000000001_o'; Title = 'O'; Kind = 'User'; Path = 'Direct > M365Group:aaaaaaaa-0000-0000-0000-000000000001/owners' }
        @{ Login = 'c:0t.c|tenant|bbbbbbbb-0000-0000-0000-000000000001'; Title = 'S'; Kind = 'User'; Path = 'Direct > SecurityGroup:bbbbbbbb-0000-0000-0000-000000000001' }
        @{ Login = 'c:0-.f|rolemanager|spo-grid-all-users/111'; Title = 'EEEU'; Kind = 'BroadGrant'; Path = 'Direct' }
        @{ Login = 'c:0(.s|true'; Title = 'Everyone'; Kind = 'BroadGrant'; Path = 'Direct' }
        @{ Login = 'SharingLinks.a.AnonymousView.b'; Title = 'SharingLinks.a.AnonymousView.b'; Kind = 'AnonymousLink'; Path = 'Direct' }
        @{ Login = 'SharingLinks.a.OrganizationEdit.b'; Title = 'SharingLinks.a.OrganizationEdit.b'; Kind = 'OrganizationLink'; Path = 'Direct' }
        @{ Login = 'SHAREPOINT\system'; Title = 'System'; Kind = 'SystemOrApp'; Path = 'Direct' }
        @{ Login = 'c:0u.c|tenant|a.68e55f997a4b45f689726a76bf5248b0'; Title = 'SLinkClaim.1.2'; Kind = 'SharingLinkClaim'; Path = 'Direct' }
        @{ Login = 'weird|claim'; Title = 'W'; Kind = 'Unresolved'; Path = 'Direct' }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ Login = $Login; Title = $Title; Kind = $Kind; Path = $Path } {
            $r = @(Resolve-TarSpPrincipal -Context ([pscustomobject]@{ GraphBase = 'g' }) -Cache (New-TarResolverCache) -Login $Login -Title $Title)
            $r.Count | Should -Be 1 -Because 'groups must be filtered to users'
            $r[0].Kind | Should -Be $Kind; $r[0].Path | Should -Be $Path
        }
    }
}

Describe 'Get-TarSiteDrift' {
    BeforeAll {
        $script:site = @{ Url = 's'; Kind = 'ChannelSite'; Label = 'T / C' }
        function Row { param($Kind = 'User', $UserId, $Upn, $Scope = 'Web', $Limited = 'False', $UserType = 'Member') [pscustomobject]@{ Kind = $Kind; UserId = $UserId; Upn = $Upn; DisplayName = $Upn; UserType = $UserType; Scope = $Scope; ObjectPath = 'p'; Roles = 'Edit'; LimitedAccessOnly = $Limited; ExpansionPath = 'Direct' } }
    }
    It 'does not flag Teams members, ignores Limited Access, ranks guests as High' {
        InModuleScope TeamsAccessReport -Parameters @{ site = $script:site; rows = @((Row -UserId 'u1' -Upn 'anna'), (Row -UserId 'u2' -Upn 'bo' -Limited 'True'), (Row -UserId 'u9' -Upn 'c#EXT#@x' -UserType 'Guest')); teams = @([pscustomobject]@{ UserId = 'u1'; Email = 'anna'; DisplayName = 'Anna'; Scope = 'Channel'; Role = 'Owner'; AccessPath = 'DirectChannelMember' }) } {
            $d = @(Get-TarSiteDrift -Site $site -TeamsUsers $teams -SiteRows $rows)
            $d.Count | Should -Be 1
            $d[0].UserId | Should -Be 'u9'; $d[0].Severity | Should -Be 'High'; $d[0].Finding | Should -Be 'SiteAccessNotInTeams'
        }
    }
    It 'reports a Teams member missing from SharePoint as Info' {
        InModuleScope TeamsAccessReport -Parameters @{ site = $script:site; rows = @((Row -UserId 'u1' -Upn 'anna')); teams = @([pscustomobject]@{ UserId = 'u1' }, [pscustomobject]@{ UserId = 'u5'; Email = 'e'; DisplayName = 'E'; Scope = 'Channel'; Role = 'Member'; AccessPath = 'DirectChannelMember' }) } {
            $d = @(Get-TarSiteDrift -Site $site -TeamsUsers $teams -SiteRows $rows)
            $d.Count | Should -Be 1; $d[0].Finding | Should -Be 'InTeamsNotInSharePoint'; $d[0].Severity | Should -Be 'Info'
        }
    }
    It 'broad grants are always High' {
        InModuleScope TeamsAccessReport -Parameters @{ site = $script:site; rows = @((Row -Kind 'BroadGrant' -UserId '*' -Upn 'Everyone')) } {
            $d = @(Get-TarSiteDrift -Site $site -TeamsUsers @() -SiteRows $rows)
            $d[0].Finding | Should -Be 'BroadGrant'; $d[0].Severity | Should -Be 'High'
        }
    }
}

Describe 'Get-TarDriftCause' {
    BeforeAll {
        $script:pub = @{ Url = 'https://t/sites/a'; Kind = 'TeamSite'; Visibility = 'public' }
        function New-EeeuRow { param($AssociatedGroup = 'Member', $Scope = 'Web', $WebUrl = 'https://t/sites/a', $Upn = 'EveryoneExceptExternalUsers')
            [pscustomobject]@{ Kind = 'BroadGrant'; Upn = $Upn; Scope = $Scope; AssociatedGroup = $AssociatedGroup; WebUrl = $WebUrl } }
    }
    It 'explains EEEU in the associated Members group of a public team site' {
        InModuleScope TeamsAccessReport -Parameters @{ site = $script:pub; row = (New-EeeuRow) } { Get-TarDriftCause -Site $site -Row $row | Should -Be 'PublicTeamDefault' }
    }
    It 'does not explain <Name>' -ForEach @(
        @{ Name = 'a private team'; Site = @{ Url = 'https://t/sites/a'; Kind = 'TeamSite'; Visibility = 'private' }; Row = @{} }
        @{ Name = 'a channel site'; Site = @{ Url = 'https://t/sites/a'; Kind = 'ChannelSite'; Visibility = 'public' }; Row = @{} }
        @{ Name = 'a custom group'; Site = $null; Row = @{ AssociatedGroup = '' } }
        @{ Name = 'the Visitors group'; Site = $null; Row = @{ AssociatedGroup = 'Visitor' } }
        @{ Name = 'a list-level grant'; Site = $null; Row = @{ Scope = 'List' } }
        @{ Name = 'a subweb'; Site = $null; Row = @{ WebUrl = 'https://t/sites/a/sub' } }
        @{ Name = 'Everyone (incl. external)'; Site = $null; Row = @{ Upn = 'Everyone' } }
    ) {
        $s = if ($Site) { $Site } else { $script:pub }
        $r = New-EeeuRow @Row
        InModuleScope TeamsAccessReport -Parameters @{ site = $s; row = $r } { Get-TarDriftCause -Site $site -Row $row | Should -Be '' }
    }
}

Describe 'Selection.json round trip' {
    BeforeAll {
        $script:ctx = InModuleScope TeamsAccessReport -Parameters @{ p = (Join-Path $TestDrive 'sel') } { $c = New-TarContext -OutputPath $p; $c.TenantId = 't1'; $null = New-Item -ItemType Directory -Path $p -Force; $c }
    }
    It 'writes ChannelIds as an array even for one channel, and AllChannels as a boolean' {
        InModuleScope TeamsAccessReport -Parameters @{ ctx = $script:ctx } {
            $f = @{ A = [System.Collections.Generic.HashSet[string]]::new([string[]]@('c1')) }
            Save-TarSelection -Context $ctx -Selection @{ Teams = @(@{ id = 'A'; displayName = 'a' }, @{ id = 'B'; displayName = 'b' }); ChannelFilter = $f }
            $raw = Get-Content (Get-TarStorePath $ctx 'Selection.json') -Raw
            $raw | Should -Match '"ChannelIds":\s*\[\s*"c1"\s*\]'
            $raw | Should -Match '"ChannelIds":\s*\[\s*\]'
            $j = $raw | ConvertFrom-Json
            $j.SchemaVersion | Should -Be 2
            $j.Teams[0].AllChannels | Should -BeFalse; $j.Teams[1].AllChannels | Should -BeTrue
            $back = Import-TarSelection -Context $ctx
            $back.ChannelFilter.Keys | Should -Be @('A'); $back.ChannelFilter['A'].Contains('c1') | Should -BeTrue
        }
    }
    It 'still imports schema 1 (<Name>)' -ForEach @(
        @{ Name = 'ALL'; Ids = '"ALL"'; Expect = $null }
        @{ Name = 'single id as string'; Ids = '"c1"'; Expect = 'c1' }
        @{ Name = 'array'; Ids = '["c1","c2"]'; Expect = 'c2' }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ ctx = $script:ctx; ids = $Ids; expect = $Expect } {
            Set-Content (Get-TarStorePath $ctx 'Selection.json') "{`"SchemaVersion`":1,`"TenantId`":`"t1`",`"Teams`":[{`"Id`":`"A`",`"Name`":`"a`",`"ChannelIds`":$ids}]}"
            $f = (Import-TarSelection -Context $ctx).ChannelFilter
            if ($expect) { $f['A'].Contains($expect) | Should -BeTrue } else { $f.ContainsKey('A') | Should -BeFalse }
        }
    }
    It 'refuses a newer schema' {
        InModuleScope TeamsAccessReport -Parameters @{ ctx = $script:ctx } {
            Set-Content (Get-TarStorePath $ctx 'Selection.json') '{"SchemaVersion":99,"TenantId":"t1","Teams":[]}'
            { Import-TarSelection -Context $ctx } | Should -Throw -ExpectedMessage '*newer than this module*'
        }
    }
}

Describe 'CSV column contract' {
    It 'writes a header-only file for an empty set and keeps it readable as zero rows' {
        InModuleScope TeamsAccessReport -Parameters @{ p = (Join-Path $TestDrive 'hdr') } {
            $c = New-TarContext -OutputPath $p; $null = New-Item -ItemType Directory -Path $p -Force
            Set-TarRecord -Context $c -Name $script:TarFile.ExternalAssoc -InputObject @()
            Get-TarCsvHeader -Path (Get-TarStorePath $c $script:TarFile.ExternalAssoc) | Should -Be $script:TarColumns[$script:TarFile.ExternalAssoc]
            @(Read-TarRecord -Context $c -Name $script:TarFile.ExternalAssoc).Count | Should -Be 0
            Write-TarRecord -Context $c -Name $script:TarFile.ExternalAssoc -InputObject @([pscustomobject]@{ UserId = 'u'; HostTenantId = 'h'; HostTeamId = 't'; HostTeamName = 'n' })
            @(Read-TarRecord -Context $c -Name $script:TarFile.ExternalAssoc).Count | Should -Be 1
        }
    }
    It 'refuses -Resume into a folder written with other columns' {
        InModuleScope TeamsAccessReport -Parameters @{ p = (Join-Path $TestDrive 'old') } {
            $c = New-TarContext -OutputPath $p; $null = New-Item -ItemType Directory -Path $c.StateDir -Force
            Set-Content (Join-Path $p 'Selection.json') '{}'
            Set-Content (Join-Path $p $script:TarFile.Drift) '"SiteUrl","Site","SiteKind","Finding","Severity","UserId"'
            { Initialize-TarStore -Context $c -Resume } | Should -Throw -ExpectedMessage '*incompatible version*'
        }
    }
}

Describe 'ConvertTo-TarTenantQuery' {
    It '<In> -> <Out>' -ForEach @(
        @{ In = 'contoso'; Out = 'contoso.onmicrosoft.com' }
        @{ In = ' Contoso.OnMicrosoft.com '; Out = 'contoso.onmicrosoft.com' }
        @{ In = 'contoso.com'; Out = 'contoso.com' }
        @{ In = 'admin@Contoso.com'; Out = 'contoso.com' }
        @{ In = 'https://login.microsoftonline.com/'; Out = 'login.microsoftonline.com' }
        @{ In = '3C28C7A0-9A0A-435D-B574-9345D1A794F7'; Out = '3c28c7a0-9a0a-435d-b574-9345d1a794f7' }
        @{ In = '"etherlinked"'; Out = 'etherlinked.onmicrosoft.com' }
    ) { InModuleScope TeamsAccessReport -Parameters @{ i = $In } { ConvertTo-TarTenantQuery -InputText $i } | Should -Be $Out }
    It 'rejects <_>' -ForEach @('', '   ', 'two words', 'bad_name', '-contoso', 'contoso..com') {
        InModuleScope TeamsAccessReport -Parameters @{ i = $_ } { ConvertTo-TarTenantQuery -InputText $i } | Should -BeNullOrEmpty
    }
}

Describe 'Resolve-TarTenant' {
    It 'reads the tenant ID from the issuer' {
        InModuleScope TeamsAccessReport {
            Mock Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 200; Body = [pscustomobject]@{ issuer = 'https://login.microsoftonline.com/3C28C7A0-9A0A-435D-B574-9345D1A794F7/v2.0'; tenant_region_scope = 'EU'; cloud_instance_name = 'microsoftonline.com' } } }
            $r = Resolve-TarTenant -Tenant 'etherlinked'
            $r.TenantId | Should -Be '3c28c7a0-9a0a-435d-b574-9345d1a794f7'; $r.Query | Should -Be 'etherlinked.onmicrosoft.com'; $r.Region | Should -Be 'EU'
        }
    }
    It 'retries transient failures with backoff, then gives up as Unavailable' {
        InModuleScope TeamsAccessReport {
            Mock Start-Sleep { }
            Mock Write-TarLog { }
            Mock Invoke-TarHttpRaw { throw 'No such host is known.' }
            $e = $null
            try { Resolve-TarTenant -Tenant 'contoso' -MaxRetries 3 } catch { $e = $_ }
            $e.Exception.Data['TarReason'] | Should -Be 'Unavailable'
            Should -Invoke Invoke-TarHttpRaw -Times 3 -Exactly
            Should -Invoke Start-Sleep -Times 2 -Exactly
        }
    }
    It 'passes the input through when the lookup is unavailable, but not when the tenant does not exist' {
        InModuleScope TeamsAccessReport {
            Mock Start-Sleep { }
            Mock Write-TarLog { }
            Mock Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 503; Body = $null } }
            (Resolve-TarTenantOrPassThrough -Tenant 'contoso').TenantId | Should -Be 'contoso.onmicrosoft.com'
            Mock Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 400; Body = $null } }
            { Resolve-TarTenantOrPassThrough -Tenant 'contoso' } | Should -Throw -ExpectedMessage '*not found*'
        }
    }
}

Describe 'Profile record validation' {
    It 'normalizes thumbprint and client ID' {
        InModuleScope TeamsAccessReport {
            $r = New-TarProfileRecord -Name 'a' -TenantId 't' -AuthMode AppOnly -ClientId ' 9D35D536-264B-43BE-B838-1E34F2702154 ' -CertificateThumbprint "`u{200E}c5 f0 51 e8 5f c0 59 df 7b 30 84 b4 36 03 7e e2 51 62 b1 81"
            $r.ClientId | Should -Be '9d35d536-264b-43be-b838-1e34f2702154'
            $r.CertificateThumbprint | Should -Be 'C5F051E85FC059DF7B3084B436037EE25162B181'
        }
    }
    It 'rejects <Name>' -ForEach @(
        @{ Name = 'a bad profile name'; Splat = @{ Name = 'a b'; AuthMode = 'Delegated' }; Pattern = '*invalid*' }
        @{ Name = 'app-only without certificate'; Splat = @{ Name = 'a'; AuthMode = 'AppOnly'; ClientId = '9d35d536-264b-43be-b838-1e34f2702154' }; Pattern = '*both*' }
        @{ Name = 'a short thumbprint'; Splat = @{ Name = 'a'; AuthMode = 'AppOnly'; ClientId = '9d35d536-264b-43be-b838-1e34f2702154'; CertificateThumbprint = 'ABC' }; Pattern = '*40 hexadecimal*' }
        @{ Name = 'a non-GUID client ID'; Splat = @{ Name = 'a'; AuthMode = 'AppOnly'; ClientId = 'app'; CertificateThumbprint = ('AB' * 20) }; Pattern = '*must be a GUID*' }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ s = $Splat; pattern = $Pattern } { { New-TarProfileRecord -TenantId 't' @s } | Should -Throw -ExpectedMessage $pattern }
    }
    It 'stores no secrets and round-trips through the store atomically' {
        InModuleScope TeamsAccessReport -Parameters @{ path = (Join-Path $TestDrive 'p' 'profiles.json') } {
            $env:TEAMSACCESSREPORT_PROFILE_PATH = $path
            try {
                Mock Write-TarLog { }
                Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'b' -TenantId 't2' -AuthMode Delegated)
                Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'a' -TenantId 't1' -AuthMode AppOnly -ClientId '9d35d536-264b-43be-b838-1e34f2702154' -CertificateThumbprint ('AB' * 20))
                Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'b' -TenantId 't3' -AuthMode Delegated)
                $all = @(Read-TarProfileStore)
                $all.Name | Should -Be @('a', 'b'); ($all | Where-Object Name -eq 'b').TenantId | Should -Be 't3'
                (Get-Content $path -Raw) | Should -Not -Match '(?i)secret|password|privatekey'
                @(Get-ChildItem (Split-Path $path) -Filter '*.tmp').Count | Should -Be 0
            } finally { Remove-Item Env:TEAMSACCESSREPORT_PROFILE_PATH }
        }
    }
}

Describe 'Resolve-TarConnection precedence' {
    BeforeAll {
        $script:store = Join-Path $TestDrive 'prec.json'
        $env:TEAMSACCESSREPORT_PROFILE_PATH = $script:store
        InModuleScope TeamsAccessReport {
            Mock Write-TarLog { }
            Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'one' -TenantId 'aaaaaaaa-0000-0000-0000-000000000001' -AuthMode AppOnly -ClientId '9d35d536-264b-43be-b838-1e34f2702154' -CertificateThumbprint ('AB' * 20))
        }
    }
    AfterAll { Remove-Item Env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }
    It 'resume: uses the single saved profile for the resumed tenant' {
        InModuleScope TeamsAccessReport {
            Mock Assert-TarCertificate { }
            $c = Resolve-TarConnection -ResumeSelection ([pscustomobject]@{ TenantId = 'aaaaaaaa-0000-0000-0000-000000000001' })
            $c.ProfileName | Should -Be 'one'; $c.AuthMode | Should -Be 'AppOnly'
        }
    }
    It 'resume: falls back to delegated for a tenant without a profile' {
        InModuleScope TeamsAccessReport {
            $c = Resolve-TarConnection -ResumeSelection ([pscustomobject]@{ TenantId = 'bbbbbbbb-0000-0000-0000-000000000002' })
            $c.AuthMode | Should -Be 'Delegated'; $c.TenantId | Should -Be 'bbbbbbbb-0000-0000-0000-000000000002'
        }
    }
    It 'no parameters and no console: delegated, tenant chosen at sign-in' {
        InModuleScope TeamsAccessReport { $c = Resolve-TarConnection; $c.AuthMode | Should -Be 'Delegated'; $c.TenantId | Should -BeNullOrEmpty }
    }
    It 'only one of -ClientId/-CertificateThumbprint fails' {
        InModuleScope TeamsAccessReport { { Resolve-TarConnection -ClientId '9d35d536-264b-43be-b838-1e34f2702154' -TenantId 't' } | Should -Throw -ExpectedMessage '*both -ClientId and -CertificateThumbprint*' }
    }
}

Describe 'Assert-TarCertificate' {
    It '<Name>' -ForEach @(
        @{ Name = 'no private key'; Row = @{ HasPrivateKey = $false; Valid = $true }; Pattern = '*no private key*' }
        @{ Name = 'expired'; Row = @{ HasPrivateKey = $true; Valid = $false }; Pattern = '*not valid now*' }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ row = $Row; pattern = $Pattern } {
            Mock Get-TarCertificateCandidate { [pscustomobject](@{ Id = 'X'; Subject = 'CN=x'; Expires = '2020-01-01'; Store = 'CurrentUser\My' } + $row) }
            { Assert-TarCertificate -Thumbprint 'X' } | Should -Throw -ExpectedMessage $pattern
        }
    }
    It 'lists real store certificates without throwing (smoke)' { InModuleScope TeamsAccessReport { { @(Get-TarCertificateCandidate) } | Should -Not -Throw } }
}

Describe 'Interactive connection flow (simulated answers)' {
    BeforeAll {
        $env:TEAMSACCESSREPORT_PROFILE_PATH = Join-Path $TestDrive 'ui.json'
        $script:okDiscovery = { [pscustomobject]@{ StatusCode = 200; Body = [pscustomobject]@{ issuer = 'https://login.microsoftonline.com/3c28c7a0-9a0a-435d-b574-9345d1a794f7/v2.0'; tenant_region_scope = 'EU'; cloud_instance_name = 'microsoftonline.com' } } }
    }
    AfterAll { Remove-Item Env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }
    BeforeEach { Remove-Item $env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }

    It 'tenant prompt: retries after an unknown tenant, then accepts a short name' {
        InModuleScope TeamsAccessReport -Parameters @{ ok = $script:okDiscovery } {
            Mock Write-Host { }
            $script:answers = [System.Collections.Generic.Queue[string]]::new([string[]]@('', 'nope', 'etherlinked'))
            Mock Read-Host { $script:answers.Dequeue() }
            Mock Invoke-TarHttpRaw -ParameterFilter { $Uri -like '*/nope.onmicrosoft.com/*' } { [pscustomobject]@{ StatusCode = 400; Body = $null } }
            Mock Invoke-TarHttpRaw -ParameterFilter { $Uri -like '*/etherlinked.onmicrosoft.com/*' } $ok
            (Read-TarTenantInput).TenantId | Should -Be '3c28c7a0-9a0a-435d-b574-9345d1a794f7'
            Should -Invoke Read-Host -Times 3 -Exactly
        }
    }
    It 'tenant prompt: q cancels; Enter means "choose at sign-in" when allowed' {
        InModuleScope TeamsAccessReport {
            Mock Read-Host { 'q' }
            Read-TarTenantInput | Should -BeNullOrEmpty
            Mock Read-Host { '' }
            $t = Read-TarTenantInput -AllowEmpty
            $t | Should -Not -BeNullOrEmpty; $t.TenantId | Should -BeNullOrEmpty
        }
    }
    It 'client ID prompt: rejects a non-GUID, then accepts' {
        InModuleScope TeamsAccessReport {
            Mock Write-Host { }
            $script:answers = [System.Collections.Generic.Queue[string]]::new([string[]]@('app', '9D35D536-264B-43BE-B838-1E34F2702154'))
            Mock Read-Host { $script:answers.Dequeue() }
            Read-TarGuidInput -Prompt 'x' | Should -Be '9d35d536-264b-43be-b838-1e34f2702154'
            Mock Read-Host { 'q' }
            Read-TarGuidInput -Prompt 'x' | Should -BeNullOrEmpty
        }
    }
    It 'certificate picker: picks from the store, or falls back to typing when the store is empty' {
        InModuleScope TeamsAccessReport {
            Mock Write-Host { }
            Mock Get-TarCertificateCandidate { [pscustomobject]@{ Id = 'C5F051E85FC059DF7B3084B436037EE25162B181'; Subject = 'CN=TeamsAccessReport-Test'; Expires = '2027-11-05'; Store = 'CurrentUser\My'; HasPrivateKey = $true; Valid = $true } }
            Mock Select-TarItem { $Rows[0] }
            Select-TarCertificate -SelectionMode Menu | Should -Be 'C5F051E85FC059DF7B3084B436037EE25162B181'
            Should -Invoke Select-TarItem -ParameterFilter { $Single } -Times 1 -Exactly
            Mock Get-TarCertificateCandidate { }
            $script:answers = [System.Collections.Generic.Queue[string]]::new([string[]]@('short', ('ab' * 20)))
            Mock Read-Host { $script:answers.Dequeue() }
            Select-TarCertificate -SelectionMode Menu | Should -Be ('AB' * 20)
        }
    }
    It 'start screen without saved profiles runs the wizard and saves the profile' {
        InModuleScope TeamsAccessReport -Parameters @{ ok = $script:okDiscovery } {
            Mock Write-Host { }
            Mock Write-TarLog { }
            Mock Invoke-TarHttpRaw $ok
            Mock Get-TarCertificateCandidate { [pscustomobject]@{ Id = ('AB' * 20); Subject = 'CN=t'; Expires = '2099-01-01'; Store = 'CurrentUser\My'; HasPrivateKey = $true; Valid = $true } }
            # Select-TarItem: 1st call = sign-in method -> AppOnly, 2nd = certificate.
            Mock Select-TarItem { $Rows[0] }
            # Read-Host: tenant, client ID, save? (Read-TarYesNo), profile name (Enter = default).
            $script:answers = [System.Collections.Generic.Queue[string]]::new([string[]]@('etherlinked', '9d35d536-264b-43be-b838-1e34f2702154', 'y', ''))
            Mock Read-Host { $script:answers.Dequeue() }
            $c = Select-TarConnection -SelectionMode Menu
            $c.AuthMode | Should -Be 'AppOnly'; $c.TenantId | Should -Be '3c28c7a0-9a0a-435d-b574-9345d1a794f7'; $c.ProfileName | Should -Be 'etherlinked'
            (@(Read-TarProfileStore)).Name | Should -Be @('etherlinked')
        }
    }
    It 'wizard without saving returns a connection without a profile name' {
        InModuleScope TeamsAccessReport -Parameters @{ ok = $script:okDiscovery } {
            Mock Write-Host { }
            Mock Invoke-TarHttpRaw $ok
            Mock Select-TarItem { $Rows | Where-Object Id -eq 'Delegated' }
            $script:answers = [System.Collections.Generic.Queue[string]]::new([string[]]@('etherlinked', 'n'))
            Mock Read-Host { $script:answers.Dequeue() }
            $c = Select-TarConnection -SelectionMode Menu
            $c.AuthMode | Should -Be 'Delegated'; $c.ProfileName | Should -BeNullOrEmpty
            @(Read-TarProfileStore).Count | Should -Be 0
        }
    }
    It 'start screen with saved profiles: <Pick>' -ForEach @(
        @{ Pick = 'profile:saved'; Mode = 'AppOnly'; Profile = 'saved' }
        @{ Pick = 'delegated'; Mode = 'Delegated'; Profile = $null }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ pick = $Pick; mode = $Mode; prof = $Profile } {
            Mock Write-TarLog { }
            Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'saved' -TenantId 'aaaaaaaa-0000-0000-0000-000000000001' -AuthMode AppOnly -ClientId '9d35d536-264b-43be-b838-1e34f2702154' -CertificateThumbprint ('AB' * 20))
            Mock Select-TarItem { $Rows | Where-Object Id -eq $pick }
            Mock Read-Host { '' }
            $c = Select-TarConnection -SelectionMode Menu
            $c.AuthMode | Should -Be $mode; $c.ProfileName | Should -Be $prof
            Should -Invoke Select-TarItem -ParameterFilter { $Single -and $Title -like 'Connection*' } -Times 1 -Exactly
        }
    }
    It 'start screen: nothing picked cancels' {
        InModuleScope TeamsAccessReport {
            Mock Write-TarLog { }
            Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'saved' -TenantId 't' -AuthMode Delegated)
            Mock Select-TarItem { }
            Select-TarConnection -SelectionMode Menu | Should -BeNullOrEmpty
        }
    }
    It 'New-TeamsAccessReportProfile without a console and without enough parameters fails clearly' {
        Mock -ModuleName TeamsAccessReport Test-TarCanPrompt { $false }
        { New-TeamsAccessReportProfile -Tenant contoso } | Should -Throw -ExpectedMessage '*No interactive console*'
        { New-TeamsAccessReportProfile -Tenant contoso -Delegated -ClientId x } | Should -Throw -ExpectedMessage '*cannot be combined*'
    }
    It 'Remove-TeamsAccessReportProfile -WhatIf keeps the profile' {
        InModuleScope TeamsAccessReport { Mock Write-TarLog { }; Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'keep' -TenantId 't' -AuthMode Delegated) }
        Remove-TeamsAccessReportProfile -Name keep -WhatIf
        (Get-TeamsAccessReportProfile -Name 'ke*').Name | Should -Be 'keep'
    }
}

Describe 'Test-TarStartFormInput (start screen validation)' {
    BeforeAll {
        $env:TEAMSACCESSREPORT_PROFILE_PATH = Join-Path $TestDrive 'form.json'
        function New-State { param([hashtable]$Change = @{})
            $s = @{ SelectedProfile = ''; Tenant = 'etherlinked'; AuthMode = 'AppOnly'; ClientId = '9d35d536-264b-43be-b838-1e34f2702154'; CertificateThumbprint = ('AB' * 20)
                PnPClientId = ''; SaveProfile = $true; ProfileName = ''; IncludeSharePoint = $true; IncludeUserAssociations = $true; SelectChannels = $false
                OutputPath = (Join-Path $TestDrive 'out1'); ConnectionOnly = $false }
            foreach ($k in $Change.Keys) { $s[$k] = $Change[$k] }
            $s
        }
    }
    AfterAll { Remove-Item Env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }
    BeforeEach {
        Remove-Item $env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue
        Mock -ModuleName TeamsAccessReport Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 200; Body = [pscustomobject]@{ issuer = 'https://login.microsoftonline.com/3c28c7a0-9a0a-435d-b574-9345d1a794f7/v2.0'; tenant_region_scope = 'EU'; cloud_instance_name = 'microsoftonline.com' } } }
        Mock -ModuleName TeamsAccessReport Get-TarCertificateCandidate { [pscustomobject]@{ Id = ('AB' * 20); Subject = 'CN=t'; Expires = '2099-01-01'; Store = 'CurrentUser\My'; HasPrivateKey = $true; Valid = $true } }
        Mock -ModuleName TeamsAccessReport Write-TarLog { }
    }
    It 'valid app-only input: connection, options and a profile named after the tenant' {
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State) } { Test-TarStartFormInput -State $s }
        $r.Errors.Count | Should -Be 0
        $r.Connection.TenantId | Should -Be '3c28c7a0-9a0a-435d-b574-9345d1a794f7'; $r.Connection.ProfileName | Should -Be 'etherlinked'
        $r.ProfileRecord.Options.IncludeSharePoint | Should -BeTrue
        $r.Options.IncludeUserAssociations | Should -BeTrue
    }
    It 'collects every error at once' {
        Mock -ModuleName TeamsAccessReport Get-TarCertificateCandidate { }
        $state = New-State @{ ClientId = 'app'; PnPClientId = 'x'; OutputPath = '' }
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = $state } { Test-TarStartFormInput -State $s }
        @($r.Errors.Keys) | Should -Be @('ClientId', 'Certificate', 'PnPClientId', 'OutputPath')
        $r.ProfileRecord | Should -BeNullOrEmpty
    }
    It 'unknown tenant is reported on the Tenant field' {
        Mock -ModuleName TeamsAccessReport Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 400; Body = $null } }
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State) } { Test-TarStartFormInput -State $s }
        $r.Errors.Tenant | Should -Match 'not found'
    }
    It 'delegated: no tenant needed, user associations forced off, SharePoint requires a PnP app' {
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State @{ AuthMode = 'Delegated'; Tenant = ''; SaveProfile = $false }) } { Test-TarStartFormInput -State $s }
        @($r.Errors.Keys) | Should -Be @('SharePoint')
        $r.Options.IncludeUserAssociations | Should -BeFalse
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State @{ AuthMode = 'Delegated'; Tenant = ''; SaveProfile = $false; PnPClientId = '9d35d536-264b-43be-b838-1e34f2702154' }) } { Test-TarStartFormInput -State $s }
        $r.Errors.Count | Should -Be 0; $r.Connection.TenantId | Should -BeNullOrEmpty
    }
    It 'refuses an output folder that already contains a run' {
        $o = Join-Path $TestDrive 'used'; $null = New-Item -ItemType Directory $o -Force; Set-Content (Join-Path $o 'Selection.json') '{}'
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State @{ OutputPath = $o }) } { Test-TarStartFormInput -State $s }
        $r.Errors.OutputPath | Should -Match 'already contains a run'
    }
    It 'refuses to overwrite another profile, but updates the selected one' {
        InModuleScope TeamsAccessReport { Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'etherlinked' -TenantId '3c28c7a0-9a0a-435d-b574-9345d1a794f7' -AuthMode Delegated) }
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State) } { Test-TarStartFormInput -State $s }
        $r.Errors.ProfileName | Should -Match 'already exists'
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State @{ SelectedProfile = 'etherlinked' }) } { Test-TarStartFormInput -State $s }
        $r.Errors.Count | Should -Be 0; $r.ProfileRecord.AuthMode | Should -Be 'AppOnly'
    }
    It 'keeps the profile name for an unchanged selected profile that is not re-saved' {
        InModuleScope TeamsAccessReport { Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'p1' -TenantId '3c28c7a0-9a0a-435d-b574-9345d1a794f7' -AuthMode AppOnly -ClientId '9d35d536-264b-43be-b838-1e34f2702154' -CertificateThumbprint ('AB' * 20)) }
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State @{ SelectedProfile = 'p1'; SaveProfile = $false }) } { Test-TarStartFormInput -State $s }
        $r.Connection.ProfileName | Should -Be 'p1'
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State @{ SelectedProfile = 'p1'; SaveProfile = $false; ClientId = '11111111-2222-3333-4444-555555555555' }) } { Test-TarStartFormInput -State $s }
        $r.Connection.ProfileName | Should -BeNullOrEmpty
    }
    It 'connection-only mode ignores report options and stores no options' {
        $r = InModuleScope TeamsAccessReport -Parameters @{ s = (New-State @{ ConnectionOnly = $true; OutputPath = '' }) } { Test-TarStartFormInput -State $s }
        $r.Errors.Count | Should -Be 0; $r.ProfileRecord.Options | Should -BeNullOrEmpty
    }
}

Describe 'Start screen window (Windows Forms)' -Skip:(-not $IsWindows) {
    BeforeAll {
        $env:TEAMSACCESSREPORT_PROFILE_PATH = Join-Path $TestDrive 'win.json'
        InModuleScope TeamsAccessReport { Add-Type -AssemblyName System.Windows.Forms, System.Drawing }
        $script:profiles = @(
            [pscustomobject]@{ Name = 'old'; TenantId = 't-old'; TenantQuery = 'old.onmicrosoft.com'; AuthMode = 'Delegated'; ClientId = $null; CertificateThumbprint = $null; PnPClientId = $null; Options = $null; UpdatedAt = '2026-01-01T00:00:00' }
            [pscustomobject]@{ Name = 'etherlinked'; TenantId = '3c28c7a0-9a0a-435d-b574-9345d1a794f7'; TenantQuery = 'etherlinked.onmicrosoft.com'; AuthMode = 'AppOnly'; ClientId = '9d35d536-264b-43be-b838-1e34f2702154'; CertificateThumbprint = ('AB' * 20); PnPClientId = $null; Options = [pscustomobject]@{ IncludeSharePoint = $true; IncludeUserAssociations = $true; SelectChannels = $false }; UpdatedAt = '2026-11-05T00:00:00' }
        )
        $script:certs = @([pscustomobject]@{ Id = ('AB' * 20); Subject = 'CN=TeamsAccessReport-Test'; Expires = '2027-11-05'; Store = 'CurrentUser\My'; HasPrivateKey = $true; Valid = $true })
    }
    AfterAll { Remove-Item Env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }
    It 'pre-selects the newest profile and fills every field from it' {
        InModuleScope TeamsAccessReport -Parameters @{ p = $script:profiles; c = $script:certs } {
            $ui = New-TarStartForm -Profiles $p -Certificates $c -Defaults @{ OutputPath = 'C:\out' }
            try {
                $ui.Profile.SelectedIndex | Should -Be 2
                $ui.Tenant.Text | Should -Be 'etherlinked.onmicrosoft.com'; $ui.ClientId.Text | Should -Be '9d35d536-264b-43be-b838-1e34f2702154'
                $ui.Certificate.SelectedIndex | Should -Be 0; $ui.SharePoint.Checked | Should -BeTrue; $ui.Associations.Checked | Should -BeTrue
                $ui.Save.Text | Should -Be 'Update profile'; $ui.Output.Text | Should -Be 'C:\out'
                (Get-TarStartFormState -Ui $ui).CertificateThumbprint | Should -Be ('AB' * 20)
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'command-line switches win over profile options' {
        InModuleScope TeamsAccessReport -Parameters @{ p = $script:profiles; c = $script:certs } {
            $ui = New-TarStartForm -Profiles $p -Certificates $c -Defaults @{ OutputPath = 'C:\out'; IncludeSharePoint = $false }
            try { $ui.SharePoint.Checked | Should -BeFalse; $ui.Associations.Checked | Should -BeTrue } finally { $ui.Form.Dispose() }
        }
    }
    It 'delegated disables app fields and user associations' {
        InModuleScope TeamsAccessReport -Parameters @{ p = $script:profiles; c = $script:certs } {
            $ui = New-TarStartForm -Profiles $p -Certificates $c
            try {
                $ui.Delegated.Checked = $true; Update-TarStartFormState -Ui $ui
                $ui.ClientId.Enabled | Should -BeFalse; $ui.Certificate.Enabled | Should -BeFalse
                $ui.Associations.Enabled | Should -BeFalse; $ui.Associations.Checked | Should -BeFalse
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'new connection with a single certificate pre-selects it; a typed thumbprint is used as typed' {
        InModuleScope TeamsAccessReport -Parameters @{ c = $script:certs } {
            $ui = New-TarStartForm -Certificates $c
            try {
                $ui.Profile.SelectedIndex | Should -Be 0; $ui.Certificate.SelectedIndex | Should -Be 0
                $ui.Certificate.SelectedIndex = -1; $ui.Certificate.Text = 'cd' * 20
                (Get-TarStartFormState -Ui $ui).CertificateThumbprint | Should -Be ('cd' * 20)
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'Start keeps the window open on errors and closes with OK when valid' {
        InModuleScope TeamsAccessReport -Parameters @{ p = $script:profiles; c = $script:certs; out = (Join-Path $TestDrive 'winout') } {
            Mock Test-TarStartFormInput { [pscustomobject]@{ Errors = [ordered]@{ Tenant = 'Tenant: bad' }; Connection = $null; Options = $null; ProfileRecord = $null } }
            $ui = New-TarStartForm -Profiles $p -Certificates $c -Defaults @{ OutputPath = $out }
            try {
                Invoke-TarStartFormSubmit -Ui $ui | Should -BeFalse
                $ui.Error.Text | Should -Be 'Tenant: bad'; $ui.Form.DialogResult | Should -Be 'None'
                Mock Test-TarStartFormInput { [pscustomobject]@{ Errors = [ordered]@{}; Connection = 'c'; Options = 'o'; ProfileRecord = $null } }
                Invoke-TarStartFormSubmit -Ui $ui | Should -BeTrue
                $ui.Form.DialogResult | Should -Be 'OK'; $ui.Result.Connection | Should -Be 'c'
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'Check shows the resolved tenant or the error' {
        InModuleScope TeamsAccessReport {
            Mock Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 200; Body = [pscustomobject]@{ issuer = 'https://login.microsoftonline.com/3c28c7a0-9a0a-435d-b574-9345d1a794f7/v2.0'; tenant_region_scope = 'EU' } } }
            $ui = New-TarStartForm
            try {
                $ui.Tenant.Text = 'etherlinked'; Invoke-TarStartFormCheck -Ui $ui
                $ui.TenantStatus.Text | Should -Match '3c28c7a0'; $ui.ProfileName.Text | Should -Be 'etherlinked'
                Mock Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 400; Body = $null } }
                Invoke-TarStartFormCheck -Ui $ui
                $ui.TenantStatus.Text | Should -Match 'not found'
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'connection-only window hides report options and always saves' {
        InModuleScope TeamsAccessReport {
            $ui = New-TarStartForm -ConnectionOnly
            try { $ui.SharePoint.Parent.Parent.Visible | Should -BeFalse; $ui.Save.Checked | Should -BeTrue; $ui.Save.Enabled | Should -BeFalse; $ui.Start.Text | Should -Be 'Save' }
            finally { $ui.Form.Dispose() }
        }
    }
}

Describe 'Select-TarItem routing' {
    It 'Gui multi-select uses the check-box list and passes pre-checked IDs' {
        InModuleScope TeamsAccessReport {
            Mock Test-TarGuiAvailable { $true }
            Mock Show-TarCheckList { $Rows | Where-Object { $_.Id -in $CheckedId } }
            $rows = @([pscustomobject]@{ Id = 'a'; Name = 'A' }, [pscustomobject]@{ Id = 'b'; Name = 'B' })
            @(Select-TarItem -Rows $rows -Title 't' -Columns 'Name' -Mode Gui -CheckedId 'b').Id | Should -Be @('b')
            Should -Invoke Show-TarCheckList -Times 1 -Exactly
        }
    }
}

Describe 'Check-box list (Windows Forms)' -Skip:(-not $IsWindows) {
    BeforeAll {
        InModuleScope TeamsAccessReport { Add-Type -AssemblyName System.Windows.Forms, System.Drawing }
        $script:teams = @(1..250 | ForEach-Object { [pscustomobject]@{ Id = "id$_"; Name = $(if ($_ % 50 -eq 0) { "Finance $_" } else { "Team $_" }); Visibility = 'private' } })
    }
    It 'starts with the pre-checked IDs, counts and enables OK accordingly' {
        InModuleScope TeamsAccessReport -Parameters @{ rows = $script:teams } {
            $ui = New-TarCheckListForm -Rows $rows -Title 't' -Columns 'Name', 'Visibility'
            try {
                $ui.List.Items.Count | Should -Be 250; $ui.Count.Text | Should -Be '0 of 250 selected'; $ui.Ok.Enabled | Should -BeFalse
            } finally { $ui.Form.Dispose() }
            $ui = New-TarCheckListForm -Rows $rows -Title 't' -Columns 'Name' -CheckedId 'id1', 'id2'
            try { $ui.Count.Text | Should -Be '2 of 250 selected'; $ui.Ok.Enabled | Should -BeTrue; $ui.List.Items[0].Checked | Should -BeTrue } finally { $ui.Form.Dispose() }
        }
    }
    It 'filter + Select all checks only visible rows and keeps them when the filter is cleared' {
        InModuleScope TeamsAccessReport -Parameters @{ rows = $script:teams } {
            $ui = New-TarCheckListForm -Rows $rows -Title 't' -Columns 'Name', 'Visibility'
            try {
                $ui.Filter.Text = 'finance'
                $ui.List.Items.Count | Should -Be 5
                Set-TarCheckListAll -Ui $ui -Checked $true
                $ui.Count.Text | Should -Be '5 of 250 selected (245 hidden by filter)'
                $ui.Filter.Text = ''
                $ui.List.Items.Count | Should -Be 250
                @($ui.List.Items | Where-Object Checked).Count | Should -Be 5
                $ui.Checked.Contains('id50') | Should -BeTrue
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'a single row keeps all its columns (no array unrolling)' {
        InModuleScope TeamsAccessReport {
            $ui = New-TarCheckListForm -Rows @([pscustomobject]@{ Id = 'x'; Name = 'Only'; Visibility = 'public' }) -Title 't' -Columns 'Name', 'Visibility'
            try { $ui.List.Items.Count | Should -Be 1; $ui.List.Items[0].SubItems[1].Text | Should -Be 'public' } finally { $ui.Form.Dispose() }
        }
    }
    It 'checking an item updates the selection; Clear removes visible ones' {
        InModuleScope TeamsAccessReport -Parameters @{ rows = $script:teams } {
            $ui = New-TarCheckListForm -Rows $rows -Title 't' -Columns 'Name'
            try {
                $ui.List.Items[2].Checked = $true
                $ui.Checked.Contains('id3') | Should -BeTrue; $ui.Count.Text | Should -Be '1 of 250 selected'
                Set-TarCheckListAll -Ui $ui -Checked $false
                $ui.Checked.Count | Should -Be 0; $ui.Ok.Enabled | Should -BeFalse
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'certificate text in the start window has no CN= prefix and a wide drop-down' {
        InModuleScope TeamsAccessReport {
            $c = @([pscustomobject]@{ Id = ('AB' * 20); Subject = 'CN=TeamsAccessReport-Test'; Expires = '2027-11-05'; Store = 'CurrentUser\My'; HasPrivateKey = $true; Valid = $true })
            $ui = New-TarStartForm -Certificates $c
            try { $ui.Certificate.Items[0] | Should -Be 'TeamsAccessReport-Test  (expires 2027-11-05, CurrentUser\My)'; $ui.Certificate.DropDownWidth | Should -BeGreaterOrEqual 560 }
            finally { $ui.Form.Dispose() }
        }
    }
}

Describe 'PnP app for delegated SharePoint' {
    BeforeAll { $env:TEAMSACCESSREPORT_PROFILE_PATH = Join-Path $TestDrive 'pnp.json' }
    AfterAll { Remove-Item Env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }
    It 'guide text explains why, what happens, least privilege and requirements' {
        $t = InModuleScope TeamsAccessReport { Get-TarPnPGuideText -Tenant 'etherlinked.onmicrosoft.com' -PnPInstalled $true }
        $t | Should -Match '9 September 2024'; $t | Should -Match 'ONE delegated permission: SharePoint AllSites.FullControl'
        $t | Should -Match 'http://localhost'; $t | Should -Match 'Application Administrator'; $t | Should -Not -Match 'NOT installed'
        InModuleScope TeamsAccessReport { Get-TarPnPGuideText -Tenant 't' -PnPInstalled $false } | Should -Match 'Install-Module PnP.PowerShell'
    }
    It 'fails clearly when PnP.PowerShell is missing' {
        InModuleScope TeamsAccessReport { Mock Test-TarPnPInstalled { $false }; { New-TarPnPApp -Tenant 't' -Confirm:$false } | Should -Throw -ExpectedMessage '*Install-Module PnP.PowerShell*' }
    }
    It 'returns the client ID written by the worker and cleans up the result file' {
        InModuleScope TeamsAccessReport {
            Mock Test-TarPnPInstalled { $true }; Mock Write-TarLog { }
            Mock Invoke-TarPnPRegistrationWorker { $script:lastResult = $ResultPath; '{"ok":true,"clientId":"AAAAAAAA-0000-0000-0000-000000000009"}' | Set-Content $ResultPath }
            New-TarPnPApp -Tenant 't' -Confirm:$false | Should -Be 'aaaaaaaa-0000-0000-0000-000000000009'
            Test-Path $script:lastResult | Should -BeFalse
        }
    }
    It 'reports <Name>' -ForEach @(
        @{ Name = 'a worker error'; Json = '{"ok":false,"error":"Insufficient privileges"}'; Pattern = '*registration failed: Insufficient privileges*' }
        @{ Name = 'a closed window'; Json = $null; Pattern = '*closed before it finished*' }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ json = $Json; pattern = $Pattern } {
            Mock Test-TarPnPInstalled { $true }; Mock Write-TarLog { }
            Mock Invoke-TarPnPRegistrationWorker { if ($json) { $json | Set-Content $ResultPath } }
            { New-TarPnPApp -Tenant 't' -Confirm:$false } | Should -Throw -ExpectedMessage $pattern
        }
    }
    It 'New-TeamsAccessReportPnPApp stores the ID in the profile and keeps its options' {
        InModuleScope TeamsAccessReport {
            Mock Write-TarLog { }
            Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'd' -TenantId 't1' -TenantQuery 'd.onmicrosoft.com' -AuthMode Delegated -Options @{ IncludeSharePoint = $true })
            Mock New-TarPnPApp { 'bbbbbbbb-0000-0000-0000-000000000001' }
        }
        New-TeamsAccessReportPnPApp -ProfileName d -InformationAction SilentlyContinue | Should -Be 'bbbbbbbb-0000-0000-0000-000000000001'
        Should -Invoke -ModuleName TeamsAccessReport New-TarPnPApp -ParameterFilter { $Tenant -eq 'd.onmicrosoft.com' }
        $p = Get-TeamsAccessReportProfile -Name d
        $p.PnPClientId | Should -Be 'bbbbbbbb-0000-0000-0000-000000000001'; $p.Options.IncludeSharePoint | Should -BeTrue
    }
}

Describe 'Worker/Register-PnPApp.ps1 (separate process, fake PnP module)' {
    BeforeAll {
        $script:mods = Join-Path $TestDrive 'mods'
        $dir = Join-Path $script:mods 'PnP.PowerShell'
        $null = New-Item -ItemType Directory -Path $dir -Force
        Set-Content (Join-Path $dir 'PnP.PowerShell.psm1') @'
function Register-PnPEntraIDAppForInteractiveLogin {
    param($ApplicationName, $Tenant, [string[]]$SharePointDelegatePermissions)
    if ($Tenant -eq 'fail.onmicrosoft.com') { throw 'Insufficient privileges to complete the operation.' }
    if ($SharePointDelegatePermissions -ne 'AllSites.FullControl') { throw 'unexpected permissions' }
    $o = [pscustomobject]@{}; $o | Add-Member -NotePropertyName 'AzureAppId/ClientId' -NotePropertyValue 'cccccccc-0000-0000-0000-000000000003'; $o
}
'@
        $script:worker = Join-Path $PSScriptRoot '..' 'Worker' 'Register-PnPApp.ps1'
        $script:pwsh = (Get-Process -Id $PID).Path
        function Invoke-Worker { param($Tenant)
            $out = Join-Path $TestDrive "r-$([guid]::NewGuid()).json"
            $old = $env:PSModulePath
            try { $env:PSModulePath = $script:mods + [IO.Path]::PathSeparator + $old; $null = & $script:pwsh -NoProfile -File $script:worker -Tenant $Tenant -ApplicationName 'X' -ResultPath $out 2>&1 }
            finally { $env:PSModulePath = $old }
            Get-Content $out -Raw | ConvertFrom-Json
        }
    }
    It 'extracts the client ID from the AzureAppId/ClientId property and requests only AllSites.FullControl' {
        $r = Invoke-Worker -Tenant 'ok.onmicrosoft.com'
        $r.ok | Should -BeTrue; $r.clientId | Should -Be 'cccccccc-0000-0000-0000-000000000003'
    }
    It 'writes the error when registration fails' {
        $r = Invoke-Worker -Tenant 'fail.onmicrosoft.com'
        $r.ok | Should -BeFalse; $r.error | Should -Match 'Insufficient privileges'
    }
}

Describe 'PnP controls in the start window (Windows Forms)' -Skip:(-not $IsWindows) {
    BeforeAll { InModuleScope TeamsAccessReport { Add-Type -AssemblyName System.Windows.Forms, System.Drawing } }
    It 'hint follows sign-in method, SharePoint choice and filled ID' {
        InModuleScope TeamsAccessReport {
            $ui = New-TarStartForm
            try {
                $ui.PnPHint.Text | Should -Match 'Not needed for app-only'
                $ui.Delegated.Checked = $true; $ui.SharePoint.Checked = $true; Update-TarStartFormState -Ui $ui
                $ui.PnPHint.Text | Should -Match "Required.*Create"
                $ui.PnPClientId.Text = 'x'; Update-TarStartFormState -Ui $ui
                $ui.PnPHint.Text | Should -Match 'Used for SharePoint'
            } finally { $ui.Form.Dispose() }
        }
    }
    It 'Create without a tenant asks for the tenant first' {
        InModuleScope TeamsAccessReport {
            $ui = New-TarStartForm
            try { Invoke-TarStartFormPnPCreate -Ui $ui; $ui.Error.Text | Should -Be 'Enter the tenant first.' } finally { $ui.Form.Dispose() }
        }
    }
    It 'guide window disables Create when PnP.PowerShell is missing' {
        InModuleScope TeamsAccessReport {
            $g = New-TarPnPGuideForm -Tenant 't' -PnPInstalled $false
            try { $g.Create.Enabled | Should -BeFalse; $g.Text.Text | Should -Match 'NOT installed' } finally { $g.Form.Dispose() }
            $g = New-TarPnPGuideForm -Tenant 't' -PnPInstalled $true
            try { $g.Create.Enabled | Should -BeTrue } finally { $g.Form.Dispose() }
        }
    }
}

Describe 'Get-TarChannelAccessPath' {
    It '<Type> channel, source <Src> -> <Path>' -ForEach @(
        @{ Type = 'private'; Src = 'HOST'; Path = 'DirectChannelMember' }
        @{ Type = 'private'; Src = ''; Path = 'DirectChannelMember' }
        @{ Type = 'shared'; Src = ''; Path = 'DirectChannelMember' }
        @{ Type = 'shared'; Src = 'HOST'; Path = 'IndirectViaHostTeam' }
        @{ Type = 'shared'; Src = 'OTHER'; Path = 'IndirectViaSharedTeam' }
    ) {
        InModuleScope TeamsAccessReport -Parameters @{ t = $Type; s = $Src } { Get-TarChannelAccessPath -ChannelType $t -SourceTeamId $s -HostTeamId 'HOST' } | Should -Be $Path
    }
}

Describe 'Graph enum handling' {
    It 'asks Graph for evolvable enum members (shared channels are otherwise unknownFutureValue)' {
        InModuleScope TeamsAccessReport {
            Mock Invoke-MgGraphRequest { @{ value = @() } }
            $null = Invoke-TarGraphRaw -Uri 'https://graph.microsoft.com/v1.0/teams/x/channels'
            Should -Invoke Invoke-MgGraphRequest -Times 1 -Exactly -ParameterFilter { $Headers['Prefer'] -eq 'include-unknown-enum-members' }
        }
    }
    It 'treats an unknown channel type as a channel site, not as the team site' {
        InModuleScope TeamsAccessReport {
            $m = Get-TarSiteMap -Channels @([pscustomobject]@{ TeamId = 't'; TeamName = 'T'; ChannelId = 'c'; ChannelName = 'Delad'; ChannelType = 'unknownFutureValue'; SiteUrl = 'https://x/teams/T-Delad' })
            $m['https://x/teams/T-Delad'].Kind | Should -Be 'ChannelSite'
        }
    }
}
