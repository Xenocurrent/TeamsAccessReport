BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'TeamsAccessReport.psd1') -Force
}

Describe 'Module surface' {
    It 'exports the report command and the profile commands' {
        (Get-Command -Module TeamsAccessReport).Name | Sort-Object | Should -Be @('Get-TeamsAccessReportProfile', 'Invoke-TeamsAccessReport', 'New-TeamsAccessReportProfile', 'Remove-TeamsAccessReportProfile')
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
