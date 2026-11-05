BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'TeamsAccessReport.psd1') -Force
}

Describe 'Modulyta' {
    It 'exporterar exakt en publik funktion' {
        (Get-Command -Module TeamsAccessReport).Name | Should -Be @('Invoke-TeamsAccessReport')
    }
    It 'har parameter sets Delegated och AppOnly' {
        (Get-Command Invoke-TeamsAccessReport).ParameterSets.Name | Sort-Object | Should -Be @('AppOnly', 'Delegated')
    }
    It 'worker-processen refererar aldrig Graph SDK (assemblykonflikt)' {
        foreach ($f in 'Worker/Invoke-SharePointScan.ps1', 'Private/SharePointScan.ps1', 'Private/Store.ps1', 'Private/Util.ps1', 'Private/Logging.ps1') {
            $tokens = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..' $f), [ref]$tokens, [ref]$null)
            @($tokens | Where-Object { $_.Kind -in 'Generic', 'Identifier' -and $_.Text -match 'MgGraph|-Mg[A-Z]' }).Count | Should -Be 0 -Because $f
        }
    }
}

Describe 'Get-TarHttpStatusFromError' {
    It 'tolkar <Message> som <Expected>' -ForEach @(
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
    It 'Retry-After har företräde' {
        InModuleScope TeamsAccessReport { $d = Get-TarBackoffDelay -Attempt 5 -BaseDelaySeconds 3 -RetryAfterSeconds 7; $d | Should -BeGreaterOrEqual 7000; $d | Should -BeLessThan 8000 }
    }
    It 'har tak på 120 s + jitter' {
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
    It 'neutraliserar formelprefix och lämnar övrigt orört' {
        InModuleScope TeamsAccessReport {
            $o = [pscustomobject]@{ A = '=SUM(1)'; B = '+1'; C = '-x'; D = '@y'; E = 'Anna'; F = 5 } | Protect-TarCsvValue
            $o.A | Should -Be "'=SUM(1)"; $o.B | Should -Be "'+1"; $o.C | Should -Be "'-x"; $o.D | Should -Be "'@y"; $o.E | Should -Be 'Anna'; $o.F | Should -Be 5
        }
    }
}

Describe 'ConvertTo-TarSpAssignmentRecord' {
    It 'läser roller i nometadata- och verbose-format och flaggar Limited Access' -ForEach @(
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
            $r.Count | Should -Be 1 -Because 'grupper ska filtreras till användare'
            $r[0].Kind | Should -Be $Kind; $r[0].Path | Should -Be $Path
        }
    }
}

Describe 'Get-TarSiteDrift' {
    BeforeAll {
        $script:site = @{ Url = 's'; Kind = 'ChannelSite'; Label = 'T / C' }
        function Row { param($Kind = 'User', $UserId, $Upn, $Scope = 'Web', $Limited = 'False', $UserType = 'Member') [pscustomobject]@{ Kind = $Kind; UserId = $UserId; Upn = $Upn; DisplayName = $Upn; UserType = $UserType; Scope = $Scope; ObjectPath = 'p'; Roles = 'Edit'; LimitedAccessOnly = $Limited; ExpansionPath = 'Direct' } }
    }
    It 'flaggar inte Teams-medlemmar, ignorerar Limited Access, rankar gäst som High' {
        InModuleScope TeamsAccessReport -Parameters @{ site = $script:site; rows = @((Row -UserId 'u1' -Upn 'anna'), (Row -UserId 'u2' -Upn 'bo' -Limited 'True'), (Row -UserId 'u9' -Upn 'c#EXT#@x' -UserType 'Guest')); teams = @([pscustomobject]@{ UserId = 'u1'; Email = 'anna'; DisplayName = 'Anna'; Scope = 'Channel'; Role = 'Owner'; AccessPath = 'DirectChannelMember' }) } {
            $d = @(Get-TarSiteDrift -Site $site -TeamsUsers $teams -SiteRows $rows)
            $d.Count | Should -Be 1
            $d[0].UserId | Should -Be 'u9'; $d[0].Severity | Should -Be 'High'; $d[0].Finding | Should -Be 'SiteAccessNotInTeams'
        }
    }
    It 'rapporterar Teams-medlem som saknas i SharePoint som Info' {
        InModuleScope TeamsAccessReport -Parameters @{ site = $script:site; rows = @((Row -UserId 'u1' -Upn 'anna')); teams = @([pscustomobject]@{ UserId = 'u1' }, [pscustomobject]@{ UserId = 'u5'; Email = 'e'; DisplayName = 'E'; Scope = 'Channel'; Role = 'Member'; AccessPath = 'DirectChannelMember' }) } {
            $d = @(Get-TarSiteDrift -Site $site -TeamsUsers $teams -SiteRows $rows)
            $d.Count | Should -Be 1; $d[0].Finding | Should -Be 'InTeamsNotInSharePoint'; $d[0].Severity | Should -Be 'Info'
        }
    }
    It 'breda tilldelningar är alltid High' {
        InModuleScope TeamsAccessReport -Parameters @{ site = $script:site; rows = @((Row -Kind 'BroadGrant' -UserId '*' -Upn 'Everyone')) } {
            $d = @(Get-TarSiteDrift -Site $site -TeamsUsers @() -SiteRows $rows)
            $d[0].Finding | Should -Be 'BroadGrant'; $d[0].Severity | Should -Be 'High'
        }
    }
}
