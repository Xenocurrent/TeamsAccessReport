# Integration tests: the full flow (phases 1-4) against deterministic responses via the adapter seams
# Invoke-TarGraphRaw and Invoke-TarSpRaw. Phase 2 runs InProcess (the same code the worker process runs).

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'TeamsAccessReport.psd1') -Force
    . (Join-Path $PSScriptRoot 'TestData.ps1')
    if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
        function global:Connect-PnPOnline { param($Url, $ClientId, $Thumbprint, $Tenant, $Connection, [switch]$Interactive, [switch]$ReturnConnection) }
    }
    InModuleScope TeamsAccessReport { $script:TarSharePointExecutionMode = 'InProcess' }

    # Profiles go to TestDrive, never to the real per-user store.
    $env:TEAMSACCESSREPORT_PROFILE_PATH = Join-Path $TestDrive 'profiles.json'
    $script:AppArgs = @{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '22222222-2222-2222-2222-222222222222'; CertificateThumbprint = 'AB' * 20; Force = $true; InformationAction = 'SilentlyContinue'; WarningAction = 'SilentlyContinue' }

    function Set-TarTestMocks {
        Mock -ModuleName TeamsAccessReport Invoke-TarGraphRaw { Get-TestGraphResponse -Uri $Uri }
        Mock -ModuleName TeamsAccessReport Invoke-TarSpRaw { Get-TestSpResponse -Url $Url -Connection $Connection }
        Mock -ModuleName TeamsAccessReport Connect-PnPOnline { if ($global:TarTest.FailSite -and $Url -like "*$($global:TarTest.FailSite)") { throw 'Simulated (403) Forbidden' }; [pscustomobject]@{ Url = $Url } }
        Mock -ModuleName TeamsAccessReport Connect-MgGraph { }
        Mock -ModuleName TeamsAccessReport Disconnect-MgGraph { }
        Mock -ModuleName TeamsAccessReport Set-MgRequestContext { }
        Mock -ModuleName TeamsAccessReport Get-MgContext { [pscustomobject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; Account = $null; AppName = 'TestApp'; AuthType = 'AppOnly'; Scopes = @('Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All', 'ChannelMember.Read.All', 'Files.Read.All', 'GroupMember.Read.All', 'User.Read.All') } }
        Mock -ModuleName TeamsAccessReport Start-Sleep { }
        Mock -ModuleName TeamsAccessReport Test-TarCanPrompt { $false }
        Mock -ModuleName TeamsAccessReport Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 200; Body = [pscustomobject]@{ issuer = 'https://login.microsoftonline.com/11111111-1111-1111-1111-111111111111/v2.0'; tenant_region_scope = 'EU'; cloud_instance_name = 'microsoftonline.com' } } }
        Mock -ModuleName TeamsAccessReport Get-TarCertificateCandidate { [pscustomobject]@{ Id = 'AB' * 20; Subject = 'CN=Test'; Expires = '2099-01-01'; Store = 'CurrentUser\My'; HasPrivateKey = $true; Valid = $true; NotAfter = [datetime]'2099-01-01' } }
    }
}

AfterAll { Remove-Item Env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }

Describe 'Full run (app-only, SharePoint, external associations)' {
    BeforeAll {
        Reset-TarTestState; Set-TarTestMocks
        $script:out = Join-Path $TestDrive 'full'
        $script:summary = Invoke-TeamsAccessReport @AppArgs -AllTeams -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:out
        $script:drift = @(Import-Csv (Join-Path $script:out 'DriftReport.csv'))
    }
    It 'returns exactly one object (no leak into the output stream)' { @($summary).Count | Should -Be 1 }
    It 'returns Complete without errors' { $summary.ResultStatus | Should -Be 'Complete'; $summary.Errors | Should -Be 0 }
    It 'follows pagination (2 teams across 2 pages)' { $summary.Teams | Should -Be 2 }
    It 'falls back to the team document library when a standard channel folder is not provisioned' {
        $c4 = Import-Csv (Join-Path $out 'Channels.csv') | Where-Object ChannelId -eq 'C4'
        $c4.SiteSource | Should -Be 'GroupDriveFallback'
        $c4.SiteUrl | Should -Be 'https://contoso.sharepoint.com/sites/IT'
    }
    It 'labels a private channel member whose source is the host team as direct (real-tenant case)' {
        $payroll = @(Import-Csv (Join-Path $out 'TeamsAccess.csv') | Where-Object { $_.ChannelId -eq 'C2' })
        $payroll.Count | Should -Be 1; $payroll[0].AccessPath | Should -Be 'DirectChannelMember'
    }
    It 'counts channel types, archived teams, indirect and external rows' {
        $summary.Channels | Should -Be 4; $summary.PrivateChannels | Should -Be 1; $summary.SharedChannels | Should -Be 1
        $summary.ArchivedTeams | Should -Be 1; $summary.IndirectRows | Should -Be 1; $summary.ExternalTenantRows | Should -Be 1
        $summary.IncomingCrossTenantChannels | Should -Be 1
    }
    It 'recovers from 503 via the outer retry' { Should -Invoke -ModuleName TeamsAccessReport Start-Sleep -Times 1 -Exactly -Scope Describe }
    It 'scans subwebs and paginated list items' {
        Should -Invoke -ModuleName TeamsAccessReport Invoke-TarSpRaw -ParameterFilter { $Url -like '*/Archive/_api/web?*' } -Times 1 -Scope Describe
        Should -Invoke -ModuleName TeamsAccessReport Invoke-TarSpRaw -ParameterFilter { $Url -like '*skiptoken=2*' } -Times 1 -Scope Describe
    }
    It 'populates Roles and flags Limited Access' {
        $sp = @(Import-Csv (Join-Path $out 'SPAssignments.csv'))
        @($sp | Where-Object { -not $_.Roles }).Count | Should -Be 0
        @($sp | Where-Object { $_.PrincipalTitle -like 'SharingLinks*Flexible*' -and $_.LimitedAccessOnly -eq 'True' }).Count | Should -Be 4
    }
    It 'finds expected drift: <Finding> for <Upn>' -ForEach @(
        @{ Finding = 'BroadGrant'; Upn = 'EveryoneExceptExternalUsers'; Severity = 'High' }
        @{ Finding = 'AnonymousLink'; Upn = 'SharingLinks.abc.AnonymousView.xyz'; Severity = 'High' }
        @{ Finding = 'SiteAccessNotInTeams'; Upn = 'cecilia_ext.com#EXT#@contoso.onmicrosoft.com'; Severity = 'High' }
        @{ Finding = 'SiteCollectionAdminNotInTeams'; Upn = 'ghost@contoso.se'; Severity = 'Medium' }
        @{ Finding = 'SiteAccessNotInTeams'; Upn = 'bo@contoso.se'; Severity = 'Medium' }
    ) {
        @($drift | Where-Object { $_.Finding -eq $Finding -and $_.Upn -eq $Upn -and $_.Severity -eq $Severity }).Count | Should -BeGreaterOrEqual 1
    }
    It 'does not flag team members on the team site' {
        @($drift | Where-Object { $_.SiteKind -eq 'TeamSite' -and $_.Upn -in 'anna@contoso.se', 'bo@contoso.se' }).Count | Should -Be 0
    }
    It 'attributes EEEU on the public team site to the Microsoft default, but not the direct grant on a private site' {
        $it = @($drift | Where-Object { $_.Finding -eq 'BroadGrant' -and $_.SiteUrl -like '*/IT' })
        $it.Count | Should -Be 1; $it[0].Cause | Should -Be 'PublicTeamDefault'; $it[0].Severity | Should -Be 'High'
        $fin = @($drift | Where-Object { $_.Finding -eq 'BroadGrant' -and $_.SiteUrl -like '*/Finance' })
        $fin.Count | Should -Be 1; $fin[0].Cause | Should -BeNullOrEmpty
    }
    It '<_> header matches the column contract' -ForEach @('Teams.csv', 'TeamsAccess.csv', 'Channels.csv', 'IncomingSharedChannels.csv', 'SPAssignments.csv', 'SPGroupMembers.csv', 'EffectiveSharePointAccess.csv', 'DriftReport.csv', 'ExternalAssociations.csv', 'Errors.csv') {
        $name = $_
        $expected = InModuleScope TeamsAccessReport -Parameters @{ n = $name } { $script:TarColumns[$n] }
        $actual = InModuleScope TeamsAccessReport -Parameters @{ p = (Join-Path $out $name) } { Get-TarCsvHeader -Path $p }
        $actual | Should -Be $expected
    }
    It 'writes Errors.csv header-only when there are no errors' { @(Import-Csv (Join-Path $out 'Errors.csv')).Count | Should -Be 0 }
    It 'finds external associations for internal users' { @(Import-Csv (Join-Path $out 'ExternalAssociations.csv')).Count | Should -Be 2 }
    It 'neutralizes formula injection in CSV' { Get-Content (Join-Path $out 'TeamsAccess.csv') -Raw | Should -Match ([regex]::Escape('"''=HYPERLINK')) }
    It 'writes a structured log (JSON Lines)' { Get-Content (Join-Path $out 'run.jsonl') | Select-Object -First 3 | ForEach-Object { { $_ | ConvertFrom-Json } | Should -Not -Throw } }
}

Describe 'Interruption + -Resume produces the same result as a full run' {
    BeforeAll {
        Reset-TarTestState; Set-TarTestMocks
        $script:full = Join-Path $TestDrive 'ref'
        $null = Invoke-TeamsAccessReport @AppArgs -AllTeams -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:full

        Reset-TarTestState; $global:TarTest.FailAllMembers = $true; $global:TarTest.FailSite = 'Finance-Payroll'
        $script:out = Join-Path $TestDrive 'resume'
        $script:first = Invoke-TeamsAccessReport @AppArgs -AllTeams -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:out
        $global:TarTest.FailAllMembers = $false; $global:TarTest.FailSite = $null
        $resumeArgs = $AppArgs.Clone(); $resumeArgs.Remove('Force')
        $script:second = Invoke-TeamsAccessReport @resumeArgs -Resume -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:out
    }
    It 'first run is Incomplete with one partial team' { $first.ResultStatus | Should -Be 'Incomplete'; $first.TeamsPartial | Should -Be 1; $first.Errors | Should -BeGreaterOrEqual 2 }
    It 'resumed run becomes Complete' { $second.ResultStatus | Should -Be 'Complete'; $second.Errors | Should -Be 0 }
    It 'archives old errors' { @(Get-ChildItem (Join-Path $out 'state') -Filter 'Errors_*.csv').Count | Should -Be 1 }
    It '<_> is identical to the reference run' -ForEach @('Teams.csv', 'TeamsAccess.csv', 'Channels.csv', 'IncomingSharedChannels.csv', 'SPAssignments.csv', 'SPGroupMembers.csv', 'EffectiveSharePointAccess.csv', 'DriftReport.csv', 'ExternalAssociations.csv') {
        # -ErrorAction Stop: a missing file must NEVER be compared as two empty strings (false green).
        $a = (Get-Content (Join-Path $full $_) -ErrorAction Stop | Sort-Object) -join "`n"
        $b = (Get-Content (Join-Path $out $_) -ErrorAction Stop | Sort-Object) -join "`n"
        $a | Should -Not -BeNullOrEmpty
        $b | Should -Be $a
    }
}

Describe 'Validation and safeguards' {
    BeforeAll { Reset-TarTestState; Set-TarTestMocks }
    It '<Name>' -ForEach @(
        @{ Name = '-ChannelId without -TeamId'; Splat = @{ ChannelId = 'C1'; AllTeams = $true }; Pattern = 'ChannelId requires' }
        @{ Name = '-TeamId and -AllTeams'; Splat = @{ TeamId = 'x'; AllTeams = $true }; Pattern = 'either -TeamId' }
        @{ Name = '-Resume with -AllTeams'; Splat = @{ Resume = $true; AllTeams = $true }; Pattern = 'Resume reads' }
        @{ Name = 'SharePoint delegated without PnP app'; Splat = @{ AllTeams = $true; IncludeSharePoint = $true }; Pattern = 'New-TeamsAccessReportPnPApp' }
    ) {
        { Invoke-TeamsAccessReport @Splat -OutputPath (Join-Path $TestDrive ([guid]::NewGuid())) -WarningAction SilentlyContinue } | Should -Throw -ExpectedMessage "*$Pattern*"
    }
    It 'app-only without a tenant fails clearly when nobody can be prompted' {
        { Invoke-TeamsAccessReport -ClientId $AppArgs.ClientId -CertificateThumbprint $AppArgs.CertificateThumbprint -AllTeams -OutputPath (Join-Path $TestDrive 'notenant') } |
            Should -Throw -ExpectedMessage '*needs a tenant*'
    }
    It '-ProfileName cannot be combined with explicit connection parameters' {
        { Invoke-TeamsAccessReport -ProfileName x -TenantId contoso -AllTeams -OutputPath (Join-Path $TestDrive 'mix') } | Should -Throw -ExpectedMessage '*either -ProfileName*'
    }
    It 'refuses to overwrite an existing run' {
        $p = Join-Path $TestDrive 'dup'
        $null = Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath $p
        { Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath $p } | Should -Throw -ExpectedMessage '*already contains*'
    }
    It '-WhatIf creates nothing' {
        $p = Join-Path $TestDrive 'whatif'
        $null = Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath $p -WhatIf
        Test-Path $p | Should -BeFalse
    }
    It 'delegated + -IncludeUserAssociations skips the step with a warning' {
        Mock -ModuleName TeamsAccessReport Get-MgContext { [pscustomobject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; Account = 'admin@contoso.se'; AppName = $null; AuthType = 'Delegated'; Scopes = @() } }
        $p = Join-Path $TestDrive 'deleg'
        $null = Invoke-TeamsAccessReport -AllTeams -Force -IncludeUserAssociations -OutputPath $p -InformationAction SilentlyContinue -WarningVariable w -WarningAction SilentlyContinue
        ($w -join ' ') | Should -Match 'requires app-only'
        Test-Path (Join-Path $p 'ExternalAssociations.csv') | Should -BeFalse
    }
    It 'refuses to run when PnP.PowerShell is already loaded' {
        Mock -ModuleName TeamsAccessReport Get-Module -ParameterFilter { $Name -eq 'PnP.PowerShell' -and -not $ListAvailable } { [pscustomobject]@{ Name = 'PnP.PowerShell' } }
        { Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath (Join-Path $TestDrive 'pnp') } | Should -Throw -ExpectedMessage '*PnP.PowerShell is loaded*'
    }
}

Describe 'Connection profiles end to end' {
    BeforeAll { Reset-TarTestState; Set-TarTestMocks; Remove-Item $env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }
    It 'creates a profile from a short tenant name, resolving it to the tenant ID' {
        $p = New-TeamsAccessReportProfile -Tenant 'contoso' -ClientId $AppArgs.ClientId -CertificateThumbprint ("`u{200E}" + ('ab ' * 20)) -InformationAction SilentlyContinue
        $p.Name | Should -Be 'contoso'; $p.TenantId | Should -Be '11111111-1111-1111-1111-111111111111'; $p.TenantQuery | Should -Be 'contoso.onmicrosoft.com'
        $p.CertificateThumbprint | Should -Be ('AB' * 20)
        Should -Invoke -ModuleName TeamsAccessReport Invoke-TarHttpRaw -ParameterFilter { $Uri -eq 'https://login.microsoftonline.com/contoso.onmicrosoft.com/v2.0/.well-known/openid-configuration' }
        (Get-TeamsAccessReportProfile).Name | Should -Be @('contoso')
    }
    It 'refuses to overwrite without -Force' {
        { New-TeamsAccessReportProfile -Tenant 'contoso' -ClientId $AppArgs.ClientId -CertificateThumbprint $AppArgs.CertificateThumbprint } | Should -Throw -ExpectedMessage '*already exists*'
    }
    It 'runs with -ProfileName only and records it for -Resume and the rerun command' {
        $out = Join-Path $TestDrive 'byprofile'
        $s = Invoke-TeamsAccessReport -ProfileName contoso -AllTeams -Force -OutputPath $out -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        $s.ResultStatus | Should -Be 'Complete'; $s.SharePointIncluded | Should -BeFalse   # profile without saved options
        Should -Invoke -ModuleName TeamsAccessReport Connect-MgGraph -ParameterFilter { $TenantId -eq '11111111-1111-1111-1111-111111111111' -and $ClientId -eq '22222222-2222-2222-2222-222222222222' }
        $sel = Get-Content (Join-Path $out 'Selection.json') -Raw | ConvertFrom-Json
        $sel.ProfileName | Should -Be 'contoso'
        $sel.RerunCommand | Should -Match "-ProfileName 'contoso'"; $sel.RerunCommand | Should -Not -Match 'CertificateThumbprint'
    }
    It 'report options saved in the profile apply with -ProfileName only' {
        InModuleScope TeamsAccessReport -Parameters @{ c = $AppArgs.ClientId; th = $AppArgs.CertificateThumbprint } {
            Save-TarProfile -ProfileRecord (New-TarProfileRecord -Name 'withopts' -TenantId '11111111-1111-1111-1111-111111111111' -TenantQuery 'contoso.onmicrosoft.com' -AuthMode AppOnly -ClientId $c -CertificateThumbprint $th -Options @{ IncludeSharePoint = $true }) -Confirm:$false
        }
        $out = Join-Path $TestDrive 'withopts'
        $s = Invoke-TeamsAccessReport -ProfileName withopts -AllTeams -Force -OutputPath $out -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        $s.SharePointIncluded | Should -BeTrue
        Should -Invoke -ModuleName TeamsAccessReport Connect-PnPOnline -Scope It
        Remove-TeamsAccessReportProfile -Name withopts -Confirm:$false -InformationAction SilentlyContinue
    }
    It '-Resume without any connection parameter picks the profile from Selection.json' {
        $out = Join-Path $TestDrive 'byprofile'
        $s = Invoke-TeamsAccessReport -Resume -OutputPath $out -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        $s.ResultStatus | Should -Be 'Complete'
        (Get-Content (Join-Path $out 'run.jsonl') -Raw) | Should -Match "via profile 'contoso'"
    }
    It 'a missing certificate fails before sign-in with the reason' {
        Mock -ModuleName TeamsAccessReport Get-TarCertificateCandidate { }
        { Invoke-TeamsAccessReport -ProfileName contoso -AllTeams -Force -OutputPath (Join-Path $TestDrive 'nocert') } | Should -Throw -ExpectedMessage '*was not found in CurrentUser*'
    }
    It 'an unknown tenant fails with a clear message' {
        Mock -ModuleName TeamsAccessReport Invoke-TarHttpRaw { [pscustomobject]@{ StatusCode = 400; Body = [pscustomobject]@{ error = 'invalid_tenant' } } }
        { New-TeamsAccessReportProfile -Name nope -Tenant 'doesnotexist' -Delegated } | Should -Throw -ExpectedMessage "*'doesnotexist.onmicrosoft.com' was not found*"
    }
    It 'removes a profile' {
        Remove-TeamsAccessReportProfile -Name contoso -Confirm:$false -InformationAction SilentlyContinue
        @(Get-TeamsAccessReportProfile).Count | Should -Be 0
        { Invoke-TeamsAccessReport -ProfileName contoso -AllTeams -OutputPath (Join-Path $TestDrive 'gone') } | Should -Throw -ExpectedMessage "*'contoso' not found*"
    }
}

Describe 'Start screen drives the whole run (no console input)' {
    BeforeAll { Reset-TarTestState; Set-TarTestMocks; Remove-Item $env:TEAMSACCESSREPORT_PROFILE_PATH -ErrorAction SilentlyContinue }
    It 'uses the connection and options from the window, saves the profile and asks the rest with pop-ups' {
        $out = Join-Path $TestDrive 'fromform'
        Mock -ModuleName TeamsAccessReport Test-TarCanPrompt { $true }
        Mock -ModuleName TeamsAccessReport Test-TarGuiAvailable { $true }
        Mock -ModuleName TeamsAccessReport Read-Host { throw 'Console input must not be needed' }
        Mock -ModuleName TeamsAccessReport Show-TarMessageBox { $true }
        Mock -ModuleName TeamsAccessReport Select-TarItem { $Rows }
        $record = InModuleScope TeamsAccessReport -Parameters @{ c = $AppArgs } {
            New-TarProfileRecord -Name 'contoso' -TenantId '11111111-1111-1111-1111-111111111111' -TenantQuery 'contoso.onmicrosoft.com' -AuthMode AppOnly -ClientId $c.ClientId -CertificateThumbprint $c.CertificateThumbprint -Options @{ IncludeSharePoint = $true }
        }
        $result = [pscustomobject]@{
            Errors = [ordered]@{}; ProfileRecord = $record
            Connection = [pscustomobject]@{ AuthMode = 'AppOnly'; TenantId = '11111111-1111-1111-1111-111111111111'; ClientId = $AppArgs.ClientId; CertificateThumbprint = $AppArgs.CertificateThumbprint; PnPClientId = $null; ProfileName = 'contoso' }
            Options = [pscustomobject]@{ IncludeSharePoint = $true; IncludeUserAssociations = $true; SelectChannels = $false; OutputPath = $out }
        }
        Mock -ModuleName TeamsAccessReport Show-TarStartForm { $result }.GetNewClosure()
        $s = Invoke-TeamsAccessReport -SelectionMode Gui -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        $s.ResultStatus | Should -Be 'Complete'
        Test-Path (Join-Path $out 'SPAssignments.csv') | Should -BeTrue
        Test-Path (Join-Path $out 'ExternalAssociations.csv') | Should -BeTrue
        (Get-TeamsAccessReportProfile).Name | Should -Be @('contoso')
        Should -Invoke -ModuleName TeamsAccessReport Show-TarMessageBox -Times 1 -Exactly -ParameterFilter { $Text -like 'Continue with tenant*' }
        Should -Invoke -ModuleName TeamsAccessReport Read-Host -Times 0 -Exactly
    }
    It 'cancelling the window ends the run without creating the output folder' {
        Mock -ModuleName TeamsAccessReport Test-TarCanPrompt { $true }
        Mock -ModuleName TeamsAccessReport Test-TarGuiAvailable { $true }
        Mock -ModuleName TeamsAccessReport Show-TarStartForm { }
        $out = Join-Path $TestDrive 'cancelled'
        $s = Invoke-TeamsAccessReport -SelectionMode Gui -OutputPath $out -InformationAction SilentlyContinue -WarningAction SilentlyContinue
        $s | Should -BeNullOrEmpty; Test-Path $out | Should -BeFalse
    }
}
