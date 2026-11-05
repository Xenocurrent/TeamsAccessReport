# Integration tests: the full flow (phases 1-4) against deterministic responses via the adapter seams
# Invoke-TarGraphRaw and Invoke-TarSpRaw. Phase 2 runs InProcess (the same code the worker process runs).

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '..' 'TeamsAccessReport.psd1') -Force
    . (Join-Path $PSScriptRoot 'TestData.ps1')
    if (-not (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
        function global:Connect-PnPOnline { param($Url, $ClientId, $Thumbprint, $Tenant, [switch]$ReturnConnection) }
    }
    InModuleScope TeamsAccessReport { $script:TarSharePointExecutionMode = 'InProcess' }

    $script:AppArgs = @{ TenantId = 'contoso.onmicrosoft.com'; ClientId = 'app'; CertificateThumbprint = 'thumb'; Force = $true; InformationAction = 'SilentlyContinue'; WarningAction = 'SilentlyContinue' }

    function Set-TarTestMocks {
        Mock -ModuleName TeamsAccessReport Invoke-TarGraphRaw { Get-TestGraphResponse -Uri $Uri }
        Mock -ModuleName TeamsAccessReport Invoke-TarSpRaw { Get-TestSpResponse -Url $Url -Connection $Connection }
        Mock -ModuleName TeamsAccessReport Connect-PnPOnline { if ($global:TarTest.FailSite -and $Url -like "*$($global:TarTest.FailSite)") { throw 'Simulated (403) Forbidden' }; [pscustomobject]@{ Url = $Url } }
        Mock -ModuleName TeamsAccessReport Connect-MgGraph { }
        Mock -ModuleName TeamsAccessReport Disconnect-MgGraph { }
        Mock -ModuleName TeamsAccessReport Set-MgRequestContext { }
        Mock -ModuleName TeamsAccessReport Get-MgContext { [pscustomobject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; Account = $null; AppName = 'TestApp'; AuthType = 'AppOnly'; Scopes = @('Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All', 'ChannelMember.Read.All', 'Files.Read.All', 'GroupMember.Read.All', 'User.Read.All') } }
        Mock -ModuleName TeamsAccessReport Start-Sleep { }
    }
}

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
        @{ Name = 'SharePoint delegated without PnP app'; Splat = @{ AllTeams = $true; IncludeSharePoint = $true }; Pattern = 'dedicated Entra app' }
    ) {
        { Invoke-TeamsAccessReport @Splat -OutputPath (Join-Path $TestDrive ([guid]::NewGuid())) -WarningAction SilentlyContinue } | Should -Throw -ExpectedMessage "*$Pattern*"
    }
    It 'app-only requires TenantId (parameter set metadata; not invoked because it would prompt interactively)' {
        $attr = (Get-Command Invoke-TeamsAccessReport).Parameters['TenantId'].Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] -and $_.ParameterSetName -eq 'AppOnly' }
        $attr.Mandatory | Should -BeTrue
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
