# Integrationstester: hela flödet (fas 1-4) mot deterministiska svar via adaptersömmarna
# Invoke-TarGraphRaw och Invoke-TarSpRaw. Fas 2 körs InProcess (samma kod som worker-processen kör).

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
        Mock -ModuleName TeamsAccessReport Connect-PnPOnline { if ($global:TarTest.FailSite -and $Url -like "*$($global:TarTest.FailSite)") { throw 'Simulerad (403) Forbidden' }; [pscustomobject]@{ Url = $Url } }
        Mock -ModuleName TeamsAccessReport Connect-MgGraph { }
        Mock -ModuleName TeamsAccessReport Disconnect-MgGraph { }
        Mock -ModuleName TeamsAccessReport Set-MgRequestContext { }
        Mock -ModuleName TeamsAccessReport Get-MgContext { [pscustomobject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; Account = $null; AppName = 'TestApp'; AuthType = 'AppOnly'; Scopes = @('Team.ReadBasic.All', 'Channel.ReadBasic.All', 'TeamMember.Read.All', 'ChannelMember.Read.All', 'Files.Read.All', 'GroupMember.Read.All', 'User.Read.All') } }
        Mock -ModuleName TeamsAccessReport Start-Sleep { }
    }
}

Describe 'Hel körning (app-only, SharePoint, externa kopplingar)' {
    BeforeAll {
        Reset-TarTestState; Set-TarTestMocks
        $script:out = Join-Path $TestDrive 'full'
        $script:summary = Invoke-TeamsAccessReport @AppArgs -AllTeams -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:out
        $script:drift = @(Import-Csv (Join-Path $script:out 'DriftReport.csv'))
    }
    It 'returnerar exakt ett objekt (ingen läcka i output-strömmen)' { @($summary).Count | Should -Be 1 }
    It 'returnerar Complete utan fel' { $summary.ResultStatus | Should -Be 'Complete'; $summary.Errors | Should -Be 0 }
    It 'följer paginering (2 team över 2 sidor)' { $summary.Teams | Should -Be 2 }
    It 'räknar kanaltyper, arkiv, indirekta och externa rader' {
        $summary.Channels | Should -Be 4; $summary.PrivateChannels | Should -Be 1; $summary.SharedChannels | Should -Be 1
        $summary.ArchivedTeams | Should -Be 1; $summary.IndirectRows | Should -Be 1; $summary.ExternalTenantRows | Should -Be 1
        $summary.IncomingCrossTenantChannels | Should -Be 1
    }
    It 'återhämtar sig från 503 via yttre retry' { Should -Invoke -ModuleName TeamsAccessReport Start-Sleep -Times 1 -Exactly -Scope Describe }
    It 'skannar underwebbar och paginerade listobjekt' {
        Should -Invoke -ModuleName TeamsAccessReport Invoke-TarSpRaw -ParameterFilter { $Url -like '*/Arkiv/_api/web?*' } -Times 1 -Scope Describe
        Should -Invoke -ModuleName TeamsAccessReport Invoke-TarSpRaw -ParameterFilter { $Url -like '*skiptoken=2*' } -Times 1 -Scope Describe
    }
    It 'fyller Roles och flaggar Limited Access' {
        $sp = @(Import-Csv (Join-Path $out 'SPAssignments.csv'))
        @($sp | Where-Object { -not $_.Roles }).Count | Should -Be 0
        @($sp | Where-Object { $_.PrincipalTitle -like 'SharingLinks*Flexible*' -and $_.LimitedAccessOnly -eq 'True' }).Count | Should -Be 4
    }
    It 'hittar förväntade avvikelser: <Finding> för <Upn>' -ForEach @(
        @{ Finding = 'BroadGrant'; Upn = 'EveryoneExceptExternalUsers'; Severity = 'High' }
        @{ Finding = 'AnonymousLink'; Upn = 'SharingLinks.abc.AnonymousView.xyz'; Severity = 'High' }
        @{ Finding = 'SiteAccessNotInTeams'; Upn = 'cecilia_ext.com#EXT#@contoso.onmicrosoft.com'; Severity = 'High' }
        @{ Finding = 'SiteCollectionAdminNotInTeams'; Upn = 'ghost@contoso.se'; Severity = 'Medium' }
        @{ Finding = 'SiteAccessNotInTeams'; Upn = 'bo@contoso.se'; Severity = 'Medium' }
    ) {
        @($drift | Where-Object { $_.Finding -eq $Finding -and $_.Upn -eq $Upn -and $_.Severity -eq $Severity }).Count | Should -BeGreaterOrEqual 1
    }
    It 'flaggar inte teammedlemmar på teamsajten' {
        @($drift | Where-Object { $_.SiteKind -eq 'TeamSite' -and $_.Upn -in 'anna@contoso.se', 'bo@contoso.se' }).Count | Should -Be 0
    }
    It 'hittar externa kopplingar för interna användare' { @(Import-Csv (Join-Path $out 'ExternalAssociations.csv')).Count | Should -Be 2 }
    It 'neutraliserar formelinjektion i CSV' { Get-Content (Join-Path $out 'TeamsAccess.csv') -Raw | Should -Match ([regex]::Escape('"''=HYPERLINK')) }
    It 'skriver strukturerad logg (JSON Lines)' { Get-Content (Join-Path $out 'run.jsonl') | Select-Object -First 3 | ForEach-Object { { $_ | ConvertFrom-Json } | Should -Not -Throw } }
}

Describe 'Avbrott + -Resume ger samma resultat som hel körning' {
    BeforeAll {
        Reset-TarTestState; Set-TarTestMocks
        $script:full = Join-Path $TestDrive 'ref'
        $null = Invoke-TeamsAccessReport @AppArgs -AllTeams -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:full

        Reset-TarTestState; $global:TarTest.FailAllMembers = $true; $global:TarTest.FailSite = 'Ekonomi-Lon'
        $script:out = Join-Path $TestDrive 'resume'
        $script:first = Invoke-TeamsAccessReport @AppArgs -AllTeams -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:out
        $global:TarTest.FailAllMembers = $false; $global:TarTest.FailSite = $null
        $resumeArgs = $AppArgs.Clone(); $resumeArgs.Remove('Force')
        $script:second = Invoke-TeamsAccessReport @resumeArgs -Resume -IncludeSharePoint -IncludeUserAssociations -OutputPath $script:out
    }
    It 'första körningen är Incomplete med ett ofullständigt team' { $first.ResultStatus | Should -Be 'Incomplete'; $first.TeamsPartial | Should -Be 1; $first.Errors | Should -BeGreaterOrEqual 2 }
    It 'återupptaget blir Complete' { $second.ResultStatus | Should -Be 'Complete'; $second.Errors | Should -Be 0 }
    It 'arkiverar gamla fel' { @(Get-ChildItem (Join-Path $out 'state') -Filter 'Errors_*.csv').Count | Should -Be 1 }
    It '<_> är identisk med referenskörningen' -ForEach @('Teams.csv', 'TeamsAccess.csv', 'Channels.csv', 'IncomingSharedChannels.csv', 'SPAssignments.csv', 'SPGroupMembers.csv', 'EffectiveSharePointAccess.csv', 'DriftReport.csv', 'ExternalAssociations.csv') {
        # -ErrorAction Stop: en saknad fil får ALDRIG jämföras som två tomma strängar (falskt grönt).
        $a = (Get-Content (Join-Path $full $_) -ErrorAction Stop | Sort-Object) -join "`n"
        $b = (Get-Content (Join-Path $out $_) -ErrorAction Stop | Sort-Object) -join "`n"
        $a | Should -Not -BeNullOrEmpty
        $b | Should -Be $a
    }
}

Describe 'Validering och skydd' {
    BeforeAll { Reset-TarTestState; Set-TarTestMocks }
    It '<Name>' -ForEach @(
        @{ Name = '-ChannelId utan -TeamId'; Splat = @{ ChannelId = 'C1'; AllTeams = $true }; Pattern = 'ChannelId kräver' }
        @{ Name = '-TeamId och -AllTeams'; Splat = @{ TeamId = 'x'; AllTeams = $true }; Pattern = 'antingen -TeamId' }
        @{ Name = '-Resume med -AllTeams'; Splat = @{ Resume = $true; AllTeams = $true }; Pattern = 'Resume läser' }
        @{ Name = 'SharePoint delegerat utan PnP-app'; Splat = @{ AllTeams = $true; IncludeSharePoint = $true }; Pattern = 'egen Entra-app' }
    ) {
        { Invoke-TeamsAccessReport @Splat -OutputPath (Join-Path $TestDrive ([guid]::NewGuid())) -WarningAction SilentlyContinue } | Should -Throw -ExpectedMessage "*$Pattern*"
    }
    It 'app-only kräver TenantId (parameter set-metadata, anropas inte – skulle annars fråga interaktivt)' {
        $attr = (Get-Command Invoke-TeamsAccessReport).Parameters['TenantId'].Attributes | Where-Object { $_ -is [System.Management.Automation.ParameterAttribute] -and $_.ParameterSetName -eq 'AppOnly' }
        $attr.Mandatory | Should -BeTrue
    }
    It 'vägrar skriva över en befintlig körning' {
        $p = Join-Path $TestDrive 'dup'
        $null = Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath $p
        { Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath $p } | Should -Throw -ExpectedMessage '*innehåller redan*'
    }
    It '-WhatIf skapar ingenting' {
        $p = Join-Path $TestDrive 'whatif'
        $null = Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath $p -WhatIf
        Test-Path $p | Should -BeFalse
    }
    It 'delegerat + -IncludeUserAssociations hoppar över steget med varning' {
        Mock -ModuleName TeamsAccessReport Get-MgContext { [pscustomobject]@{ TenantId = '11111111-1111-1111-1111-111111111111'; Account = 'admin@contoso.se'; AppName = $null; AuthType = 'Delegated'; Scopes = @() } }
        $p = Join-Path $TestDrive 'deleg'
        $null = Invoke-TeamsAccessReport -AllTeams -Force -IncludeUserAssociations -OutputPath $p -InformationAction SilentlyContinue -WarningVariable w -WarningAction SilentlyContinue
        ($w -join ' ') | Should -Match 'kräver app-only'
        Test-Path (Join-Path $p 'ExternalAssociations.csv') | Should -BeFalse
    }
    It 'vägrar köra om PnP.PowerShell redan är laddad' {
        Mock -ModuleName TeamsAccessReport Get-Module -ParameterFilter { $Name -eq 'PnP.PowerShell' -and -not $ListAvailable } { [pscustomobject]@{ Name = 'PnP.PowerShell' } }
        { Invoke-TeamsAccessReport @AppArgs -AllTeams -OutputPath (Join-Path $TestDrive 'pnp') } | Should -Throw -ExpectedMessage '*PnP.PowerShell är laddad*'
    }
}
