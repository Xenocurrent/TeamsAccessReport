# TeamsAccessReport

Audit report of **actual** access in Microsoft Teams and the associated SharePoint sites:
team and channel membership (direct/indirect, cross-tenant), SharePoint permissions down to item level,
resolution of group claims to users, and drift between Teams and SharePoint.

```powershell
Import-Module ./TeamsAccessReport/TeamsAccessReport.psd1
Invoke-TeamsAccessReport            # Windows: one window for connection + options, then pop-ups only
```

## Connecting to a tenant

Nothing has to be edited in a command. On **Windows**, `Invoke-TeamsAccessReport` without connection parameters opens
one window with everything an interactive run needs – no console input:

| Section | Fields |
|---|---|
| Connection | Saved connection (newest pre-selected) · Tenant + **Check** · App-only / Delegated · App (client) ID · Certificate (drop-down from the store, or type a thumbprint) · optional PnP app ID · Save/Update profile |
| Report | SharePoint (phases 2-4) · users' teams in other tenants · choose channels per team · output folder + **Browse** |

**Start** validates everything at once (tenant lookup, GUID, certificate present with private key, output folder
unused) and lists all problems in the window instead of failing later. Tenant confirmation follows as a pop-up; team
and channel selection use a check-box list (filter, Select all / Clear on visible rows, "n of m selected"; filtering
never drops a selection, so large tenants can be narrowed down step by step). Report options are remembered per profile; switches on the command line win.
`New-TeamsAccessReportProfile` opens the same window without the report section.

Elsewhere (Linux/macOS, or `-SelectionMode Menu`) the same questions are asked in the console. Saved connections,
**New connection…** or **Sign in as myself** are offered first; a new connection asks for:

1. **Tenant** – type it the way you know it: `contoso`, `contoso.com`, `admin@contoso.com` or the tenant ID. It is
   resolved to the tenant ID through Microsoft's public OpenID discovery document (no sign-in) and shown for confirmation.
2. **Sign-in method** – app-only (certificate) or delegated.
3. **App ID** and **certificate** (app-only) – the certificate is picked from `CurrentUser\My` / `LocalMachine\My`.
4. **Save as profile?** – next time: `Invoke-TeamsAccessReport -ProfileName contoso` (tab-completes).

Pickers are pop-ups (Out-GridView) on Windows and numbered menus elsewhere; `-SelectionMode Menu` forces the console.

| Command | Purpose |
|---|---|
| `New-TeamsAccessReportProfile` | Create a profile (interactive, or fully with `-Tenant -ClientId -CertificateThumbprint`) |
| `Get-TeamsAccessReportProfile` | List profiles |
| `Remove-TeamsAccessReportProfile -Name x` | Delete a profile (the app and certificate are not touched) |
| `New-TeamsAccessReportPnPApp -ProfileName x` | Create the PnP app for **delegated** SharePoint scanning and store it in the profile (also the **Create...** button in the window). Not needed for app-only. |

Profiles hold identifiers only – tenant ID, app ID, certificate thumbprint – never secrets; the private key stays in the
certificate store. Location: `%LOCALAPPDATA%\TeamsAccessReport\profiles.json` (Windows) or
`~/.local/share/TeamsAccessReport/profiles.json`; override with `TEAMSACCESSREPORT_PROFILE_PATH`.

Unattended runs still take explicit parameters, and `-TenantId` accepts the same short forms:

```powershell
$s = Invoke-TeamsAccessReport -ProfileName contoso -AllTeams -Force -IncludeSharePoint -IncludeUserAssociations
$s = Invoke-TeamsAccessReport -TenantId contoso -ClientId '<APP_ID>' -CertificateThumbprint '<THUMBPRINT>' -AllTeams -Force
$s.ResultStatus   # Complete | Incomplete
```

`-Resume` needs no connection parameters: it reuses the profile recorded in `Selection.json` (or the only saved profile
for that tenant). Run `Get-Help Invoke-TeamsAccessReport -Full` for every parameter.

## Architecture

```
Invoke-TeamsAccessReport (Public)          orchestration, validation, try/finally
 ├─ Resolve-TarConnection       Connection.ps1 profile / parameters / start screen → tenant ID, app, certificate
 ├─ Connect-TarGraph            Graph.ps1      sign-in, tenant discovery, SDK retry
 ├─ Get/Import-TarTeamSelection Selection.ps1  interactive selection (UI isolated from logic)
 ├─ Phase 1 Invoke-TarTeamInventory   Inventory.ps1      Graph → Teams/TeamsAccess/Channels/Incoming.csv
 ├─ Phase 2 Start-TarSharePointScan   SharePointScan.ps1 PnP   → SPAssignments/SPGroupMembers.csv
 │          └─ pwsh Worker/Invoke-SharePointScan.ps1   (separate process)
 ├─ Phase 3 Invoke-TarPrincipalResolution Resolution.ps1 Graph → EffectiveSharePointAccess.csv
 ├─ Phase 4 Invoke-TarDriftAnalysis / ExternalAssociationScan  Analysis.ps1 → DriftReport/ExternalAssociations.csv
 └─ Export-TarSummary                                            → Summary.json + return value
Shared: Constants.ps1 (data contract), Store.ps1 (CSV store, checkpoints), Util.ps1 (pure functions), Logging.ps1
```

**Rules that keep the structure clean**

| Rule | Why |
|---|---|
| All Graph IO goes through `Invoke-TarGraphRaw`; all SharePoint IO through `Invoke-TarSpRaw`; tenant discovery through `Invoke-TarHttpRaw` | One seam per external service; tests mock only these |
| No script-scope variables in logic; everything flows through `-Context` | Explicit dependencies, testable, PSScriptAnalyzer can verify parameter usage |
| Phases communicate only through the files defined in `Constants.ps1` (`$script:TarFile`) | Each phase can be re-run or resumed in isolation; the files are a data contract |
| Functions returning collections emit elements to the pipeline; callers wrap with `@(...)` | Avoids PowerShell's unary-comma/`$null` pitfalls (three v1 bugs came from here) |
| UI lives only in `Selection.ps1` (console) and `Gui.ps1` (Windows Forms); validation of the window is `Test-TarStartFormInput` in `Connection.ps1`; `Test-TarCanPrompt` is the only console check | Everything else runs unattended and is tested without a console |
| The worker never loads the Graph SDK | Known PnP ↔ Graph assembly conflict; a unit test verifies this at token level |

## Design decisions (short ADRs)

1. **Separate process for PnP.** The Graph SDK and PnP.PowerShell load incompatible assemblies in the same session. Alternatives (PnP only or Graph only) were rejected: Graph does not expose SharePoint role assignments, and `inheritedFrom` is not returned for SharePoint.
2. **CSV as intermediate store instead of in-memory objects.** Enables checkpoint/resume and auditability. Cost: schemas are implicit → centralized in `Constants.ps1` and guarded by an integration test (byte-identical comparison).
3. **Two retry layers.** The SDK handler (429/503/504, `Retry-After`) is configured via `Set-MgRequestContext`; an outer loop (max 3) handles 500/502 and exhausted SDK retries. Upper bound: 3 × (MaxRetries + 1) attempts per call.
4. **Sequential execution.** Predictable throttling and simple checkpoints. Parallelism waits for measurements from a real tenant.

## Permissions

| Mode | Graph | SharePoint (`-IncludeSharePoint`) |
|---|---|---|
| App-only (recommended) | Team.ReadBasic.All, Channel.ReadBasic.All, TeamMember.Read.All, ChannelMember.Read.All, Files.Read.All, GroupMember.Read.All, User.Read.All | Sites.FullControl.All (reading role assignments requires Full Control) or Sites.Selected + fullcontrol per site |
| Delegated | Same, but User.ReadBasic.All + User.Read | AllSites.FullControl + a dedicated PnP app (`-PnPClientId`) |

## Output (data contract)

| File | Key | Content |
|---|---|---|
| `Teams.csv` | TeamId | Inventory + `Completeness` / `FailedStages` |
| `TeamsAccess.csv` | TeamId, ChannelId, UserId | Membership, `AccessPath` (TeamMembership / DirectChannelMember / IndirectViaHostTeam / IndirectViaSharedTeam) |
| `Channels.csv` | ChannelId | Channel type, `SiteUrl` (input for phase 2) and `SiteSource` (FilesFolder / GroupDriveFallback) |
| `IncomingSharedChannels.csv` | TeamId, ChannelId | Shared channels from other teams/tenants, `CrossTenant` |
| `SPAssignments.csv` | SiteUrl, Scope, ObjectPath | Raw role assignments (SiteCollectionAdmin/Web/List/Item); `AssociatedGroup` = Owner/Member/Visitor for the site's default groups |
| `SPGroupMembers.csv` | SiteUrl, GroupTitle | Members of each SharePoint group |
| `EffectiveSharePointAccess.csv` | SiteUrl, UserId | Resolved access with `ExpansionPath` |
| `DriftReport.csv` | SiteUrl, Finding, UserId | Drift findings, `Severity` High/Medium/Low/Info, `Cause` (see below) |
| `ExternalAssociations.csv` | UserId, HostTenantId | Internal users' teams in other tenants (`-IncludeUserAssociations`) |
| `Errors.csv` | – | Failed steps; archived to `state/` on `-Resume` |
| `Selection.json` | – | Selected teams; schema 2: `AllChannels` (bool) + `ChannelIds` (always an array) |
| `run.jsonl` | – | Structured log (one JSON object per line) |

File presence is part of the contract: **present with only a header row** = the phase ran and found 0 rows;
**absent** = the phase did not run (for example `-IncludeSharePoint` was not set). Column order is defined once in
`Private/Constants.ps1` (`$script:TarColumns`); `-Resume` refuses folders written with a different column set.

`Cause` in `DriftReport.csv` names a known by-design explanation; empty means no default explains the finding.
Severity is not lowered by a cause – the access is real, the cause only tells you where it comes from.

| Cause | Meaning |
|---|---|
| `PublicTeamDefault` | "Everyone except external users" in the associated Members group of a **public** team site. Microsoft adds it automatically; making the team private removes it. All conditions must hold (team site, public, root web, associated Members group), so a manually added grant is never attributed to the default. |

## Development

```powershell
./build.ps1           # PSScriptAnalyzer (0 findings allowed) + Pester + coverage ≥ 80 %
./build.ps1 -Format   # format according to PSScriptAnalyzerSettings.psd1 first
```

Tests: `Tests/Unit.Tests.ps1` (pure functions, claim parsing, drift logic) and `Tests/Integration.Tests.ps1`
(the full flow through the adapter seams; interruption + `-Resume` is compared byte for byte with a full run).
Phase 2 runs `InProcess` in tests – the process boundary is covered by a smoke test, not by Pester.
The Windows Forms window is built (not shown) and exercised in tests on Windows only; `Gui.ps1` is excluded from the
coverage figure on other platforms.

Contributing: create a branch, open a pull request, and wait for green CI on both platforms (`main` is protected).

## Known limitations

- Not yet verified against a real tenant. `Summary.json` → `ResultStatus` and `Teams.csv` → `Completeness` reveal incomplete data.
- PnP's error format on 429 is not verified (parsed from the error message; the Graph side uses the status code).
- Hidden lists (`Hidden eq true`) are not scanned.
- Delegated PnP sign-in may prompt per site (hypothesis) – use app-only.
- SharePoint role names are localized; `Limited Access` detection covers English and Swedish only (extend `$script:TarLimitedAccessRoles`).
