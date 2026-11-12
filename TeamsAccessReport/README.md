# TeamsAccessReport

Audit report of **actual** access in Microsoft Teams and the associated SharePoint sites:
team and channel membership (direct/indirect, cross-tenant), SharePoint permissions down to item level,
resolution of group claims to users, every sharing link with its URL and creator, and drift findings between Teams
and SharePoint (broad grants, organization/anonymous links, access outside the team, internal users who are also guests).

**Requirements:** PowerShell 7.2+ (the SharePoint worker needs 7.4.6+, PnP.PowerShell 3.0+; tested on 7.7),
Microsoft.Graph.Authentication, and PnP.PowerShell for `-IncludeSharePoint`. Read-only against the tenant; verified in a
real tenant app-only (4 teams, 4 sites, item-level permissions).

```powershell
Import-Module ./TeamsAccessReport/TeamsAccessReport.psd1
Invoke-TeamsAccessReport            # Windows: one window for connection + options, then pop-ups only
```

## Setting up a tenant (once per tenant)

`New-TeamsAccessReportApp` creates everything app-only needs, so nothing is set up by hand:

```powershell
New-TeamsAccessReportApp                                      # one tenant: taken from the admin sign-in, you confirm
New-TeamsAccessReportApp -Tenant contoso                      # one tenant, named explicitly
Import-Csv .\tenants.csv | New-TeamsAccessReportApp |          # many tenants: CSV with Tenant[,ProfileName]
    Export-Csv .\setup-results.csv
```

Per tenant you sign in once as an admin in the browser; then the command

1. creates a certificate in `CurrentUser\My` with a **non-exportable** private key (RSA 2048, SHA-256, 1-3 years),
2. registers a single-tenant app with exactly the application permissions below (IDs looked up in the tenant),
3. creates the service principal and grants admin consent,
4. saves a connection profile (SharePoint and user associations on) → `Invoke-TeamsAccessReport`, pick the profile in
   the start window, or `Invoke-TeamsAccessReport -ProfileName contoso`.

Without `-Tenant` the command signs you in first, then shows account, tenant name, initial domain and tenant ID and
creates nothing until you answer Yes (default No). The profile is named after the initial `*.onmicrosoft.com` domain.

One app and one certificate **per tenant**, so a lost certificate affects one tenant only. Each tenant produces one
result row; a failure never stops the batch:

| Status | Meaning |
|---|---|
| `Ready` | App, certificate, consent and profile are in place |
| `ConsentPending` | App kept; open `ConsentUrl` as an admin to finish |
| `Failed` | `Error` holds the reason (for example missing admin role) |
| `Cancelled` | You answered No, or closed/cancelled the sign-in window |
| `Skipped` | `-WhatIf` |

On Windows the sign-in uses Web Account Manager (WAM); its window can open **behind** the terminal. A closed or hidden
window ends as `Cancelled` – run the command again.
Required role of the admin: **Privileged Role Administrator** (consent to Microsoft Graph application permissions)
plus **Application Administrator** or **Cloud Application Administrator** – or Global Administrator. The command never
deletes anything except a certificate it created when the app itself could not be created. In the start window the
**Create app...** button next to the App ID does the same for the tenant entered there.

New permissions can take a few minutes to apply before the first report run works.

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
 ├─ Phase 2 Start-TarSharePointScan   SharePointScan.ps1 PnP   → SPAssignments/SPGroupMembers/SharingLinks.csv
 │          └─ pwsh Worker/Invoke-SharePointScan.ps1   (separate process)
 ├─ Phase 3 Invoke-TarPrincipalResolution Resolution.ps1 Graph → EffectiveSharePointAccess.csv
 ├─ Phase 4 Invoke-TarDriftAnalysis / ExternalAssociationScan  Analysis.ps1 → DriftReport/ExternalAssociations.csv
 └─ Export-TarSummary                                            → Summary.json + return value
Shared: Constants.ps1 (data contract), Store.ps1 (CSV store, checkpoints), Util.ps1 (pure functions), Logging.ps1
```

**Rules that keep the structure clean**

| Rule | Why |
|---|---|
| The report only reads. Every request is GET except SharePoint `GetSharingInformation`, a read that SharePoint only accepts as POST (`Invoke-TarSpRaw -Method Post`). The only writes are in `Provisioning.ps1` (`New-TeamsAccessReportApp`) through `Invoke-TarGraphWriteRaw` | The read-only guarantee is checkable in one place |
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
| App-only (recommended) | Team.ReadBasic.All, Channel.ReadBasic.All, TeamMember.Read.All, ChannelMember.Read.All, Files.Read.All, GroupMember.Read.All, User.Read.All, Organization.Read.All (tenant name) | Sites.FullControl.All (reading role assignments requires Full Control) or Sites.Selected + fullcontrol per site |
| Delegated | Same, but User.ReadBasic.All + User.Read | AllSites.FullControl + a dedicated PnP app (`-PnPClientId`) |

## Output (data contract)

| File | Key | Content |
|---|---|---|
| `Teams.csv` | TeamId | Inventory + `Completeness` / `FailedStages` |
| `TeamsAccess.csv` | TeamId, ChannelId, UserId | Membership, `AccessPath` (TeamMembership / DirectChannelMember / IndirectViaSharedTeam) |
| `Channels.csv` | ChannelId | Channel type, `SiteUrl` (input for phase 2) and `SiteSource` (FilesFolder / GroupDriveFallback) |
| `IncomingSharedChannels.csv` | TeamId, ChannelId | Shared channels from other teams/tenants, `CrossTenant` |
| `SPAssignments.csv` | SiteUrl, Scope, ObjectPath | Raw role assignments (SiteCollectionAdmin/Web/List/Item); `AssociatedGroup` = Owner/Member/Visitor for the site's default groups |
| `SPGroupMembers.csv` | SiteUrl, GroupTitle | Members of each SharePoint group |
| `SharingLinks.csv` | SiteUrl, ObjectPath, ShareId | One row per sharing link on an item or folder: `LinkKind`, `Access` (View/Edit/Review), `Url`, `Created`/`CreatedBy`, `LastModified`/`LastModifiedBy`, `Expiration`, `RequiresPassword`, `BlocksDownload`, `HasExternalGuestInvitees`, `Recipients`; `LinkGroup` joins to `DriftReport.csv` `DisplayName` (see below) |
| `EffectiveSharePointAccess.csv` | SiteUrl, UserId | Resolved access with `ExpansionPath` |
| `DriftReport.csv` | SiteUrl, Finding, UserId | Drift findings, `Severity` High/Medium/Low/Info, `Cause` (see [Drift findings](#drift-findings)) |
| `ExternalAssociations.csv` | UserId, HostTenantId | Internal users' teams in other tenants (`-IncludeUserAssociations`) |
| `Errors.csv` | – | Failed steps; archived to `state/` on `-Resume` |
| `Selection.json` | – | Selected teams; schema 2: `AllChannels` (bool) + `ChannelIds` (always an array) |
| `run.jsonl` | – | Structured log (one JSON object per line) |

File presence is part of the contract: **present with only a header row** = the phase ran and found 0 rows;
**absent** = the phase did not run (for example `-IncludeSharePoint` was not set). Column order is defined once in
`Private/Constants.ps1` (`$script:TarColumns`); `-Resume` refuses folders written with a different column set.

### Drift findings

Phase 4 compares, per site, who should have access (team or channel members) with who actually has it (resolved
SharePoint permissions). Limited Access rows are ignored.

| Finding | Severity | Meaning |
|---|---|---|
| `BroadGrant` | High | "Everyone" or "Everyone except external users" has access (see `Cause`) |
| `AnonymousLink` | High | An "Anyone with the link" sharing link exists |
| `OrganizationLink` | High | A "People in your organization" sharing link exists |
| `SiteCollectionAdminNotInTeams` | Medium (High for guests) | Site collection admin who is not a team/channel member |
| `SiteAccessNotInTeams` | Medium (High for guests) | Site-level access without team/channel membership |
| `ItemOrListAccessNotInTeams` | Low (High for guests) | Access only to lists or items, without membership |
| `InternalUserAlsoGuest` | Medium | Guest whose name matches an internal account on the same site (below) |
| `UnresolvedPrincipal` | Medium | A principal the resolver could not map to a user or group |
| `InTeamsNotInSharePoint` | Info | Team/channel member without site access |

For link findings `Upn` is empty and `DisplayName` holds the `SharingLinks.<id>.<kind>.<ShareId>` group – the join key
to `SharingLinks.csv` (`LinkGroup`), which has the URL and creator.

`Cause` names a known by-design explanation; empty means no default explains the finding. Severity is not lowered by a
cause – the access is real, the cause only tells you where it comes from.

| Cause | Meaning |
|---|---|
| `PublicTeamDefault` | "Everyone except external users" in the associated Members group of a **public** team site. Microsoft adds it automatically; making the team private removes it. Set on the root web and on lists/items with broken inheritance that grant that same group directly (their ACL is a copy of the parent's). A grant through any other group, a nested path, a subweb or a private team never gets this cause. |

`InternalUserAlsoGuest` flags a guest whose display name matches an internal account that has access to the **same**
site or team – typically one person with both an internal and a guest account, so offboarding and Conditional Access
only cover one of them. Names are compared without case, diacritics or the separators `. _ -` and need at least two
words (`gisela.zinn` matches `Gisela Zinn`; `admin` matches nothing). A name match is a lead, not proof: the row is the
guest; `ExpansionPath` names the internal account(s) (`NameMatch > Internal:<upn> (<id>)`) and `Scope` lists where the
guest has access (`Teams:Team`, `SharePoint:Web`, ...). Verify in Entra before removing either account.

### Sharing links

`SharingLinks.csv` lists the actual links behind the `SharingLinks.*` groups. For every item or folder whose role
assignments contain such a group, phase 2 calls SharePoint `GetSharingInformation` (`$expand=permissionsInformation`),
documented in [MS-CSOMSPT SharingInformation](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-csomspt/159d6427-e2b0-440f-bdf0-a6368ee72edd).
It needs no extra permission (the existing SharePoint Full Control) and, unlike Graph `driveItem/permissions`, returns
creator and creation date. Graph also returns link URLs only to callers that may create the link
([List permissions](https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions?view=graph-rest-1.0)).

- `LinkKind` is SharePoint's `SharingLinkKind`: `Direct`, `OrganizationView`, `OrganizationEdit`, `AnonymousView`,
  `AnonymousEdit`, `Flexible` (specific people) ([SharingLinkKind](https://learn.microsoft.com/en-us/dotnet/api/microsoft.sharepoint.client.sharinglinkkind?view=sharepoint-csom)).
- Placeholder entries (empty `ShareId` and no `Url`, i.e. link types never created) are skipped.
- `LinkGroup` is the item's `SharingLinks.<id>.<kind>.<ShareId>` principal; filter `DriftReport.csv` on
  `DisplayName = LinkGroup` to get the URL and creator of an `OrganizationLink`/`AnonymousLink` finding.
- A failure on one item is logged (WARN) and written to `Errors.csv` (`SPSharingLinks`); the scan continues.
- **The file contains working links.** An `Anonymous*` URL gives anyone who has it access without signing in. Store
  and share the output folder like credentials.

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

- Verified in one real tenant (app-only, PowerShell 7.7). `Summary.json` → `ResultStatus` and `Teams.csv` →
  `Completeness` reveal incomplete data on every run.
- `SharingLinks.csv`: that app-only calls return the link `Url` is not yet confirmed in a real tenant. Rows without a
  `Url` still carry kind, creator and dates.
- `InternalUserAlsoGuest` only sees accounts with access to a scanned site or team; duplicate guests without such access
  (for example an unused guest account) are not reported.
- PnP's error format on 429 is not verified (parsed from the error message; the Graph side uses the status code).
- Hidden lists (`Hidden eq true`) are not scanned.
- SharePoint role names are localized; `Limited Access` detection covers English and Swedish only (extend `$script:TarLimitedAccessRoles`).
