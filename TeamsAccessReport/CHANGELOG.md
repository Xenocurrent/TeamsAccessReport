# Changelog

## 2.5.0 – 2026-11-06
### Added
- `New-TeamsAccessReportApp`: sets up a tenant in one step - non-exportable certificate in CurrentUser\My, single-tenant
  app with exactly the report's application permissions (IDs looked up in the tenant), service principal, admin consent
  and a connection profile. Accepts many tenants via the pipeline (CSV), one result row per tenant, failures isolated.
  Retries throttling and directory replication delays; if consent fails the app is kept and a ConsentUrl is returned.
- Start window: **Create app...** button next to the App ID runs the same setup in a separate window and fills in the
  app ID and certificate.
- `Organization.Read.All` is part of the app-only permission set (tenant name in the report).

### Changed
- Permission names live once in `$script:TarAppPermissions` (scope check and setup use the same list).
- Worker windows share one seam (`Invoke-TarWorkerWindow`).

## 2.4.2 – 2026-11-06
### Fixed
- The owner of a shared channel was labelled `IndirectViaHostTeam` (added in 2.4.1). In the real tenant Graph sets
  `originalSourceMembershipUrl` to the host team for direct owners of both private and shared channels, so a host-team
  source is now `DirectChannelMember`. `IndirectViaHostTeam` is removed; only a source in another team is
  `IndirectViaSharedTeam`.

## 2.4.1 – 2026-11-06
Found in the TAR-Test run against a real tenant.
### Fixed
- Shared channels were reported as `unknownFutureValue`: Graph only returns `shared` with the request header
  `Prefer: include-unknown-enum-members`, which is now sent on every Graph call. As a consequence the shared channel's
  site was treated as a team site (false `InTeamsNotInSharePoint`) and its members were not read. Any channel type
  other than `standard` is now handled as a channel with its own site and member list, with a WARN for unknown types.
- A private channel member was labelled `IndirectViaSharedTeam` because Graph sets `originalSourceMembershipUrl` to
  the host team. Private channel members are now always `DirectChannelMember`; shared channels distinguish
  `IndirectViaHostTeam` (new) from `IndirectViaSharedTeam`.
- The token principal of a sharing link (`c:0u.c|tenant|…`, title `SLinkClaim.*`) was reported twice as
  `UnresolvedPrincipal`. It is now `SharingLinkClaim` and not a separate finding (the link itself is reported).
- `Web-Only Limited Access` is treated as limited access like `Limited Access`.

## 2.4.0 – 2026-11-05
### Added
- Guided PnP app setup for delegated SharePoint scanning. The start window has a **Create...** button next to the PnP
  app ID with a guide (why it is needed, exactly what is registered, requirements) and a one-click registration:
  `Register-PnPEntraIDAppForInteractiveLogin` runs in a separate PowerShell window (`Worker/Register-PnPApp.ps1`) with
  one delegated permission (SharePoint `AllSites.FullControl`), and the client ID is filled in and saved with the profile.
- `New-TeamsAccessReportPnPApp -ProfileName <name>` does the same from the console and updates the profile.
- A hint under the PnP field says when it is needed (delegated + SharePoint) and when not (app-only).

## 2.3.0 – 2026-11-05
### Added
- Windows check-box list for team and channel selection (replaces Out-GridView multi-select): a check box per row,
  click anywhere on a row to toggle, filter that keeps selections, Select all / Clear on the visible rows,
  "n of m selected (k hidden by filter)", OK disabled while nothing is selected. Built for large tenants: cell texts
  are computed once and the list is filled in one batch.
- The channel list starts with every channel checked (the default is all channels).

### Fixed
- Certificate text in the start window was cut off: shorter text without `CN=` and a wider drop-down.

## 2.2.0 – 2026-11-05
### Added
- Windows start window (Windows Forms): connection (saved profile, tenant with live check, sign-in method, app ID,
  certificate drop-down, optional PnP app ID, save/update profile) and report options (SharePoint, user associations,
  channel selection, output folder with folder picker) in one place. All errors are validated and shown at once.
- Tenant confirmation and the remaining yes/no questions are pop-ups in Gui mode. With the window, an interactive run
  needs no console input.
- Profiles remember the last report options (`Options`); command-line switches still win.
- `New-TeamsAccessReportProfile` uses the same window (connection part only) on Windows.

### Changed
- `build.ps1` measures `Gui.ps1` coverage on Windows only.

## 2.1.0 – 2026-11-05
### Added
- Connection profiles: `New-TeamsAccessReportProfile`, `Get-TeamsAccessReportProfile`, `Remove-TeamsAccessReportProfile`
  and `Invoke-TeamsAccessReport -ProfileName` (tab completion). Profiles store identifiers only, never secrets.
- Start screen when `Invoke-TeamsAccessReport` runs without connection parameters: saved connection, new connection
  (tenant, sign-in method, app ID, certificate picked from the store, optional save) or delegated sign-in.
- Tenant input accepts `contoso`, `contoso.com`, `admin@contoso.com` or a tenant ID; it is resolved to the tenant ID via
  the public OpenID discovery document (3 attempts with backoff + jitter). If the service is unreachable the input is
  passed on to the sign-in with a WARN; an unknown tenant fails immediately.
- App-only certificates are checked before sign-in (present, private key, valid now) with a precise error.
- `-Resume` without connection parameters reuses the profile recorded in `Selection.json`.

### Changed
- `Invoke-TeamsAccessReport` has no parameter sets and no mandatory parameters; combinations are validated in code.
  Existing explicit calls (`-TenantId -ClientId -CertificateThumbprint`) work unchanged.
- `-ClientId` must be a GUID and `-CertificateThumbprint` 40 hex characters; spaces and invisible characters copied
  from certmgr are removed.
- `Selection.json` records `ProfileName`; `RerunCommand` uses `-ProfileName` instead of IDs when a profile was used.
- Log entries from a new command no longer append to the previous run's `run.jsonl` in the same session.

## 2.0.4 – 2026-11-05
### Added
- `DriftReport.csv` has a new column `Cause`. `PublicTeamDefault` marks "Everyone except external users" in the
  associated Members group of a public team site – a Microsoft default, not a misconfiguration. Severity stays High.
- `SPAssignments.csv` and `EffectiveSharePointAccess.csv` have a new column `AssociatedGroup` (Owner/Member/Visitor),
  read once per site from `_api/web?$expand=AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup`.
  If that call fails the scan continues with a WARN and `Cause` stays empty for the site.

### Changed
- CSV files of phases that ran are always written, with only a header row when there are 0 rows (previously
  `ExternalAssociations.csv`, `Errors.csv` and others were missing, which was indistinguishable from "not run").
  Files of phases that did not run remain absent. Column order lives in `$script:TarColumns`.
- `Selection.json` schema 2: `AllChannels` (bool) and `ChannelIds` (always an array; previously a single selected
  channel was written as a string). Schema 1 files are still imported.
- `-Resume` checks existing CSV headers and refuses folders written by an incompatible version with a clear message.
  Runs from 2.0.3 with `-IncludeSharePoint` cannot be resumed with 2.0.4; start a new run.

## 2.0.3 – 2026-11-05
### Fixed
- Graph returns HTTP 404 "Folder location for this channel is not ready yet" for channels whose files folder has
  not been provisioned (found in the first real-tenant run). Standard channels now fall back to the team's document
  library (`/groups/{id}/drive/root` → `sharepointIds.siteUrl`), because standard channels always live on the team site.
  Private/shared channels fail with an actionable message (open the Files tab once, then `-Resume`).
- Graph errors carry `StatusCode` and `GraphMessage` in `Exception.Data` so callers branch without parsing text.

### Changed
- `Channels.csv` has a new column `SiteSource` (`FilesFolder` | `GroupDriveFallback`). Runs started with 2.0.2 or
  earlier cannot be resumed with 2.0.3; start a new run.

## 2.0.2 – 2026-11-05
### Fixed
- `Get-TarProperty` failed on generic `Dictionary<TKey,TValue>` (the type the Graph SDK returns for response
  headers) with "Cannot find an overload for Contains". This broke every Graph call against a real tenant.
  Root cause: `IDictionary.Contains` is an explicit interface implementation; the value is now cast inline at the call.
- Header lookup is now case-insensitive.

### Tests
- New unit tests for `Get-TarProperty` with Hashtable, generic Dictionary and PSCustomObject.
- Retry-After parsing extracted to the pure function `ConvertFrom-TarRetryAfterHeader`, tested with the SDK's real
  header type (the integration tests mock the whole adapter and therefore never exercised it).

## 2.0.1 – 2026-11-05
### Changed
- All code comments, log messages, error messages, prompts, tests and documentation are now in English.
- Interactive yes/no prompts accept `y`/`yes`/`n`/`no` (previously Swedish `j`/`ja`/`nej` as well).
- Interactive selection columns renamed: `Name`, `Visibility`, `Description`, `Channel`, `Type`.
- Full comment-based help for every parameter of `Invoke-TeamsAccessReport`.

### Not changed
- The SharePoint role name `Begränsad åtkomst` remains in `$script:TarLimitedAccessRoles`. It is a data value
  (the Swedish localization of "Limited Access" returned by SharePoint), not UI text.

## 2.0.0 – 2026-11-05
Restructured from a monolithic script (930 lines, 26 functions, ~450 lines of logic in the script body) into a module.

### Changed
- Module `TeamsAccessReport` with a single public function `Invoke-TeamsAccessReport`; `Get-TeamsAccessReport.ps1` remains as a thin wrapper (`@args`).
- An explicit `$Context` object replaces implicit script-scope variables (PSScriptAnalyzer previously flagged `-SelectionMode` as unused because of this).
- Parameter sets `Delegated` / `AppOnly` instead of manual validation of ClientId/Thumbprint/TenantId.
- Graph status codes are read deterministically via `-SkipHttpErrorCheck -StatusCodeVariable` (previously: parsing error messages).
- The phase 2 worker is its own file (`Worker/Invoke-SharePointScan.ps1`) instead of a hidden parameter mode in the main script.
- Structured log `run.jsonl` (JSON Lines). No `Write-Host` outside the interactive menu.
- `-WhatIf`/`-Confirm` supported (`SupportsShouldProcess`).
- The function returns the summary object (pipeline friendly) in addition to the files.

### Fixed (found during review)
- Worker process stdout leaked into the output stream (9 junk objects before the summary).
- `$x = if (...) { @() }` produced `$null` → drift analysis crashed for sites without matching Teams users.
- Empty strings vs `$null` were written differently in CSV → a resumed run was not byte-identical.
- Test comparisons of missing files produced a false green (two empty strings).

### Quality gate
- `build.ps1`: PSScriptAnalyzer (0 findings required) + Pester (68 tests) + code coverage ≥ 80 %.
