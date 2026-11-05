# Changelog

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
