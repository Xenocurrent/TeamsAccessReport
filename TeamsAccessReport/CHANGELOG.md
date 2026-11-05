# Changelog

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
