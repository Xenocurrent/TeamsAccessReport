# TeamsAccessReport

[![CI](https://github.com/Xenocurrent/TeamsAccessReport/actions/workflows/ci.yml/badge.svg)](https://github.com/Xenocurrent/TeamsAccessReport/actions/workflows/ci.yml)

PowerShell module that reports **actual** access in Microsoft Teams and the associated SharePoint sites.
Full documentation: [TeamsAccessReport/README.md](TeamsAccessReport/README.md) · Changes: [CHANGELOG](TeamsAccessReport/CHANGELOG.md)

- Team and channel membership (owner/member/guest, private and shared channels, cross-tenant)
- SharePoint permissions down to item level, with group claims resolved to users
- Every sharing link with URL, kind, creator, dates and expiry (`SharingLinks.csv`)
- Drift findings: broad grants, organization/anonymous links, access outside the team, internal users who are also guests
- One-step tenant setup (app, non-exportable certificate, admin consent, profile); read-only against the tenant

Requires PowerShell 7.2+ (7.4.6+ and PnP.PowerShell 3.0+ for the SharePoint scan), Microsoft.Graph.Authentication.

```powershell
Import-Module ./TeamsAccessReport/TeamsAccessReport.psd1
New-TeamsAccessReportApp                              # once per tenant: sign in as admin, confirm
Invoke-TeamsAccessReport -ProfileName <profile>       # or without parameters: start window (Windows)
./TeamsAccessReport/build.ps1                         # PSScriptAnalyzer + Pester + coverage (same as CI)
```

Output is one folder per run with CSV files, `Summary.json` and `run.jsonl`. It contains personal data and, in
`SharingLinks.csv`, working links – store it like credentials.

CI runs `build.ps1` on Ubuntu and Windows on every push to `main` and on pull requests.
`main` is protected: changes require a pull request with green CI on both platforms.
