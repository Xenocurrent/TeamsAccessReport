# TeamsAccessReport

[![CI](https://github.com/Xenocurrent/TeamsAccessReport/actions/workflows/ci.yml/badge.svg)](https://github.com/Xenocurrent/TeamsAccessReport/actions/workflows/ci.yml)

PowerShell module that reports **actual** access in Microsoft Teams and the associated SharePoint sites.
Full documentation: [TeamsAccessReport/README.md](TeamsAccessReport/README.md) · Changes: [CHANGELOG](TeamsAccessReport/CHANGELOG.md)

```powershell
Import-Module ./TeamsAccessReport/TeamsAccessReport.psd1
Invoke-TeamsAccessReport            # interactive
./TeamsAccessReport/build.ps1       # PSScriptAnalyzer + Pester + coverage (same as CI)
```

CI runs `build.ps1` on Ubuntu and Windows on every push to `main` and on pull requests.
`main` is protected: changes require a pull request with green CI on both platforms.
