# TeamsAccessReport

[![CI](https://github.com/Xenocurrent/TeamsAccessReport/actions/workflows/ci.yml/badge.svg)](https://github.com/Xenocurrent/TeamsAccessReport/actions/workflows/ci.yml)

PowerShell-modul som rapporterar faktisk åtkomst i Microsoft Teams och tillhörande SharePoint-webbplatser.
Fullständig dokumentation: [TeamsAccessReport/README.md](TeamsAccessReport/README.md) · Ändringar: [CHANGELOG](TeamsAccessReport/CHANGELOG.md)

```powershell
Import-Module ./TeamsAccessReport/TeamsAccessReport.psd1
Invoke-TeamsAccessReport            # interaktivt
./TeamsAccessReport/build.ps1       # PSScriptAnalyzer + Pester + täckning (samma som CI)
```

CI kör `build.ps1` på Ubuntu och Windows vid push till `main` och vid pull requests.
