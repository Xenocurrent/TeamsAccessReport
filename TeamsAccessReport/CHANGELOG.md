# Changelog

## 2.0.0 – 2026-11-05
Omstrukturering från monolitiskt skript (930 rader, 26 funktioner, ~450 rader logik i skriptkroppen) till modul.

### Ändrat
- Modul `TeamsAccessReport` med en publik funktion `Invoke-TeamsAccessReport`; `Get-TeamsAccessReport.ps1` finns kvar som tunn wrapper (`@args`).
- Explicit `$Context`-objekt ersätter implicita script-globala variabler (PSSA flaggade tidigare `-SelectionMode` som oanvänd p.g.a. detta).
- Parameter sets `Delegated` / `AppOnly` i stället för manuell validering av ClientId/Thumbprint/TenantId.
- Graph-statuskod läses deterministiskt via `-SkipHttpErrorCheck -StatusCodeVariable` (tidigare: tolkning av felmeddelanden).
- Fas 2-workern är en egen fil (`Worker/Invoke-SharePointScan.ps1`) i stället för ett dolt parameter-läge i huvudskriptet.
- Strukturerad logg `run.jsonl` (JSON Lines). Inga `Write-Host` utanför interaktiv meny.
- `-WhatIf`/`-Confirm` stöds (`SupportsShouldProcess`).
- Funktionen returnerar sammanfattningsobjektet (pipeline-vänligt) utöver filerna.

### Rättat (hittat under granskning)
- Worker-processens stdout läckte in i returströmmen (9 skräpobjekt före sammanfattningen).
- `$x = if (...) { @() }` gav `$null` → avvikelseanalysen kraschade för webbplatser utan matchande Teams-användare.
- Tomma strängar vs `$null` skrevs olika i CSV → återupptagen körning var inte byte-identisk.
- Testjämförelser av saknade filer gav falskt grönt (två tomma strängar).

### Kvalitetsgrind
- `build.ps1`: PSScriptAnalyzer (0 anmärkningar krävs) + Pester (68 tester) + kodtäckning ≥ 80 %.
