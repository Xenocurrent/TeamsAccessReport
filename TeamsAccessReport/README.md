# TeamsAccessReport

Revisionsrapport över **faktisk** åtkomst i Microsoft Teams och tillhörande SharePoint-webbplatser:
team- och kanalmedlemskap (direkt/indirekt, cross-tenant), SharePoint-behörigheter ner på objektnivå,
upplösning av grupp-claims till användare och avvikelser mellan Teams och SharePoint.

```powershell
Import-Module ./TeamsAccessReport/TeamsAccessReport.psd1
$s = Invoke-TeamsAccessReport -TenantId '<TENANT_ID>' -ClientId '<APP_ID>' -CertificateThumbprint '<THUMBPRINT>' `
        -AllTeams -Force -IncludeSharePoint -IncludeUserAssociations
$s.ResultStatus   # Complete | Incomplete
```

## Arkitektur

```
Invoke-TeamsAccessReport (Public)          orkestrering, validering, parameter sets, try/finally
 ├─ Connect-TarGraph            Graph.ps1      inloggning, tenant-identifiering, SDK-retry
 ├─ Get/Import-TarTeamSelection Selection.ps1  interaktivt urval (UI isolerat från logik)
 ├─ Fas 1 Invoke-TarTeamInventory   Inventory.ps1     Graph  → Teams/TeamsAccess/Channels/Incoming.csv
 ├─ Fas 2 Start-TarSharePointScan   SharePointScan.ps1 PnP  → SPAssignments/SPGroupMembers.csv
 │        └─ pwsh Worker/Invoke-SharePointScan.ps1   (separat process)
 ├─ Fas 3 Invoke-TarPrincipalResolution Resolution.ps1 Graph → EffectiveSharePointAccess.csv
 ├─ Fas 4 Invoke-TarDriftAnalysis / ExternalAssociationScan  Analysis.ps1 → DriftReport/ExternalAssociations.csv
 └─ Export-TarSummary                                            → Summary.json + returvärde
Delade: Constants.ps1 (datakontrakt), Store.ps1 (CSV-lager, checkpoint), Util.ps1 (rena funktioner), Logging.ps1
```

**Regler som håller strukturen ren**

| Regel | Varför |
|---|---|
| All IO mot Graph går via `Invoke-TarGraphRaw`, all IO mot SharePoint via `Invoke-TarSpRaw` | En söm per extern tjänst; tester mockar bara dessa |
| Inga script-globala variabler i logik; allt via `-Context` | Explicita beroenden, testbart, PSSA kan verifiera parameteranvändning |
| Faserna kommunicerar endast via filer i `Constants.ps1` (`$script:TarFile`) | Varje fas kan köras om/återupptas isolerat; filerna är ett datakontrakt |
| Funktioner som returnerar samlingar skickar element i pipeline; anropare wrappar `@(...)` | Undviker PowerShells unary-comma/`$null`-fällor (tre buggar i v1 kom härifrån) |
| UI (`Read-Host`, meny) endast i `Selection.ps1` | Resten körs oövervakat och testas utan konsol |
| Worker laddar aldrig Graph SDK | Känd assemblykonflikt PnP ↔ Graph; enhetstest verifierar detta på tokennivå |

## Designbeslut (ADR-kort)

1. **Separat process för PnP.** Graph SDK och PnP.PowerShell laddar inkompatibla assemblies i samma session. Alternativ (bara PnP eller bara Graph) förkastade: Graph saknar SharePoint-rollkopplingar, och `inheritedFrom` returneras inte för SharePoint.
2. **CSV som mellanlager i stället för minnesobjekt.** Ger checkpoint/återupptag och granskningsbarhet. Kostnad: scheman är implicita → samlade i `Constants.ps1` och skyddade av integrationstest (byte-identisk jämförelse).
3. **Två retry-lager.** SDK-hanteraren (429/503/504, `Retry-After`) konfigureras via `Set-MgRequestContext`; yttre loop (max 3) tar 500/502 och uttömd SDK-retry. Totalt tak: 3 × (MaxRetries + 1) försök per anrop.
4. **Sekventiell körning.** Förutsägbar throttling och enkel checkpoint. Parallellism avvaktar mätdata från riktig tenant.

## Behörigheter

| Läge | Graph | SharePoint (`-IncludeSharePoint`) |
|---|---|---|
| App-only (rekommenderas) | Team.ReadBasic.All, Channel.ReadBasic.All, TeamMember.Read.All, ChannelMember.Read.All, Files.Read.All, GroupMember.Read.All, User.Read.All | Sites.FullControl.All (läsning av rollkopplingar kräver Full Control) eller Sites.Selected + fullcontrol per webbplats |
| Delegerat | Samma men User.ReadBasic.All + User.Read | AllSites.FullControl + egen PnP-app (`-PnPClientId`) |

## Utdata (datakontrakt)

| Fil | Nyckel | Innehåll |
|---|---|---|
| `Teams.csv` | TeamId | Inventering + `Completeness` / `FailedStages` |
| `TeamsAccess.csv` | TeamId, ChannelId, UserId | Medlemskap, `AccessPath` (TeamMembership / DirectChannelMember / IndirectViaSharedTeam) |
| `Channels.csv` | ChannelId | Kanaltyp och `SiteUrl` (källa för fas 2) |
| `SPAssignments.csv` | SiteUrl, Scope, ObjectPath | Rå rollkopplingar (SiteCollectionAdmin/Web/List/Item) |
| `EffectiveSharePointAccess.csv` | SiteUrl, UserId | Upplöst åtkomst med `ExpansionPath` |
| `DriftReport.csv` | SiteUrl, Finding, UserId | Avvikelser, `Severity` High/Medium/Low/Info |
| `Errors.csv` | – | Misslyckade steg; arkiveras till `state/` vid `-Resume` |
| `run.jsonl` | – | Strukturerad logg (en JSON per rad) |

## Utveckling

```powershell
./build.ps1           # PSScriptAnalyzer (0 tillåtna) + Pester + täckning ≥ 80 %
./build.ps1 -Format   # formatera enligt PSScriptAnalyzerSettings.psd1 först
```

Tester: `Tests/Unit.Tests.ps1` (rena funktioner, claim-parsning, drift-logik), `Tests/Integration.Tests.ps1`
(hela flödet via adaptersömmarna, avbrott + `-Resume` jämförs byte för byte mot hel körning).
Fas 2 körs `InProcess` i tester – processgränsen täcks av röktest, inte Pester.

## Kända begränsningar

- Inte verifierat mot riktig tenant. `Summary.json` → `ResultStatus` och `Teams.csv` → `Completeness` visar ofullständig data.
- PnP:s felformat vid 429 är inte verifierat (tolkas ur felmeddelande; Graph-sidan använder statuskod).
- Dolda listor (`Hidden eq true`) skannas inte.
- Delegerad PnP-inloggning kan fråga per webbplats (hypotes) – använd app-only.
