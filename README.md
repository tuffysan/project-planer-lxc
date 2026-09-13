# Project Planer LXC v5.0.0 – Intelligent Project Platform

This is a cumulative release. It contains all functionality from the preceding roadmap releases.

## Added in v5.0.0

- Intelligent Project Manager
- natural-language style project questions using local project data
- go-live blocker analysis
- risk summaries
- weekly change summaries
- meeting notes → actions/decisions
- assistant query history
- no external AI data transfer by default

## Upgrade an existing LXC

```bash
CTID=<your-ctid> VERSION=9.0.2 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## New install

```bash
VERSION=9.0.2 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Run `./VERIFY.sh` before publishing. Keep a database backup before production upgrades.

## Navigation hotfix

Rebuilt navigation: no dynamic `undefined` labels, compact dropdown navigation, responsive header and account menu.

## v5.0.1 hotfix

This release fixes the HTTP 500 error on the start page introduced in v5.0.0.

Cause: the global navigation rendered `url_for('delivery_center')`, while the `delivery_center` route requires `project_id`.

Fixes:
- removes the invalid global Delivery link
- exposes Delivery only from project context where `project_id` is available
- adds `VERIFY-PYTHON.py` to detect template `url_for()` calls missing required route parameters
- adds `SMOKE-TEST.sh` for post-upgrade HTTP/journal verification


## v5.0.2 – Full App Stability & Responsive Edition

Full static audit of the application with fixes for Project Control, automation notifications,
intelligence action creation, archive filtering and Report Pack. Mobile/tablet styling has also
been hardened across the app.

Run before publishing:

```bash
./VERIFY.sh
python3 VERIFY-PYTHON.py
```

After updating an LXC:

```bash
/opt/project-plan/current/SMOKE-TEST.sh
```

## v5.0.3 – Visual UX Edition

Adds RAG project health, portfolio visual dashboard, visual project cockpit, health score, progress bars, status chips, attention panels, milestones and responsive visual components for desktop, tablet and mobile.

## v5.0.4 – Visual UX Refinement

Second visual UX pass: fixes dashboard data binding, makes RAG status visible directly on the main dashboard, adds portfolio distribution, status reasons, trend indicators, search/filter/sort and a mobile bottom navigation.


## v7.0.1 Full Quality Audit

See `AUDIT-v7.0.1.md`. This release focuses on permissions, database reliability, startup migrations, health checks, static validation and responsive hardening rather than new features.


## Excel Round-trip (v8.1.0)

From Project Workspace, choose **Exportera Excel** or **Importera Excel**.

The full workbook supports offline editing for project information, tasks, risks, change requests, resources, costs, task dependencies, decisions, meetings, actions and benefits. Import always runs through a preview. Stable hidden IDs identify existing records; snapshot hashes detect concurrent changes. Missing spreadsheet rows do not delete data. A database backup is created before commit.

## Excel Pro (v9.0.2)

Excel Center now offers three explicit workflows: a round-trip Project Workbook, a presentation-oriented Status Report workbook, and a Portfolio workbook. The project workbook includes a visual overview, formatted Excel tables, editable-field highlighting, filters, validations, conditional formatting, charts, print setup and hidden technical metadata for safe re-import.


## Senaste versionen automatiskt

Från och med v14.0.3 behöver du normalt inte ange versionsnummer.

### Smart deploy – rekommenderad

Samma kommando installerar om CTID inte finns och uppgraderar om CTID redan finns:

```bash
CTID=200 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/deploy-lxc.sh)"
```

### Uppgradera befintlig CT till senaste release

```bash
CTID=200 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

### Ny installation av senaste release

```bash
CTID=200 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

`VERSION=x.y.z` kan fortfarande anges om en specifik release ska installeras.

## v15.0.0 – Professional PPM Edition

v15 samlar kärnarbetet kring fyra huvudytor: **Idag**, **Mitt arbete**, **Projekt** och **Portfolio**.

Nytt i v15 är bland annat Action Inbox, nytt Project Cockpit, leveransprognos,
prioriterad Portfolio Cockpit och What-if-påverkansanalys för aktivitetsberoenden.


## v15.1.0 – Planning & Resource Engine

Planeringsmotor för beroenden/kritisk linje samt kapacitets-heatmap och samlad projekthälsa.


## v15.2.0 – Controlled Planning Edition

Rättad språkväxlare, kontrollerad automatisk omplanering med undo/audit, resursutjämning och explicit Azure DevOps tvåvägssynk.


## v15.2.1 – Language Encoding Fix

Rättar grundorsaken till den trasiga språkmenyn i Windows-publishflödet och ersätter emoji-flaggor med stabila språkbadges (SV/EN/DE/NO/DA/FI).


## v15.2.2 – Complete Excel Template

Nytt val **Komplett Excel-mall** exporterar en enda arbetsbok med hela projektets planerings-, styrnings-, leverans-, resurs-, ekonomi-, integrations- och historikdata. Round-trip-kompatibla blad kan fortsatt importeras tillbaka via Excel-importen.


## v15.2.3 – Create Project Fix

Fixar HTTP 500 på **Skapa projekt** genom att rätta den felaktiga CSRF-renderingen i äldre mallar. Skapaflödet är dessutom ombyggt med v15-UX, validering, automatisk projektroll och audit-logg.


## v15.2.4 – Excel Discoverability

Excel är nu en förstaklassfunktion i projektvyn med tydlig **Excel ▾**-meny, permanent Excel-flik, snabbpanel och rekommenderad komplett Excel-export högst upp i Excel Center.


## v15.2.5 – Excel Start Center

Excel kan nu användas utan ett befintligt projekt. Huvudmenyn har en permanent **Excel**-genväg för tom komplett mall och för att importera en ifylld mall som ett nytt projekt.


## v15.2.6 – Excel Navigation Fix

Den globala **Excel**-genvägen ligger nu direkt i huvudnavigationen mellan **Portfolio** och **Sök** och visas för alla inloggade användare. `/health/ui` kan användas för att verifiera att rätt UI-template är installerad.


## v15.2.7 – XLSX Integrity

Den kompletta tomma Excel-mallen normaliseras och valideras innan den skickas till webbläsaren. Installationen kör dessutom ett riktigt XLSX-smoketest med release-versionens Python/openpyxl-miljö.

## v15.2.8 – Clean Excel Template Rebuild
Den globala tomma Excel-mallen byggs nu från en helt ny workbook utan Excel Tables, charts eller defined names. Det minimerar Microsoft Excel-specifika reparationsproblem och behåller full importkompatibilitet för de centrala projektbladen.


## v15.2.9 – Excel MergedCell Fix

Fixar installationsfelet i v15.2.8 där `Projektinformation` försökte skriva till `B4/B5`, som ingick i den sammanslagna underrubriken `A4:F5`. Projektfälten börjar nu på rad 7 och kontrolleras innan XLSX-filen skapas.
