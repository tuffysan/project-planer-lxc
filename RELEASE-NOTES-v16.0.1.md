# Project Planer v16.0.1 — Unified UX Runtime Fix

## Fixat

### Single-project Excel runtime
v16.0.0 kunde inte aktiveras eftersom installationssmoketestet stoppade på:

`NameError: name 'dm' is not defined`

Orsaken var att Multi-Project-fliklogik av misstag även hade lagts i
`excel_blank_complete_workbook_v1525()`. Single-project-mallen använder
variabeln `dd` för Datamodell och har inte Start/Gantt/Kontroll.

v16.0.1 återställer single-project-mallens korrekta avslut och döljer
Datamodell via `dd`.

### Multi-Project Excel
- Multi-Project-mallen får den avsedda flikordningen med **Start** först.
- Datamodell döljs som avsett.
- Multi-Project-importen laddar nu även en `data_only=True`-arbetsbok
  (`wb_values`) så formelbaserade projektkoder kan läsas robust.

### Starkare installationstest
Installationen verifierar nu uttryckligen:
- single-project: Datamodell finns och är dold
- multi-project: Start, Projekt, Uppgifter, Gantt och Kontroll finns
- Start är första fliken
- Datamodell är dold

v16.0.0 aktiverades inte när felet uppstod, så tidigare aktiv release
ska ha lämnats orörd av rollback/installationsskyddet.
