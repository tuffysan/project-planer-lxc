# Project Planer v17.0.0 — Simple Planning Edition

## Mål
Göra Project Planer enklare att använda i både webbappen och Excel utan att ta bort avancerade funktioner.

## Ny Plan-vy i webbappen
Varje projekt får `/projects/<id>/plan`.

Den visar bara det användaren normalt behöver:
- Aktivitetsnummer
- Aktivitet
- Nivå
- Start
- Varaktighet
- Slut
- Ansvarig
- Status
- Progress

### Datumväljare i webbappen
Startdatum använder webbläsarens riktiga `input type="date"` och visar därför operativsystemets/webbläsarens kalenderdatumväljare.

### Ingen manuell WBS
Användaren väljer:
- Huvudaktivitet
- Underaktivitet
- Detaljaktivitet

Project Planer skapar aktivitetsnummer automatiskt, t.ex. `1`, `1.1`, `1.1.1`.

### Automatisk slutdag
Start + Varaktighet räknar automatiskt fram slutdatum på serversidan.

### Direktredigering
Aktiviteter kan uppdateras direkt i Plan-tabellen utan att öppna det fullständiga uppgiftsformuläret.

## Excel är förenklat
Excel-sidan visar nu tre primära val:

1. **Ladda ner tom Excel-fil**
2. **Hämta projekt till Excel**
3. **Importera Excel**

Den tomma Excel-mallen finns alltså kvar som ett tydligt förstahandsval.

## En Excel-import
Ny `/excel/import` känner automatiskt igen:
- `new_project_complete`
- `multi_project_complete` / Connected Projects

Befintligt single-project Round-trip finns fortfarande inne i respektive projekt.

## Projektöversikt
`Plan` ersätter `Uppgifter` som primärt planeringsval i den förenklade projektvyn.

## Navigation
Mobilnavigation och kommandopalett använder nu `Start` konsekvent och Excel-länken går direkt till Excel Center.

## Bakåtkompatibilitet
- gamla task-formulär finns kvar
- Gantt finns kvar
- Round-trip finns kvar
- Multi-Project Excel finns kvar
- Connected Projects finns kvar
- avancerade PM/PPM-funktioner tas inte bort
