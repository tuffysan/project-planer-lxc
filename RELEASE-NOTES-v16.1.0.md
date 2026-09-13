# Project Planer v16.1.0 — Excel Connected Projects

## Projektlistor som fungerar direkt i Excel
Projekt som finns på bladet **Projekt** blir direkt valbara på importerbara blad.

Dropdownen visar exempelvis:

`PRJ-001 – test`

v16.1 använder en dold hjälplista på samma blad som dropdownen. Det undviker den tidigare direkta cross-sheet Data Validation-referensen som kunde ge en tom lista i desktop Excel.

## Befintliga projekt från webbappen
Excel Center har en ny knapp **Befintliga projekt i Excel**.

Connected Projects-filen innehåller användarens tillgängliga projekt på bladet Projekt med:
- stabil kod `APP-000123`
- projektnamn, kund, projektledare, beskrivning och datum
- dolt `_ProjectID`
- blå markering för rader som kommer från Project Planer

Nya projekt kan läggas till på tomma rader i samma fil.

## Import utan projektdubbletter
Vid import:
- befintligt `_ProjectID` uppdaterar projektets grunddata
- nya projekt skapas
- ProjectID verifieras mot användarens projektbehörighet
- projektval som `APP-000123 – Projektnamn` normaliseras automatiskt

Nya aktiviteter/risker/resurser/kostnader som anges i Connected-filen läggs till i valt befintligt eller nytt projekt.

## Avgränsning
Full tvåvägsredigering av redan befintliga projektposter görs fortsatt via single-project Round-trip. v16.1 fokuserar på Connected Projects, projektval och att lägga till planering i befintliga projekt.
