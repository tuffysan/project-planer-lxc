# Project Planer v16.2.1 — Excel Visual UX Runtime Fix

## Fixat installationsfelet i v16.2.0

v16.2.0 stoppades korrekt av installationssmoketestet med:

`ValueError: too many values to unpack (expected 4)`

Felet låg i `excel_blank_complete_workbook_v1525()`.

Datumvalideringsloopen från Excel Visual UX hade av misstag lagts i
single-project-generatorn. Där består `core` av fem värden per rad:

`(sheet, headers, dates, money, percent)`

Visual UX-loopen förväntade sig fyra värden och hör egentligen hemma i
Multi-Project-generatorn.

## v16.2.1

- den felplacerade loopen är borttagen från single-project Excel
- datumvalideringen ligger nu i Multi-Project-generatorn där `core` har fyra fält
- Plan start / Plan slut / Faktisk start / Faktiskt slut behåller v16.2-valideringen
- datumkolumner på Risker, Resurser, Kostnader, Beslut, Möten, Åtgärder och Nyttor valideras
- statusfärger, progress bars, förseningsmarkering, Gantt och färglegend finns kvar
- Connected Projects från v16.1 finns kvar
- installationskontrollen verifierar nu specifikt att denna kod inte kan hamna i fel generator igen

v16.2.0 aktiverades inte eftersom felet inträffade före aktiveringssteget.
