# Project Planer v15.2.8 — Clean Excel Template Rebuild

## Huvudfix
Den globala kompletta Excel-mallen är nu **ombyggd från grunden** i stället för att återanvända Project Planers round-trip-arbetsbok.

Detta tar bort den klass av Excel-strukturer som kunde ge Microsoft Excels dialog **“We found a problem with some content…”** trots att openpyxl kunde läsa filen.

## Ny XLSX-strategi
- Ny `Workbook()` varje gång den tomma projektmallen skapas.
- Inga Excel Tables i den tomma mallen.
- Inga charts, drawing-relations eller defined names.
- Inga tabeller med 0 datarader.
- Endast standardceller, format, säkra AutoFilters och enkla literal-listor för datavalidering.
- `_Metadata` är vanlig hidden, inte veryHidden.
- Importkompatibla rubriker är exakt desamma som Project Planers importer för:
  - Uppgifter
  - Beroenden
  - Risker
  - Ändringsärenden
  - Resurser
  - Kostnader
  - Beslut
  - Möten
  - Åtgärder
  - Nyttor
- Kompletta referensblad finns fortfarande med för resten av projektmodellen.
- XLSX-filen går fortfarande genom save → reload → save och integritetskontroll innan nedladdning.

## Excel Start Center
- Global Excel-genväg finns kvar.
- Skapa tom komplett Excel-mall finns kvar.
- Importera Excel och skapa projekt finns kvar.
