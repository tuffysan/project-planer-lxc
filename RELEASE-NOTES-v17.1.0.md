# Project Planer v17.1.0 — Unified Excel Edition

## Ett Excel-format
Project Planer Excel stöder alltid ett eller flera projekt.

## Tom Excel
Excel → Ladda ner tom Excel skapar nu alltid Unified/Multi-Project-arbetsboken.

## Exportera befintliga projekt
Ny sida `/excel/export-projects` låter användaren välja ett eller flera projekt.
Inne i ett projekt finns `/projects/<id>/excel/unified` för direkt export av just det projektet.

## Data som följer med
Projekt, Uppgifter, Risker, Ändringsärenden, Resurser, Kostnader, Beslut, Möten, Åtgärder, Nyttor och Beroenden.

## Säker återimport
Importerbara befintliga rader får dolda `_ID`-fält.
Rad med giltigt ID uppdateras, rad utan ID skapas, och felaktigt projekt-ID stoppar importen.

## Bakåtkompatibilitet
Äldre Single-Project Round-trip, komplett referensarbetsbok, statusrapport och Portfolio Excel finns kvar under avancerade verktyg.
