# Project Planer v15.5.0 — Excel UX Edition

## Nytt Start-blad
- Tre tydliga steg: Projekt → Uppgifter → Kontroll/import
- Snabbvägar
- KPI för projekt, aktiviteter, risker, resurser och kontrollfel
- Färglegend och tips

## Enklare Uppgifter
- Nivå visas som `1 - Huvudaktivitet`, `2 - Underaktivitet`, `3 - Detaljaktivitet`
- Aktivitetsnummer genereras automatiskt
- Ny `Varaktighet dagar`
- `Plan slut` räknas automatiskt från start + varaktighet och kan skrivas över
- Avancerade kolumner finns kvar men är dolda som standard

## Gantt
- Automatisk 26-veckors planvy
- Hierarkiskt visade aktiviteter
- Filtrering på Projektkod
- Första 100 aktiviteterna visas

## Kontroll före import
Kontrollerar bland annat:
- saknade projektkoder
- saknade ansvariga
- saknade datum
- slutdatum före startdatum
- strukturproblem

## Renare flikordning
`Start → Projekt → Uppgifter → Gantt → Risker → Resurser → Kostnader → Beroenden → Kontroll`

Avancerade/referenceblad är dolda som standard men finns kvar.

## Stabilitet
- Fixar latent `wb_values`-problem i multi-project-importen.
- Servern räknar själv nivåstruktur och slutdatum vid import.
- Installationssmoketest kräver Start, Gantt och Kontroll innan release aktiveras.
- Single-project Excel är oförändrat.
