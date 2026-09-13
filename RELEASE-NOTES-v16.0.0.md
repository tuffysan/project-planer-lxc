# Project Planer v16.0.0 — Unified UX Edition

## Mål
Webbappen och Excel ska använda samma enkla arbetssätt och samma begrepp.

## Ny Start
Ny `/start`-sida med:
- Vad behöver du göra idag?
- Mina aktiviteter
- Mina projekt
- Öppna risker
- Notiser
- Snabbknappar för Skapa projekt och Excel

## Projektöversikt
Ny `/projects/<id>/overview` med:
- progress
- antal aktiviteter
- klara aktiviteter
- öppna risker
- nästa milstolpar
- projektteam
- kostnadssammanfattning

## Gemensamt språk med Excel
- WBS visas som **Aktivitetsnummer**
- Nivåerna presenteras som **Huvudaktivitet / Underaktivitet / Detaljaktivitet**
- Projektöversikt, Uppgifter, Gantt, Risker, Resurser, Kostnader och Excel använder samma ord

## Förenklad navigation
Primärt:
`Start → Mitt arbete → Projekt → Portfolio → Excel`

Projekt:
`Översikt → Uppgifter → Gantt → Risker → Resurser → Kostnader → Avancerat`

## Progressive disclosure
RAID, ändringar, dokument och leverans ligger under **Avancerat** i projektvyn.
Funktionerna finns kvar men ligger inte i vägen för vanliga användare.

## Mobil
Ny Unified UX-layout är responsiv och projektflikarna kan scrollas horisontellt på små skärmar.

## Kompatibilitet
- Ingen databasmodell tas bort.
- Befintliga routes och funktioner bevaras.
- Excel v15.5-flödet ligger kvar.
