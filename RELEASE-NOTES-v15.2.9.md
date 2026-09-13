# Project Planer v15.2.9 — Excel MergedCell Fix

## Root cause
v15.2.8 byggde den rena Excel-mallen från `Workbook()`, men bladet
`Projektinformation` använde samma rader som den sammanslagna underrubriken.

`title(..., subtitle=...)` skapar:
- `A3:F3` för rubriken
- `A4:F5` för underrubriken

Projektfälten började samtidigt på rad 4. Skrivning till `B4`/`B5` träffade därför
`MergedCell` och installationens XLSX-smoketest stoppade releasen med:

`AttributeError: 'MergedCell' object attribute 'value' is read-only`

## Fix
- Projektinformation börjar nu på rad 7.
- Ny defensiv kontroll verifierar att inmatningscellerna `B7:B12` inte är merged/read-only.
- Det befintliga riktiga XLSX-smoketestet i `install-app.sh` ligger kvar och måste passera innan releasen aktiveras.
- Inga databas- eller migreringsändringar krävs.

## Viktigt
v15.2.8 aktiverades inte när smoketestet misslyckades. Den tidigare fungerande releasen ska därför fortfarande vara aktiv.
