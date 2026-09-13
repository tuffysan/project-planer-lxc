# Project Planer v17.1.1 — Unified Excel Installer Hotfix

## Fix
Installationen av v17.1.0 stoppades av en gammal v17.0-verifiering som fortfarande letade efter UI-texten `Ladda ner tom Excel-fil`.

v17.1.0 bytte texten till `Ladda ner tom Excel`, men kontrollen i `install-app.sh` och `VERIFY.sh` uppdaterades inte.

## Ändrat
- Installeringskontrollen accepterar nu den nya Unified Excel-texten.
- VERIFY.sh använder samma kontroll.
- Ingen funktionalitet i Unified Excel tas bort eller ändras.
- v17.1.1 innehåller samma v17.1 Unified Excel-funktioner som v17.1.0.
