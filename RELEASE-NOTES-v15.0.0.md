# Project Planer v15.0.0 — Professional PPM Edition

v15.0.0 fokuserar på att göra appen enklare att arbeta i, inte bara större.

## Ny kärnupplevelse
- Ny **Idag**-vy med en riktig Action Inbox och prioriterad portfölj.
- Ny **Mitt arbete**-vy som samlar aktiviteter, actions, förändringar och notifieringar.
- Ny **Project Cockpit** med leveransprognos, nästa beslut, milstolpar, team och styrning.
- Ny **Portfolio Cockpit** som prioriterar projekt efter risk och leveranshälsa.
- Ny **What-if / Påverkansanalys** som följer aktivitetsberoenden utan att ändra projektdata.

## Produktivitetsförbättringar
- Direktmarkering av aktivitet/action som klar från Mitt arbete.
- Direktmarkering av notifiering som läst.
- Projekthälsa väger ihop försenade aktiviteter, blockeringar, höga risker, budget och projektslut.
- Batched SQL på startsidan för bättre skalning än tidigare projekt-för-projekt-frågor.
- Prognos för projektslut baserat på aktuell försening.
- Tydligare svensk, responsiv och mer konsekvent UX.

## Deployment
- Behåller tag-pinned deploy från v14.0.6.
- `PUBLISH.ps1` synkroniserar nu automatiskt `VERSION` och `APP_VERSION`, även vid auto-increment.

## Rekommenderat kommando
```bash
CTID=200 bash -c "$(curl -fsSL -H 'Cache-Control: no-cache' "https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/deploy-lxc.sh?nocache=$(date +%s)")"
```
