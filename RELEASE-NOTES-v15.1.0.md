# Project Planer v15.1.0 — Planning & Resource Engine

Bygger vidare på Professional PPM-grunden i v15.0.

## Nytt
- Planeringsmotor som räknar prognos från FS/SS/FF/SF, lag och arbetsdagar.
- Indikativ kritisk linje och slack.
- Upptäcker beroendecykler.
- Visar aktiviteter vars prognos flyttats jämfört med sparad plan.
- Ny kapacitets-heatmap för resurser per vecka.
- Ny samlad Projekthälsa med planeringsprognos.
- Project Cockpit länkar direkt till planeringsmotorn.
- Befintlig What-if finns kvar och kompletterar planeringsmotorn.

Planeringsmotorn är medvetet en förhandsberäkning i v15.1. Den skriver inte automatiskt över
projektplanen; det minskar risken att en prognos oavsiktligt ändrar basplanen.
