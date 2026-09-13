# Project Planer v15.2.7 — Excel XLSX Integrity Fix

## Fixat
- Fixar Excel-varningen **“We found a problem with some content…”** för den globala kompletta projektmallen.
- Den tomma projektmallen lägger inte längre till exempelrader efter att tabeller/filter har finaliserats.
- Header-only AutoFilters tas bort från helt tomma templateblad.
- Tomma blad får aldrig zero-row Excel Tables.
- Arbetsboken körs genom save → reload → save innan nedladdning för att normalisera XLSX-relationer.
- Alla XML- och `.rels`-delar i XLSX-paketet parsas före nedladdning.
- Tabellreferenser valideras.
- Samma XLSX-integritetskontroll används även för **Komplett Excel** från befintliga projekt.
- Installationen genererar nu en riktig blank Excel-mall i release-venv och stoppar releasen om XLSX-smoketestet misslyckas.

## Kvar
- Excel Start Center och den globala Excel-genvägen från v15.2.5/v15.2.6 finns kvar.
- Importera Excel och skapa projekt finns kvar.
