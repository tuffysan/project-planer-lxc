# Project Planer v15.3.0 — Multi-Project Excel

## Nytt
- En Excel-fil kan innehålla **flera projekt**.
- Ny komplett mall: `Project-Planer-MULTI-PROJECT-projektmall.xlsx`.
- Nytt blad **Projekt** med en rad per projekt.
- **Projektkod** används som stabil nyckel i hela arbetsboken.
- Importbladen Uppgifter, Beroenden, Risker, Ändringsärenden, Resurser, Kostnader, Beslut, Möten, Åtgärder och Nyttor har Projektkod.
- WBS behöver bara vara unik inom respektive projekt.
- Samma WBS kan därför förekomma i flera projekt.
- Ny import: **Excel → Importera flera projekt**.
- Importen skapar alla projekt, lägger användaren som PM och importerar projektposter.
- Beroenden löses separat inom varje projekt.
- Importen är atomisk: ett fel gör rollback på hela importen.
- Importresultatet visar skapade projekt samt antal importerade poster per blad.
- Max 100 projekt och 10 000 importerade projektposter per arbetsbok.
- Runtime-smoketest verifierar både single-project och multi-project XLSX innan release aktiveras.

## Bakåtkompatibilitet
- Den befintliga single-project-mallen och single-project-importen finns kvar.
- Multi-Project Excel är ett separat flöde och påverkar inte befintliga projektfiler.
