# Project Planer v15.4.0 — Simple Excel Planning

## Enklare Uppgifter
Användaren behöver inte längre förstå eller skriva WBS i Multi-Project Excel.

Bladet **Uppgifter** använder:
- Projektkod
- Aktivitet
- Nivå (1, 2 eller 3)
- Aktivitetsnummer — skapas automatiskt
- övriga vanliga planeringsfält

Exempel:
- Installation · Nivå 1 → 1
- Installera server · Nivå 2 → 1.1
- Konfigurera IIS · Nivå 3 → 1.1.1

Project Planer räknar om den tekniska WBS-strukturen vid import och använder
inte Excel-formelns cache som källa till sanningen.

## Enklare beroenden
Bladet **Beroenden** använder nu:
- Föregående aktivitet
- Efterföljande aktivitet

Du kan ange antingen aktivitetsnummer (`2.1`) eller exakt aktivitetsnamn
(`Installera server`). Om samma namn finns flera gånger kräver importen
aktivitetsnumret.

## Säkerhet/validering
- Nivå 2 måste ha en föregående Nivå 1 i samma projekt.
- Nivå 3 måste ha en föregående Nivå 2 i samma projekt.
- Projektens aktivitetsnummer räknas separat.
- Importen förblir atomisk.
- Single-project Excel är oförändrat.
- Installationssmoketest för single- och multi-project XLSX ligger kvar.
