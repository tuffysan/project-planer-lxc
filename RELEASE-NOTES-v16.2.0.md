# Project Planer v16.2.0 — Excel Visual UX

## Datum i Excel
Planeringsdatum har nu riktig Excel-datumvalidering:
- Plan start
- Plan slut
- Faktisk start
- Faktiskt slut
- datumkolumner på övriga importerbara blad

Microsoft Excel-versioner som erbjuder en datumväljare för validerade datumfält kan visa den när cellen aktiveras. På Excel-versioner utan popup-kalender får användaren fortfarande:
- strikt datumvalidering
- input-hjälp
- konsekvent format `ÅÅÅÅ-MM-DD`
- tydligt felmeddelande vid ogiltigt datum

Ingen makro-/ActiveX-lösning används.

## Enhetlig färgkodning
Start-bladet innehåller nu en komplett legend:

- Ljusgrön: fyll i
- Ljusblå: kommer från Project Planer
- Grå: beräknas automatiskt
- Gul: varning
- Ljus röd: fel eller blockerad
- Blå status: Pågår
- Grön status: Klar
- Orange status: I riskzonen
- Grå status: Ej påbörjad

Färg är aldrig den enda informationsbäraren; status visas alltid som text.

## Uppgifter
- Status färgkodas automatiskt
- Progress % har databar 0–100 %
- försenade aktiviteter markeras ljusrött när Plan slut har passerat och status inte är Klar
- Plan slut fortsätter att räknas från Plan start + Varaktighet dagar
- Aktivitetsnummer är fortsatt grått/automatiskt

## Kontroll
Kontroll-bladet visar:
- OK i grönt
- VARNING i gult
- FEL i rött

## Gantt
Tidslinjen har fått tydligare visuell markering av veckor som aktivitet pågår.

## Connected Projects
Alla funktioner från v16.1 finns kvar:
- `Projektkod – Projektnamn` i dropdowns
- Connected Projects-download
- dolt ProjectID för befintliga projekt
- uppdatering utan projektdubbletter
