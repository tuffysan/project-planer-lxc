# Project Planer v15.2.3 — Create Project Fix

## Fixar HTTP 500 på Skapa
Grundorsaken var CSRF-renderingen i äldre v12.x/v11.x-mallar:

- Applikationen injicerar `csrf_token` som en sträng.
- Flera mallar använde ändå `{{ csrf_token() }}` och försökte alltså anropa en sträng som funktion.
- Skapa projekt-sidan använde dessutom POST-fältnamnet `csrf_token`, medan säkerhetslagret kräver `_csrf`.

## Ändrat
- Alla kvarvarande `csrf_token()` i mallar är borttagna.
- Alla dessa formulär använder nu `_csrf` + `{{csrf_token}}`.
- Ny v15-anpassad sida för att skapa projekt.
- Validering av projektnamn och datum.
- Skaparen läggs automatiskt till som projekt-PM.
- Skapandet skrivs till audit log.
- Efter skapande öppnas Project Cockpit.
- Skapa projekt är begränsat till Admin/PM.
