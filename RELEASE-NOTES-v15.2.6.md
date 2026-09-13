# Project Planer v15.2.6 — Excel Navigation Fix

## Fixat
- **Excel** ligger nu direkt i den faktiska huvudnavigationen:
  `Idag · Mitt arbete · Projekt · Portfolio · Excel · Sök · + Skapa · Admin`
- Excel-länken är inte längre beroende av Admin/PM-blocket.
- Excel-länken är en vanlig navigationslänk utan specialklass som kan döljas av CSS.
- Stylesheet får versionsparameter (`?v=15.2.6`) så webbläsaren inte återanvänder gammal CSS efter uppgradering.
- Ny diagnostik: `/health/ui` visar om den installerade `base.html` verkligen innehåller Excel-navigationen.
- `VERIFY.sh` kontrollerar att länken finns.
- `install-app.sh` stoppar installation av v15.2.6 om rätt huvudmeny saknas.

## Excel Start Center
Funktionerna från v15.2.5 finns kvar:
- skapa tom komplett Excel-mall utan projekt
- importera Excel och skapa projekt
- `/excel` som globalt Excel Start Center
