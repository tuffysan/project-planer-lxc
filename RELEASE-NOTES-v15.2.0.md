# Project Planer v15.2.0 — Controlled Planning Edition

## Språkmeny
- Språkmenyn är ombyggd till en stabil `details`-meny.
- Menyrubriken följer nu valt språk.
- Huvudnavigation, kontoåtgärder och systemetiketter följer valt språk.
- Svenska, engelska, tyska, norska, danska och finska finns kvar.

## Kontrollerad automatisk omplanering
- Förhandsgranska planeringsmotorns datumförslag.
- Välj exakt vilka aktiviteter som får ändras.
- Ange orsak innan godkännande.
- Varje omplanering sparas som ett batch-ID.
- Omplanering kan återställas.
- Audit trail skapas för apply/undo.
- Cirkulära beroenden blockerar omplanering.

## Resursutjämning
- Ny motor letar efter överbeläggning >100%.
- Föreslår senare veckor med ledig kapacitet.
- Flytt görs först efter explicit godkännande.
- Allokeringsprocent och planerade timmar flyttas proportionellt.
- Ändringen loggas i audit trail.

## Azure DevOps tvåvägssynk
- Synkade Work Items kan länkas till en Project Planer-aktivitet.
- Explicit Project Planer → DevOps push av titel/status.
- Explicit DevOps → Project Planer pull av titel/status.
- Alla synkar loggas.
- Ingen bakgrundspush skriver över data automatiskt.

## Deployment
Tag-pinned deploy och VERSION/APP_VERSION-kontrollerna från v14/v15 finns kvar.
