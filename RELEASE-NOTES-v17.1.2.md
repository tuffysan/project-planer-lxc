# Project Planer v17.1.2 — Validator Hotfix

## Fix
Installationen av v17.1.1 stoppades av en historisk regressionskontroll som krävde exakt metadata-värde `visual_ux_version=16.2.1`.

Det var fel eftersom Unified Excel i v17.1.x medvetet har ett nyare metadata-värde.

## Ändring
- install-app.sh kontrollerar nu att Visual UX-metadata finns, inte att den är låst till en gammal versionssträng.
- VERIFY.sh använder samma versionsoberoende kontroll.
- Python-regressionskontrollen verifierar metadatafältets existens i rätt workbook-scope utan att kräva v16.2.1.
- Unified Excel-funktionerna från v17.1.1 är oförändrade.
- APP_VERSION/VERSION är 17.1.2.
