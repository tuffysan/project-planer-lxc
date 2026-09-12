# Project Planer v14.0.2 — Deployment Version Fix

Åtgärdar versionsproblemet där updateringen kunde se lyckad ut trots att en äldre app fortfarande kördes.

- GitHub-taggen verifieras före installation.
- VERSION måste matcha APP_VERSION.
- Kandidatreleasen kontrolleras innan aktivering.
- /health måste rapportera exakt begärd version efter start.
- Vid mismatch rullas releasen tillbaka.
- Aktiv release och venv skrivs ut av updateraren.
- VERIFY.sh använder aktuell VERSION i stället för gamla hårdkodade värden.
