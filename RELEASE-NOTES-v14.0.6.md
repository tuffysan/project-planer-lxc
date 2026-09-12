# Project Planer v14.0.6 — Tag-Pinned Deploy

Fixar deploy-flödet så att en ny release aldrig blandas med ett äldre script från `main`.

## Problemet
v14.0.5 installerades korrekt, men den sista verifieringen kunde fortfarande komma från en äldre cached `update-lxc.sh`.

## Lösningen
- `deploy-lxc.sh` fastställer först senaste GitHub-release.
- Därefter hämtas `update-lxc.sh` eller `install-lxc.sh` från exakt samma immutable tagg, t.ex. `v14.0.6`.
- `Cache-Control: no-cache` och cache-busting query används.
- Den valda releaseversionen skickas vidare explicit.
- Resultatet kan inte längre bli: app från v14.0.6 + updaterare från en äldre main-version.

## Rekommenderat kommando
```bash
CTID=200 bash -c "$(curl -fsSL -H 'Cache-Control: no-cache' "https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/deploy-lxc.sh?nocache=$(date +%s)")"
```
