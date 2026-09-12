# Project Planer v14.0.3 — Latest Release Deployment

Den här releasen gör installation och uppgradering enklare och säkrare.

## Nytt
- `install-lxc.sh` använder senaste GitHub-releasen automatiskt om VERSION inte anges.
- `update-lxc.sh` använder senaste GitHub-releasen automatiskt om VERSION inte anges.
- Ny `deploy-lxc.sh`:
  - installerar om CTID inte finns,
  - uppgraderar om CTID redan finns.
- `install-latest.sh` och `update-latest.sh` finns som tydliga wrappers.
- Det går fortfarande att låsa till en specifik version med `VERSION=x.y.z`.
- Versionskontrollerna från v14.0.2 finns kvar: installerad runtime måste matcha den release som begärts.

## Rekommenderat kommando
```bash
CTID=200 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/deploy-lxc.sh)"
```
