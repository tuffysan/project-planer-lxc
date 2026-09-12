# Project Planer v14.0.5 — Deploy Hardening

Fixar deploy-flödet efter v14.0.3/v14.0.4.

## Ändrat
- `update-lxc.sh` omskriven för att undvika nested quoting helt.
- Fjärrkommandon körs via `bash -s -- ... <<'REMOTE'`.
- Paketets VERSION och APP_VERSION verifieras före aktivering.
- Efter installation verifieras:
  - `/opt/project-plan/current`
  - `/opt/project-plan/current-venv`
  - `/health` runtime-version
- `deploy-lxc.sh` uppgraderar alltid om CTID redan finns.
- `install-lxc.sh` vägrar skriva över en befintlig CTID och hänvisar till deploy/update.
- Senaste GitHub-release används automatiskt om VERSION inte anges.

## Rekommenderat kommando

```bash
CTID=200 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/deploy-lxc.sh)"
```
