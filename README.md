# Project Planer LXC v2.1.2

Maintenance release that fixes the v2.1.1 deployment path bug.

## Fixed in v2.1.2

- `install-app.sh` now installs from the actual extracted GitHub release directory.
- Verifies `requirements.txt`, `app/app.py`, templates, static assets, and the systemd service before deployment.
- Uses a staging directory before activating the new release.
- Preserves `/opt/project-plan/data` and `/opt/project-plan/backups`.
- Creates the Python virtual environment inside staging.
- Installs Python packages from the correct `requirements.txt`.
- Uses atomic release switching.
- Automatically rolls back to the previous application release if systemd startup or `/health` fails.
- Removes the previous release only after a successful health check.
- Adds `rsync` as an installation dependency.
- Update script uses the same safe install path as a fresh install.
- Keeps all features from v2.1.1 including multilingual UI and branding.

## Fresh install

```bash
VERSION=2.1.2 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

## Upgrade

```bash
CTID=200 VERSION=2.1.2 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## Publish

```powershell
.\PUBLISH.cmd
```
