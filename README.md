# Project Planer LXC v2.1.1

Maintenance release for the v2.1 multilingual/branding edition.

## Fixes in v2.1.1

- Fixes fresh Debian 12 LXC installation where `python3 -m venv` failed because `python3-venv` was not installed.
- Installs `python3`, `python3-venv` and `python3-pip` before the application installer starts.
- Installs and generates `en_US.UTF-8` locale to remove Debian locale warnings.
- Replaces the problematic `HOSTNAME` environment variable with `LXC_HOSTNAME`.
- Default container hostname is now `project-planer`.
- Adds Python/venv preflight checks.
- Detects and removes an incomplete `.venv`.
- Verifies both the systemd service and `/health` before reporting success.
- Failed fresh installations automatically remove the failed LXC unless `KEEP_FAILED_CONTAINER=1` is supplied.
- Keeps all v2.1.0 language selection, branding, icon, Admin, PMO, planning and integration features.

## Fresh install

```bash
VERSION=2.1.1 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Custom hostname:

```bash
LXC_HOSTNAME=project-planer VERSION=2.1.1 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Keep a failed container for troubleshooting:

```bash
KEEP_FAILED_CONTAINER=1 VERSION=2.1.1 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

## Upgrade existing container

```bash
CTID=201 VERSION=2.1.1 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## Publish

```powershell
.\PUBLISH.cmd
```
