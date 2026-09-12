# Project Planer LXC v2.1.3

v2.1.3 fixes the virtual-environment relocation problem seen in v2.1.2.

## Deployment architecture

The application now uses immutable release directories and immutable virtual
environments:

```text
/opt/project-plan/
├── current -> /opt/project-plan/releases/v2.1.3
├── current-venv -> /opt/project-plan/venvs/v2.1.3
├── data/
├── backups/
├── releases/
│   └── v2.1.3/
└── venvs/
    └── v2.1.3/
```

The virtual environment is created directly at its final path and is never
renamed or moved. This prevents broken `gunicorn` shebangs.

## Improvements in v2.1.3

- Fixes `status=203/EXEC` / `gunicorn: No such file or directory`.
- Creates each Python venv directly at its permanent path.
- Uses `current` and `current-venv` symlinks for release activation.
- Candidate app and dependencies are tested before activation.
- Verifies that gunicorn is executable before systemd starts.
- Preserves data in `/opt/project-plan/data`.
- Supports rollback to the previous release and previous venv.
- Can migrate from the older legacy `/opt/project-plan` layout.
- Keeps the previous systemd unit for rollback during migration.
- Keeps two older releases/venvs for rollback/history.
- Uses `--no-install-recommends` and no longer installs unnecessary
  `python3-pip`/development packages on fresh LXC installation.
- Uses `C.UTF-8` during package installation to avoid pre-locale warnings.
- Keeps all multilingual UI, branding, project planning, PMO, collaboration
  and admin functionality from v2.1.2.

## Fresh install

```bash
VERSION=2.1.3 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

## Keep a failed test container

```bash
KEEP_FAILED_CONTAINER=1 VERSION=2.1.3 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

## Upgrade

```bash
CTID=200 VERSION=2.1.3 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## Publish

```powershell
.\PUBLISH.cmd
```
