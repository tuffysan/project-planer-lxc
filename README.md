# Project Planer LXC v5.0.0 – Intelligent Project Platform

This is a cumulative release. It contains all functionality from the preceding roadmap releases.

## Added in v5.0.0

- Intelligent Project Manager
- natural-language style project questions using local project data
- go-live blocker analysis
- risk summaries
- weekly change summaries
- meeting notes → actions/decisions
- assistant query history
- no external AI data transfer by default

## Upgrade an existing LXC

```bash
CTID=<your-ctid> VERSION=5.0.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## New install

```bash
VERSION=5.0.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Run `./VERIFY.sh` before publishing. Keep a database backup before production upgrades.

## Navigation hotfix

Rebuilt navigation: no dynamic `undefined` labels, compact dropdown navigation, responsive header and account menu.
