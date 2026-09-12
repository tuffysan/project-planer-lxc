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
CTID=<your-ctid> VERSION=5.0.1 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## New install

```bash
VERSION=5.0.1 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

Run `./VERIFY.sh` before publishing. Keep a database backup before production upgrades.

## Navigation hotfix

Rebuilt navigation: no dynamic `undefined` labels, compact dropdown navigation, responsive header and account menu.

## v5.0.1 hotfix

This release fixes the HTTP 500 error on the start page introduced in v5.0.0.

Cause: the global navigation rendered `url_for('delivery_center')`, while the `delivery_center` route requires `project_id`.

Fixes:
- removes the invalid global Delivery link
- exposes Delivery only from project context where `project_id` is available
- adds `VERIFY-PYTHON.py` to detect template `url_for()` calls missing required route parameters
- adds `SMOKE-TEST.sh` for post-upgrade HTTP/journal verification
