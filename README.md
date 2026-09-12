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
CTID=<your-ctid> VERSION=8.1.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```

## New install

```bash
VERSION=8.1.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
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


## v5.0.2 – Full App Stability & Responsive Edition

Full static audit of the application with fixes for Project Control, automation notifications,
intelligence action creation, archive filtering and Report Pack. Mobile/tablet styling has also
been hardened across the app.

Run before publishing:

```bash
./VERIFY.sh
python3 VERIFY-PYTHON.py
```

After updating an LXC:

```bash
/opt/project-plan/current/SMOKE-TEST.sh
```

## v5.0.3 – Visual UX Edition

Adds RAG project health, portfolio visual dashboard, visual project cockpit, health score, progress bars, status chips, attention panels, milestones and responsive visual components for desktop, tablet and mobile.

## v5.0.4 – Visual UX Refinement

Second visual UX pass: fixes dashboard data binding, makes RAG status visible directly on the main dashboard, adds portfolio distribution, status reasons, trend indicators, search/filter/sort and a mobile bottom navigation.


## v7.0.1 Full Quality Audit

See `AUDIT-v7.0.1.md`. This release focuses on permissions, database reliability, startup migrations, health checks, static validation and responsive hardening rather than new features.


## Excel Round-trip (v8.1.0)

From Project Workspace, choose **Exportera Excel** or **Importera Excel**.

The full workbook supports offline editing for project information, tasks, risks, change requests, resources, costs, task dependencies, decisions, meetings, actions and benefits. Import always runs through a preview. Stable hidden IDs identify existing records; snapshot hashes detect concurrent changes. Missing spreadsheet rows do not delete data. A database backup is created before commit.
