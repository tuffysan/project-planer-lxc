# Project Planer LXC v2.0.0 – Complete Edition

Cumulative release containing the complete feature set built through v1.3–v1.6 plus the integration/admin layer.

## Complete feature set
- Multi-user security and project isolation
- Admin Center
- Gantt, Kanban, Calendar, My Work
- FS/SS/FF/SF dependencies + lag
- Critical Path / CPM
- Auto-scheduling
- Resource capacity and workload
- Time reporting
- Budget, costs and Earned Value
- PMO, RAG status, RAID, decisions and Change Requests
- Project docs/wiki, attachments, meetings, actions and notifications
- REST API v1 with API keys
- Custom field definitions
- Webhook data model
- Admin backup/download
- Runtime/database diagnostics
- Dashboard-preference data model
- Health check, upgrade backup and rollback

## API
Create an API key in Admin Center.

```http
GET /api/v1/projects
Authorization: Bearer pp_...
```

## Install
```bash
VERSION=2.0.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/install-lxc.sh)"
```

## Upgrade
```bash
CTID=140 VERSION=2.0.0 bash -c "$(curl -fsSL https://raw.githubusercontent.com/tuffysan/project-planer-lxc/main/update-lxc.sh)"
```
