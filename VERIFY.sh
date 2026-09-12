#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

test -f "$ROOT/requirements.txt"
test -f "$ROOT/app/app.py"
test -d "$ROOT/app/templates"
test -d "$ROOT/app/static"
test -f "$ROOT/install-app.sh"
test -f "$ROOT/install-lxc.sh"
test -f "$ROOT/update-lxc.sh"
test -f "$ROOT/scripts/project-plan.service"

python3 -m py_compile "$ROOT/app/app.py"
bash -n "$ROOT/install-app.sh"
bash -n "$ROOT/install-lxc.sh"
bash -n "$ROOT/update-lxc.sh"

grep -q 'APP_VERSION = "2.1.3"' "$ROOT/app/app.py"
grep -q '^2.1.3$' "$ROOT/VERSION"
grep -q '/opt/project-plan/current-venv/bin/gunicorn' "$ROOT/scripts/project-plan.service"

echo "v2.1.3 package verification OK"
