#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
EXPECTED="$(tr -d '[:space:]' < "$ROOT/VERSION")"
python3 -m py_compile "$ROOT/app/app.py"
bash -n "$ROOT/install-app.sh"
bash -n "$ROOT/install-lxc.sh"
bash -n "$ROOT/update-lxc.sh"
grep -q "APP_VERSION = \"$EXPECTED\"" "$ROOT/app/app.py"
grep -q '/opt/project-plan/current-venv/bin/gunicorn' "$ROOT/scripts/project-plan.service"
echo "[VERIFY] VERSION=$EXPECTED och APP_VERSION matchar."
python3 "$ROOT/VERIFY-PYTHON.py"
echo "Project Planer package verification OK: v$EXPECTED"
