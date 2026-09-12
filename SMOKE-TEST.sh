#!/usr/bin/env bash
set -euo pipefail
BASE_URL="${BASE_URL:-http://127.0.0.1:8080}"

echo "[1/6] Service liveness"
curl -fsS "$BASE_URL/health/live" | grep -q '"alive"'

echo "[2/6] Application health"
curl -fsS "$BASE_URL/health" >/dev/null

echo "[3/6] Readiness including database/integrity/disk"
curl -fsS "$BASE_URL/health/ready" >/tmp/project-plan-ready.json
cat /tmp/project-plan-ready.json
grep -q '"ready"' /tmp/project-plan-ready.json

echo "[4/6] Root/login reachability"
code="$(curl -sS -o /tmp/project-plan-smoke.html -w '%{http_code}' "$BASE_URL/")"
case "$code" in
  200|301|302|303|307|308) ;;
  *) echo "Unexpected HTTP status from /: $code"; exit 1 ;;
esac

echo "[5/6] Static verification in installed release"
if [ -f /opt/project-plan/current/VERIFY-PYTHON.py ]; then
  cd /opt/project-plan/current
  /opt/project-plan/current-venv/bin/python VERIFY-PYTHON.py
fi

echo "[6/6] Recent application errors"
if command -v journalctl >/dev/null 2>&1; then
  recent="$(journalctl -u project-plan --since '5 minutes ago' --no-pager || true)"
  if echo "$recent" | grep -Eq 'BuildError|jinja2\.exceptions|sqlite3\.(OperationalError|IntegrityError)|Traceback|ERROR in app: Exception on /'; then
    echo "$recent"
    echo "Smoke test detected application errors."
    exit 1
  fi
fi
echo "v7.0.1 smoke test passed."
