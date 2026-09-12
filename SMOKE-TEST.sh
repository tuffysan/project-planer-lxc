#!/usr/bin/env bash
set -euo pipefail

BASE_URL="${BASE_URL:-http://127.0.0.1:8080}"

echo "[1/4] Health"
curl -fsS "$BASE_URL/health" >/dev/null

echo "[2/4] Readiness (if available)"
curl -fsS "$BASE_URL/health/ready" >/dev/null || true

echo "[3/4] Verify login page or redirect is reachable"
code="$(curl -sS -o /tmp/project-plan-smoke.html -w '%{http_code}' "$BASE_URL/")"
case "$code" in
  200|301|302|303|307|308) ;;
  *) echo "Unexpected HTTP status from /: $code"; exit 1 ;;
esac

echo "[4/4] Check journal for new BuildError / template errors"
if command -v journalctl >/dev/null 2>&1; then
  recent="$(journalctl -u project-plan --since '2 minutes ago' --no-pager || true)"
  if echo "$recent" | grep -Eq 'BuildError|jinja2\.exceptions|ERROR in app: Exception on /'; then
    echo "$recent"
    echo "Smoke test detected application errors."
    exit 1
  fi
fi

echo "Smoke test passed."
