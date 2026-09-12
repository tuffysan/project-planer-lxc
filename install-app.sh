#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/opt/project-plan}"
PORT="${PORT:-8080}"

echo "=== Project Planer app installer ==="
echo "App dir: $APP_DIR"
echo "Port:    $PORT"

for cmd in python3 curl; do
  command -v "$cmd" >/dev/null 2>&1 || {
    echo "FEL: Saknar $cmd"
    exit 1
  }
done

python3 -c 'import venv' >/dev/null 2>&1 || {
  echo "FEL: Python venv-modulen saknas. Installera python3-venv."
  exit 1
}

mkdir -p "$APP_DIR/data" "$APP_DIR/backups"
cd "$APP_DIR"

if [[ -d .venv && ! -x .venv/bin/python ]]; then
  echo "Tar bort trasig virtuell miljö..."
  rm -rf .venv
fi

if [[ ! -d .venv ]]; then
  python3 -m venv .venv
fi

.venv/bin/python -m pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

cp scripts/project-plan.service /etc/systemd/system/project-plan.service
sed -i "s|Environment=PORT=.*|Environment=PORT=${PORT}|" /etc/systemd/system/project-plan.service || true

systemctl daemon-reload
systemctl enable project-plan
systemctl restart project-plan

echo "Väntar på health check..."
for i in $(seq 1 30); do
  if curl -fsS "http://127.0.0.1:${PORT}/health" >/dev/null; then
    echo "Health check OK"
    exit 0
  fi
  sleep 1
done

echo "FEL: health check misslyckades."
journalctl -u project-plan --no-pager -n 100 || true
exit 1
