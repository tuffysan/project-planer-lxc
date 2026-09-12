#!/usr/bin/env bash
set -euo pipefail
APP_DIR="${APP_DIR:-/opt/project-plan}"
PORT="${PORT:-8080}"

mkdir -p "$APP_DIR/data" "$APP_DIR/backups"
cd "$APP_DIR"

python3 -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

cp scripts/project-plan.service /etc/systemd/system/project-plan.service
sed -i "s|Environment=PORT=.*|Environment=PORT=${PORT}|" /etc/systemd/system/project-plan.service || true
systemctl daemon-reload
systemctl enable --now project-plan

for i in $(seq 1 30); do
  if curl -fsS "http://127.0.0.1:${PORT}/health" >/dev/null; then
    echo "Health check OK"
    exit 0
  fi
  sleep 1
done

echo "ERROR: health check failed"
journalctl -u project-plan --no-pager -n 80 || true
exit 1
