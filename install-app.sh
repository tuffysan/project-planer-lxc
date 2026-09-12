#!/usr/bin/env bash
set -euo pipefail

APP_DIR="${APP_DIR:-/opt/project-plan}"
PORT="${PORT:-8080}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_DIR="$SCRIPT_DIR"
STAGING_DIR="${APP_DIR}.staging"
BACKUP_DIR="${APP_DIR}.previous"

echo "=== Project Planer app installer ==="
echo "Source:  $SOURCE_DIR"
echo "App dir: $APP_DIR"
echo "Port:    $PORT"

for cmd in python3 curl rsync; do
  command -v "$cmd" >/dev/null 2>&1 || {
    echo "FEL: Saknar $cmd"
    exit 1
  }
done

python3 -c 'import venv' >/dev/null 2>&1 || {
  echo "FEL: Python venv-modulen saknas. Installera python3-venv."
  exit 1
}

for required in \
  "$SOURCE_DIR/requirements.txt" \
  "$SOURCE_DIR/app/app.py" \
  "$SOURCE_DIR/app/templates" \
  "$SOURCE_DIR/app/static" \
  "$SOURCE_DIR/scripts/project-plan.service"
do
  if [[ ! -e "$required" ]]; then
    echo "FEL: Releasepaketet saknar: $required"
    exit 1
  fi
done

echo "Förbereder staging..."
rm -rf "$STAGING_DIR"
mkdir -p "$STAGING_DIR"

# copy application while excluding generated/runtime data
rsync -a \
  --exclude '.git' \
  --exclude '__pycache__' \
  --exclude '.venv' \
  --exclude 'data' \
  --exclude 'backups' \
  --exclude '*.pyc' \
  "$SOURCE_DIR/" "$STAGING_DIR/"

# Preserve runtime data from current install
if [[ -d "$APP_DIR/data" ]]; then
  mkdir -p "$STAGING_DIR/data"
  rsync -a "$APP_DIR/data/" "$STAGING_DIR/data/"
else
  mkdir -p "$STAGING_DIR/data"
fi

if [[ -d "$APP_DIR/backups" ]]; then
  mkdir -p "$STAGING_DIR/backups"
  rsync -a "$APP_DIR/backups/" "$STAGING_DIR/backups/"
else
  mkdir -p "$STAGING_DIR/backups"
fi

echo "Skapar virtuell miljö..."
python3 -m venv "$STAGING_DIR/.venv"
"$STAGING_DIR/.venv/bin/python" -m pip install --upgrade pip
"$STAGING_DIR/.venv/bin/pip" install -r "$STAGING_DIR/requirements.txt"

echo "Installerar systemd service..."
cp "$STAGING_DIR/scripts/project-plan.service" /etc/systemd/system/project-plan.service
sed -i "s|Environment=PORT=.*|Environment=PORT=${PORT}|" /etc/systemd/system/project-plan.service || true

# Ensure the service points to the final app path
sed -i "s|WorkingDirectory=.*|WorkingDirectory=${APP_DIR}|" /etc/systemd/system/project-plan.service || true
sed -i "s|ExecStart=.*|ExecStart=${APP_DIR}/.venv/bin/gunicorn -w 2 -b 0.0.0.0:${PORT} app.app:app|" /etc/systemd/system/project-plan.service || true

echo "Växlar release atomiskt..."
rm -rf "$BACKUP_DIR"
if [[ -d "$APP_DIR" ]]; then
  mv "$APP_DIR" "$BACKUP_DIR"
fi
mv "$STAGING_DIR" "$APP_DIR"

rollback() {
  echo "Återställer föregående release..."
  systemctl stop project-plan >/dev/null 2>&1 || true
  rm -rf "$APP_DIR"
  if [[ -d "$BACKUP_DIR" ]]; then
    mv "$BACKUP_DIR" "$APP_DIR"
    systemctl daemon-reload
    systemctl start project-plan >/dev/null 2>&1 || true
  fi
}

echo "Startar tjänsten..."
systemctl daemon-reload
systemctl enable project-plan >/dev/null
if ! systemctl restart project-plan; then
  echo "FEL: Kunde inte starta project-plan."
  journalctl -u project-plan --no-pager -n 100 || true
  rollback
  exit 1
fi

echo "Väntar på health check..."
healthy=0
for i in $(seq 1 45); do
  if curl -fsS "http://127.0.0.1:${PORT}/health" >/tmp/project-plan-health.json 2>/dev/null; then
    healthy=1
    break
  fi
  sleep 1
done

if [[ "$healthy" != "1" ]]; then
  echo "FEL: health check misslyckades."
  journalctl -u project-plan --no-pager -n 120 || true
  rollback
  exit 1
fi

echo "Health check OK:"
cat /tmp/project-plan-health.json || true
echo

rm -rf "$BACKUP_DIR"
echo "Project Planer installerad."
