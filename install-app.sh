#!/usr/bin/env bash
set -euo pipefail

APP_DIR=/opt/project-plan
SOURCE_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "[1/6] Installerar beroenden..."
apt-get update
DEBIAN_FRONTEND=noninteractive apt-get install -y python3 python3-venv python3-pip ca-certificates curl

echo "[2/6] Skapar användare och mappar..."
id projectplan >/dev/null 2>&1 || useradd --system --home "$APP_DIR" --shell /usr/sbin/nologin projectplan
mkdir -p "$APP_DIR"/{data,backups}

echo "[3/6] Installerar appfiler..."
cp -a "$SOURCE_DIR/app" "$APP_DIR/"
cp "$SOURCE_DIR/requirements.txt" "$APP_DIR/"
python3 -m venv "$APP_DIR/venv"
"$APP_DIR/venv/bin/pip" install --upgrade pip
"$APP_DIR/venv/bin/pip" install -r "$APP_DIR/requirements.txt"

echo "[4/6] Installerar systemd service..."
cp "$SOURCE_DIR/scripts/project-plan.service" /etc/systemd/system/project-plan.service
chown -R projectplan:projectplan "$APP_DIR"

echo "[5/6] Startar tjänsten..."
systemctl daemon-reload
systemctl enable --now project-plan

echo "[6/6] Kontrollerar..."
sleep 2
systemctl --no-pager --full status project-plan || true
echo
echo "Project Plan kör på http://$(hostname -I | awk '{print $1}'):8080"
