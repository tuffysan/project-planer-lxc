#!/usr/bin/env bash
set -euo pipefail
REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION="${VERSION:-2.1.1}"
TAG="v${VERSION}"
CTID="${CTID:-140}"
APP_DIR="/opt/project-plan"
TMP="/tmp/project-plan-update-${VERSION}.zip"
BACKUP="/tmp/project-plan-rollback-$(date +%Y%m%d-%H%M%S).tar.gz"

echo "=== Project Planer upgrade ${TAG} ==="
pct status "$CTID" >/dev/null

curl -fL "https://github.com/${REPO}/archive/refs/tags/${TAG}.zip" -o "$TMP"
pct push "$CTID" "$TMP" "/tmp/project-plan-update.zip"

pct exec "$CTID" -- bash -lc "
set -euo pipefail
apt-get update -qq
apt-get install -y -qq unzip curl python3 python3-venv python3-pip locales >/dev/null
if [ -d '$APP_DIR' ]; then
  mkdir -p '$APP_DIR/backups'
  [ -f '$APP_DIR/data/projectplan.db' ] && cp '$APP_DIR/data/projectplan.db' '$APP_DIR/backups/projectplan-\$(date +%Y%m%d-%H%M%S).db' || true
  tar -czf '$BACKUP' -C '$APP_DIR' --exclude='.venv' .
fi
rm -rf /tmp/project-plan-src
mkdir -p /tmp/project-plan-src
unzip -q /tmp/project-plan-update.zip -d /tmp/project-plan-src
SRC=\$(find /tmp/project-plan-src -mindepth 1 -maxdepth 1 -type d | head -n1)
mkdir -p '$APP_DIR'
find '$APP_DIR' -mindepth 1 -maxdepth 1 ! -name data ! -name backups ! -name .venv -exec rm -rf {} +
cp -a \"\$SRC\"/. '$APP_DIR'/
cd '$APP_DIR'
chmod +x install-app.sh
if ! ./install-app.sh; then
  echo 'Upgrade failed - rolling back application files'
  systemctl stop project-plan || true
  find '$APP_DIR' -mindepth 1 -maxdepth 1 ! -name data ! -name backups ! -name .venv -exec rm -rf {} +
  tar -xzf '$BACKUP' -C '$APP_DIR'
  systemctl daemon-reload
  systemctl restart project-plan || true
  exit 1
fi
"
echo "Upgrade complete: ${TAG}"
