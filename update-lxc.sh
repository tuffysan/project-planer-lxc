#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION="${VERSION:-1.0.8}"
CTID="${CTID:-}"
TAG="$VERSION"
[[ "$TAG" == v* ]] || TAG="v$TAG"

if [[ -z "$CTID" ]]; then
  echo "Ange CTID, exempel:"
  echo "CTID=200 VERSION=1.0.8 bash update-lxc.sh"
  exit 1
fi

ARCHIVE_URL="https://github.com/${REPO}/archive/refs/tags/${TAG}.zip"

pct status "$CTID" >/dev/null 2>&1 || { echo "LXC $CTID finns inte."; exit 1; }
curl -fsIL "$ARCHIVE_URL" >/dev/null || { echo "Release $TAG finns inte i $REPO."; exit 1; }
pct start "$CTID" >/dev/null 2>&1 || true

pct exec "$CTID" -- bash -lc "
  set -euo pipefail
  stamp=\$(date +%Y%m%d-%H%M%S)
  mkdir -p /opt/project-plan/backups
  if [ -f /opt/project-plan/data/projectplan.db ]; then
    cp /opt/project-plan/data/projectplan.db /opt/project-plan/backups/projectplan-\${stamp}.db
  fi
  apt-get update -qq
  DEBIAN_FRONTEND=noninteractive apt-get install -y -qq curl unzip ca-certificates
  rm -rf /tmp/project-planer-update
  mkdir -p /tmp/project-planer-update
  curl -fL '${ARCHIVE_URL}' -o /tmp/project-planer-update/repo.zip
  unzip -q /tmp/project-planer-update/repo.zip -d /tmp/project-planer-update
  SRC=\$(find /tmp/project-planer-update -mindepth 1 -maxdepth 1 -type d | head -1)
  test -f \"\$SRC/install-app.sh\"
  systemctl stop project-plan || true
  cp -a \"\$SRC/app/.\" /opt/project-plan/app/
  cp \"\$SRC/requirements.txt\" /opt/project-plan/
  /opt/project-plan/venv/bin/pip install -q -r /opt/project-plan/requirements.txt
  cp \"\$SRC/scripts/project-plan.service\" /etc/systemd/system/project-plan.service
  chown -R projectplan:projectplan /opt/project-plan
  systemctl daemon-reload
  systemctl enable --now project-plan
"
echo "Uppgradering till $TAG klar för CTID $CTID."
