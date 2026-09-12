#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-plan-lxc}"
BRANCH="${BRANCH:-main}"
CTID="${CTID:-}"

if [[ -z "$CTID" ]]; then
  echo "Ange CTID, exempel:"
  echo "CTID=140 bash update-lxc.sh"
  exit 1
fi

pct status "$CTID" >/dev/null 2>&1 || { echo "LXC $CTID finns inte."; exit 1; }
pct start "$CTID" >/dev/null 2>&1 || true

pct exec "$CTID" -- bash -lc "
  set -e
  stamp=\$(date +%Y%m%d-%H%M%S)
  mkdir -p /opt/project-plan/backups
  if [ -f /opt/project-plan/data/projectplan.db ]; then
    cp /opt/project-plan/data/projectplan.db /opt/project-plan/backups/projectplan-\${stamp}.db
  fi
  apt-get update -qq
  DEBIAN_FRONTEND=noninteractive apt-get install -y -qq curl unzip ca-certificates
  rm -rf /tmp/project-plan-update
  mkdir -p /tmp/project-plan-update
  curl -fsSL https://github.com/${REPO}/archive/refs/heads/${BRANCH}.zip -o /tmp/project-plan-update/repo.zip
  unzip -q /tmp/project-plan-update/repo.zip -d /tmp/project-plan-update
  SRC=\$(find /tmp/project-plan-update -mindepth 1 -maxdepth 1 -type d -name '*-${BRANCH}' | head -1)
  systemctl stop project-plan || true
  cp -a \"\$SRC/app/.\" /opt/project-plan/app/
  cp \"\$SRC/requirements.txt\" /opt/project-plan/
  /opt/project-plan/venv/bin/pip install -q -r /opt/project-plan/requirements.txt
  cp \"\$SRC/scripts/project-plan.service\" /etc/systemd/system/project-plan.service
  chown -R projectplan:projectplan /opt/project-plan
  systemctl daemon-reload
  systemctl enable --now project-plan
"
echo "Uppgradering klar för CTID $CTID."
