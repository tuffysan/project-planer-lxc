#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION="${VERSION:-2.1.2}"
CTID="${CTID:-}"

if ! command -v pct >/dev/null 2>&1; then
  echo "FEL: Kör detta script på en Proxmox VE host."
  exit 1
fi

if [[ -z "$CTID" ]]; then
  echo "FEL: Ange CTID, exempel:"
  echo "CTID=200 VERSION=2.1.2 bash -c \"\$(curl -fsSL https://raw.githubusercontent.com/${REPO}/main/update-lxc.sh)\""
  exit 1
fi

TAG="$VERSION"
[[ "$TAG" == v* ]] || TAG="v$TAG"
ARCHIVE_URL="https://github.com/${REPO}/archive/refs/tags/${TAG}.zip"

echo "=== Project Planer LXC updater ==="
echo "Repo:    $REPO"
echo "Version: $TAG"
echo "CTID:    $CTID"

if ! pct status "$CTID" >/dev/null 2>&1; then
  echo "FEL: CT $CTID finns inte."
  exit 1
fi

if ! pct status "$CTID" | grep -q running; then
  echo "Startar CT $CTID..."
  pct start "$CTID"
  sleep 3
fi

echo "Kontrollerar release..."
curl -fsIL "$ARCHIVE_URL" >/dev/null

echo "Installerar systemkrav..."
pct exec "$CTID" -- bash -lc '
  set -euo pipefail
  export DEBIAN_FRONTEND=noninteractive
  apt-get update -qq
  apt-get install -y -qq ca-certificates curl unzip rsync python3 python3-venv python3-pip locales >/dev/null
'

echo "Hämtar release..."
pct exec "$CTID" -- bash -lc "
  set -euo pipefail
  rm -rf /tmp/project-planer-update
  mkdir -p /tmp/project-planer-update
  curl -fL '${ARCHIVE_URL}' -o /tmp/project-planer-update/repo.zip
  unzip -q /tmp/project-planer-update/repo.zip -d /tmp/project-planer-update
  SRC=\$(find /tmp/project-planer-update -mindepth 1 -maxdepth 1 -type d | head -1)
  test -n \"\$SRC\"
  test -f \"\$SRC/install-app.sh\"
  chmod +x \"\$SRC/install-app.sh\"
  \"\$SRC/install-app.sh\"
"

echo "Verifierar..."
pct exec "$CTID" -- bash -lc '
  set -euo pipefail
  systemctl is-active --quiet project-plan
  curl -fsS http://127.0.0.1:8080/health
'

IP="$(pct exec "$CTID" -- hostname -I | awk '{print $1}')"

echo
echo "============================================"
echo "UPPGRADERING KLAR"
echo "Version: $TAG"
echo "CTID:    $CTID"
echo "App:     http://${IP}:8080"
echo "============================================"
