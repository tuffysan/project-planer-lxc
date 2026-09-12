#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-plan-lxc}"
BRANCH="${BRANCH:-main}"
HOSTNAME="${HOSTNAME:-project-plan}"
CORES="${CORES:-1}"
MEMORY="${MEMORY:-1024}"
SWAP="${SWAP:-512}"
DISK="${DISK:-8}"
BRIDGE="${BRIDGE:-vmbr0}"
IP_CONFIG="${IP_CONFIG:-dhcp}"
PASSWORD="${PASSWORD:-$(openssl rand -base64 18 | tr -d '/+=' | head -c 20)}"

if ! command -v pct >/dev/null 2>&1; then
  echo "FEL: Kör detta script på en Proxmox VE host."
  exit 1
fi

get_next_ctid() {
  local id="${CTID:-}"
  if [[ -n "$id" ]]; then echo "$id"; return; fi
  id=200
  while pct status "$id" >/dev/null 2>&1; do id=$((id+1)); done
  echo "$id"
}
CTID="$(get_next_ctid)"

pick_storage() {
  if [[ -n "${STORAGE:-}" ]]; then echo "$STORAGE"; return; fi
  pvesm status -content rootdir 2>/dev/null | awk 'NR>1 && $3=="active"{print $1; exit}'
}
STORAGE="$(pick_storage)"
if [[ -z "$STORAGE" ]]; then
  echo "FEL: Kunde inte hitta storage med rootdir. Ange STORAGE=..."
  exit 1
fi

pick_template_storage() {
  pvesm status -content vztmpl 2>/dev/null | awk 'NR>1 && $3=="active"{print $1; exit}'
}
TEMPLATE_STORAGE="$(pick_template_storage)"
if [[ -z "$TEMPLATE_STORAGE" ]]; then
  echo "FEL: Kunde inte hitta storage för templates."
  exit 1
fi

echo "=== Project Plan LXC installer ==="
echo "Repo:      $REPO"
echo "CTID:      $CTID"
echo "Hostname:  $HOSTNAME"
echo "Storage:   $STORAGE"
echo "Bridge:    $BRIDGE"

TEMPLATE="$(pveam list "$TEMPLATE_STORAGE" 2>/dev/null | awk '/debian-12-standard_.*amd64.tar.zst/{print $1}' | tail -1 || true)"
if [[ -z "$TEMPLATE" ]]; then
  echo "Hämtar Debian 12 template..."
  pveam update
  TEMPLATE_NAME="$(pveam available --section system | awk '/debian-12-standard_.*amd64.tar.zst/{print $2}' | tail -1)"
  if [[ -z "$TEMPLATE_NAME" ]]; then
    echo "FEL: Debian 12 template hittades inte."
    exit 1
  fi
  pveam download "$TEMPLATE_STORAGE" "$TEMPLATE_NAME"
  TEMPLATE="${TEMPLATE_STORAGE}:vztmpl/${TEMPLATE_NAME}"
fi

if [[ "$IP_CONFIG" == "dhcp" ]]; then
  NET0="name=eth0,bridge=${BRIDGE},ip=dhcp"
else
  NET0="name=eth0,bridge=${BRIDGE},ip=${IP_CONFIG}"
  if [[ -n "${GATEWAY:-}" ]]; then NET0="${NET0},gw=${GATEWAY}"; fi
fi

echo "Skapar LXC..."
pct create "$CTID" "$TEMPLATE" \
  --hostname "$HOSTNAME" \
  --cores "$CORES" \
  --memory "$MEMORY" \
  --swap "$SWAP" \
  --rootfs "${STORAGE}:${DISK}" \
  --net0 "$NET0" \
  --unprivileged 1 \
  --features nesting=1 \
  --password "$PASSWORD" \
  --start 1

echo "Väntar på nätverk..."
for i in {1..30}; do
  if pct exec "$CTID" -- bash -lc 'getent hosts github.com >/dev/null 2>&1'; then break; fi
  sleep 2
done

echo "Installerar appen från GitHub..."
pct exec "$CTID" -- bash -lc "
  set -e
  apt-get update
  DEBIAN_FRONTEND=noninteractive apt-get install -y ca-certificates curl unzip
  rm -rf /tmp/project-plan
  mkdir -p /tmp/project-plan
  curl -fsSL https://github.com/${REPO}/archive/refs/heads/${BRANCH}.zip -o /tmp/project-plan/repo.zip
  unzip -q /tmp/project-plan/repo.zip -d /tmp/project-plan
  SRC=\$(find /tmp/project-plan -mindepth 1 -maxdepth 1 -type d -name '*-${BRANCH}' | head -1)
  chmod +x \"\$SRC/install-app.sh\"
  \"\$SRC/install-app.sh\"
"

IP="$(pct exec "$CTID" -- hostname -I | awk '{print $1}')"

echo
echo "============================================"
echo "KLART"
echo "CTID:      $CTID"
echo "LXC root:  root / $PASSWORD"
echo "App:       http://${IP}:8080"
echo "============================================"
