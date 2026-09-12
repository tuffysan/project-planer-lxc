#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION="${VERSION:-2.1.1}"
LXC_HOSTNAME="${LXC_HOSTNAME:-project-planer}"
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

TAG="$VERSION"
[[ "$TAG" == v* ]] || TAG="v$TAG"

get_next_ctid() {
  if [[ -n "${CTID:-}" ]]; then
    echo "$CTID"
    return
  fi
  local id=200
  while pct status "$id" >/dev/null 2>&1; do
    id=$((id+1))
  done
  echo "$id"
}
CTID="$(get_next_ctid)"

pick_storage() {
  if [[ -n "${STORAGE:-}" ]]; then
    echo "$STORAGE"
    return
  fi
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

ARCHIVE_URL="https://github.com/${REPO}/archive/refs/tags/${TAG}.zip"

echo "=== Project Planer LXC installer ==="
echo "Repo:      $REPO"
echo "Version:   $TAG"
echo "CTID:      $CTID"
echo "Hostname:  $LXC_HOSTNAME"
echo "Storage:   $STORAGE"
echo "Bridge:    $BRIDGE"
echo "Release:   $ARCHIVE_URL"

echo "Kontrollerar att release $TAG finns..."
if ! curl -fsIL "$ARCHIVE_URL" >/dev/null; then
  echo "FEL: Release $TAG kunde inte hämtas från $REPO."
  echo "Kontrollera att GitHub-taggen finns och är publik."
  exit 1
fi

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
  [[ -n "${GATEWAY:-}" ]] && NET0="${NET0},gw=${GATEWAY}"
fi

cleanup_failed_container() {
  local rc=$?
  if [[ $rc -ne 0 ]]; then
    echo
    echo "Installationen misslyckades."
    if [[ "${KEEP_FAILED_CONTAINER:-0}" != "1" ]]; then
      echo "Tar bort misslyckad container CT $CTID..."
      pct stop "$CTID" >/dev/null 2>&1 || true
      pct destroy "$CTID" --purge >/dev/null 2>&1 || true
    else
      echo "KEEP_FAILED_CONTAINER=1 - lämnar CT $CTID kvar för felsökning."
    fi
  fi
  exit $rc
}
trap cleanup_failed_container EXIT

echo "Skapar LXC..."
pct create "$CTID" "$TEMPLATE" \
  --hostname "$LXC_HOSTNAME" \
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
  if pct exec "$CTID" -- bash -lc 'getent hosts github.com >/dev/null 2>&1'; then
    break
  fi
  sleep 2
done

echo "Installerar systemkrav..."
pct exec "$CTID" -- bash -lc "
  set -euo pipefail
  export DEBIAN_FRONTEND=noninteractive
  apt-get update
  apt-get install -y \
    ca-certificates \
    curl \
    unzip \
    python3 \
    python3-venv \
    python3-pip \
    locales

  if ! locale -a | grep -qi '^en_US\.utf8$'; then
    sed -i 's/^# *en_US.UTF-8 UTF-8/en_US.UTF-8 UTF-8/' /etc/locale.gen || true
    locale-gen en_US.UTF-8
  fi

  update-locale LANG=en_US.UTF-8
  export LANG=en_US.UTF-8
  export LC_ALL=en_US.UTF-8

  echo 'Preflight Python:'
  python3 --version
  python3 -m venv --help >/dev/null
"

echo "Installerar $TAG..."
pct exec "$CTID" -- bash -lc "
  set -euo pipefail
  export LANG=en_US.UTF-8
  export LC_ALL=en_US.UTF-8

  rm -rf /tmp/project-planer
  mkdir -p /tmp/project-planer
  curl -fL '${ARCHIVE_URL}' -o /tmp/project-planer/repo.zip
  unzip -q /tmp/project-planer/repo.zip -d /tmp/project-planer

  SRC=\$(find /tmp/project-planer -mindepth 1 -maxdepth 1 -type d | head -1)
  test -n \"\$SRC\"
  test -f \"\$SRC/install-app.sh\"

  chmod +x \"\$SRC/install-app.sh\"
  \"\$SRC/install-app.sh\"
"

echo "Verifierar tjänsten..."
pct exec "$CTID" -- bash -lc "
  set -euo pipefail
  systemctl is-active --quiet project-plan
  curl -fsS http://127.0.0.1:8080/health
"

IP="$(pct exec "$CTID" -- hostname -I | awk '{print $1}')"

trap - EXIT

echo
echo "============================================"
echo "KLART"
echo "Version:   $TAG"
echo "CTID:      $CTID"
echo "Hostname:  $LXC_HOSTNAME"
echo "LXC root:  root / $PASSWORD"
echo "App:       http://${IP}:8080"
echo "============================================"
