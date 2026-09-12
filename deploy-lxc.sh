#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION="${VERSION:-latest}"
CTID="${CTID:-200}"

command -v pct >/dev/null 2>&1 || { echo "FEL: Kör detta på Proxmox VE-host." >&2; exit 1; }

echo "=== Project Planer smart deploy ==="
echo "Repo: $REPO"
echo "CTID: $CTID"

if pct status "$CTID" >/dev/null 2>&1; then
  echo "CT $CTID finns -> uppgraderar till vald/senaste release."
  export REPO VERSION CTID
  exec bash -c "$(curl -fsSL "https://raw.githubusercontent.com/${REPO}/main/update-lxc.sh")"
else
  echo "CT $CTID finns inte -> installerar vald/senaste release."
  export REPO VERSION CTID
  exec bash -c "$(curl -fsSL "https://raw.githubusercontent.com/${REPO}/main/install-lxc.sh")"
fi
