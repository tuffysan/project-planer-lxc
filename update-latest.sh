#!/usr/bin/env bash
set -euo pipefail
REPO="${REPO:-tuffysan/project-planer-lxc}"
CTID="${CTID:-200}"
CTID="$CTID" VERSION=latest REPO="$REPO" bash -c "$(curl -fsSL "https://raw.githubusercontent.com/${REPO}/main/update-lxc.sh")"
