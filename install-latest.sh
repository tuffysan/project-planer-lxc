#!/usr/bin/env bash
set -euo pipefail
REPO="${REPO:-tuffysan/project-planer-lxc}"
CTID="${CTID:-200}"
CTID="$CTID" VERSION=latest REPO="$REPO" bash -c "$(curl -fsSL -H 'Cache-Control: no-cache' "https://raw.githubusercontent.com/${REPO}/main/deploy-lxc.sh?nocache=$(date +%s)")"
