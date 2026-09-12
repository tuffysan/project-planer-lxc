#!/usr/bin/env bash
set -euo pipefail
REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION=latest REPO="$REPO" bash -c "$(curl -fsSL "https://raw.githubusercontent.com/${REPO}/main/install-lxc.sh")"
