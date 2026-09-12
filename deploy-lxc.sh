#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-planer-lxc}"
REQUESTED_VERSION="${VERSION:-latest}"
CTID="${CTID:-200}"

command -v pct >/dev/null 2>&1 || {
  echo "FEL: Kör detta script på Proxmox VE-host." >&2
  exit 1
}

resolve_version() {
  local requested="$1"
  if [[ -z "$requested" || "$requested" == "latest" || "$requested" == "LATEST" ]]; then
    local tag
    tag="$(
      curl -fsSL \
        -H 'Accept: application/vnd.github+json' \
        -H 'Cache-Control: no-cache' \
        "https://api.github.com/repos/${REPO}/releases/latest?nocache=$(date +%s)" \
      | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' \
      | head -1
    )"
    [[ -n "$tag" ]] || {
      echo "FEL: Kunde inte läsa senaste GitHub-release för ${REPO}." >&2
      exit 1
    }
    echo "${tag#v}"
  else
    echo "${requested#v}"
  fi
}

VERSION="$(resolve_version "$REQUESTED_VERSION")"
TAG="v${VERSION}"

echo "=== Project Planer smart deploy ==="
echo "Repo:    $REPO"
echo "Version: $TAG"
echo "CTID:    $CTID"

# Critical: run deployment script from the immutable tag we are deploying.
if pct status "$CTID" >/dev/null 2>&1; then
  SCRIPT_URL="https://raw.githubusercontent.com/${REPO}/${TAG}/update-lxc.sh?nocache=$(date +%s)"
  echo "CT $CTID finns -> uppgraderar med updateraren från $TAG."
else
  SCRIPT_URL="https://raw.githubusercontent.com/${REPO}/${TAG}/install-lxc.sh?nocache=$(date +%s)"
  echo "CT $CTID finns inte -> installerar med installeraren från $TAG."
fi

echo "Deploy-script: $SCRIPT_URL"
SCRIPT="$(curl -fsSL -H 'Cache-Control: no-cache' "$SCRIPT_URL")"

# Sanity check so a stale or wrong script can never silently run.
printf '%s\n' "$SCRIPT" | grep -q 'Project Planer' || {
  echo "FEL: Hämtat deploy-script ser inte giltigt ut." >&2
  exit 1
}

export REPO VERSION CTID
exec bash -c "$SCRIPT"
