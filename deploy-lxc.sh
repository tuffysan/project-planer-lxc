#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION="${VERSION:-latest}"
CTID="${CTID:-200}"

resolve_version() {
  local requested="${1:-latest}"
  if [[ -z "$requested" || "$requested" == "latest" || "$requested" == "LATEST" ]]; then
    local latest_tag
    latest_tag="$(
      curl -fsSL -H 'Accept: application/vnd.github+json' \
        "https://api.github.com/repos/${REPO}/releases/latest" \
      | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' \
      | head -1
    )"
    [[ -n "$latest_tag" ]] || {
      echo "FEL: Kunde inte fastställa senaste GitHub-release för ${REPO}." >&2
      exit 1
    }
    echo "${latest_tag#v}"
  else
    echo "${requested#v}"
  fi
}

VERSION="$(resolve_version "$VERSION")"

echo "=== Project Planer LXC deploy ==="
echo "Repo:    $REPO"
echo "Version: v$VERSION"
echo "CTID:    $CTID"

if pct status "$CTID" >/dev/null 2>&1; then
  echo "CT $CTID finns redan -> uppgraderar till senaste release."
  CTID="$CTID" VERSION="$VERSION" REPO="$REPO" \
    bash -c "$(curl -fsSL "https://raw.githubusercontent.com/${REPO}/main/update-lxc.sh")"
else
  echo "CT $CTID finns inte -> installerar senaste release."
  CTID="$CTID" VERSION="$VERSION" REPO="$REPO" \
    bash -c "$(curl -fsSL "https://raw.githubusercontent.com/${REPO}/main/install-lxc.sh")"
fi
