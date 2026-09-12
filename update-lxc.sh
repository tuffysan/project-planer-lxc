#!/usr/bin/env bash
set -euo pipefail

REPO="${REPO:-tuffysan/project-planer-lxc}"
VERSION="${VERSION:-latest}"
CTID="${CTID:-200}"

resolve_version() {
  local requested="${1:-latest}"
  if [[ -z "$requested" || "$requested" == "latest" || "$requested" == "LATEST" ]]; then
    local tag
    tag="$(
      curl -fsSL -H 'Accept: application/vnd.github+json' \
        "https://api.github.com/repos/${REPO}/releases/latest?nocache=$(date +%s)" \
      | sed -n 's/.*"tag_name"[[:space:]]*:[[:space:]]*"\([^"]*\)".*/\1/p' \
      | head -1
    )"
    [[ -n "$tag" ]] || { echo "FEL: Kunde inte läsa senaste GitHub-release." >&2; exit 1; }
    echo "${tag#v}"
  else
    echo "${requested#v}"
  fi
}

VERSION="$(resolve_version "$VERSION")"
TAG="v${VERSION}"
ARCHIVE_URL="https://github.com/${REPO}/archive/refs/tags/${TAG}.zip"

command -v pct >/dev/null 2>&1 || { echo "FEL: Kör detta script på Proxmox VE-host." >&2; exit 1; }
pct status "$CTID" >/dev/null 2>&1 || { echo "FEL: CT $CTID finns inte." >&2; exit 1; }

echo "=== Project Planer LXC updater ==="
echo "Repo:    $REPO"
echo "Version: $TAG"
echo "CTID:    $CTID"

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
export LANG=C.UTF-8
export LC_ALL=C.UTF-8
apt-get update -qq
apt-get install -y -qq --no-install-recommends ca-certificates curl unzip rsync python3 python3-venv locales >/dev/null
'

echo "Hämtar release..."
pct exec "$CTID" -- bash -s -- "$ARCHIVE_URL" "$VERSION" <<'REMOTE'
set -euo pipefail
ARCHIVE_URL="$1"
EXPECTED_VERSION="$2"

export LANG=C.UTF-8
export LC_ALL=C.UTF-8

rm -rf /tmp/project-planer-update
mkdir -p /tmp/project-planer-update

curl -fL "$ARCHIVE_URL" -o /tmp/project-planer-update/repo.zip
unzip -q /tmp/project-planer-update/repo.zip -d /tmp/project-planer-update

SRC="$(find /tmp/project-planer-update -mindepth 1 -maxdepth 1 -type d | head -1)"
[[ -n "$SRC" ]] || { echo "FEL: Releasekatalog hittades inte."; exit 1; }
[[ -f "$SRC/install-app.sh" ]] || { echo "FEL: install-app.sh saknas."; exit 1; }
[[ -f "$SRC/VERSION" ]] || { echo "FEL: VERSION saknas."; exit 1; }
[[ -f "$SRC/app/app.py" ]] || { echo "FEL: app/app.py saknas."; exit 1; }

PACKAGE_VERSION="$(tr -d '[:space:]' < "$SRC/VERSION")"
APP_VERSION="$(sed -n 's/^APP_VERSION = "\([^"]*\)".*/\1/p' "$SRC/app/app.py" | head -1)"

echo "Taggpaketets VERSION: $PACKAGE_VERSION"
echo "Taggpaketets APP_VERSION: $APP_VERSION"

[[ "$PACKAGE_VERSION" == "$EXPECTED_VERSION" ]] || {
  echo "FEL: VERSION=$PACKAGE_VERSION, förväntat $EXPECTED_VERSION."
  exit 1
}
[[ "$APP_VERSION" == "$EXPECTED_VERSION" ]] || {
  echo "FEL: APP_VERSION=$APP_VERSION, förväntat $EXPECTED_VERSION."
  exit 1
}

chmod +x "$SRC/install-app.sh"
"$SRC/install-app.sh"
REMOTE

echo "Verifierar installerad runtime..."
pct exec "$CTID" -- bash -s -- "$VERSION" <<'REMOTE'
set -euo pipefail
EXPECTED_VERSION="$1"

systemctl is-active --quiet project-plan
test -x /opt/project-plan/current-venv/bin/gunicorn

CURRENT_RELEASE="$(readlink -f /opt/project-plan/current)"
CURRENT_VENV="$(readlink -f /opt/project-plan/current-venv)"
HEALTH="$(curl -fsS http://127.0.0.1:8080/health)"
RUNTIME_VERSION="$(printf '%s' "$HEALTH" | sed -n 's/.*"version":"\([^"]*\)".*/\1/p')"

echo "Current release: $CURRENT_RELEASE"
echo "Current venv:    $CURRENT_VENV"
echo "Health:          $HEALTH"
echo "Runtime version: $RUNTIME_VERSION"

[[ "$CURRENT_RELEASE" == "/opt/project-plan/releases/v${EXPECTED_VERSION}" ]] || {
  echo "FEL: current pekar inte på v${EXPECTED_VERSION}."
  exit 41
}
[[ "$CURRENT_VENV" == "/opt/project-plan/venvs/v${EXPECTED_VERSION}" ]] || {
  echo "FEL: current-venv pekar inte på v${EXPECTED_VERSION}."
  exit 42
}
[[ "$RUNTIME_VERSION" == "$EXPECTED_VERSION" ]] || {
  echo "FEL: Runtime version '$RUNTIME_VERSION' matchar inte '$EXPECTED_VERSION'."
  exit 43
}
REMOTE

IP="$(pct exec "$CTID" -- hostname -I | awk '{print $1}')"

echo
echo "============================================"
echo "UPPGRADERING KLAR"
echo "Version: $TAG"
echo "CTID:    $CTID"
echo "App:     http://${IP}:8080"
echo "============================================"
