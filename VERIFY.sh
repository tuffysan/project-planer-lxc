#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
EXPECTED="$(tr -d '[:space:]' < "$ROOT/VERSION")"

python3 -m py_compile "$ROOT/app/app.py"

for s in install-app.sh install-lxc.sh update-lxc.sh deploy-lxc.sh install-latest.sh update-latest.sh; do
  [[ -f "$ROOT/$s" ]] && bash -n "$ROOT/$s"
done

APP_VERSION="$(sed -n 's/^APP_VERSION = "\([^"]*\)".*/\1/p' "$ROOT/app/app.py" | head -1)"
[[ "$APP_VERSION" == "$EXPECTED" ]] || {
  echo "FEL: VERSION=$EXPECTED men APP_VERSION=$APP_VERSION"
  exit 1
}

python3 "$ROOT/VERIFY-PYTHON.py"

echo "VERIFY OK: v$EXPECTED"
