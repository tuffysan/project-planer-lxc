#!/usr/bin/env bash
set -euo pipefail

APP_ROOT="${APP_DIR:-/opt/project-plan}"
PORT="${PORT:-8080}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SOURCE_DIR="$SCRIPT_DIR"

RELEASES_DIR="$APP_ROOT/releases"
VENVS_DIR="$APP_ROOT/venvs"
DATA_DIR="$APP_ROOT/data"
BACKUPS_DIR="$APP_ROOT/backups"
CURRENT_LINK="$APP_ROOT/current"
CURRENT_VENV_LINK="$APP_ROOT/current-venv"

VERSION_VALUE="2.1.3"
if [[ -f "$SOURCE_DIR/VERSION" ]]; then
  VERSION_VALUE="$(tr -d '[:space:]' < "$SOURCE_DIR/VERSION")"
fi
[[ -n "$VERSION_VALUE" ]] || VERSION_VALUE="2.1.3"

RELEASE_KEY="v${VERSION_VALUE}"
if [[ -e "$RELEASES_DIR/$RELEASE_KEY" || -e "$VENVS_DIR/$RELEASE_KEY" ]]; then
  RELEASE_KEY="${RELEASE_KEY}-$(date +%Y%m%d%H%M%S)"
fi

RELEASE_DIR="$RELEASES_DIR/$RELEASE_KEY"
VENV_DIR="$VENVS_DIR/$RELEASE_KEY"
SERVICE_FILE="/etc/systemd/system/project-plan.service"
SERVICE_BACKUP="/etc/systemd/system/project-plan.service.pre-v213"

echo "=== Project Planer app installer ==="
echo "Source:      $SOURCE_DIR"
echo "App root:    $APP_ROOT"
echo "Release:     $RELEASE_KEY"
echo "Release dir: $RELEASE_DIR"
echo "Venv dir:    $VENV_DIR"
echo "Port:        $PORT"

for cmd in python3 curl rsync systemctl; do
  command -v "$cmd" >/dev/null 2>&1 || {
    echo "FEL: Saknar $cmd"
    exit 1
  }
done

python3 -c 'import venv' >/dev/null 2>&1 || {
  echo "FEL: Python venv-modulen saknas. Installera python3-venv."
  exit 1
}

for required in \
  "$SOURCE_DIR/requirements.txt" \
  "$SOURCE_DIR/app/app.py" \
  "$SOURCE_DIR/app/templates" \
  "$SOURCE_DIR/app/static"
do
  if [[ ! -e "$required" ]]; then
    echo "FEL: Releasepaketet saknar: $required"
    exit 1
  fi
done

mkdir -p "$APP_ROOT" "$RELEASES_DIR" "$VENVS_DIR" "$DATA_DIR" "$BACKUPS_DIR"

PREVIOUS_RELEASE=""
PREVIOUS_VENV=""
[[ -L "$CURRENT_LINK" ]] && PREVIOUS_RELEASE="$(readlink "$CURRENT_LINK" || true)"
[[ -L "$CURRENT_VENV_LINK" ]] && PREVIOUS_VENV="$(readlink "$CURRENT_VENV_LINK" || true)"

# Save the currently installed unit so the first migration from legacy layout
# can also be rolled back.
if [[ -f "$SERVICE_FILE" ]]; then
  cp -a "$SERVICE_FILE" "$SERVICE_BACKUP"
fi

cleanup_candidate() {
  rm -rf "$RELEASE_DIR" "$VENV_DIR"
}

rollback() {
  echo "Återställer föregående fungerande release..."
  systemctl stop project-plan >/dev/null 2>&1 || true

  if [[ -n "$PREVIOUS_RELEASE" && -n "$PREVIOUS_VENV" ]]; then
    ln -sfn "$PREVIOUS_RELEASE" "${CURRENT_LINK}.rollback"
    mv -Tf "${CURRENT_LINK}.rollback" "$CURRENT_LINK"
    ln -sfn "$PREVIOUS_VENV" "${CURRENT_VENV_LINK}.rollback"
    mv -Tf "${CURRENT_VENV_LINK}.rollback" "$CURRENT_VENV_LINK"
  else
    rm -f "$CURRENT_LINK" "$CURRENT_VENV_LINK"
    if [[ -f "$SERVICE_BACKUP" ]]; then
      cp -a "$SERVICE_BACKUP" "$SERVICE_FILE"
    fi
  fi

  systemctl daemon-reload || true
  systemctl start project-plan >/dev/null 2>&1 || true
  cleanup_candidate
}


# v15.2.6 release-integrity guard: do not activate a package that lacks
# the global Excel navigation promised by this release.
if [[ "$VERSION_VALUE" == "15.2.6" || "$VERSION_VALUE" == "15.2.8" ]]; then
  if ! grep -Fq 'href="/excel"' "$SOURCE_DIR/app/templates/base.html"; then
    echo "FEL: release saknar Excel-genvägen i den faktiska huvudnavigationen." >&2
    exit 1
  fi
fi

echo "Kopierar release..."
mkdir -p "$RELEASE_DIR"
rsync -a \
  --exclude '.git' \
  --exclude '__pycache__' \
  --exclude '.venv' \
  --exclude 'data' \
  --exclude 'backups' \
  --exclude 'releases' \
  --exclude 'venvs' \
  --exclude '*.pyc' \
  "$SOURCE_DIR/" "$RELEASE_DIR/"

# Persistent runtime directories. Path.resolve() in app.py resolves current ->
# release, therefore each release gets a data symlink back to APP_ROOT/data.
ln -sfn ../../data "$RELEASE_DIR/data"
ln -sfn ../../backups "$RELEASE_DIR/backups"

echo "Skapar virtuell miljö på slutlig sökväg..."
python3 -m venv "$VENV_DIR"
"$VENV_DIR/bin/python" -m pip install --upgrade pip
"$VENV_DIR/bin/pip" install -r "$RELEASE_DIR/requirements.txt"

echo "Verifierar v16.1.1 Unified UI-fixar..."
! grep -Fq 'db_connect()' "$RELEASE_DIR/app/app.py" || { echo "Ogiltig db_connect-referens finns kvar."; exit 1; }
! grep -Fq 'class="ux-v1600-nav"' "$RELEASE_DIR/app/templates/base.html" || { echo "Dubbla menyer finns kvar."; exit 1; }
grep -Fq '<a href="/start">Start</a>' "$RELEASE_DIR/app/templates/base.html" || { echo "Start-länk saknas."; exit 1; }

echo "Verifierar v17.0 Simple Planning Edition..."
grep -Fq '@app.get("/projects/<int:project_id>/plan")' "$RELEASE_DIR/app/app.py" || { echo "Simple Plan saknas."; exit 1; }
grep -Fq 'Ladda ner tom Excel-fil' "$RELEASE_DIR/app/templates/excel_start_v1525.html" || { echo "Tom Excel-mall saknas från UI."; exit 1; }
grep -Fq 'type="date" name="start_date"' "$RELEASE_DIR/app/templates/simple_plan_v1700.html" || { echo "Datumväljare saknas i Plan."; exit 1; }

echo "Verifierar v17.1 Unified Excel Edition..."
grep -Fq 'def excel_multi_project_workbook_v1530(existing_projects=None, include_project_data=False)' "$RELEASE_DIR/app/app.py" || { echo "Unified Excel saknas."; exit 1; }
grep -Fq 'Project-Planer-Excel-TOM.xlsx' "$RELEASE_DIR/app/app.py" || { echo "Tom Excel är inte Unified."; exit 1; }
grep -Fq 'excel_export_projects_v1710' "$RELEASE_DIR/app/app.py" || { echo "Export av befintliga projekt saknas."; exit 1; }

echo "Verifierar kandidat innan aktivering..."
test -x "$VENV_DIR/bin/python"
test -x "$VENV_DIR/bin/gunicorn"
"$VENV_DIR/bin/python" -c 'import flask, openpyxl, gunicorn'
(
  cd "$RELEASE_DIR"
  EXPECTED_VERSION="$VERSION_VALUE" "$VENV_DIR/bin/python" - <<'PY'
import os
from app.app import app, APP_VERSION
assert app is not None
expected = os.environ["EXPECTED_VERSION"]
print("Application import OK, version:", APP_VERSION)
if APP_VERSION != expected:
    raise SystemExit(f"VERSION MISMATCH: package VERSION={expected}, APP_VERSION={APP_VERSION}")
PY
)

echo "Verifierar v16.1 Connected Projects-källkod..."
grep -Fq 'def excel_project_ref_code_v1610' "$RELEASE_DIR/app/app.py" || { echo "Connected Projects parser saknas."; exit 1; }
grep -Fq '_Projektlista' "$RELEASE_DIR/app/app.py" || { echo "Connected Projects dropdown-helper saknas."; exit 1; }
grep -Fq 'excel_multi_connected_template_v1610' "$RELEASE_DIR/app/app.py" || { echo "Connected Projects route saknas."; exit 1; }

echo "Verifierar v16.2 Excel Visual UX-källkod..."
grep -Fq 'def add_date_validation' "$RELEASE_DIR/app/app.py" || { echo "Datumvalidering saknas."; exit 1; }
grep -Fq 'visual_ux_version","16.2.1' "$RELEASE_DIR/app/app.py" || { echo "Visual UX metadata saknas."; exit 1; }
grep -Fq 'Försenad aktivitet' "$RELEASE_DIR/app/app.py" || { echo "Förseningsmarkering saknas."; exit 1; }

python3 - "$RELEASE_DIR/app/app.py" <<'PY'
import sys
from pathlib import Path
src=Path(sys.argv[1]).read_text(encoding="utf-8")
a=src.index("def excel_blank_complete_workbook_v1525")
b=src.index("def excel_new_project_from_workbook_v1525")
single=src[a:b]
if "for _sheet_name,_headers,_dates,_money in core:" in single:
    raise SystemExit("Single-project Excel innehåller felaktig 4-fälts core-loop.")
m=src.index("def excel_multi_project_workbook_v1530")
n=src.index("def excel_multi_project_import_v1530")
multi=src[m:n]
if "for _sheet_name,_headers,_dates,_money in core:" not in multi:
    raise SystemExit("Multi-Project datumvalidering saknas.")
print("Excel Visual UX scope check OK")
PY

echo "Verifierar Excel-mallen med riktig runtime..."
(
  cd "$RELEASE_DIR"
  "$VENV_DIR/bin/python" - <<'PY'
from app.app import excel_blank_complete_workbook_v1525, excel_multi_project_workbook_v1530, excel_serialize_workbook_v1527
for label,builder in [
    ("single",excel_blank_complete_workbook_v1525),
    ("multi",excel_multi_project_workbook_v1530),
]:
    wb=builder()
    payload=excel_serialize_workbook_v1527(wb)
    if len(payload) < 5000:
        raise SystemExit(f"XLSX SMOKE TEST FAILED ({label}): generated workbook is unexpectedly small")
    if label=="single":
        if "Datamodell" not in wb.sheetnames:
            raise SystemExit("XLSX SINGLE SMOKE TEST FAILED: Datamodell missing")
        if wb["Datamodell"].sheet_state!="hidden":
            raise SystemExit("XLSX SINGLE SMOKE TEST FAILED: Datamodell must be hidden")
    if label=="multi":
        required={"Start","Projekt","Uppgifter","Gantt","Kontroll"}
        missing=required-set(wb.sheetnames)
        if missing:
            raise SystemExit("XLSX UX SMOKE TEST FAILED: missing sheets "+", ".join(sorted(missing)))
        if wb["Start"].sheet_state!="visible" or wb["Uppgifter"].sheet_state!="visible":
            raise SystemExit("XLSX UX SMOKE TEST FAILED: primary sheets are hidden")
        if wb.sheetnames[0]!="Start":
            raise SystemExit("XLSX UX SMOKE TEST FAILED: Start is not the first sheet")
        if "Datamodell" in wb.sheetnames and wb["Datamodell"].sheet_state!="hidden":
            raise SystemExit("XLSX UX SMOKE TEST FAILED: Datamodell must be hidden")
    print(f"XLSX smoke test OK ({label}):", len(payload), "bytes,", len(wb.sheetnames), "sheets")
PY
)

echo "Installerar systemd service..."
cat > "$SERVICE_FILE" <<EOF
[Unit]
Description=Project Planer
After=network.target

[Service]
Type=simple
WorkingDirectory=$CURRENT_LINK
Environment=PORT=$PORT
Environment=PYTHONUNBUFFERED=1
Environment=LANG=C.UTF-8
Environment=LC_ALL=C.UTF-8
ExecStart=$CURRENT_VENV_LINK/bin/gunicorn --workers 2 --threads 4 --bind 0.0.0.0:\${PORT} --timeout 120 app.app:app
Restart=always
RestartSec=3
User=root

[Install]
WantedBy=multi-user.target
EOF

echo "Aktiverar release..."
systemctl stop project-plan >/dev/null 2>&1 || true

ln -sfn "$RELEASE_DIR" "${CURRENT_LINK}.new"
mv -Tf "${CURRENT_LINK}.new" "$CURRENT_LINK"

ln -sfn "$VENV_DIR" "${CURRENT_VENV_LINK}.new"
mv -Tf "${CURRENT_VENV_LINK}.new" "$CURRENT_VENV_LINK"

# Explicitly prove the symlinked executable exists before systemd starts.
if [[ ! -x "$CURRENT_VENV_LINK/bin/gunicorn" ]]; then
  echo "FEL: gunicorn saknas efter aktivering: $CURRENT_VENV_LINK/bin/gunicorn"
  rollback
  exit 1
fi

if ! "$CURRENT_VENV_LINK/bin/python" -c 'import sys; print("Runtime Python:", sys.executable)'; then
  echo "FEL: Python i aktiverad venv kan inte köras."
  rollback
  exit 1
fi

systemctl daemon-reload
systemctl enable project-plan >/dev/null

if ! systemctl start project-plan; then
  echo "FEL: Kunde inte starta project-plan."
  journalctl -u project-plan --no-pager -n 120 || true
  rollback
  exit 1
fi

echo "Väntar på health check..."
healthy=0
for i in $(seq 1 45); do
  if curl -fsS "http://127.0.0.1:${PORT}/health" >/tmp/project-plan-health.json 2>/dev/null; then
    healthy=1
    break
  fi
  sleep 1
done

if [[ "$healthy" != "1" ]]; then
  echo "FEL: health check misslyckades."
  systemctl status project-plan --no-pager || true
  journalctl -u project-plan --no-pager -n 160 || true
  rollback
  exit 1
fi

echo "Health check svar:"
cat /tmp/project-plan-health.json || true
echo

echo "Verifierar att den körande appen verkligen är version $VERSION_VALUE..."
RUNTIME_VERSION="$("$CURRENT_VENV_LINK/bin/python" - <<'PY'
import json, urllib.request
with urllib.request.urlopen("http://127.0.0.1:8080/health", timeout=10) as r:
    data=json.load(r)
print(data.get("version",""))
PY
)"
if [[ "$RUNTIME_VERSION" != "$VERSION_VALUE" ]]; then
  echo "FEL: Körande app rapporterar version '$RUNTIME_VERSION', förväntat '$VERSION_VALUE'."
  echo "Current release: $(readlink -f "$CURRENT_LINK" || true)"
  echo "Current venv:    $(readlink -f "$CURRENT_VENV_LINK" || true)"
  systemctl status project-plan --no-pager || true
  journalctl -u project-plan --no-pager -n 120 || true
  rollback
  exit 1
fi
echo "Runtime version OK: $RUNTIME_VERSION"

rm -f "$SERVICE_BACKUP"

# Keep the current release plus the two most recent older releases/venvs.
# Never remove the targets currently referenced by the symlinks.
CURRENT_RELEASE_TARGET="$(readlink -f "$CURRENT_LINK" || true)"
CURRENT_VENV_TARGET="$(readlink -f "$CURRENT_VENV_LINK" || true)"

prune_dir() {
  local parent="$1"
  local keep_target="$2"
  local count=0
  while IFS= read -r candidate; do
    [[ "$candidate" == "$keep_target" ]] && continue
    count=$((count+1))
    if (( count > 2 )); then
      rm -rf "$candidate"
    fi
  done < <(find "$parent" -mindepth 1 -maxdepth 1 -type d -printf '%T@ %p\n' | sort -nr | cut -d' ' -f2-)
}

prune_dir "$RELEASES_DIR" "$CURRENT_RELEASE_TARGET"
prune_dir "$VENVS_DIR" "$CURRENT_VENV_TARGET"

echo "Project Planer $VERSION_VALUE installerad."
echo "Current: $(readlink -f "$CURRENT_LINK")"
echo "Venv:    $(readlink -f "$CURRENT_VENV_LINK")"
