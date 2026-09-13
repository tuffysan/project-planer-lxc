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


echo "[v15.2.6] Verifying global Excel navigation..."
grep -Fq 'href="/excel"' app/templates/base.html || {
  echo "ERROR: Global Excel navigation is missing from app/templates/base.html" >&2
  exit 1
}
grep -Fq 'href="/excel"' app/templates/base.html || {
  echo "ERROR: /excel navigation link is missing from app/templates/base.html" >&2
  exit 1
}

echo "[v15.2.8] Verifying XLSX integrity code..."
grep -Fq 'def excel_xlsx_integrity_check_v1527' app/app.py || {
  echo "ERROR: XLSX integrity checker missing" >&2
  exit 1
}
grep -Fq 'def excel_serialize_workbook_v1527' app/app.py || {
  echo "ERROR: XLSX normalization/serialization missing" >&2
  exit 1
}
grep -Fq 'ws.auto_filter.ref=None' app/app.py || {
  echo "ERROR: Blank-sheet AutoFilter cleanup missing" >&2
  exit 1
}


echo "[v15.2.9] Verifying Projectinformation does not overlap merged subtitle..."
grep -Fq 'enumerate(info_rows,7)' app/app.py || {
  echo "ERROR: Projectinformation must start at row 7; rows 4-5 are merged." >&2
  exit 1
}
if grep -Fq 'enumerate(info_rows,4)' app/app.py; then
  echo "ERROR: Regression: Projectinformation starts inside merged subtitle rows." >&2
  exit 1
fi

echo "[v15.3.0] Verifying Multi-Project Excel..."
grep -Fq 'def excel_multi_project_workbook_v1530' app/app.py || { echo "ERROR: Multi-Project workbook builder missing" >&2; exit 1; }
grep -Fq 'def excel_multi_project_import_v1530' app/app.py || { echo "ERROR: Multi-Project importer missing" >&2; exit 1; }
grep -Fq 'template_kind","multi_project_complete' app/app.py || { echo "ERROR: Multi-Project metadata missing" >&2; exit 1; }
grep -Fq 'Projektkod' app/templates/excel_multi_import_v1530.html || { echo "ERROR: Multi-Project UI missing" >&2; exit 1; }

echo "[v15.3.1] Verifying automatic Excel project codes..."
grep -Fq 'def excel_project_code_formula_v1531' app/app.py || { echo "ERROR: project code formula helper missing" >&2; exit 1; }
grep -Fq 'PRJ-' app/app.py || { echo "ERROR: PRJ code format missing" >&2; exit 1; }
grep -Fq "formula1=\"'Projekt'!\\$A\\$2:\\$A\\$101\"" app/app.py || { echo "ERROR: project-code dropdown missing" >&2; exit 1; }

echo "[v15.3.2] Verifying openpyxl Comment import..."
grep -Fq 'from openpyxl.comments import Comment' app/app.py || {
  echo "ERROR: Multi-Project Excel uses Comment but openpyxl Comment import is missing." >&2
  exit 1
}

echo "[v15.4.0] Verifying simplified Excel planning..."
grep -Fq 'def excel_generate_wbs_from_levels_v1540' app/app.py || { echo "ERROR: WBS generation helper missing" >&2; exit 1; }
grep -Fq '"Aktivitetsnummer"' app/app.py || { echo "ERROR: Aktivitetsnummer column missing" >&2; exit 1; }
grep -Fq '"Nivå"' app/app.py || { echo "ERROR: Nivå column missing" >&2; exit 1; }
grep -Fq '"Föregående aktivitet"' app/app.py || { echo "ERROR: simplified dependency column missing" >&2; exit 1; }
grep -Fq 'excel_resolve_activity_ref_v1540' app/app.py || { echo "ERROR: activity dependency resolver missing" >&2; exit 1; }

echo "[v15.5.0] Verifying Excel UX Edition..."
grep -Fq 'wb.create_sheet("Start")' app/app.py || { echo "ERROR: Start sheet missing" >&2; exit 1; }
grep -Fq 'wb.create_sheet("Gantt")' app/app.py || { echo "ERROR: Gantt sheet missing" >&2; exit 1; }
grep -Fq 'wb.create_sheet("Kontroll")' app/app.py || { echo "ERROR: Kontroll sheet missing" >&2; exit 1; }
grep -Fq 'Varaktighet dagar' app/app.py || { echo "ERROR: duration UX missing" >&2; exit 1; }
grep -Fq 'Huvudaktivitet' app/app.py || { echo "ERROR: friendly level labels missing" >&2; exit 1; }
grep -Fq 'wb_values=load_workbook' app/app.py || { echo "ERROR: formula-value workbook missing" >&2; exit 1; }

echo "[v16.0.0] Verifying Unified UX Edition..."
grep -Fq '@app.route("/start")' app/app.py || { echo "ERROR: unified start route missing" >&2; exit 1; }
grep -Fq '/projects/<int:project_id>/overview' app/app.py || { echo "ERROR: project overview route missing" >&2; exit 1; }
test -f app/templates/unified_start_v1600.html || { echo "ERROR: unified start template missing" >&2; exit 1; }
test -f app/templates/project_overview_v1600.html || { echo "ERROR: project overview template missing" >&2; exit 1; }
grep -Fq 'ux-v1600-nav' app/templates/base.html || { echo "ERROR: unified navigation missing" >&2; exit 1; }
grep -Fq 'Unified UX Edition' README.md || { echo "ERROR: README not updated" >&2; exit 1; }
