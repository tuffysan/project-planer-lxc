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

echo "[v16.0.1] Verifying Excel runtime regression fixes..."
grep -Fq 'dd.sheet_state="hidden"' app/app.py || { echo "ERROR: single Datamodell fix missing" >&2; exit 1; }
grep -Fq 'wb_values=load_workbook(BytesIO(payload),data_only=True)' app/app.py || { echo "ERROR: wb_values load missing" >&2; exit 1; }
grep -Fq 'wb.sheetnames[0]!="Start"' install-app.sh || { echo "ERROR: multi Start order smoke check missing" >&2; exit 1; }

echo "[v16.1.0] Verifying Excel Connected Projects..."
grep -Fq 'def excel_project_ref_code_v1610' app/app.py || { echo "ERROR: project ref parser missing" >&2; exit 1; }
grep -Fq '/excel/template/multi/connected' app/app.py || { echo "ERROR: connected download route missing" >&2; exit 1; }
grep -Fq '_ProjectID' app/app.py || { echo "ERROR: ProjectID support missing" >&2; exit 1; }
grep -Fq '_Projektlista' app/app.py || { echo "ERROR: same-sheet project dropdown helper missing" >&2; exit 1; }
grep -Fq 'updated_projects' app/app.py || { echo "ERROR: existing project update support missing" >&2; exit 1; }

echo "[v16.1.1] Verifying Unified UI hotfix..."
! grep -Fq 'db_connect()' app/app.py || { echo "ERROR: db_connect regression remains" >&2; exit 1; }
! grep -Fq 'require_project_access(project_id)' app/app.py || { echo "ERROR: invalid project access helper remains" >&2; exit 1; }
! grep -Fq 'class="ux-v1600-nav"' app/templates/base.html || { echo "ERROR: duplicate v16 navigation remains" >&2; exit 1; }
grep -Fq '<a href="/start">Start</a>' app/templates/base.html || { echo "ERROR: Start nav link missing" >&2; exit 1; }
grep -Fq 'def unified_start_v1600' app/app.py || { echo "ERROR: unified start route missing" >&2; exit 1; }
grep -Fq 'project=project_or_404(project_id)' app/app.py || { echo "ERROR: project overview access fix missing" >&2; exit 1; }

echo "[v16.2.0] Verifying Excel Visual UX..."
grep -Fq 'def add_date_validation' app/app.py || { echo "ERROR: date validation helper missing" >&2; exit 1; }
grep -Fq 'DataBarRule(start_type="num",start_value=0,end_type="num",end_value=100,color="5B9BD5")' app/app.py || { echo "ERROR: progress data bar missing" >&2; exit 1; }
grep -Fq 'Försenad aktivitet' app/app.py || { echo "ERROR: overdue activity visual rule missing" >&2; exit 1; }
grep -Fq 'Färgkodning' app/app.py || { echo "ERROR: color legend missing" >&2; exit 1; }
grep -Fq 'visual_ux_version' app/app.py || { echo "ERROR: Visual UX metadata missing" >&2; exit 1; }

echo "[compat] Verifying Excel Visual UX runtime scope fix..."
python3 - <<'PY'
from pathlib import Path
src=Path("app/app.py").read_text(encoding="utf-8")
a=src.index("def excel_blank_complete_workbook_v1525")
b=src.index("def excel_new_project_from_workbook_v1525")
single=src[a:b]
assert "for _sheet_name,_headers,_dates,_money in core:" not in single
m=src.index("def excel_multi_project_workbook_v1530")
n=src.index("def excel_multi_project_import_v1530")
multi=src[m:n]
assert "for _sheet_name,_headers,_dates,_money in core:" in multi
assert '("visual_ux_version",' in multi
print("Visual UX scope regression checks OK")
PY

echo "[v17.0.0] Verifying Simple Planning Edition..."
grep -Fq '@app.get("/projects/<int:project_id>/plan")' app/app.py || { echo "ERROR: simple Plan route missing" >&2; exit 1; }
grep -Fq 'def simple_next_wbs_v1700' app/app.py || { echo "ERROR: automatic activity numbering missing" >&2; exit 1; }
grep -Fq '@app.route("/excel/import",methods=["GET","POST"])' app/app.py || { echo "ERROR: unified Excel import missing" >&2; exit 1; }
test -f app/templates/simple_plan_v1700.html || { echo "ERROR: simple Plan template missing" >&2; exit 1; }
test -f app/templates/excel_import_unified_v1700.html || { echo "ERROR: unified Excel import template missing" >&2; exit 1; }
grep -Fq 'Ladda ner tom Excel' app/templates/excel_start_v1525.html || { echo "ERROR: blank Excel download is no longer visible" >&2; exit 1; }
grep -Fq 'url_for('\''excel_blank_template_v1525'\'')' app/templates/excel_start_v1525.html || { echo "ERROR: blank Excel route missing from Excel Center" >&2; exit 1; }
grep -Fq 'type="date" name="start_date"' app/templates/simple_plan_v1700.html || { echo "ERROR: web date picker missing" >&2; exit 1; }

echo "[v17.1.0] Verifying Unified Excel Edition..."
grep -Fq 'def excel_multi_project_workbook_v1530(existing_projects=None, include_project_data=False)' app/app.py || { echo "ERROR: unified workbook signature missing" >&2; exit 1; }
grep -Fq '@app.route("/excel/export-projects",methods=["GET","POST"])' app/app.py || { echo "ERROR: project export selector missing" >&2; exit 1; }
grep -Fq '@app.get("/projects/<int:project_id>/excel/unified")' app/app.py || { echo "ERROR: project unified export missing" >&2; exit 1; }
grep -Fq 'download_name="Project-Planer-Excel-TOM.xlsx"' app/app.py || { echo "ERROR: blank template is not Unified Excel" >&2; exit 1; }
grep -Fq 'existing_row_id=excel_int' app/app.py || { echo "ERROR: row IDs not imported" >&2; exit 1; }
grep -Fq 'excel_update(conn,spec["table"]' app/app.py || { echo "ERROR: existing child rows are not updated" >&2; exit 1; }
grep -Fq 'Exportera projekt till Excel' app/templates/excel_start_v1525.html || { echo "ERROR: export existing projects UI missing" >&2; exit 1; }
test -f app/templates/excel_export_projects_v1710.html || { echo "ERROR: export picker template missing" >&2; exit 1; }

echo "[v17.1.2] Verifying version-independent validators..."
grep -Fq "grep -Fq 'visual_ux_version' app/app.py" VERIFY.sh || { echo "ERROR: generic Visual UX validator missing" >&2; exit 1; }
grep -Fq "grep -Fq 'visual_ux_version' "\$RELEASE_DIR/app/app.py"" install-app.sh || { echo "ERROR: generic installer Visual UX validator missing" >&2; exit 1; }
grep -Fq 'APP_VERSION = "17.1.2"' app/app.py || { echo "ERROR: APP_VERSION mismatch" >&2; exit 1; }
