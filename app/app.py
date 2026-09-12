from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import sqlite3
from pathlib import Path
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "projectplan.db"

app = Flask(__name__)
app.secret_key = "project-plan-local-secret"

STATUSES = ["Ej startad", "Pågår", "Blockerad", "Klar", "Pausad"]
PRIORITIES = ["Låg", "Normal", "Hög", "Kritisk"]

def db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with db() as conn:
        conn.executescript("""
        CREATE TABLE IF NOT EXISTS projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            customer TEXT DEFAULT '',
            project_manager TEXT DEFAULT '',
            description TEXT DEFAULT '',
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            created_at TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS tasks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_id INTEGER NOT NULL,
            wbs TEXT DEFAULT '',
            title TEXT NOT NULL,
            owner TEXT DEFAULT '',
            start_date TEXT DEFAULT '',
            end_date TEXT DEFAULT '',
            status TEXT DEFAULT 'Ej startad',
            priority TEXT DEFAULT 'Normal',
            progress INTEGER DEFAULT 0,
            milestone INTEGER DEFAULT 0,
            dependencies TEXT DEFAULT '',
            notes TEXT DEFAULT '',
            sort_order INTEGER DEFAULT 0,
            FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
        );
        """)

def project_or_404(project_id):
    with db() as conn:
        row = conn.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
    if not row:
        from flask import abort
        abort(404)
    return row

@app.before_request
def ensure_db():
    init_db()

@app.route("/")
def index():
    with db() as conn:
        projects = conn.execute("""
            SELECT p.*,
                   COUNT(t.id) AS task_count,
                   COALESCE(ROUND(AVG(t.progress)), 0) AS avg_progress,
                   SUM(CASE WHEN t.status='Klar' THEN 1 ELSE 0 END) AS done_count,
                   SUM(CASE WHEN t.status='Blockerad' THEN 1 ELSE 0 END) AS blocked_count
            FROM projects p
            LEFT JOIN tasks t ON t.project_id=p.id
            GROUP BY p.id
            ORDER BY p.id DESC
        """).fetchall()
    return render_template("index.html", projects=projects)

@app.route("/projects/new", methods=["GET", "POST"])
def new_project():
    if request.method == "POST":
        with db() as conn:
            cur = conn.execute("""
                INSERT INTO projects(name, customer, project_manager, description, start_date, end_date, created_at)
                VALUES(?,?,?,?,?,?,?)
            """, (
                request.form["name"].strip(),
                request.form.get("customer","").strip(),
                request.form.get("project_manager","").strip(),
                request.form.get("description","").strip(),
                request.form.get("start_date",""),
                request.form.get("end_date",""),
                datetime.now().isoformat(timespec="seconds"),
            ))
            project_id = cur.lastrowid
        return redirect(url_for("project", project_id=project_id))
    return render_template("project_form.html", project=None)

@app.route("/projects/<int:project_id>/edit", methods=["GET", "POST"])
def edit_project(project_id):
    project = project_or_404(project_id)
    if request.method == "POST":
        with db() as conn:
            conn.execute("""
                UPDATE projects
                SET name=?, customer=?, project_manager=?, description=?, start_date=?, end_date=?
                WHERE id=?
            """, (
                request.form["name"].strip(),
                request.form.get("customer","").strip(),
                request.form.get("project_manager","").strip(),
                request.form.get("description","").strip(),
                request.form.get("start_date",""),
                request.form.get("end_date",""),
                project_id,
            ))
        flash("Projektet är uppdaterat.")
        return redirect(url_for("project", project_id=project_id))
    return render_template("project_form.html", project=project)

@app.post("/projects/<int:project_id>/delete")
def delete_project(project_id):
    with db() as conn:
        conn.execute("DELETE FROM tasks WHERE project_id=?", (project_id,))
        conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
    flash("Projektet är borttaget.")
    return redirect(url_for("index"))

@app.route("/projects/<int:project_id>")
def project(project_id):
    project = project_or_404(project_id)
    with db() as conn:
        tasks = conn.execute("""
            SELECT * FROM tasks WHERE project_id=?
            ORDER BY
              CASE WHEN wbs='' THEN 1 ELSE 0 END,
              wbs COLLATE NOCASE,
              sort_order,
              id
        """, (project_id,)).fetchall()

    avg = round(sum(t["progress"] for t in tasks) / len(tasks)) if tasks else 0
    blocked = sum(1 for t in tasks if t["status"] == "Blockerad")
    done = sum(1 for t in tasks if t["status"] == "Klar")
    return render_template("project.html", project=project, tasks=tasks,
                           avg=avg, blocked=blocked, done=done,
                           statuses=STATUSES, priorities=PRIORITIES)

@app.route("/projects/<int:project_id>/tasks/new", methods=["GET", "POST"])
def new_task(project_id):
    project = project_or_404(project_id)
    if request.method == "POST":
        with db() as conn:
            conn.execute("""
                INSERT INTO tasks(project_id,wbs,title,owner,start_date,end_date,status,priority,progress,milestone,dependencies,notes,sort_order)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                project_id,
                request.form.get("wbs","").strip(),
                request.form["title"].strip(),
                request.form.get("owner","").strip(),
                request.form.get("start_date",""),
                request.form.get("end_date",""),
                request.form.get("status","Ej startad"),
                request.form.get("priority","Normal"),
                max(0, min(100, int(request.form.get("progress","0") or 0))),
                1 if request.form.get("milestone") == "on" else 0,
                request.form.get("dependencies","").strip(),
                request.form.get("notes","").strip(),
                int(request.form.get("sort_order","0") or 0),
            ))
        return redirect(url_for("project", project_id=project_id))
    return render_template("task_form.html", project=project, task=None, statuses=STATUSES, priorities=PRIORITIES)

@app.route("/projects/<int:project_id>/tasks/<int:task_id>/edit", methods=["GET", "POST"])
def edit_task(project_id, task_id):
    project = project_or_404(project_id)
    with db() as conn:
        task = conn.execute("SELECT * FROM tasks WHERE id=? AND project_id=?", (task_id, project_id)).fetchone()
    if not task:
        from flask import abort
        abort(404)

    if request.method == "POST":
        with db() as conn:
            conn.execute("""
                UPDATE tasks SET wbs=?, title=?, owner=?, start_date=?, end_date=?, status=?, priority=?,
                                 progress=?, milestone=?, dependencies=?, notes=?, sort_order=?
                WHERE id=? AND project_id=?
            """, (
                request.form.get("wbs","").strip(),
                request.form["title"].strip(),
                request.form.get("owner","").strip(),
                request.form.get("start_date",""),
                request.form.get("end_date",""),
                request.form.get("status","Ej startad"),
                request.form.get("priority","Normal"),
                max(0, min(100, int(request.form.get("progress","0") or 0))),
                1 if request.form.get("milestone") == "on" else 0,
                request.form.get("dependencies","").strip(),
                request.form.get("notes","").strip(),
                int(request.form.get("sort_order","0") or 0),
                task_id, project_id
            ))
        return redirect(url_for("project", project_id=project_id))
    return render_template("task_form.html", project=project, task=task, statuses=STATUSES, priorities=PRIORITIES)

@app.post("/projects/<int:project_id>/tasks/<int:task_id>/delete")
def delete_task(project_id, task_id):
    with db() as conn:
        conn.execute("DELETE FROM tasks WHERE id=? AND project_id=?", (task_id, project_id))
    return redirect(url_for("project", project_id=project_id))

@app.post("/projects/<int:project_id>/tasks/<int:task_id>/quick")
def quick_task(project_id, task_id):
    status = request.form.get("status", "Ej startad")
    progress = max(0, min(100, int(request.form.get("progress","0") or 0)))
    if status == "Klar":
        progress = 100
    with db() as conn:
        conn.execute("UPDATE tasks SET status=?, progress=? WHERE id=? AND project_id=?",
                     (status, progress, task_id, project_id))
    return redirect(url_for("project", project_id=project_id))

def as_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return value

def autosize(ws):
    for col_cells in ws.columns:
        max_len = 0
        col = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            max_len = max(max_len, len(value))
        ws.column_dimensions[col].width = min(max(max_len + 2, 10), 45)

def style_header(ws, row=1):
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for cell in ws[row]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[row].height = 24

def build_project_workbook(project, tasks):
    wb = Workbook()
    ws = wb.active
    ws.title = "Projektplan"

    headers = ["WBS","Aktivitet","Ansvarig","Start","Slut","Status","Prioritet","Progress %","Milstolpe","Beroenden","Kommentar"]
    ws.append(headers)
    style_header(ws)

    for t in tasks:
        ws.append([
            t["wbs"], t["title"], t["owner"], as_date(t["start_date"]), as_date(t["end_date"]),
            t["status"], t["priority"], t["progress"], "Ja" if t["milestone"] else "Nej",
            t["dependencies"], t["notes"]
        ])

    for row in range(2, ws.max_row + 1):
        ws.cell(row, 4).number_format = "yyyy-mm-dd"
        ws.cell(row, 5).number_format = "yyyy-mm-dd"
        ws.cell(row, 8).number_format = '0"%"'
        if ws.cell(row, 6).value == "Klar":
            for col in range(1, 12):
                ws.cell(row, col).fill = PatternFill("solid", fgColor="E2F0D9")
        elif ws.cell(row, 6).value == "Blockerad":
            for col in range(1, 12):
                ws.cell(row, col).fill = PatternFill("solid", fgColor="FCE4D6")

    if ws.max_row >= 2:
        ws.conditional_formatting.add(
            f"H2:H{ws.max_row}",
            DataBarRule(start_type="num", start_value=0, end_type="num", end_value=100, showValue=True)
        )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize(ws)

    summary = wb.create_sheet("Sammanfattning")
    summary.append(["Projekt", project["name"]])
    summary.append(["Kund", project["customer"]])
    summary.append(["Projektledare", project["project_manager"]])
    summary.append(["Startdatum", as_date(project["start_date"])])
    summary.append(["Slutdatum", as_date(project["end_date"])])
    summary.append(["Beskrivning", project["description"]])
    summary.append([])
    summary.append(["Nyckeltal", "Värde"])
    style_header(summary, 8)

    total = len(tasks)
    done = sum(1 for t in tasks if t["status"] == "Klar")
    blocked = sum(1 for t in tasks if t["status"] == "Blockerad")
    milestones = sum(1 for t in tasks if t["milestone"])
    avg = round(sum(t["progress"] for t in tasks) / total) if total else 0

    summary.append(["Antal aktiviteter", total])
    summary.append(["Klara aktiviteter", done])
    summary.append(["Blockerade aktiviteter", blocked])
    summary.append(["Milstolpar", milestones])
    summary.append(["Genomsnittlig progress", avg / 100])
    summary["B13"].number_format = "0%"
    summary["B4"].number_format = "yyyy-mm-dd"
    summary["B5"].number_format = "yyyy-mm-dd"
    summary.column_dimensions["A"].width = 28
    summary.column_dimensions["B"].width = 70
    summary["A1"].font = Font(bold=True, size=14)
    summary["B1"].font = Font(bold=True, size=14)

    ms = wb.create_sheet("Milstolpar")
    ms.append(["WBS","Milstolpe","Ansvarig","Datum","Status","Progress %"])
    style_header(ms)
    for t in tasks:
        if t["milestone"]:
            ms.append([t["wbs"], t["title"], t["owner"], as_date(t["end_date"]), t["status"], t["progress"]])
    for row in range(2, ms.max_row + 1):
        ms.cell(row, 4).number_format = "yyyy-mm-dd"
        ms.cell(row, 6).number_format = '0"%"'
    ms.freeze_panes = "A2"
    ms.auto_filter.ref = ms.dimensions
    autosize(ms)

    return wb

@app.route("/projects/<int:project_id>/export.xlsx")
def export_project(project_id):
    project = project_or_404(project_id)
    with db() as conn:
        tasks = conn.execute("SELECT * FROM tasks WHERE project_id=? ORDER BY wbs COLLATE NOCASE, sort_order, id",
                             (project_id,)).fetchall()
    wb = build_project_workbook(project, tasks)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    safe = "".join(c if c.isalnum() or c in "-_" else "_" for c in project["name"]).strip("_") or f"projekt_{project_id}"
    return send_file(output,
                     as_attachment=True,
                     download_name=f"{safe}_projektplan.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/export-all.xlsx")
def export_all():
    with db() as conn:
        projects = conn.execute("SELECT * FROM projects ORDER BY id").fetchall()
        tasks = conn.execute("SELECT t.*, p.name project_name FROM tasks t JOIN projects p ON p.id=t.project_id ORDER BY p.id, t.wbs, t.id").fetchall()

    wb = Workbook()
    ws = wb.active
    ws.title = "Alla aktiviteter"
    ws.append(["Projekt","WBS","Aktivitet","Ansvarig","Start","Slut","Status","Prioritet","Progress %","Milstolpe","Beroenden","Kommentar"])
    style_header(ws)
    for t in tasks:
        ws.append([
            t["project_name"], t["wbs"], t["title"], t["owner"], as_date(t["start_date"]), as_date(t["end_date"]),
            t["status"], t["priority"], t["progress"], "Ja" if t["milestone"] else "Nej", t["dependencies"], t["notes"]
        ])
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 5).number_format = "yyyy-mm-dd"
        ws.cell(row, 6).number_format = "yyyy-mm-dd"
        ws.cell(row, 9).number_format = '0"%"'
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize(ws)

    ps = wb.create_sheet("Projekt")
    ps.append(["Projekt","Kund","Projektledare","Start","Slut","Beskrivning"])
    style_header(ps)
    for p in projects:
        ps.append([p["name"], p["customer"], p["project_manager"], as_date(p["start_date"]), as_date(p["end_date"]), p["description"]])
    for row in range(2, ps.max_row + 1):
        ps.cell(row, 4).number_format = "yyyy-mm-dd"
        ps.cell(row, 5).number_format = "yyyy-mm-dd"
    ps.freeze_panes = "A2"
    ps.auto_filter.ref = ps.dimensions
    autosize(ps)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, as_attachment=True, download_name="alla_projekt.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    init_db()
    app.run(host="0.0.0.0", port=8080, debug=False)
