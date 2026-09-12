from pathlib import Path
import ast, re, sqlite3, sys
from jinja2 import Environment, FileSystemLoader

root=Path(__file__).resolve().parent
app_py=root/"app"/"app.py"
templates=root/"app"/"templates"
source=app_py.read_text(encoding="utf-8")
tree=ast.parse(source)
errors=[]; warnings=[]

# 1. Route inventory, duplicate endpoint/path detection and url_for validation.
routes={}; paths={}
for node in tree.body:
    if not isinstance(node,(ast.FunctionDef,ast.AsyncFunctionDef)): continue
    required=set(); route_paths=[]
    for d in node.decorator_list:
        if isinstance(d,ast.Call) and isinstance(d.func,ast.Attribute) and d.func.attr in {"route","get","post","put","patch","delete"}:
            if d.args and isinstance(d.args[0],ast.Constant) and isinstance(d.args[0].value,str):
                path=d.args[0].value; route_paths.append(path)
                required.update(re.findall(r"<(?:[^:>]+:)?([^>]+)>",path))
                paths.setdefault(path,[]).append(node.name)
    if route_paths: routes[node.name]={"required":required,"paths":route_paths}

for path,names in paths.items():
    if len(set(names))>1: warnings.append(f"duplicate route path {path}: {', '.join(names)}")

url_for_re=re.compile(r"url_for\(\s*['\"]([^'\"]+)['\"]([^)]*)\)")
env=Environment(loader=FileSystemLoader(str(templates)))
template_count=0
for tpl in templates.rglob("*.html"):
    template_count+=1
    text=tpl.read_text(encoding="utf-8")
    try: env.get_template(str(tpl.relative_to(templates)))
    except Exception as ex: errors.append(f"{tpl.relative_to(root)}: Jinja error: {ex}")
    for m in url_for_re.finditer(text):
        ep,args=m.group(1),m.group(2)
        if ep=="static": continue
        if ep not in routes:
            errors.append(f"{tpl.relative_to(root)}: unknown endpoint {ep}"); continue
        missing=[x for x in routes[ep]["required"] if re.search(rf"\b{re.escape(x)}\s*=",args) is None]
        if missing: errors.append(f"{tpl.relative_to(root)}: url_for('{ep}') missing {', '.join(sorted(missing))}")

# 2. Build database schema from init_db executescript plus static CREATE TABLE calls.
db=sqlite3.connect(":memory:")
db.row_factory=sqlite3.Row
def exec_schema(sql,where):
    try: db.executescript(sql)
    except Exception as ex: errors.append(f"{where}: schema error: {ex}")

for node in ast.walk(tree):
    if isinstance(node,ast.Call) and isinstance(node.func,ast.Attribute) and node.func.attr=="executescript" and node.args:
        a=node.args[0]
        if isinstance(a,ast.Constant) and isinstance(a.value,str):
            exec_schema(a.value,f"app.py:{getattr(node,'lineno',0)}")
    if isinstance(node,ast.Call) and isinstance(node.func,ast.Attribute) and node.func.attr=="execute" and node.args:
        a=node.args[0]
        if isinstance(a,ast.Constant) and isinstance(a.value,str) and a.value.strip().upper().startswith("CREATE TABLE"):
            try: db.execute(a.value)
            except Exception as ex: errors.append(f"app.py:{getattr(node,'lineno',0)} CREATE error: {ex}")

# Apply ensure_column migrations that use literal args.
for node in ast.walk(tree):
    if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and node.func.id=="ensure_column" and len(node.args)>=4:
        vals=[a.value if isinstance(a,ast.Constant) else None for a in node.args[1:4]]
        table,column,definition=vals
        if table and column and definition:
            try: db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
            except sqlite3.OperationalError as ex:
                if "duplicate column" not in str(ex).lower(): errors.append(f"migration {table}.{column}: {ex}")

# 3. Compile static SQL and validate placeholder count for literal tuple/list parameters.
skip=("PRAGMA","CREATE","ALTER","VACUUM")
sql_checks=0
for node in ast.walk(tree):
    if not (isinstance(node,ast.Call) and isinstance(node.func,ast.Attribute) and node.func.attr in {"execute","executemany"} and node.args): continue
    a=node.args[0]
    if not (isinstance(a,ast.Constant) and isinstance(a.value,str)): continue
    sql=a.value.strip()
    if not sql or sql.upper().startswith(skip): continue
    sql_checks+=1
    qcount=sql.count("?")
    try: db.execute("EXPLAIN "+sql,[None]*qcount)
    except Exception as ex: errors.append(f"app.py:{getattr(node,'lineno',0)} SQL error: {ex} :: {' '.join(sql.split())[:170]}")
    if len(node.args)>=2 and isinstance(node.args[1],(ast.Tuple,ast.List)):
        argc=len(node.args[1].elts)
        if qcount!=argc:
            errors.append(f"app.py:{getattr(node,'lineno',0)} placeholder mismatch: {qcount} placeholders vs {argc} literal args")

# 4. Quality/security regression checks for known high-risk areas.
checks={
    "write permission quick add": 'project_or_404(project_id,write=True)' in source,
    "task experience scoped": 't.project_id IN ({marks})' in source,
    "wizard role guard": '@role_required("admin","pm")\ndef project_wizard_v58' in source,
    "liveness endpoint": '@app.get("/health/live")' in source,
    "500 handler": '@app.errorhandler(500)' in source,
    "WAL/busy timeout": 'PRAGMA busy_timeout=5000' in source and 'PRAGMA journal_mode=WAL' in source,
}
for name,ok in checks.items():
    if not ok: errors.append(f"quality regression: {name}")

base=(templates/"base.html").read_text(encoding="utf-8")
if 'name="viewport"' not in base: errors.append("base.html missing mobile viewport meta")
css=(root/"app"/"static"/"style.css").read_text(encoding="utf-8")
if "@media" not in css: errors.append("style.css has no responsive media queries")

if errors:
    print("VERIFY FAILED")
    for e in errors: print(" -",e)
    if warnings:
        print("WARNINGS")
        for w in warnings: print(" -",w)
    sys.exit(1)
print(f"VERIFY PASSED: {len(routes)} endpoints, {template_count} templates, {sql_checks} static SQL statements.")
if warnings:
    print("WARNINGS")
    for w in warnings: print(" -",w)
