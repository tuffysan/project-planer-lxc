from pathlib import Path
import ast, re, sqlite3, sys
from jinja2 import Environment, FileSystemLoader

root = Path(__file__).resolve().parent
app_py = root / "app" / "app.py"
templates = root / "app" / "templates"
source = app_py.read_text(encoding="utf-8")
tree = ast.parse(source)
errors=[]

# Routes and url_for validation
routes={}
for node in tree.body:
    if isinstance(node,(ast.FunctionDef,ast.AsyncFunctionDef)):
        required=set()
        routed=False
        for d in node.decorator_list:
            if isinstance(d,ast.Call) and isinstance(d.func,ast.Attribute) and d.func.attr in {"route","get","post","put","patch","delete"}:
                if d.args and isinstance(d.args[0],ast.Constant) and isinstance(d.args[0].value,str):
                    routed=True
                    required.update(re.findall(r"<(?:[^:>]+:)?([^>]+)>",d.args[0].value))
        if routed: routes[node.name]=required

for tpl in templates.rglob("*.html"):
    text=tpl.read_text(encoding="utf-8")
    for m in re.finditer(r"url_for\(\s*['\"]([^'\"]+)['\"]([^)]*)\)",text):
        ep,args=m.group(1),m.group(2)
        if ep=="static": continue
        if ep not in routes:
            errors.append(f"{tpl.relative_to(root)}: unknown endpoint {ep}")
            continue
        missing=[x for x in routes[ep] if re.search(rf"\b{re.escape(x)}\s*=",args) is None]
        if missing:
            errors.append(f"{tpl.relative_to(root)}: url_for('{ep}') missing {', '.join(sorted(missing))}")

# Jinja syntax/extends/includes
env=Environment(loader=FileSystemLoader(str(templates)))
for tpl in templates.rglob("*.html"):
    try: env.get_template(str(tpl.relative_to(templates)))
    except Exception as ex: errors.append(f"{tpl.relative_to(root)}: Jinja error: {ex}")

# Build an in-memory SQLite schema from CREATE TABLE + ensure_column migrations
db=sqlite3.connect(":memory:")
for m in re.finditer(r"CREATE TABLE IF NOT EXISTS\s+\w+\s*\(.*?\);",source,re.S|re.I):
    try: db.execute(m.group(0))
    except Exception as ex: errors.append(f"Schema create error: {ex}")
for node in ast.walk(tree):
    if isinstance(node,ast.Call) and isinstance(node.func,ast.Name) and node.func.id=="ensure_column" and len(node.args)>=4:
        vals=[]
        for a in node.args[1:4]:
            vals.append(a.value if isinstance(a,ast.Constant) else None)
        table,column,definition=vals
        if table and column and definition:
            try: db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {definition}")
            except sqlite3.OperationalError as ex:
                if "duplicate column" not in str(ex).lower():
                    errors.append(f"Migration {table}.{column}: {ex}")

# Compile all static SQL used by conn.execute/executemany
for node in ast.walk(tree):
    if not (isinstance(node,ast.Call) and isinstance(node.func,ast.Attribute) and node.func.attr in {"execute","executemany"} and node.args):
        continue
    a=node.args[0]
    if not (isinstance(a,ast.Constant) and isinstance(a.value,str)): continue
    sql=a.value.strip()
    if not sql or sql.upper().startswith(("PRAGMA","CREATE","ALTER")): continue
    try: db.execute("EXPLAIN "+sql,[None]*sql.count("?"))
    except Exception as ex:
        errors.append(f"app.py:{getattr(node,'lineno',0)} SQL error: {ex} :: {' '.join(sql.split())[:160]}")

if errors:
    print("VERIFY FAILED")
    for e in errors: print(" -",e)
    sys.exit(1)
print(f"VERIFY PASSED: {len(routes)} routes, {len(list(templates.rglob('*.html')))} templates, static SQL and url_for checks.")
