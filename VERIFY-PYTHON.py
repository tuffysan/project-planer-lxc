from pathlib import Path
import ast, re, sys

root = Path(__file__).resolve().parent
app_py = root / "app" / "app.py"
templates = root / "app" / "templates"

src = app_py.read_text(encoding="utf-8")
tree = ast.parse(src)

routes = {}
pending = []
for node in tree.body:
    if isinstance(node, ast.FunctionDef):
        route_patterns = []
        methods = None
        for d in node.decorator_list:
            if isinstance(d, ast.Call) and isinstance(d.func, ast.Attribute):
                if d.func.attr in {"route","get","post","put","patch","delete"}:
                    if d.args and isinstance(d.args[0], ast.Constant) and isinstance(d.args[0].value, str):
                        route_patterns.append(d.args[0].value)
        if route_patterns:
            required = set()
            for r in route_patterns:
                required.update(re.findall(r"<(?:[^:>]+:)?([^>]+)>", r))
            routes[node.name] = required

errors = []
for tpl in templates.rglob("*.html"):
    text = tpl.read_text(encoding="utf-8")
    for m in re.finditer(r"url_for\(\s*['\"]([^'\"]+)['\"]([^)]*)\)", text):
        endpoint = m.group(1)
        args = m.group(2)
        if endpoint not in routes:
            continue
        required = routes[endpoint]
        missing = [p for p in required if re.search(rf"\b{re.escape(p)}\s*=", args) is None]
        if missing:
            errors.append(f"{tpl.relative_to(root)}: url_for('{endpoint}') missing {', '.join(sorted(missing))}")

if errors:
    print("Template route verification FAILED:")
    for e in errors:
        print(" -", e)
    sys.exit(1)

print(f"Template route verification passed ({len(routes)} routed endpoints checked).")
