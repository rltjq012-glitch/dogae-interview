"""dogae/ 모듈 이름 검사 — 고친 뒤 `python tools/check_modules.py` 로 실행하세요.

나눈 모듈마다: 쓰는 이름이 모두 (자기 모듈 · import · 아래 층 모듈 · 내장 이름)에서 오는지 확인."""
import ast, sys, os, symtable, builtins, json
OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LAYERS = ["base", "ai", "storage", "exam_data", "questions", "word_docs", "screens"]
exports, problems = {}, {}
B = set(dir(builtins)) | {"__file__", "__name__", "__all__"}
def names_used(src, fn):
    top = symtable.symtable(src, fn, "exec")
    used = set()
    def walk(t):
        for s in t.get_symbols():
            if t.get_type() == "module":
                if s.is_referenced():
                    used.add(s.get_name())
            elif s.is_global() or (s.is_free() and False):
                used.add(s.get_name())
        for c in t.get_children():
            walk(c)
    walk(top)
    return used
for i, L in enumerate(LAYERS):
    fn = os.path.join(OUT, "dogae", f"{L}.py")
    src = open(fn, encoding="utf-8").read()
    tree = ast.parse(src)
    own = set()
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)): own.add(n.name)
        elif isinstance(n, (ast.Assign, ast.AnnAssign)):
            for t in (n.targets if isinstance(n, ast.Assign) else [n.target]):
                own |= {x.id for x in ast.walk(t) if isinstance(x, ast.Name)}
        elif isinstance(n, (ast.Import, ast.ImportFrom)):
            own |= {(a.asname or a.name).split(".")[0] for a in n.names if a.name != "*"}
    allv = next(ast.literal_eval(n.value) for n in tree.body if isinstance(n, ast.Assign) and getattr(n.targets[0], "id", "") == "__all__")
    missing_all = sorted(x for x in own if x not in allv and not x.startswith("__") and x not in
                         {a for n in tree.body if isinstance(n, (ast.Import, ast.ImportFrom)) for a in [(z.asname or z.name).split(".")[0] for z in n.names]})
    avail = own | B | {x for M in LAYERS[:i] for x in exports[M]}
    comp_targets = {x.id for n in ast.walk(tree) if isinstance(n, ast.comprehension)
                    for x in ast.walk(n.target) if isinstance(x, ast.Name)}
    bad = sorted(names_used(src, fn) - avail - comp_targets)
    exports[L] = set(allv)
    if bad or missing_all:
        problems[L] = {"unresolved": bad, "not_exported": missing_all}
# app.py: 화면 흐름이 쓰는 이름
app = open(os.path.join(OUT, "app.py"), encoding="utf-8").read()
atree = ast.parse(app)
aown = set()
for n in ast.walk(atree):
    if isinstance(n, (ast.FunctionDef, ast.ClassDef)): aown.add(n.name)
    elif isinstance(n, ast.Name) and isinstance(n.ctx, ast.Store): aown.add(n.id)
    elif isinstance(n, (ast.Import, ast.ImportFrom)):
        aown |= {(a.asname or a.name).split(".")[0] for a in n.names if a.name != "*"}
    elif isinstance(n, ast.arg): aown.add(n.arg)
    elif isinstance(n, (ast.ExceptHandler,)) and n.name: aown.add(n.name)
allx = set().union(*exports.values())
abad = sorted(names_used(app, "app.py") - aown - allx - B)
if abad: problems["app.py"] = {"unresolved": abad}
# 같은 이름을 두 모듈이 내보내는지
seen = {}
for L in LAYERS:
    for x in exports[L]:
        seen.setdefault(x, []).append(L)
dup = {k: v for k, v in seen.items() if len(v) > 1}
print(json.dumps({"problems": problems, "duplicate_exports": dup}, ensure_ascii=False, indent=1))
sys.exit(1 if problems or dup else 0)
