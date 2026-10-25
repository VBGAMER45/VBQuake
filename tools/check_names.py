"""
check_names.py - finds names that clash in VB6 (it ignores case) before the
compiler does, one at a time.

  python -I tools/check_names.py

Reports, case-insensitively:
* a module-level name (Sub, Function, variable, Const, Type, Declare) defined
  twice in one module, Private or Public
* a Public name defined in two modules
* a module name equal to a Public name
* a Public name of another module that a module also defines Private
  (legal in VB, but a likely mistake when porting)
* a parameter or local variable spelled like a Public name or a name of its
  own module in another case: C's yaw and YAW are two names, VB's one, and
  inside the procedure the local wins
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DECL = re.compile(
    r"^\s*(?:(Public|Private|Global|Dim|Friend)\s+)?"
    r"(?:(Const|Declare\s+(?:Sub|Function)|Sub|Function|Property\s+(?:Get|Let|Set)|Type|Enum)\s+)?"
    r"([A-Za-z_]\w*)", re.I)
END_PROC = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)
START_PROC = re.compile(r"^\s*(?:(?:Public|Private|Friend|Static)\s+)*(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)", re.I)
TYPE_START = re.compile(r"^\s*(?:(?:Public|Private)\s+)?(Type|Enum)\s+([A-Za-z_]\w*)", re.I)
TYPE_END = re.compile(r"^\s*End\s+(Type|Enum)\b", re.I)
KEYWORDS = {"option", "attribute", "end", "rem", "#if", "#end"}


def module_files():
    files = []
    for sub in ("src", "stubs"):
        d = os.path.join(ROOT, sub)
        if os.path.isdir(d):
            files += [os.path.join(d, n) for n in sorted(os.listdir(d)) if n.lower().endswith((".bas", ".frm"))]
    return files


def split_decl_list(rest):
    """names in 'a As Long, b() As Byte, c As X' (after the first)"""
    names = []
    depth = 0
    part = ""
    for ch in rest:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            names.append(part)
            part = ""
        else:
            part += ch
    names.append(part)
    out = []
    for p in names[1:]:
        m = re.match(r"\s*([A-Za-z_]\w*)", p)
        if m:
            out.append(m.group(1))
    return out


PARAM_WORDS = {"byval", "byref", "optional", "paramarray"}


def param_names(code, start):
    i = code.find("(", start)
    if i < 0:
        return []
    depth = 0
    j = i
    for j in range(i, len(code)):
        if code[j] == "(":
            depth += 1
        elif code[j] == ")":
            depth -= 1
            if depth == 0:
                break
    out = []
    for part in code[i + 1:j].split(","):
        words = [w for w in re.findall(r"[A-Za-z_]\w*", part.split(" As ")[0]) if w.lower() not in PARAM_WORDS]
        if words:
            out.append(words[0])
    return out


def scan(path):
    """module name, its module-level names [(name, public, kind, line)] and
    the parameters and locals of its procedures [(name, procedure, line)]"""
    modname = os.path.splitext(os.path.basename(path))[0]
    names = []
    locals_ = []
    proc = None
    in_proc = False
    in_type = False
    with open(path, encoding="latin-1") as f:
        lines = f.read().split("\n")
    for ln in lines:
        if ln.startswith("Attribute VB_Name"):
            modname = ln.split('"')[1]
            break
    is_form = path.lower().endswith(".frm")
    joined = []
    buf = ""
    for ln in lines:
        ln = ln.rstrip("\r")
        if buf:
            ln = buf + " " + ln.strip()
            buf = ""
        if ln.rstrip().endswith(" _"):
            buf = ln.rstrip()[:-2]
            continue
        joined.append(ln)
    for no, ln in enumerate(joined, 1):
        code = re.sub(r"\"[^\"]*\"", '""', ln)
        code = code.split("'")[0]
        if code.startswith("Attribute VB_Name"):
            modname = ln.split('"')[1]
            continue
        if in_type:
            if TYPE_END.match(code):
                in_type = False
            continue
        if in_proc:
            if END_PROC.match(code):
                in_proc = False
                continue
            m = re.match(r"^\s*(Dim|Static)\s+([A-Za-z_]\w*)", code, re.I)
            if m:
                locals_.append((m.group(2), proc, no))
                for extra in split_decl_list(code[m.end():]):
                    locals_.append((extra, proc, no))
            continue
        m = TYPE_START.match(code)
        if m:
            public = not code.strip().lower().startswith("private")
            names.append((m.group(2), public, m.group(1).lower(), no))
            in_type = True
            continue
        m = START_PROC.match(code)
        if m:
            public = not re.match(r"^\s*Private\b", code, re.I) and not is_form
            names.append((m.group(2), public, "proc", no))
            in_proc = True
            proc = m.group(2)
            for pn in param_names(code, m.end()):
                locals_.append((pn, proc, no))
            continue
        m = DECL.match(code)
        if not m or not m.group(1):
            continue
        scope, kind, name = m.group(1).lower(), (m.group(2) or "var").split()[0].lower(), m.group(3)
        if name.lower() in KEYWORDS:
            continue
        public = scope in ("public", "global") and not is_form
        names.append((name, public, kind, no))
        if kind in ("var", "const"):
            rest = code[m.end():]
            for extra in split_decl_list(rest):
                names.append((extra, public, kind, no))
    return modname, names, locals_


def main():
    problems = 0
    publics = {}
    modules = {}
    scanned = []
    for path in module_files():
        modname, names, locals_ = scan(path)
        rel = os.path.relpath(path, ROOT)
        modules[modname.lower()] = rel
        scanned.append((rel, modname, names, locals_))
        seen = {}
        for name, public, kind, no in names:
            key = name.lower()
            if key in seen:
                print("%s:%d: %s clashes with %s (line %d) in the same module" % (rel, no, name, seen[key][0], seen[key][1]))
                problems += 1
            else:
                seen[key] = (name, no)
            if public:
                if key in publics and publics[key][0] != rel:
                    print("%s:%d: Public %s is also Public in %s:%d" % (rel, no, name, publics[key][0], publics[key][1]))
                    problems += 1
                else:
                    publics[key] = (rel, no, name)
    for key, (rel, no, name) in publics.items():
        if key in modules:
            print("%s:%d: Public %s has the name of module %s" % (rel, no, name, modules[key]))
            problems += 1
    for rel, modname, names, locals_ in scanned:
        for name, public, kind, no in names:
            if not public and name.lower() in publics and publics[name.lower()][0] != rel:
                p = publics[name.lower()]
                print("%s:%d: Private %s hides Public %s of %s:%d" % (rel, no, name, p[2], p[0], p[1]))
                problems += 1
        own = {}
        for name, public, kind, no in names:
            own[name.lower()] = (rel, no, name)
        for name, proc, no in locals_:
            key = name.lower()
            other = own.get(key) or publics.get(key)
            if other and other[2] != name:
                print("%s:%d: %s in %s is spelled %s in %s:%d" % (rel, no, name, proc, other[2], other[0], other[1]))
                problems += 1
    print("%d problem(s)" % problems)
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
