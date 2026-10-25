"""
check_with.py - finds GoTo statements that jump out of a With block on an
array element.

  python -I tools/check_with.py

VB6 locks the array for the With block and unlocks it at End With. A GoTo
to a label outside the block skips the unlock, and the lock stays after the
procedure returns, so the next ReDim of the array fails with run time error
10 ("This array is fixed or temporarily locked"). (Exit Sub / Exit Function
inside the block are unlocked properly.) RecursiveLightPoint's
With mod_surfaces(surf) ... GoTo nexti was such a leak: the next map's
ReDim Preserve of mod_surfaces failed after a demo had been drawn.
This reports every GoTo inside a With <name>(...) block whose label is not
inside the same block.
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

WITH_ARRAY = re.compile(r"^\s*With\s+[\w.]*\w\s*\(", re.I)
WITH_ANY = re.compile(r"^\s*With\b", re.I)
END_WITH = re.compile(r"^\s*End\s+With\b", re.I)
GOTO = re.compile(r"\bGoTo\s+(\w+)", re.I)
LABEL = re.compile(r"^\s*(\w+):(\s|$)")
PROC_END = re.compile(r"^\s*End\s+(Sub|Function)\b", re.I)


def code_part(line):
    instr = False
    for i, ch in enumerate(line):
        if ch == '"':
            instr = not instr
        elif ch == "'" and not instr:
            return line[:i]
    return line


def check(path):
    problems = []
    with open(path, encoding="latin-1") as f:
        lines = [code_part(l.rstrip("\r\n")) for l in f]
    stack = []                           # open With blocks: [start line, is array, gotos, labels]
    for n, line in enumerate(lines, 1):
        if PROC_END.match(line):
            stack = []
            continue
        if WITH_ANY.match(line):
            stack.append([n, bool(WITH_ARRAY.match(line)), [], set()])
            continue
        if END_WITH.match(line):
            if stack:
                start, is_array, gotos, labels = stack.pop()
                if is_array:
                    for gn, label in gotos:
                        if label.lower() not in labels:
                            problems.append("%s:%d: GoTo %s leaves the With block of line %d"
                                            % (os.path.relpath(path, ROOT), gn, label, start))
                # labels and gotos of an inner block also belong to the outer one
                if stack:
                    stack[-1][2].extend(gotos)
                    stack[-1][3] |= labels
            continue
        if stack:
            m = LABEL.match(line)
            if m and m.group(1).lower() not in ("case", "else"):
                stack[-1][3].add(m.group(1).lower())
            for g in GOTO.finditer(line):
                if g.group(1).lower() != "0":
                    stack[-1][2].append((n, g.group(1)))
    return problems


def main():
    problems = []
    for folder in ("src", "stubs"):
        d = os.path.join(ROOT, folder)
        for name in sorted(os.listdir(d)):
            if name.lower().endswith((".bas", ".frm")):
                problems += check(os.path.join(d, name))
    for p in problems:
        print(p)
    print("%d problem(s)" % len(problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
