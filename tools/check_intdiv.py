"""
check_intdiv.py - finds VB integer divisions that a later * or / takes apart.

  python -I tools/check_intdiv.py

In VB, \\ binds looser than * and /: a * b \\ c * d is (a * b) \\ (c * d), where
C's a * b / c * d is ((a * b) / c) * d. So C's left-to-right integer
arithmetic must be written with parentheses: (a * b \\ c) * d. This reports
every \\ that is followed, at the same parenthesis depth and before a + - Mod
comparison And Or or the end of the expression, by a * or /.
(model.bas's texture size, w * h \\ 64 * 85, was such a bug.)
"""
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

STOP = re.compile(r'\s*(\+|-|=|<|>|<>|<=|>=|\bAnd\b|\bOr\b|\bXor\b|\bMod\b|\bThen\b|\bTo\b|\bStep\b|,|:|\))', re.I)


def code_of(line):
    """the line without strings and the comment"""
    out = []
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            j = line.find('"', i + 1)
            while j != -1 and j + 1 < len(line) and line[j + 1] == '"':
                j = line.find('"', j + 2)
            if j == -1:
                break
            out.append('""')
            i = j + 1
            continue
        if ch == "'":
            break
        out.append(ch)
        i += 1
    return ''.join(out)


def check_line(code):
    hits = []
    for m in re.finditer(r'\\', code):
        depth = 0
        i = m.end()
        while i < len(code):
            ch = code[i]
            if ch == '(':
                depth += 1
            elif ch == ')':
                if depth == 0:
                    break
                depth -= 1
            elif depth == 0:
                if ch in '*/':
                    hits.append(m.start())
                    break
                if ch == '\\':
                    break
                if STOP.match(code, i) and code[i] not in ' ':
                    break
            i += 1
    return hits


def main():
    problems = 0
    for sub in ('src', 'stubs'):
        d = os.path.join(ROOT, sub)
        for name in sorted(os.listdir(d)):
            if not name.lower().endswith('.bas'):
                continue
            with open(os.path.join(d, name), encoding='latin-1') as f:
                for no, line in enumerate(f, 1):
                    code = code_of(line.rstrip('\r\n'))
                    if check_line(code):
                        print('%s\\%s:%d: %s' % (sub, name, no, line.strip()))
                        problems += 1
    print('%d problem(s)' % problems)
    return 1 if problems else 0


if __name__ == '__main__':
    sys.exit(main())
