"""
vbtest.py - builds and runs the VB unit tests (tests/*.bas).

  python -I tools/vbtest.py

Copies src/ and stubs/ to build_test/, renames the engine's Sub Main so the
one in tests/vbtests.bas starts instead, compiles, runs the exe and prints
build_test/vbtests.log. Regenerate tests/printf_cases.bas with
tools/gen_printf_tests.py.
"""
import os
import shutil
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
import build  # noqa: E402


def main():
    outdir = os.path.join(ROOT, "build_test")
    if os.path.exists(outdir):
        shutil.rmtree(outdir)
    for sub in ("src", "stubs"):
        shutil.copytree(os.path.join(ROOT, sub), os.path.join(outdir, sub))
    for name in os.listdir(os.path.join(ROOT, "tests")):
        if name.lower().endswith(".bas"):
            shutil.copy(os.path.join(ROOT, "tests", name), os.path.join(outdir, "src", name))
    path = os.path.join(outdir, "src", "sys_win.bas")
    with open(path, encoding="latin-1", newline="") as f:
        s = f.read()
    if s.count("Public Sub Main()") != 1:
        print("sys_win.bas: Sub Main not found")
        return 1
    with open(path, "w", encoding="latin-1", newline="") as f:
        f.write(s.replace("Public Sub Main()", "Public Sub Main_Engine()"))
    if build.main(outdir) != 0:
        return 1
    log = os.path.join(outdir, "vbtests.log")
    subprocess.run([os.path.join(outdir, "QuakeVB.exe")], cwd=outdir, timeout=120)
    with open(log, encoding="latin-1") as f:
        text = f.read()
    print(text.strip())
    return 0 if " 0 failed" in text else 1


if __name__ == "__main__":
    sys.exit(main())
