"""Compile the VB6 feasibility probe in three configurations and run each one.

usage: python -I build_probe.py <probe dir>
"""
import os
import subprocess
import sys

VB6 = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
CONFIGS = [("chk", "0", "2"),      # bounds checks on, no optimization (DoomVB's debug default)
           ("nochk", "-1", "2"),   # bounds checks off, no optimization (DoomVB --release)
           ("fast", "-1", "0")]    # bounds checks off, optimize for fast code

VBP = r"""Type=Exe
Reference=*\G{{00020430-0000-0000-C000-000000000046}}#2.0#0#C:\Windows\SysWOW64\stdole2.tlb#OLE Automation
Module=probe; probe.bas
Startup="Sub Main"
ExeName32="probe_{name}.exe"
Name="probe_{name}"
CompilationType=0
OptimizationType={opt}
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck={bounds}
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
"""


def main(folder):
    os.chdir(folder)
    with open("probe.bas", "rb") as f:
        data = f.read()
    with open("probe.bas", "wb") as f:
        f.write(data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n"))
    for name, bounds, opt in CONFIGS:
        with open("probe_%s.vbp" % name, "w", newline="\r\n") as f:
            f.write(VBP.format(name=name, bounds=bounds, opt=opt))
        subprocess.run([VB6, "/make", "probe_%s.vbp" % name, "/out", "build_%s.txt" % name])
        with open("build_%s.txt" % name, encoding="latin-1") as f:
            result = f.read().strip()
        print(result)
        if "succeeded" not in result:
            return 1
        subprocess.run([os.path.join(folder, "probe_%s.exe" % name), name], timeout=120)
        with open("probe_%s.log" % name, encoding="latin-1") as f:
            print(f.read())
    return 0


if __name__ == "__main__":
    sys.exit(main(os.path.abspath(sys.argv[1])))
