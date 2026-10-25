"""
build.py - builds QuakeVB.exe with the Visual Basic 6 command line compiler.

  python -I tools/build.py [--release] [project folder]

* writes QuakeVB.vbp from the modules in src/, plus the stub of every cluster
  that is not ported yet (stubs/)
* normalises every source file to CRLF line endings (VB6 requires them)
* compiles res/quakevb.rc -> res/quakevb.res with the VB6 resource compiler,
  if there is one
* runs VB6.EXE /make and prints the compile log (VB6 reports only the first
  error, with the file and line)

--release removes the array bounds checks.
"""
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VB6 = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
RC = r"C:\Program Files (x86)\Microsoft Visual Studio\VB98\Wizards\RC.EXE"

# cluster -> (stub module, the src/ files that replace it)
CLUSTERS = {
    "A": ("stubs_A", ["model", "pr_edict", "pr_exec", "pr_cmds", "sv_main", "sv_phys", "sv_move", "sv_user",
                      "world"]),
    "B": ("stubs_B", ["net_main", "net_loop", "cl_main", "cl_parse", "cl_input", "cl_demo", "cl_tent", "chase",
                      "view", "in_win"]),
    "C": ("stubs_C", ["r_main", "r_bsp", "r_edge", "r_surf", "r_draw", "r_light", "r_efrag", "r_misc", "r_sky",
                      "r_vars", "d_init", "d_modech", "d_edge", "d_scan", "d_surf", "d_sky", "d_vars", "d_fill",
                      "d_zpoint", "nonintel"]),
    "D": ("stubs_D", ["r_alias", "r_aclip", "d_polyse", "r_sprite", "d_sprite", "r_part", "d_part"]),
    "E": ("stubs_E", ["sbar", "menu"]),
    "F": ("stubs_F", ["snd_dma", "snd_mem", "snd_mix", "snd_win", "cd_win"]),
}

VBP_TAIL = """{resline}Startup="Sub Main"
HelpFile=""
Title="QuakeVB"
ExeName32="QuakeVB.exe"
Command32=""
Name="QuakeVB"
HelpContextID="0"
CompatibleMode="0"
MajorVer=0
MinorVer=1
RevisionVer=0
AutoIncrementVer=0
ServerSupportFiles=0
VersionComments="Visual Basic 6 port of the Quake source release (GNU GPL v2)"
VersionCompanyName="QuakeVB"
VersionFileDescription="QuakeVB"
VersionProductName="QuakeVB"
CompilationType=0
OptimizationType=2
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=0
BoundsCheck={bounds}
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
"""


def normalise(folder):
    for name in os.listdir(folder):
        if not name.lower().endswith((".bas", ".frm", ".cls", ".vbp", ".ctl")):
            continue
        path = os.path.join(folder, name)
        with open(path, "rb") as f:
            data = f.read()
        fixed = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
        if fixed != data:
            with open(path, "wb") as f:
                f.write(fixed)


def module_name(path):
    with open(path, "r", encoding="latin-1") as f:
        for line in f:
            if line.startswith("Attribute VB_Name"):
                return line.split('"')[1]
    return os.path.splitext(os.path.basename(path))[0]


def write_vbp(root, vbp, release, has_res):
    src = os.path.join(root, "src")
    present = {os.path.splitext(n)[0].lower() for n in os.listdir(src)}
    lines = ["Type=Exe",
             "Reference=*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#"
             "C:\\Windows\\SysWOW64\\stdole2.tlb#OLE Automation"]
    for n in sorted(os.listdir(src)):
        if n.lower().endswith(".frm"):
            lines.append("Form=src\\" + n)
    for n in sorted(os.listdir(src)):
        if n.lower().endswith(".bas"):
            lines.append("Module=%s; src\\%s" % (module_name(os.path.join(src, n)), n))
    for key, (stub, mods) in sorted(CLUSTERS.items()):
        have = [m for m in mods if m in present]
        if have and len(have) != len(mods):
            print("WARNING: cluster %s is incomplete: missing %s" %
                  (key, ", ".join(m for m in mods if m not in present)))
        if not have:
            lines.append("Module=%s; stubs\\%s.bas" % (stub, stub))
    resline = 'ResFile32="res\\quakevb.res"\n' if has_res else ""
    lines.append(VBP_TAIL.format(bounds="-1" if release else "0", resline=resline))
    with open(vbp, "w", newline="\r\n") as f:
        f.write("\n".join(lines))


def main(root=None):
    release = "--release" in sys.argv
    if root is None:
        args = [a for a in sys.argv[1:] if not a.startswith("--")]
        root = os.path.abspath(args[0]) if args else ROOT
    vbp = os.path.join(root, "QuakeVB.vbp")

    resdir = os.path.join(root, "res")
    rcfile = os.path.join(resdir, "quakevb.rc")
    has_res = os.path.exists(rcfile)
    if has_res:
        r = subprocess.run([RC, "/r", "/fo", os.path.join(resdir, "quakevb.res"), rcfile],
                           cwd=resdir, capture_output=True, text=True)
        if r.returncode != 0:
            print("RC failed:", r.stdout, r.stderr)
            return 1

    write_vbp(root, vbp, release, has_res)
    for sub in ("src", "stubs"):
        normalise(os.path.join(root, sub))
    normalise(root)

    log = os.path.join(root, "build.log")
    if os.path.exists(log):
        os.remove(log)
    subprocess.run([VB6, "/make", vbp, "/out", log], cwd=root)
    with open(log, "r", encoding="latin-1") as f:
        text = f.read().strip()
    print(text)
    return 0 if "succeeded" in text else 1


if __name__ == "__main__":
    sys.exit(main())
