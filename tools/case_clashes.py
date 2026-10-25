"""List C identifiers in the ported WinQuake files that differ only in case.

usage: python -I case_clashes.py <ref/WinQuake>
"""
import collections
import glob
import os
import re
import sys

SKIP = ("gl_", "glquake", "vid_dos", "vid_x", "vid_sun", "vid_svga", "vid_vga", "vid_ext", "sys_dos",
        "sys_linux", "sys_sun", "snd_dos", "snd_gus", "snd_linux", "snd_sun", "snd_next", "in_dos", "in_sun",
        "cd_linux", "net_dos", "net_bsd", "net_bw", "net_comx", "net_ipx", "net_mp", "net_ser", "net_udp",
        "mplib", "mplpc", "dos_v2", "net_wipx", "net_wso", "vregset", "sys_null", "vid_null", "snd_null",
        "in_null", "cd_null", "net_none")
STRIP = re.compile(r'//[^\n]*|/\*.*?\*/|"(?:\\.|[^"\\])*"', re.S)


def main(folder):
    groups = collections.defaultdict(set)
    for path in glob.glob(os.path.join(folder, "*.[ch]")):
        if os.path.basename(path).startswith(SKIP):
            continue
        with open(path, encoding="latin-1") as f:
            src = STRIP.sub(" ", f.read())
        for word in re.findall(r"\b[A-Za-z_]\w*\b", src):
            groups[word.lower()].add(word)
    clashes = sorted(sorted(v) for v in groups.values() if len(v) > 1)
    print(len(clashes), "case-insensitive clashes")
    for v in clashes:
        print("  " + " / ".join(v))


if __name__ == "__main__":
    main(sys.argv[1])
