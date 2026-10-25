"""Check a Quake game dir against the limits of the 1996 WinQuake engine.

usage: python -I inspect_paks.py <dir containing pak*.pak>
"""
import collections
import glob
import os
import struct
import sys

# WinQuake limits (quakedef.h, bspfile.h, progs.h, protocol.h)
MAX_MAP_LEAFS = 8192
MAX_MAP_CLIPNODES = 32767
MAX_MAP_PLANES = 32767
MAX_MAP_VERTS = 65535
MAX_MAP_FACES = 65535
MAX_MAP_MODELS = 256
BSP_LUMPS = ["entities", "planes", "textures", "vertexes", "visibility", "nodes", "texinfo", "faces",
             "lighting", "clipnodes", "leafs", "marksurfaces", "edges", "surfedges", "models"]


def read_pak(path):
    with open(path, "rb") as f:
        data = f.read()
    ident, ofs, size = struct.unpack_from("<4sii", data, 0)
    assert ident == b"PACK", path
    files = {}
    for i in range(size // 64):
        name, fpos, flen = struct.unpack_from("<56sii", data, ofs + i * 64)
        files[name.split(b"\0")[0].decode("latin-1")] = data[fpos:fpos + flen]
    return files


def bsp_info(name, b):
    version = struct.unpack_from("<i", b, 0)[0]
    if version != 29:
        return {"name": name, "version": b[:4]}
    lumps = {}
    for i, ln in enumerate(BSP_LUMPS):
        ofs, ln_len = struct.unpack_from("<ii", b, 4 + i * 8)
        lumps[ln] = (ofs, ln_len)
    sizes = {"planes": 20, "vertexes": 12, "nodes": 24, "texinfo": 40, "faces": 20, "clipnodes": 8,
             "leafs": 28, "marksurfaces": 2, "edges": 4, "surfedges": 4, "models": 64}
    counts = {k: lumps[k][1] // v for k, v in sizes.items()}
    ents = b[lumps["entities"][0]:lumps["entities"][0] + lumps["entities"][1]]
    nent = ents.count(b"{")
    # world bounds from model 0
    mo = lumps["models"][0]
    mins = struct.unpack_from("<3f", b, mo)
    maxs = struct.unpack_from("<3f", b, mo + 12)
    return {"name": name, "version": 29, "counts": counts, "entities": nent, "mins": mins, "maxs": maxs,
            "lighting": lumps["lighting"][1]}


def progs_info(b):
    hdr = struct.unpack_from("<ii" + "ii" * 6 + "i", b, 0)
    version, crc = hdr[0], hdr[1]
    ofs_statements, numstatements = hdr[2], hdr[3]
    ofs_functions, numfunctions = hdr[8], hdr[9]
    ofs_strings, numstrings = hdr[10], hdr[11]
    entityfields = hdr[14]
    opcodes = collections.Counter()
    for i in range(numstatements):
        op = struct.unpack_from("<H", b, ofs_statements + i * 8)[0]
        opcodes[op] += 1
    builtins = set()
    strings = b[ofs_strings:ofs_strings + numstrings]
    for i in range(numfunctions):
        first_statement, _, _, s_name = struct.unpack_from("<iiii", b, ofs_functions + i * 36)
        if first_statement < 0:
            nm = strings[s_name:strings.index(b"\0", s_name)].decode("latin-1")
            builtins.add((-first_statement, nm))
    return version, crc, entityfields, opcodes, sorted(builtins)


def main(folder):
    files = {}
    for pak in sorted(glob.glob(os.path.join(folder, "pak*.pak"))):
        pf = read_pak(pak)
        print(f"{os.path.basename(pak)}: {len(pf)} files, {os.path.getsize(pak) / 1e6:.1f} MB")
        files.update(pf)
    exts = collections.Counter(os.path.splitext(n)[1].lower() for n in files)
    print("file types:", dict(exts))
    for n in sorted(files):
        if not n.startswith(("maps/", "sound/", "progs/", "gfx/")) or n.endswith((".bsp", ".mdl", ".wav", ".spr")):
            continue
        print("  other:", n)

    print("\n== progs.dat")
    version, crc, fields, opcodes, builtins = progs_info(files["progs.dat"])
    print(f"version {version}  crc {crc} (WinQuake expects 6 / 5927)  entityfields {fields}")
    print(f"opcodes used: max {max(opcodes)} (vanilla 0..65)", "EXTENDED!" if max(opcodes) > 65 else "ok")
    print(f"builtins: {len(builtins)}, max #{max(b[0] for b in builtins)} (WinQuake has 1..78)")
    for num, nm in builtins:
        if num > 78:
            print(f"   extension builtin #{num} {nm}")

    print("\n== maps")
    over = 0
    for n in sorted(files):
        if not n.endswith(".bsp"):
            continue
        info = bsp_info(n, files[n])
        if info["version"] != 29:
            print(f"  {n}: NOT BSP29 ({info['version']})")
            over += 1
            continue
        c = info["counts"]
        problems = []
        if c["leafs"] > MAX_MAP_LEAFS:
            problems.append(f"leafs {c['leafs']}")
        if c["clipnodes"] > MAX_MAP_CLIPNODES:
            problems.append(f"clipnodes {c['clipnodes']}")
        if c["models"] > MAX_MAP_MODELS:
            problems.append(f"models {c['models']}")
        if info["entities"] > 600:
            problems.append(f"entities {info['entities']} (>600 edicts)")
        ext = max(max(abs(v) for v in info["mins"]), max(abs(v) for v in info["maxs"]))
        if ext > 4096:
            problems.append(f"extent {ext:.0f} (>4096 protocol coords)")
        if problems:
            over += 1
        if n.count("/") == 1:
            print(f"  {n:28s} ents {info['entities']:4d} leafs {c['leafs']:5d} faces {c['faces']:5d} "
                  f"clip {c['clipnodes']:5d} {'; '.join(problems)}")
    print(f"maps over vanilla limits: {over}")

    print("\n== sounds")
    rates = collections.Counter()
    for n, b in files.items():
        if n.endswith(".wav"):
            if b[:4] != b"RIFF":
                rates["not RIFF"] += 1
                continue
            fmt, ch, rate, _, _, bits = struct.unpack_from("<HHIIHH", b, 20)
            rates[(fmt, ch, rate, bits)] += 1
    print("(format, channels, rate, bits):", dict(rates))
    music = [n for n in files if n.startswith("music/") or n.endswith((".ogg", ".mp3"))]
    print("music files in paks:", len(music), music[:5])


if __name__ == "__main__":
    main(sys.argv[1])
