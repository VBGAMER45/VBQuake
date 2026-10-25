"""
typekeys.py - turns text into QuakeVB -key arguments for scripted test runs.

  python -I tools/typekeys.py [--file keys.txt] <start frame> <step> <item> [<item> ...]

An item is text to type (one key per character) or a key name in angle
brackets: <ENTER> <TAB> <UPARROW> <BACKSPACE> <ESCAPE> <SPACE> ... A space
in the text is typed as SPACE. Each key goes down at its frame and up one
frame later; the next key follows <step> frames after. Prints the arguments
on one line, for example:

  QuakeVB.exe -hidden $(python -I tools/typekeys.py 10 2 "echo hi" "<ENTER>")

The command line keeps only 50 arguments (16 keys); --file writes the keys
to a file for -keyfile instead, one "<frame> <key>" per line.
"""
import sys

NAMES = {" ": "SPACE", ";": "SEMICOLON"}


def keys(items):
    for item in items:
        if item.startswith("<") and item.endswith(">") and len(item) > 2:
            yield item[1:-1]
            continue
        for ch in item:
            yield NAMES.get(ch, ch)


def main():
    args = sys.argv[1:]
    path = None
    if args[0] == "--file":
        path = args[1]
        args = args[2:]
    frame = int(args[0])
    step = int(args[1])
    out = []
    for k in keys(args[2:]):
        out.append((frame, k))
        frame += step
    if path:
        # CRLF: VB's Line Input does not end a line at a bare LF
        with open(path, "w", newline="\r\n") as f:
            f.write("".join("%d %s\n" % (fr, k) for fr, k in out))
        print("%d keys, last at frame %d" % (len(out), out[-1][0]))
    else:
        print(" ".join("-key %d %s" % (fr, k) for fr, k in out))


if __name__ == "__main__":
    main()
