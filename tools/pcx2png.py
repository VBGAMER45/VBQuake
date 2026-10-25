"""
pcx2png.py - converts a Quake screenshot (8-bit PCX, as SCR_ScreenShot_f
writes it) to PNG, optionally scaled up with nearest-neighbour pixels.

  python -I tools/pcx2png.py in.pcx out.png [scale]
"""
import struct
import sys
import zlib


def read_pcx(path):
    with open(path, "rb") as f:
        data = f.read()
    if data[0] != 0x0A or data[3] != 8:
        raise ValueError("not an 8-bit PCX file")
    xmin, ymin, xmax, ymax = struct.unpack_from("<4H", data, 4)
    width, height = xmax - xmin + 1, ymax - ymin + 1
    bpl = struct.unpack_from("<H", data, 66)[0]
    if data[-769] != 0x0C:
        raise ValueError("no 256 color palette")
    palette = data[-768:]
    pixels = bytearray()
    p = 128
    for _ in range(height):
        row = bytearray()
        while len(row) < bpl:
            b = data[p]
            p += 1
            if b & 0xC0 == 0xC0:
                row += bytes([data[p]]) * (b & 0x3F)
                p += 1
            else:
                row.append(b)
        pixels += row[:width]
    return width, height, pixels, palette


def write_png(path, width, height, pixels, palette, scale):
    def chunk(tag, body):
        return struct.pack(">I", len(body)) + tag + body + struct.pack(">I", zlib.crc32(tag + body) & 0xFFFFFFFF)

    raw = bytearray()
    for y in range(height):
        row = pixels[y * width:(y + 1) * width]
        if scale > 1:
            row = bytes(b for b in row for _ in range(scale))
        for _ in range(scale):
            raw.append(0)
            raw += row
    png = b"\x89PNG\r\n\x1a\n"
    png += chunk(b"IHDR", struct.pack(">IIBBBBB", width * scale, height * scale, 8, 3, 0, 0, 0))
    png += chunk(b"PLTE", bytes(palette))
    png += chunk(b"IDAT", zlib.compress(bytes(raw), 9))
    png += chunk(b"IEND", b"")
    with open(path, "wb") as f:
        f.write(png)


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 1
    scale = int(sys.argv[3]) if len(sys.argv) > 3 else 1
    width, height, pixels, palette = read_pcx(sys.argv[1])
    write_png(sys.argv[2], width, height, pixels, palette, scale)
    print("%s: %dx%d -> %s (x%d)" % (sys.argv[1], width, height, sys.argv[2], scale))
    return 0


if __name__ == "__main__":
    sys.exit(main())
