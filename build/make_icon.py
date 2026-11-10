# -*- coding: utf-8 -*-
"""生成 app.ico（纯标准库，无第三方依赖）。"""
import os
import struct
import zlib


def png_bytes(w, h, rgba):
    raw = b"".join(b"\x00" + bytes(rgba[y * w * 4:(y + 1) * w * 4]) for y in range(h))

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 9))
            + chunk(b"IEND", b""))


def seg_dist(px, py, ax, ay, bx, by):
    vx, vy = bx - ax, by - ay
    wx, wy = px - ax, py - ay
    L = vx * vx + vy * vy
    t = 0.0 if L == 0 else max(0.0, min(1.0, (wx * vx + wy * vy) / L))
    dx, dy = px - (ax + t * vx), py - (ay + t * vy)
    return (dx * dx + dy * dy) ** 0.5


def render(size):
    S = size
    r = S * 0.22
    stroke = max(1.4, S * 0.085)
    ax1, ay1 = S * 0.27, S * 0.52
    ax2, ay2 = S * 0.44, S * 0.69
    ax3, ay3 = S * 0.75, S * 0.33
    out = bytearray(S * S * 4)
    SS = 4
    for y in range(S):
        for x in range(S):
            cov = 0
            white = 0
            for sy in range(SS):
                for sx in range(SS):
                    px = x + (sx + 0.5) / SS
                    py = y + (sy + 0.5) / SS
                    # 圆角矩形
                    dx = max(r - px, px - (S - r), 0)
                    dy = max(r - py, py - (S - r), 0)
                    if (dx * dx + dy * dy) ** 0.5 <= r:
                        cov += 1
                        d = min(seg_dist(px, py, ax1, ay1, ax2, ay2),
                                seg_dist(px, py, ax2, ay2, ax3, ay3))
                        if d <= stroke / 2:
                            white += 1
            if cov == 0:
                continue
            i = (y * S + x) * 4
            t = y / max(1, S - 1)
            rr = int(59 + (37 - 59) * t)
            gg = int(130 + (99 - 130) * t)
            bb = int(246 + (235 - 246) * t)
            a = int(255 * cov / (SS * SS))
            w = white / (SS * SS)
            out[i] = int(rr * (1 - w) + 255 * w)
            out[i + 1] = int(gg * (1 - w) + 255 * w)
            out[i + 2] = int(bb * (1 - w) + 255 * w)
            out[i + 3] = a
    return out


def main():
    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = [(s, png_bytes(s, s, render(s))) for s in sizes]
    head = struct.pack("<HHH", 0, 1, len(images))
    offset = 6 + 16 * len(images)
    dirs = b""
    for s, data in images:
        dirs += struct.pack("<BBBBHHII", s % 256, s % 256, 0, 0, 1, 32, len(data), offset)
        offset += len(data)
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "assets", "app.ico")
    with open(out, "wb") as fh:
        fh.write(head + dirs + b"".join(d for _, d in images))
    print("icon ->", os.path.abspath(out), os.path.getsize(out))


if __name__ == "__main__":
    main()
