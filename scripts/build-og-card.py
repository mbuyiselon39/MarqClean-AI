#!/usr/bin/env python3
"""
Regenerate public/og-card.png for the cobalt design system.

The previous card was a truncated PNG: its IDAT chunk declared 13,594 bytes
but the file ended at 12,475, with no IEND chunk at the correct offset. The
CI gate only checked the PNG signature and the IHDR dimensions, so the broken
file passed while every social platform failed to render it.

Requires Pillow (not a project dependency -- run this in a throwaway venv):

    python3 -m venv /tmp/venv && /tmp/venv/bin/pip install pillow
    /tmp/venv/bin/python scripts/build-og-card.py

Colours match the `.mc-next-hero` rules in index.css so the share image and
the homepage hero stay visually identical. Copy is taken verbatim from the
og:title / og:description tags in index.html and the existing privacy
statement, so the card asserts nothing the site does not already say.
"""

import math
import struct
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "public" / "og-card.png"
MANROPE = ROOT / "public" / "fonts" / "manrope-latin-wght-normal.woff2"
DM_SANS = ROOT / "public" / "fonts" / "dm-sans-latin-wght-normal.woff2"

W, H = 1200, 630

# --- palette (mirrors .mc-next-hero) ------------------------------------
BASE = (19, 25, 65)          # #131941
GLOW_A = (45, 61, 159)       # #2d3d9f
GLOW_B = (31, 44, 131)       # #1f2c83
DOT = (212, 217, 244)        # #d4d9f4
INK = (249, 249, 253)        # #f9f9fd
MUTED = (179, 185, 219)      # #b3b9db
ACCENT = (198, 204, 243)     # #c6ccf3

EYEBROW = "BROWSER-FIRST DATA OPERATIONS"
TITLE = "Clean data. Clear decisions."
BODY = (
    "Clean, validate, transform and reconcile Excel, CSV and PDF data "
    "with a browser-first data operations platform."
)
PRIVACY = "Processed in your browser. Not uploaded to a server."
DOMAIN = "marqcleanai.vertexsg.co.za"


def font(path: str, size: int, weight: int) -> ImageFont.FreeTypeFont:
    f = ImageFont.truetype(path, size)
    # The self-hosted files are variable fonts whose default instance is light;
    # request the weight we actually want.
    try:
        f.set_variation_by_axes([weight])
    except Exception:
        pass
    return f


def wrap(draw, text, fnt, max_width):
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if draw.textlength(trial, font=fnt) <= max_width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def main() -> None:
    img = Image.new("RGB", (W, H), BASE)
    # Two cobalt glows, positioned as in the hero's :before rule.
    overlay = Image.new("RGB", (W, H), (0, 0, 0))
    opx = overlay.load()
    glows = [((int(W * 0.81), int(H * 0.52)), int(W * 0.53), GLOW_A, 0.95),
             ((int(W * 0.08), int(H * 0.95)), int(W * 0.45), GLOW_B, 0.9)]
    for (cx, cy), radius, colour, peak in glows:
        for y in range(H):
            dy2 = (y - cy) ** 2
            for x in range(W):
                d2 = (x - cx) ** 2 + dy2
                if d2 >= radius * radius:
                    continue
                a = peak * (1 - math.sqrt(d2) / radius) ** 2
                r, g, b = opx[x, y]
                opx[x, y] = (
                    min(255, int(r + colour[0] * a)),
                    min(255, int(g + colour[1] * a)),
                    min(255, int(b + colour[2] * a)),
                )
    img = Image.blend(img, overlay, 0.55)

    draw = ImageDraw.Draw(img, "RGBA")

    # Dot grid, as in the hero's :after rule.
    for y in range(24, H, 28):
        for x in range(24, W, 28):
            fade = max(0.0, min(1.0, (x - 120) / (W - 120)))
            if fade <= 0.02:
                continue
            draw.ellipse([x - 0.6, y - 0.6, x + 0.6, y + 0.6],
                         fill=(*DOT, int(46 * fade)))

    f_eye = font(str(MANROPE), 21, 700)
    f_title = font(str(MANROPE), 84, 800)
    f_body = font(str(DM_SANS), 27, 400)
    f_mark = font(str(MANROPE), 30, 700)
    f_small = font(str(DM_SANS), 21, 400)

    mx = 88

    # Wordmark
    draw.text((mx, 62), "MarqClean", font=f_mark, fill=INK)
    mw = draw.textlength("MarqClean", font=f_mark)
    draw.text((mx + mw + 8, 62), "AI", font=font(str(MANROPE), 30, 400), fill=ACCENT)

    # Eyebrow
    y = 176
    draw.text((mx, y), EYEBROW, font=f_eye, fill=ACCENT)

    # Title
    y += 52
    for line in wrap(draw, TITLE, f_title, W - mx * 2):
        draw.text((mx, y), line, font=f_title, fill=INK)
        y += 92

    # Body
    y += 14
    for line in wrap(draw, BODY, f_body, W - mx * 2 - 120):
        draw.text((mx, y), line, font=f_body, fill=MUTED)
        y += 40

    # Footer rule + privacy + domain
    fy = H - 92
    draw.line([(mx, fy - 26), (W - mx, fy - 26)], fill=(*ACCENT, 60), width=1)
    draw.ellipse([mx, fy - 5, mx + 12, fy + 7], fill=(120, 214, 162))
    draw.text((mx + 24, fy - 8), PRIVACY, font=f_small, fill=MUTED)
    dw = draw.textlength(DOMAIN, font=f_small)
    draw.text((W - mx - dw, fy - 8), DOMAIN, font=f_small, fill=ACCENT)

    img.save(OUT, "PNG", optimize=True)

    # Prove the result is a complete, decodable PNG before handing it over.
    raw = OUT.read_bytes()
    assert raw[:8] == b"\x89PNG\r\n\x1a\n", "bad signature"
    i, end_at = 8, None
    while i < len(raw) - 8:
        (ln,) = struct.unpack(">I", raw[i:i + 4])
        typ = raw[i + 4:i + 8]
        if typ == b"IHDR":
            w, h = struct.unpack(">II", raw[i + 8:i + 16])
            assert (w, h) == (W, H), f"wrong dimensions {w}x{h}"
        if typ == b"IEND":
            end_at = i + 12 + ln
            break
        i += 12 + ln
    assert end_at == len(raw), f"IEND not at EOF ({end_at} vs {len(raw)})"
    Image.open(OUT).verify()
    print(f"wrote {OUT.relative_to(ROOT)}  {W}x{H}  {len(raw)} bytes  (chunks valid, decodes)")


if __name__ == "__main__":
    main()
