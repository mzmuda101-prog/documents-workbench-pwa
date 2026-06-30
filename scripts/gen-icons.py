#!/usr/bin/env python3
"""Generate blue Word-style PWA icons."""
from pathlib import Path
from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent / "assets" / "images"
BLUE = (24, 90, 189)  # #185ABD — Word blue
WHITE = (255, 255, 255)


def draw_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    r = int(size * 0.22)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=BLUE)
    pad = int(size * 0.24)
    w = size - pad * 2
    h = int(size * 0.62)
    y = int(size * 0.19)
    d.rounded_rectangle([pad, y, pad + w, y + h], radius=int(size * 0.06), outline=WHITE, width=max(2, size // 22))
    lw = max(2, size // 24)
    y1 = y + int(h * 0.28)
    y2 = y + int(h * 0.48)
    x1 = pad + int(w * 0.18)
    x2 = pad + int(w * 0.82)
    d.line([(x1, y1), (x2, y1)], fill=WHITE, width=lw)
    d.line([(x1, y2), (x2, y2)], fill=WHITE, width=lw)
    return img


def main():
    ROOT.mkdir(parents=True, exist_ok=True)
    for name, size in [("favicon.png", 64), ("apple-touch-icon.png", 180), ("icon-512.png", 512)]:
        draw_icon(size).save(ROOT / name, "PNG")
    print("✅  icons generated in assets/images/")


if __name__ == "__main__":
    main()
