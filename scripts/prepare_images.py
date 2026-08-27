#!/usr/bin/env python3
"""Resize the brand artwork into the small PNGs the dashboard embeds.

    pip install Pillow
    python scripts/prepare_images.py

Run this only when the artwork in `Design Assets` changes. The results are
committed to `src/images/`, and `build_dashboard.py` merely base64-encodes
them, which keeps the build deterministic: PNG compression differs slightly
between Pillow and zlib versions, so resizing on every build made CI produce
a byte-different file from a local build every single time.
"""

from __future__ import annotations

from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "Design Assets"
OUT = ROOT / "src" / "images"

# Source artwork is far larger than it is ever displayed — the agency logo is
# 2987px wide and shown at 22px tall. Heights here are 2x the CSS size, which
# stays sharp on retina screens without carrying a megabyte of base64.
IMAGES = {
    "sph-icon.png":  ("SPH Icon.png", 88),
    "favicon.png":   ("SPH Icon.png", 64),
    "cc-colour.png": ("Chaos Created Colour Logo@2x.png", 44),
    "cc-white.png":  ("Chaos Created White Logo@2x.png", 44),
}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    for target, (source, height) in IMAGES.items():
        with Image.open(ASSETS / source) as img:
            img = img.convert("RGBA")
            width = round(img.width * height / img.height)
            img = img.resize((width, height), Image.LANCZOS)
            img.save(OUT / target, format="PNG", optimize=True)
        size = (OUT / target).stat().st_size
        print(f"{target:16} {width}x{height}  {size / 1024:.1f} KB  <- {source}")


if __name__ == "__main__":
    main()
