#!/usr/bin/env python3
"""
make_print_sheets.py
====================

Lay out the downloaded album covers onto printable 8.5 x 11" sheets, four to a
page, each sized to exactly 4 x 4 inches, and write a single multi-page PDF you
can send straight to a printer.

Why a PDF at a fixed DPI (and not just resizing pixels)
-------------------------------------------------------
The covers come back at all sorts of pixel dimensions (1400², 3000², …). A raw
"pixels -> inches" guess would print them at different physical sizes. Instead we
pick a print resolution (DPI) and compute the pixel box for a *physical* size:

    4 inches * 300 DPI = 1200 px

Every cover is fit into that 1200x1200 box, the page is built at
8.5*300 x 11*300 = 2550 x 3300 px, and the PDF is saved with that same DPI so
each square lands at a true 4 x 4 inches on paper, no matter its source pixels.

Usage
-----
    python make_print_sheets.py                      # uses HighRes_Covers/ -> PDF
    python make_print_sheets.py --dpi 600            # sharper (needs bigger files)
    python make_print_sheets.py --fit contain        # pad instead of crop
    python make_print_sheets.py --no-guides          # no cut lines
    python make_print_sheets.py --input some/dir --output sheets.pdf
"""

import argparse
import os
import sys

try:
    from PIL import Image, ImageDraw
except ImportError:  # pragma: no cover - friendly guidance
    sys.exit("Missing dependency 'Pillow'. Install with: pip install Pillow")


# --------------------------------------------------------------------------- #
# Defaults (all overridable on the command line)                              #
# --------------------------------------------------------------------------- #

INPUT_FOLDER = "HighRes_Covers"          # where the approved covers live
OUTPUT_FILE = "Album_Covers_Print.pdf"   # printable result

PAGE_W_IN, PAGE_H_IN = 8.5, 11.0         # US Letter photo paper
IMAGE_IN = 4.0                           # each cover: 4" x 4"
COLS, ROWS = 2, 2                        # 4 per sheet
GUTTER_IN = 0.25                         # gap between covers
DEFAULT_DPI = 300                        # print resolution

IMAGE_EXTS = (".jpg", ".jpeg", ".png")   # what we treat as a cover


def collect_images(folder: str) -> list[str]:
    """Return cover image paths, sorted alphabetically (case-insensitive).

    The files are named "Artist - Album.jpg", so a plain name sort keeps them
    in the same artist/album order you reviewed them in.
    """
    if not os.path.isdir(folder):
        sys.exit(f"Input folder not found: {folder}\n"
                 "Run the artwork fetcher first, or pass --input <dir>.")
    names = [n for n in os.listdir(folder)
             if n.lower().endswith(IMAGE_EXTS) and not n.startswith(".")]
    names.sort(key=str.lower)
    return [os.path.join(folder, n) for n in names]


def load_square(path: str, box_px: int, fit: str) -> Image.Image:
    """Open an image and return a ``box_px`` x ``box_px`` RGB tile.

    PNG transparency is flattened onto white. ``fit`` is either:
      - "cover":   center-crop to square, then scale to fill the box (default)
      - "contain": scale to fit inside the box, padding with white (no crop)
    """
    img = Image.open(path)
    img.load()

    # Flatten any transparency (the "couple of PNGs") onto a white background.
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        rgba = img.convert("RGBA")
        bg = Image.new("RGBA", rgba.size, (255, 255, 255, 255))
        img = Image.alpha_composite(bg, rgba).convert("RGB")
    else:
        img = img.convert("RGB")

    w, h = img.size
    if fit == "contain":
        scale = min(box_px / w, box_px / h)
        new = img.resize((max(1, round(w * scale)), max(1, round(h * scale))),
                         Image.LANCZOS)
        tile = Image.new("RGB", (box_px, box_px), (255, 255, 255))
        tile.paste(new, ((box_px - new.width) // 2, (box_px - new.height) // 2))
        return tile

    # "cover": center-crop to a square, then resize to the exact box.
    side = min(w, h)
    left, top = (w - side) // 2, (h - side) // 2
    square = img.crop((left, top, left + side, top + side))
    return square.resize((box_px, box_px), Image.LANCZOS)


def build_pages(paths: list[str], dpi: int, fit: str, guides: bool) -> list[Image.Image]:
    """Compose the covers into a list of full-page RGB images."""
    page_w = round(PAGE_W_IN * dpi)
    page_h = round(PAGE_H_IN * dpi)
    box = round(IMAGE_IN * dpi)
    gutter = round(GUTTER_IN * dpi)
    per_page = COLS * ROWS

    # Center the 2x2 block of covers on the page.
    grid_w = COLS * box + (COLS - 1) * gutter
    grid_h = ROWS * box + (ROWS - 1) * gutter
    origin_x = (page_w - grid_w) // 2
    origin_y = (page_h - grid_h) // 2

    pages: list[Image.Image] = []
    for start in range(0, len(paths), per_page):
        page = Image.new("RGB", (page_w, page_h), (255, 255, 255))
        draw = ImageDraw.Draw(page) if guides else None
        for slot, path in enumerate(paths[start:start + per_page]):
            row, col = divmod(slot, COLS)
            x = origin_x + col * (box + gutter)
            y = origin_y + row * (box + gutter)
            try:
                tile = load_square(path, box, fit)
            except Exception as exc:  # noqa: BLE001 - skip unreadable files
                print(f"[warn] Skipping {os.path.basename(path)}: {exc}")
                continue
            page.paste(tile, (x, y))
            if draw is not None:
                # Thin light-gray outline as a 4x4 cut guide.
                draw.rectangle([x, y, x + box - 1, y + box - 1],
                               outline=(200, 200, 200), width=1)
        pages.append(page)
    return pages


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Lay album covers onto printable 8.5x11 sheets (4 per page).")
    p.add_argument("--input", default=INPUT_FOLDER,
                   help=f"Folder of cover images (default: {INPUT_FOLDER}).")
    p.add_argument("--output", default=OUTPUT_FILE,
                   help=f"Output PDF path (default: {OUTPUT_FILE}).")
    p.add_argument("--dpi", type=int, default=DEFAULT_DPI,
                   help=f"Print resolution (default: {DEFAULT_DPI}).")
    p.add_argument("--fit", choices=("cover", "contain"), default="cover",
                   help="'cover' center-crops to fill 4x4 (default); "
                        "'contain' pads with white so nothing is cropped.")
    p.add_argument("--no-guides", dest="guides", action="store_false",
                   help="Don't draw the thin cut-line border around each cover.")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    paths = collect_images(args.input)
    if not paths:
        sys.exit(f"No images ({', '.join(IMAGE_EXTS)}) found in {args.input}.")

    pages = build_pages(paths, args.dpi, args.fit, args.guides)
    if not pages:
        sys.exit("Nothing to print (all images were unreadable).")

    # Multi-page PDF; resolution embeds the true physical size for the printer.
    pages[0].save(
        args.output, "PDF", save_all=True, append_images=pages[1:],
        resolution=float(args.dpi),
    )
    sheets = len(pages)
    print(f"Wrote {args.output}: {len(paths)} covers across {sheets} "
          f"sheet{'s' if sheets != 1 else ''} "
          f"({COLS}x{ROWS} per page, {IMAGE_IN:g}\"x{IMAGE_IN:g}\" each, "
          f"{args.dpi} DPI).")
    print("Print it at 100% / 'Actual size' (no scaling) so the covers stay "
          "exactly 4x4 inches.")


if __name__ == "__main__":
    main()
