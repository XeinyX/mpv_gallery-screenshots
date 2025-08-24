#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gallery_xlsx_export.py
- Reads images from --images-dir (e.g. images/<video_name>)
- Creates an XLSX with an image column and a timecode column
- Centers image and text horizontally and vertically
- Sets row heights based on actual image height
- Optionally physically resizes images using Pillow to keep XLSX small

Usage:
  python gallery_xlsx_export.py --images-dir ... --out ... --fps 25 --scale 0.20 --video-name NAME --resize physical --center 1

Requirements:
  - xlsxwriter (pip install xlsxwriter)
  - Optional: Pillow for physical resize (pip install pillow)
"""

import argparse
import os
import re
import math
import sys
import tempfile
import shutil

try:
    import xlsxwriter
except ImportError:
    print("Missing 'xlsxwriter'. Install with: pip install xlsxwriter", file=sys.stderr)
    sys.exit(1)

try:
    from PIL import Image  # Optional, for physical resize
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

def list_images_sorted(images_dir):
    try:
        files = sorted(os.listdir(images_dir))
    except FileNotFoundError:
        return []
    out = []
    for name in files:
        lower = name.lower()
        if lower.endswith((".jpg", ".jpeg", ".png")):
            out.append(name)
    return out

def parse_frame_from_filename(name):
    m = re.match(r"^f(\d{8})\.(?:jpe?g|png)$", name, re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    m = re.search(r"(\d{8})", name)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None

def seconds_to_hhmmss_floor(seconds: float) -> str:
    if not seconds or seconds < 0:
        return "00:00:00"
    s = int(math.floor(seconds))
    h = s // 3600
    m = (s % 3600) // 60
    r = s % 60
    return f"{h:02d}:{m:02d}:{r:02d}"

# Approximations based on Excel defaults (Calibri 11)
def pixels_to_col_width(pixels: int) -> float:
    # Excel approximation: pixels ≈ 7 * width + 5 → width ≈ (pixels - 5) / 7
    return max(1.0, (pixels - 5) / 7.0)

def col_width_to_pixels(width: float) -> int:
    return int(7.0 * width + 5)

def pixels_to_row_points(pixels: int) -> float:
    # Excel approx: pixels ≈ 4/3 * points → points ≈ 3/4 * pixels
    return max(1.0, (3.0 / 4.0) * pixels)

def get_image_size(path: str):
    try:
        with Image.open(path) as im:
            return im.width, im.height
    except Exception:
        return None, None

def physical_resize_image(src: str, dst: str, scale: float) -> tuple[int, int]:
    with Image.open(src) as im:
        w, h = im.width, im.height
        nw = max(1, int(round(w * scale)))
        nh = max(1, int(round(h * scale)))
        im = im.resize((nw, nh), Image.LANCZOS)
        # Preserve format if possible; default to PNG for lossless/compat.
        ext = os.path.splitext(src)[1].lower()
        if ext in (".jpg", ".jpeg"):
            im.save(dst, format="JPEG", quality=85, optimize=True)
        elif ext == ".png":
            im.save(dst, format="PNG", optimize=True)
        else:
            im.save(dst, format="PNG", optimize=True)
        return nw, nh

def main():
    p = argparse.ArgumentParser(description="Export screenshots to XLSX (image + timecode).")
    p.add_argument("--images-dir", required=True, help="Path to images/<video_name>")
    p.add_argument("--out", required=True, help="Output XLSX path")
    p.add_argument("--fps", type=float, default=0.0, help="Video FPS (for timecode)")
    p.add_argument("--scale", type=float, default=0.25, help="Scale factor for images (0.15–0.35 recommended)")
    p.add_argument("--video-name", default="", help="Video name (info only)")
    p.add_argument("--resize", choices=["physical", "scale-only"], default="physical",
                   help="physical: resize images using Pillow (smaller XLSX); scale-only: visual scale in Excel (bigger XLSX)")
    p.add_argument("--center", default="1", help="Center image and text (1=yes, 0=no)")
    p.add_argument("--pad-x", type=int, default=0, help="Horizontal padding (pixels) inside image cell")
    p.add_argument("--pad-y", type=int, default=0, help="Vertical padding (pixels) inside image cell")
    args = p.parse_args()

    images_dir = os.path.abspath(args.images_dir)
    out_path = os.path.abspath(args.out)
    fps = float(args.fps or 0.0)
    scale = float(args.scale or 0.25)
    do_center = str(args.center).strip() not in ("0", "false", "False", "no", "No")

    imgs = list_images_sorted(images_dir)
    if not imgs:
        print("No images to export.", file=sys.stderr)
        sys.exit(2)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)

    # Prepare temp dir for physically resized images (if used)
    tmpdir = None
    use_physical = (args.resize == "physical" and PIL_AVAILABLE)
    if args.resize == "physical" and not PIL_AVAILABLE:
        print("Pillow not available, falling back to scale-only mode.", file=sys.stderr)
        use_physical = False
    if use_physical:
        tmpdir = tempfile.mkdtemp(prefix="xlsx_gallery_")

    try:
        # Preprocess images: get sizes and optionally physically resize
        processed = []  # list of tuples (path, w, h)
        for name in imgs:
            src = os.path.join(images_dir, name)
            if use_physical:
                dst = os.path.join(tmpdir, name)
                try:
                    nw, nh = physical_resize_image(src, dst, scale)
                    processed.append((dst, nw, nh))
                except Exception as e:
                    # Fallback: no resize; insert original with visual scale
                    w, h = get_image_size(src)
                    processed.append((src, w or 0, h or 0))
            else:
                w, h = get_image_size(src)
                processed.append((src, int(round((w or 0) * scale)), int(round((h or 0) * scale))))

        # Compute largest width/height after processing (to size the column and offsets)
        max_w = max((w for _, w, _ in processed), default=0)
        max_h = max((h for _, _, h in processed), default=0)

        wb = xlsxwriter.Workbook(out_path)
        ws = wb.add_worksheet("Gallery")

        # Formats
        header_fmt = wb.add_format({"bold": True, "bg_color": "#EEEEEE", "align": "center", "valign": "vcenter"})
        text_fmt   = wb.add_format({"align": "center", "valign": "vcenter"})

        # Column widths: A = image, B = timecode
        # Set column A width based on max_w + horizontal padding.
        colA_pixels = max_w + 2 * args.pad_x
        colA_width  = pixels_to_col_width(colA_pixels)
        ws.set_column("A:A", colA_width)  # width in "Excel width units", not pixels
        ws.set_column("B:B", 16)          # reasonable width for timecode

        # Header
        ws.write(0, 0, "Image", header_fmt)
        ws.write(0, 1, "Timecode", header_fmt)
        ws.freeze_panes(1, 0) 

        row = 1
        for i, name in enumerate(imgs):
            orig_path = os.path.join(images_dir, name)
            if use_physical:
                img_path, w, h = processed[i]
                x_scale = 1.0
                y_scale = 1.0
            else:
                img_path, w, h = processed[i]
                # When using visual scale-only, we still insert the original image
                # and apply x/y scale. Recompute w/h for layout/row height.
                ow, oh = get_image_size(img_path)
                if ow and oh and ow > 0 and oh > 0:
                    w = int(round(ow * scale))
                    h = int(round(oh * scale))
                x_scale = scale
                y_scale = scale

            # Timecode
            frame = parse_frame_from_filename(name)
            tc = ""
            if frame is not None and fps > 0.0:
                tc = seconds_to_hhmmss_floor(frame / fps)

            # Set row height to image height + vertical padding (convert to points)
            row_pixels = h + 2 * args.pad_y
            ws.set_row(row, pixels_to_row_points(row_pixels), text_fmt if do_center else None)

            # Compute offsets to center image in the cell (if requested)
            insert_opts = {}
            if not use_physical:
                # When visual scaling, size in sheet is determined by x/y_scale
                # For offsets and row height we computed target 'w'/'h' already.
                pass

            if do_center:
                col_pixels = colA_pixels
                x_off = max(0, int((col_pixels - w) / 2))
                y_off = max(0, int((row_pixels - h) / 2))
                insert_opts["x_offset"] = x_off
                insert_opts["y_offset"] = y_off

            # Insert image
            try:
                if use_physical:
                    ws.insert_image(row, 0, img_path, insert_opts)
                else:
                    insert_opts["x_scale"] = x_scale
                    insert_opts["y_scale"] = y_scale
                    ws.insert_image(row, 0, img_path, insert_opts)
            except Exception:
                # Fallback: write path instead of image
                ws.write(row, 0, orig_path, text_fmt if do_center else None)

            # Write timecode, centered
            ws.write(row, 1, tc, text_fmt if do_center else None)

            row += 1

        wb.close()
        print(f"OK: {out_path}")
        return 0
    finally:
        if tmpdir and os.path.isdir(tmpdir):
            shutil.rmtree(tmpdir, ignore_errors=True)

if __name__ == "__main__":
    sys.exit(main())
