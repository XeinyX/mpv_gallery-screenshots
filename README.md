# MPV Gallery Screenshots Overlay

A Lua user script for mpv that lets you:
- Save screenshots to per‑video folders
- Browse screenshots in an in‑player tiled gallery overlay
- Select, delete, and export as contact sheets (PNG), CSV, and XLSX
- Render timecode labels on thumbnails
- Keep XLSX files small via physical image resizing (Pillow), or fall back to visual scaling

Works on Linux, macOS, and Windows.

---

## Features

- Screenshots saved to `./images/<video_name>/f########.jpg` (or .png)
- In-player gallery overlay:
  - Grid navigation with paging
  - Click to seek to a screenshot's timestamp
- Exports:
  - Contact sheets (per page or all pages) → `./exports/<video>_contact_sheet_pXX.png`
  - CSV (filename, timecode, seconds, frame) → `./exports/<video>_gallery.csv`
  - XLSX via Python (`xlsxwriter`, optional `Pillow`) → `./exports/<video>.xlsx`
    - Horizontal and vertical centering of images and text
    - Row height adjusted to the image height
    - Physical image resizing with Pillow (smaller XLSX); fallback to visual scaling if Pillow is not available
- Timecode labels rendered into thumbnails (via ffmpeg filters)
- Cross-platform; ffmpeg used where available, mpv fallback where possible

---

## What’s new in this version

- Safe margins are now relative to the viewport size:
  - `SAFE_MARGIN_*_REL` are fractions of width/height; `GRID_GAP` remains in pixels.
- Fullscreen alignment fix:
  - Thumbnails are positioned using absolute OSD coordinates, so they match the clickable zones regardless of video vs. screen aspect ratio.
- Screenshot resolution control:
  - Configurable scale and minimum dimensions; preserves aspect ratio and never upscales above the source.
- Clear limit for gallery size:
  - Maximum thumbnails per page: 63 (example: 9×7 grid).

All settings are configured at the top of the script.

---

## Requirements

- mpv (0.35+ recommended; older builds may still work)
- ffmpeg in PATH (for labeled thumbnails and contact sheet composition)
- Python 3 (only for XLSX export)
  - `xlsxwriter` is required
  - `Pillow` (optional) enables physical resizing of images for smaller XLSX files

Install Python packages:
```bash
pip install xlsxwriter pillow
```

---

## Installation

1. Place `gallery_screenshots.lua` into your mpv `scripts` folder.
2. Place `gallery_xlsx_export.py` next to the Lua file (same folder).
3. Restart mpv.

Python virtual environment (recommended):
- Linux/macOS:
  ```bash
  python3 -m venv ~/venv-mpv
  source ~/venv-mpv/bin/activate
  pip install xlsxwriter pillow
  which python
  # Then set PYTHON_PATH in the Lua script to the path printed by `which python`
  ```
- Windows (PowerShell):
  ```powershell
  py -m venv C:\venv-mpv
  C:\venv-mpv\Scripts\Activate.ps1
  pip install xlsxwriter pillow
  # Then set PYTHON_PATH = "C:\\venv-mpv\\Scripts\\python.exe"
  ```

---

## Configuration

Open the Lua file and adjust the user‑configurable variables at the top section. Defaults shown below may differ from your copy.

### Paths and Python
- `PYTHON_PATH` (string or nil): Absolute path to Python interpreter. Leave `nil` to auto‑detect. It is recommended to set it up according to the instructions above.
- `XLSX_SCRIPT_NAME` (string): Python exporter script name (kept next to the Lua script).

### Gallery grid and layout
- `GRID_COLS` (int): Number of columns per page.
- `GRID_ROWS` (int): Number of rows per page.
- `GRID_GAP` (int, px): Gap between tiles and edges.
- `SAFE_MARGIN_TOP_REL` (0–1, of viewport height)
- `SAFE_MARGIN_BOTTOM_REL` (0–1, of viewport height)
- `SAFE_MARGIN_LEFT_REL` (0–1, of viewport width)
- `SAFE_MARGIN_RIGHT_REL` (0–1, of viewport width)

Notes:
- Maximum thumbnails per page is 63. Ensure `GRID_COLS × GRID_ROWS ≤ 63` (e.g., 9×7).
- Margins scale with the window; the gap remains constant in pixels.

### Screenshot output
- `SCREENSHOT_SCALE` (0.01–1.0): Base downscale factor for saved screenshots (e.g., `0.5` → 50%).
- `SCREENSHOT_MIN_WIDTH` (int, px): Minimum width; enforced without upscaling.
- `SCREENSHOT_MIN_HEIGHT` (int, px): Minimum height; enforced without upscaling.

Behavior:
- Preserves aspect ratio.
- Applies scale and then ensures the result meets the minimum width/height.
- Never upscales beyond the source dimensions.

### Time labels and selection visuals
- `ENABLE_TIME_LABELS` (bool): Draw HH:MM:SS into thumbnails.
- `LABEL_BOX_ALPHA` (0–1): Opacity of the time label background box.
- `LABEL_MARGIN_X`, `LABEL_MARGIN_Y` (int, px): Text padding inside the box.
- `LABEL_REL_SIZE` (0–1): Relative font size based on tile size.
- `SELECT_BOX_ALPHA` (0–1): Opacity of the red bottom strip indicating selection.

### Directories
- `DIR_IMAGES_BASE` (string): Per‑video screenshots folder: `./images/<video_name>/`.
- `DIR_EXPORTS_NAME` (string): Sibling exports folder: `./exports/`.

### XLSX export
- `XLSX_IMG_SCALE` (0.15–0.35 recommended): Image scale for XLSX export.
  - With Pillow: physical resizing (smaller XLSX).
  - Without Pillow: Excel visual scaling (larger XLSX, same look).

---

## Usage

Default key bindings (can be customized):
```
s                  Save screenshot → ./images/<video_name>/f########.jpg
g                  Toggle gallery overlay on/off

In gallery:
  Click            Seek to clicked thumbnail's time
  m                Toggle selection on a tile
  [  /  ]          Previous / Next page
  a                Select/Unselect all (current page)
  d                Delete selected (files + cache)
  Shift+c          Export contact sheet (current page) → ./exports/<video>_contact_sheet_pXX.png
  c                Export contact sheets for all pages
  e                Export CSV → ./exports/<video>_gallery.csv
  x                Export XLSX → ./exports/<video>.xlsx
```

- Thumbnails are cached as raw BGRA in `./images/<video>/.gallery_bgra/` and rebuilt when window size changes.

---

## Exports

- Contact sheets (PNG): ffmpeg `tile` filter composes the page (padding/margins from your grid settings).
- CSV: `filename`, `timecode` (HH:MM:SS), `seconds` (float), `frame` (if derivable from filename).
- XLSX (Python):
  - Images in column A, timecodes in column B
  - Horizontal and vertical centering
  - Row height set to the image height
  - Physical resizing with Pillow; otherwise visual scaling

---

## File structure

For a video at `.../movie.ext`:
```
.../images/movie/               # per-video screenshots
.../images/movie/.gallery_bgra  # cached BGRA tiles (auto-managed)
.../exports/                    # contact sheets, CSV, XLSX
```

Screenshots are named `f########.jpg` by default (based on estimated frame number).

---

## Troubleshooting

- ffmpeg not found:
  - Install ffmpeg and ensure it’s in PATH.
- Python not found (XLSX):
  - Set `PYTHON_PATH` in the Lua script or ensure `python3`/`python` is in PATH.
- XLSX too large:
  - Install Pillow to enable physical resizing; lower `XLSX_IMG_SCALE`.
- Gallery looks offset in fullscreen:
  - Fixed in this version (absolute OSD coordinates). If you still see issues, ensure you’re running the latest script.
- Logging:
  - Run mpv with logging to capture script output:
    ```bash
    mpv --log-file=mpv.log --msg-level=script=trace yourvideo.mkv
    ```

---

## License

MIT License. See `LICENSE`.

---

## Credits

- mpv authors and contributors
- ffmpeg project
- xlsxwriter and Pillow authors
- Community testers and issue reporters
