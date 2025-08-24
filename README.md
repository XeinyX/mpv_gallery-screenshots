# MPV Gallery Screenshots Overlay

A Lua user script for mpv that lets you:
- Save screenshots to per-video folders
- Browse screenshots in an in-player tiled gallery overlay
- Select, delete, and export as contact sheets (PNG), CSV, and XLSX
- Insert timecode labels on thumbnails
- Keep XLSX files small via physical image resizing (Pillow), or fall back to visual scaling

Works on Linux, macOS, and Windows.

---

## Features

- Screenshots saved to `../images/<video_name>/f########.jpg` (or .png)
- In-player gallery overlay:
  - Grid navigation with paging
  - Click to seek to a screenshot's timestamp
- Exports:
  - Contact sheets (per page or all pages) → `../exports/<video>_contact_sheet_pXX.png`
  - CSV (filename, timecode, seconds, frame) → `../exports/<video>_gallery.csv`
  - XLSX via Python (`xlsxwriter`, optional `Pillow`) → `../exports/<video>.xlsx`
    - Horizontal and vertical centering of images and text
    - Row height adjusted to the image height
    - Physical image resizing with Pillow (smaller XLSX); fallback to visual scaling if Pillow is not available
- Timecode labels rendered into thumbnails (via ffmpeg filters)
- Cross-platform; ffmpeg used where available, mpv fallback where possible

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

1. Download `mpv-gallery-screenshots.lua` and place it in your mpv `scripts` folder.
2. Place `gallery_xlsx_export.py` next to the Lua file (same folder).
3. Restart mpv.

Python virtual environment (recommended):
- Linux/macOS:
  ```bash
  python3 -m venv ~/venv-mpv
  source ~/venv-mpv/bin/activate
  pip install xlsxwriter pillow
  which python
  # Then set PYTHON_PATH to output of which python
  # on Linux local PYTHON_PATH = "/home/<your_name>/venv-mpv/bin/python"
  # on macOS local PYTHON_PATH = "/Users/<your_name>/venv-mpv/bin/python"
  
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

Open the Lua file and adjust the user-configurable variables at the top:

```lua
-- Python interpreter (absolute path). Leave nil to auto-detect.
local PYTHON_PATH = nil

-- Scale factor for images in XLSX (0.15–0.35 recommended).
-- If Pillow is available, images will be physically resized (smaller XLSX).
-- Otherwise, a visual scale will be applied by Excel (larger XLSX).
local XLSX_IMG_SCALE = 0.20

-- Gallery grid
local GRID_COLS = 5
local GRID_ROWS = 4
local GRID_GAP  = 12
```

---

## Usage

Default key bindings:
```
s                  Save screenshot → ../images/<video_name>/f########.jpg
g                  Toggle gallery overlay on/off

In gallery:
  Click            Seek to clicked thumbnail's time
  m                Toggle selection on a tile
  [  /  ]          Go to previous / next page
  a  /  u          Select all / Unselect all (current page)
  d                Delete selected (files + cache)
  Shift+c          Export contact sheet (current page) → ../exports/<video>_contact_sheet_pXX.png
  c                Export contact sheets for all pages
  e                Export CSV → ../exports/<video>_gallery.csv
  x                Export XLSX → ../exports/<video>.xlsx
```

Notes:
- Thumbnails are cached as BGRA raw frames in `../images/<video>/.gallery_bgra/`.
- On significant cell size changes, the cache is purged and regenerated.

---

## Export Details

- Contact sheets (PNG): use ffmpeg to compose the page tiles via `tile` filter (with padding/margins).
- CSV: includes `filename`, `timecode` (HH:MM:SS), `seconds` (float), and `frame` number (if derivable from filename).
- XLSX (Python):
  - Places images in column A, timecodes in column B
  - Centers image and text both horizontally and vertically
  - Sets row height to match the scaled image height
  - With `--resize physical` and Pillow, actual files are resized to reduce XLSX size
  - Without Pillow, falls back to Excel visual scaling (larger XLSX but visually identical)

---

## Troubleshooting

- "Python not found":
  - Set `PYTHON_PATH` in the Lua script to your interpreter
  - Ensure Python is in PATH (`python3`, `python`, or `py`)

- XLSX is too large:
  - Install Pillow and use physical resizing (default behavior if Pillow is available)
  - Lower `XLSX_IMG_SCALE` (e.g., `0.18` or `0.16`)

- No timecode labels on thumbnails:
  - Ensure `ffmpeg` is installed and in PATH

- Errors about `mp.get_script_file`:
  - The script includes a safe fallback that uses `debug.getinfo`; make sure you’re using the latest script

- Logging:
  - Run mpv with a log file to capture script output:
    ```bash
    mpv --log-file=mpv.log --msg-level=script=trace yourvideo.mkv
    ```

---

## License

MIT License. See `LICENSE` file.

---

## Credits

- mpv authors and contributors
- ffmpeg project
- xlsxwriter and Pillow authors
- Community testers and issue reporters
