-- gallery-screenshots.lua
-- Features and hotkeys:
--  s                  - save screenshot to images/<video_name>/f########.jpg
--  g                  - toggle gallery on/off
--    Click            - seek to the clicked thumbnail time
--    m                - toggle selection mark on a tile
--    [                - previous page
--    ]                - next page
--    a                - select all   (on current page)
--    u                - unselect all (on current page)
--    d                - delete selected (disk + cache)
--    C (shift + c)    - export contact sheet (current page) > ../exports/<video>_contact_sheet_pXX.png
--    c                - export contact sheets for all pages
--    e                - export CSV (whole video) > ../exports/<video>_gallery.csv
--    x                - export XLSX via Python/xlsxwriter > ../exports/<video>.xlsx
--
-- Notes:
--  - Thumbnails (BGRA) are stored in images/<video>/.gallery_bgra and are rebuilt on cell size change.
--  - Exports (PNG/CSV/XLSX) go to the sibling "exports" directory next to the video (NOT into images/).
--  - The Python exporter script is looked up next to this Lua file (SCRIPT_DIR).

local mp    = require "mp"
local utils = require "mp.utils"
local msg   = require "mp.msg"

-- ===================== User-configurable (Python/XLSX) =====================

-- Set a specific Python interpreter if you want (e.g., a virtualenv). RECOMMENDED.
-- Example (Windows): "C:\\Users\\you\\venv-mpv\\Scripts\\python.exe"
-- Example (Linux): "/home/you/venv-mpv/bin/python"
-- Example (macOS): "/Users/you/venv-mpv/bin/python"
-- Leave as nil to auto-detect ("python3", "python", "py").
local PYTHON_PATH = nil

-- Scale factor for images in XLSX (0.15–0.35 recommended).
-- If Pillow is available, images will be physically resized (smaller XLSX).
-- Otherwise, a visual scale will be applied by Excel (larger XLSX).
local XLSX_IMG_SCALE = 0.2

-- Python exporter script name. The script is expected to sit next to this Lua script.
local XLSX_SCRIPT_NAME = "gallery_xlsx_export.py"

-- ===================== SETTINGS =====================

-- Gallery grid (per page)
local GRID_COLS = 5
local GRID_ROWS = 4
local GRID_GAP  = 12              -- px gap between tiles and edges

-- Safe margins so the gallery does not cover OSD/OSC
local SAFE_MARGIN_TOP    = 100
local SAFE_MARGIN_BOTTOM = 150
local SAFE_MARGIN_LEFT   = 16
local SAFE_MARGIN_RIGHT  = 16

-- Directories
local DIR_IMAGES_BASE = "images"   -- per-video: images/<video_name>/
-- Exports are stored in sibling directory next to the video (not in images/)
-- e.g.: ../exports/video_contact_sheet_p01.png
local DIR_EXPORTS_NAME = "exports"

-- Time labels (via ffmpeg/drawtext)
local ENABLE_TIME_LABELS = true
local LABEL_BOX_ALPHA    = 0.60
local LABEL_MARGIN_X     = 10
local LABEL_MARGIN_Y     = 6
local LABEL_REL_SIZE     = 0.13    -- text size ~ 13% of cell height

-- Selection rendering (overdraw the tile)
-- Draw a red strip across the bottom (same height as the timecode bar) + optional filled square at bottom-left.
local SELECT_BOX_ALPHA   = 0.45
local SELECT_SQUARE      = true

-- Key bindings
local KEY_SCREENSHOT       = "s"
local KEY_TOGGLE_GALLERY   = "g"
local KEY_CLICK            = "MBTN_LEFT"
local KEY_MARK             = "m"
local KEY_SELECT_ALL       = "a"
local KEY_UNSELECT_ALL     = "u"
local KEY_DELETE_SELECTED  = "d"
local KEY_PAGE_NEXT        = "]"
local KEY_PAGE_PREV        = "["
local KEY_EXPORT_CONTACT   = "C"
local KEY_EXPORT_CONTACT_ALL = "c"
local KEY_EXPORT_CSV       = "e"
local KEY_EXPORT_XLSX      = "x"

-- ===========================================================================

-- Resolve absolute path of this Lua script directory (to locate the Python script).
local function get_script_dir()
    -- Prefer newer API if available
    if type(mp.get_script_file) == "function" then
        local sf = mp.get_script_file()
        if sf and sf ~= "" then
            local d, _ = utils.split_path(sf)
            if d and d ~= "" then
                return d
            end
        end
    end

    -- Fallback: use Lua debug info to get current file path (@/path/to/script.lua)
    local info = debug.getinfo(1, "S")
    local src = info and info.source or nil
    if src and src:sub(1, 1) == "@" then
        local path = src:sub(2)
        local d, _ = utils.split_path(path)
        if d and d ~= "" then
            return d
        end
    end

    -- Last resort: working directory
    return mp.get_property("working-directory") or "."
end

local SCRIPT_DIR = get_script_dir()

-- Optional: hardcode full script path if needed; otherwise it uses SCRIPT_DIR/XLSX_SCRIPT_NAME
local XLSX_SCRIPT_PATH = nil

-- ===================== State =====================

local state = {
    active = false,         -- gallery visible?
    basedir = nil,          -- video directory
    video_name = nil,       -- name without extension
    images_dir = nil,       -- images/<video_name>
    bgra_dir   = nil,       -- images/<video_name>/.gallery_bgra
    exports_dir = nil,      -- ../exports (sibling to video)

    files_all = {},         -- all screenshots (filenames) sorted
    page = 1,               -- current page (1-index)
    rects = {},             -- tiles on current page
    overlays = {},          -- active overlay IDs

    last_cell_w = nil,      -- for cache invalidation on resize
    last_cell_h = nil,

    -- per-video memory of last visited page
    last_page_by_video = {},
}

-- ===================== Utilities =====================

local function is_windows()
    return package.config:sub(1,1) == "\\"
end

-- Safe join that accepts 2+ segments and folds them via utils.join_path
local function join(a, b, ...)
    -- utils.join_path needs 2 arguments
    local p = utils.join_path(a or "", b or "")
    local n = select('#', ...)
    for i = 1, n do
        local seg = select(i, ...)
        if seg and seg ~= "" then
            p = utils.join_path(p, seg)
        end
    end
    return p
end

local function strip_ext(name)
    return (name:gsub("%.[^%.]+$", ""))
end

local function basename(path)
    local _d, f = utils.split_path(path)
    return f
end

local function ensure_dir(path)
    local info = utils.file_info(path)
    if info and info.is_dir then return true end
    if info then return false end
    local args
    if is_windows() then args = {"cmd", "/c", "mkdir", path}
    else args = {"mkdir", "-p", path} end
    local res = utils.subprocess({args = args, cancellable = false})
    return res and res.status == 0
end

local function file_exists(path)
    local f = io.open(path, "rb")
    if f then f:close() return true end
    return false
end

local function remove_file(path)
    local ok, err = os.remove(path)
    if not ok then
        msg.warn("Failed to delete: " .. tostring(path) .. " (" .. tostring(err) .. ")")
    end
end

local function readdir_files(dir)
    return utils.readdir(dir, "files") or {}
end

local function list_images_sorted(dir)
    local t = {}
    for _, name in ipairs(readdir_files(dir)) do
        local lower = name:lower()
        if lower:match("%.jpe?g$") or lower:match("%.png$") then
            table.insert(t, name)
        end
    end
    table.sort(t)
    return t
end

local function get_osd_dims()
    local od = mp.get_property_native("osd-dimensions")
    if od then
        local ml = od.ml or 0
        local mt = od.mt or 0
        local vw = (od.w or 0) - ml - (od.mr or 0)
        local vh = (od.h or 0) - mt - (od.mb or 0)
        return ml, mt, vw, vh
    end
    local w, h = mp.get_osd_size()
    return 0, 0, w, h
end

local function parse_frame_from_filename(name)
    local f = name:match("^f(%d%d%d%d%d%d%d%d)%.jpe?g$")
    if f then return tonumber(f) end
    f = name:match("^f(%d%d%d%d%d%d%d%d)%.png$")
    if f then return tonumber(f) end
    f = name:match("[^%d](%d%d%d%d%d%d%d%d)[^%d]*%.jpe?g$")
    if f then return tonumber(f) end
    return nil
end

local function get_fps()
    return mp.get_property_native("estimated-vf-fps")
        or mp.get_property_native("fps")
        or mp.get_property_native("container-fps")
        or 0
end

local function seconds_to_hhmmss_floor(sec)
    if not sec or sec < 0 then return "00:00:00" end
    local s = math.floor(sec) -- always floor to seconds
    local h = math.floor(s / 3600)
    local m = math.floor((s % 3600) / 60)
    local r = s % 60
    return string.format("%02d:%02d:%02d", h, m, r)
end

local function find_executable(name)
    local PATH = os.getenv("PATH") or ""
    local sep = is_windows() and ";" or ":"
    for dir in string.gmatch(PATH, "([^"..sep.."]+)") do
        local cand = join(dir, name)
        if is_windows() then
            if file_exists(cand) or file_exists(cand .. ".exe") or file_exists(cand .. ".bat") then
                return cand
            end
        else
            if file_exists(cand) then return cand end
        end
    end
    return nil
end

-- Helper: find Python in PATH (used if PYTHON_PATH is nil)
local function find_python()
    local candidates = {"python3", "python", "py"}
    for _, n in ipairs(candidates) do
        local p = find_executable(n)
        if p then return p end
    end
    return nil
end

-- ===================== Binaries =====================

local BIN_FFMPEG = nil
local BIN_MPV    = nil

local function init_binaries()
    if BIN_FFMPEG == nil then BIN_FFMPEG = find_executable("ffmpeg") end
    if BIN_MPV == nil then BIN_MPV = find_executable("mpv") or "mpv" end
end

local function ensure_binaries_ready()
    init_binaries()
end

-- ===================== Per-video paths =====================

local function ensure_paths()
    local path = mp.get_property("path")
    if not path then return false end
    local dir, file = utils.split_path(path)
    if not dir or dir == "" then dir = mp.get_property("working-directory") or "." end
    local video_name = strip_ext(file or "video")
    local images_dir = join(dir, DIR_IMAGES_BASE, video_name)
    local bgra_dir   = join(images_dir, ".gallery_bgra")
    local exports_dir = join(dir, DIR_EXPORTS_NAME)

    ensure_dir(images_dir)
    ensure_dir(bgra_dir)
    ensure_dir(exports_dir)

    state.basedir    = dir
    state.video_name = video_name
    state.images_dir = images_dir
    state.bgra_dir   = bgra_dir
    state.exports_dir= exports_dir
    return true
end

-- ===================== Screenshot ("s") =====================

local function take_screenshot()
    if not ensure_paths() then
        mp.osd_message("Video path unavailable", 2)
        return
    end
    local frame = mp.get_property_native("estimated-frame-number") or 0
    local frame_str = string.format("%08d", frame)
    local filename = "f" .. frame_str .. ".jpg"
    local path = join(state.images_dir, filename)

    local ok, err = pcall(function()
        mp.commandv("screenshot-to-file", path, "video")
    end)

    if ok then
        msg.info("Saved screenshot: " .. path)
        mp.osd_message("Saved: " .. filename, 1.2)
    else
        msg.error("Failed to save screenshot: " .. tostring(err))
        mp.osd_message("Save error: " .. filename, 3.0)
    end
end

-- ===================== BGRA generation and cache =====================

-- BGRA names are fixed (no dimensions) → on cell size change, purge entire .gallery_bgra
local function bgra_name_for(src_name, selected)
    local base = strip_ext(src_name)
    return string.format("%s_%s.bgra", base, selected and "SEL" or "L0")
end

local function generate_bgra_with_ffmpeg(src, dst, w, h, label_text, selected)
    local scale_pad = string.format(
        "scale=%d:%d:force_original_aspect_ratio=decrease,pad=%d:%d:(ow-iw)/2:(oh-ih)/2:color=black",
        w, h, w, h
    )
    local fontsize = math.max(10, math.floor(h * LABEL_REL_SIZE))
    local box_h = fontsize + 2 * LABEL_MARGIN_Y
    local vf = scale_pad
    local ops = {}

    if label_text and label_text ~= "" then
        local box = string.format("drawbox=x=0:y=ih-%d:w=iw:h=%d:color=black@%.2f:t=fill",
                                  box_h, box_h, LABEL_BOX_ALPHA)
        local esc_text = label_text:gsub(":", "\\:")
        local text = string.format("drawtext=text='%s':x=w-tw-%d:y=h-%d+(%d-th)/2:fontcolor=white:fontsize=%d:shadowx=1:shadowy=1:borderw=1",
                                   esc_text, LABEL_MARGIN_X, box_h, box_h, fontsize)
        table.insert(ops, box)
        table.insert(ops, text)
    end

    if selected then
        -- Red strip same height as the timecode box (bottom)
        local sel_box = string.format("drawbox=x=0:y=ih-%d:w=iw:h=%d:color=red@%.2f:t=fill",
                                      box_h, box_h, SELECT_BOX_ALPHA)
        table.insert(ops, sel_box)
        if SELECT_SQUARE then
            -- Filled square at bottom-left – size box_h
            local sq = string.format("drawbox=x=0:y=ih-%d:w=%d:h=%d:color=red@1.0:t=fill",
                                     box_h, box_h, box_h)
            table.insert(ops, sq)
        end
    end

    if #ops > 0 then
        vf = vf .. "," .. table.concat(ops, ",")
    end

    local args = {
        BIN_FFMPEG, "-hide_banner", "-loglevel", "error", "-y",
        "-i", src,
        "-frames:v", "1",
        "-vf", vf,
        "-pix_fmt", "bgra",
        "-f", "rawvideo",
        dst
    }
    local res = utils.subprocess({args=args, cancellable=false})
    return res and res.status == 0
end

local function generate_bgra_with_mpv(src, dst, w, h)
    local args = {
        BIN_MPV,
        "--msg-level=all=no",
        "--no-config",
        "--no-audio",
        src,
        "--frames=1",
        ("--vf=scale=%d:%d:force_original_aspect_ratio=decrease"):format(w, h),
        "--vf-add=format=bgra",
        "--of=rawvideo",
        "--ovc=rawvideo",
        "--o=" .. dst
    }
    local res = utils.subprocess({args=args, cancellable=false})
    return res and res.status == 0
end

-- Generate (or overwrite) BGRA with labels and optional selection
local function ensure_bgra_thumbnail(src_path, src_name, w, h, label_text, selected)
    ensure_binaries_ready()
    ensure_dir(state.bgra_dir)
    local dst_name = bgra_name_for(src_name, selected)
    local dst_path = join(state.bgra_dir, dst_name)

    if file_exists(dst_path) then
        return dst_path, true
    end

    local ok = false
    if BIN_FFMPEG then
        ok = generate_bgra_with_ffmpeg(src_path, dst_path, w, h, label_text, selected)
    end
    if not ok then
        -- Fallback without labels and selection
        ok = generate_bgra_with_mpv(src_path, dst_path, w, h)
    end
    if not ok then
        msg.error("Failed to generate BGRA: " .. tostring(src_name))
        return nil, false
    end
    return dst_path, false
end

-- Remove all BGRA variants for a given source (L0 and SEL)
local function purge_bgra_for_source(src_name)
    local base = strip_ext(src_name)
    for _, name in ipairs(readdir_files(state.bgra_dir)) do
        if name:match("^" .. base:gsub("%W","%%%0") .. "_(L0|SEL)%.bgra$") then
            remove_file(join(state.bgra_dir, name))
        end
    end
end

-- Purge whole BGRA cache (on cell size change)
local function purge_bgra_all()
    for _, name in ipairs(readdir_files(state.bgra_dir)) do
        if name:lower():match("%.bgra$") then
            remove_file(join(state.bgra_dir, name))
        end
    end
end

-- ===================== File management =====================

local function refresh_files_list()
    state.files_all = list_images_sorted(state.images_dir)
end

-- ===================== Geometry and drawing =====================

local function clear_overlays()
    for _, id in ipairs(state.overlays) do
        pcall(function() mp.command_native({"overlay-remove", id}) end)
    end
    state.overlays = {}
    state.rects = {}
end

local function compute_cell_size(vw, vh)
    -- Reserve safe margins
    local avail_w = math.max(0, vw - SAFE_MARGIN_LEFT - SAFE_MARGIN_RIGHT)
    local avail_h = math.max(0, vh - SAFE_MARGIN_TOP - SAFE_MARGIN_BOTTOM)
    local cols, rows, gap = GRID_COLS, GRID_ROWS, GRID_GAP
    local cell_w = math.max(2, math.floor((avail_w - (cols + 1) * gap) / cols))
    local cell_h = math.max(2, math.floor((avail_h - (rows + 1) * gap) / rows))
    return cell_w, cell_h
end

local function build_and_draw_page()
    clear_overlays()

    local ml, mt, vw, vh = get_osd_dims()
    local cell_w, cell_h = compute_cell_size(vw, vh)

    -- On cell size change: purge cache and rewrite
    if state.last_cell_w ~= cell_w or state.last_cell_h ~= cell_h then
        purge_bgra_all()
        state.last_cell_w = cell_w
        state.last_cell_h = cell_h
    end

    -- Page slicing
    local page_size = GRID_COLS * GRID_ROWS
    local total = #state.files_all
    if total == 0 then
        mp.osd_message("No images in " .. DIR_IMAGES_BASE .. "/" .. (state.video_name or "?"), 2.0)
        return false
    end
    local max_page = math.max(1, math.ceil(total / page_size))
    if state.page > max_page then state.page = max_page end
    if state.page < 1 then state.page = 1 end

    -- remember current page for this video
    state.last_page_by_video[state.video_name or ""] = state.page

    local start_idx = (state.page - 1) * page_size + 1
    local end_idx = math.min(start_idx + page_size - 1, total)

    local fps = tonumber(get_fps()) or 0

    local draw_count = 0
    for i = start_idx, end_idx do
        local visible_idx = i - start_idx + 1
        local idx0 = visible_idx - 1
        local r = math.floor(idx0 / GRID_COLS)
        local c = idx0 % GRID_COLS

        local x_rel = SAFE_MARGIN_LEFT + GRID_GAP + c * (cell_w + GRID_GAP)
        local y_rel = SAFE_MARGIN_TOP  + GRID_GAP + r * (cell_h + GRID_GAP)

        local x_abs = ml + x_rel
        local y_abs = mt + y_rel

        local src_name = state.files_all[i]
        local src_path = join(state.images_dir, src_name)
        local frame = parse_frame_from_filename(src_name)
        local label_text = nil
        if ENABLE_TIME_LABELS and fps > 0 and frame then
            label_text = seconds_to_hhmmss_floor(frame / fps)
        end

        local selected = false

        local bgra_path, from_cache = ensure_bgra_thumbnail(src_path, src_name, cell_w, cell_h, label_text, selected)
        if not bgra_path then
            mp.osd_message("Gallery: thumbnail generation failed", 2.0)
            return false
        end

        local id = 1 + (visible_idx - 1)  -- ID 1..page_size (mpv overlay limit 0..63)
        if id > 63 then
            msg.error("Overlay ID limit (0..63) exceeded")
            return false
        end

        local ok, err = pcall(function()
            mp.command_native({"overlay-add", id, x_rel, y_rel, bgra_path, 0, "bgra", cell_w, cell_h, cell_w * 4})
        end)
        if not ok then
            msg.error("overlay-add failed: " .. tostring(err))
            mp.osd_message("Gallery: overlay-add failed", 2.0)
            return false
        end

        table.insert(state.overlays, id)
        table.insert(state.rects, {
            x_abs=x_abs, y_abs=y_abs, x_rel=x_rel, y_rel=y_rel,
            w=cell_w, h=cell_h,
            path=src_path, name=src_name,
            frame=frame, selected=false, id=id
        })
        draw_count = draw_count + 1
    end

    msg.info(string.format("Gallery: page %d, tiles %d, cell %dx%d, gap %d",
        state.page, draw_count, state.last_cell_w, state.last_cell_h, GRID_GAP))
    return true
end

-- ===================== Clicks and interactions =====================

local function rect_at(x, y)
    for idx, r in ipairs(state.rects) do
        if x >= r.x_abs and x < r.x_abs + r.w and y >= r.y_abs and y < r.y_abs + r.h then
            return idx, r
        end
    end
    return nil, nil
end

local function redraw_tile(idx)
    local r = state.rects[idx]
    if not r then return end
    local fps = tonumber(get_fps()) or 0
    local label_text = nil
    if ENABLE_TIME_LABELS and fps > 0 and r.frame then
        label_text = seconds_to_hhmmss_floor(r.frame / fps)
    end
    local bgra_path = ensure_bgra_thumbnail(r.path, r.name, r.w, r.h, label_text, r.selected)
    if not bgra_path then return end
    pcall(function()
        mp.command_native({"overlay-add", r.id, r.x_rel, r.y_rel, bgra_path, 0, "bgra", r.w, r.h, r.w * 4})
    end)
end

local function on_click_normal()
    if not state.active then return end
    local pos = mp.get_property_native("mouse-pos")
    if not pos or not pos.x or not pos.y then return end
    local idx, r = rect_at(pos.x, pos.y)
    if not r then return end

    local fps = tonumber(get_fps()) or 0
    local target_time = nil
    if r.frame and fps > 0 then target_time = r.frame / fps end

    -- Hide gallery before seeking (keep last page in map)
    state.active = false
    clear_overlays()
    mp.osd_message("Gallery hidden", 0.6)
    mp.remove_key_binding("gallery-click")
    mp.remove_key_binding("gallery-ctrl-click")
    mp.remove_key_binding("gallery-select-all")
    mp.remove_key_binding("gallery-unselect-all")
    mp.remove_key_binding("gallery-delete-selected")
    mp.remove_key_binding("gallery-page-next")
    mp.remove_key_binding("gallery-page-prev")
    mp.remove_key_binding("gallery-export-contact")
    mp.remove_key_binding("gallery-export-contact-all")
    mp.remove_key_binding("gallery-export-csv")
    mp.remove_key_binding("gallery-export-xlsx")

    if target_time then
        mp.commandv("seek", string.format("%.6f", target_time), "absolute", "exact")
    else
        mp.osd_message("Cannot determine time from filename: " .. (r.path or ""), 2.5)
    end
end

local function on_click_ctrl()
    if not state.active then return end
    local pos = mp.get_property_native("mouse-pos")
    if not pos or not pos.x or not pos.y then return end
    local idx, r = rect_at(pos.x, pos.y)
    if not r then return end

    r.selected = not r.selected
    redraw_tile(idx)
end

local function select_all_visible(flag)
    for i=1,#state.rects do
        state.rects[i].selected = flag
        redraw_tile(i)
    end
end

local function delete_selected_now()
    local to_delete_names = {}
    for _, r in ipairs(state.rects) do
        if r.selected then table.insert(to_delete_names, r.name) end
    end

    if #to_delete_names == 0 then
        mp.osd_message("Nothing selected", 1.0)
        return
    end

    for _, name in ipairs(to_delete_names) do
        purge_bgra_for_source(name)
        remove_file(join(state.images_dir, name))
    end

    refresh_files_list()
    build_and_draw_page()
    mp.osd_message(string.format("Deleted %d screenshots", #to_delete_names), 1.2)
end

local function next_page()
    state.page = state.page + 1
    build_and_draw_page()
end

local function prev_page()
    state.page = math.max(1, state.page - 1)
    build_and_draw_page()
end

-- ===================== Exports =====================

-- Export current page to PNG (contact sheet)
local function export_contact_sheet_current_page()
    if #state.rects == 0 then
        mp.osd_message("No thumbnails on this page", 1.0)
        return false
    end
    ensure_binaries_ready()

    local out = join(state.exports_dir, string.format("%s_contact_sheet_p%02d.png", state.video_name, state.page))

    -- 1) Convert each tile (current state – with label and optional selection) into PNG
    local tmp_dir = join(state.bgra_dir, ".cs_tmp")
    ensure_dir(tmp_dir)
    -- cleanup previous temp pngs
    for _, name in ipairs(readdir_files(tmp_dir)) do
        remove_file(join(tmp_dir, name))
    end

    for i, r in ipairs(state.rects) do
        local fps = tonumber(get_fps()) or 0
        local label_text = nil
        if ENABLE_TIME_LABELS and fps > 0 and r.frame then
            label_text = seconds_to_hhmmss_floor(r.frame / fps)
        end
        local bgra_path = ensure_bgra_thumbnail(r.path, r.name, r.w, r.h, label_text, r.selected)
        if not bgra_path then
            mp.osd_message("Export failed (tile BGRA)", 2.0)
            return false
        end
        local png_out = join(tmp_dir, string.format("cs_%03d.png", i))
        local args = {
            BIN_FFMPEG, "-hide_banner", "-loglevel", "error", "-y",
            "-f", "rawvideo", "-pix_fmt", "bgra", "-s", string.format("%dx%d", r.w, r.h),
            "-i", bgra_path,
            "-frames:v", "1", png_out
        }
        local res = utils.subprocess({args=args, cancellable=false})
        if not (res and res.status == 0) then
            mp.osd_message("Export failed (BGRA→PNG)", 2.0)
            return false
        end
    end

    -- 2) Compose PNG tiles into a grid using tile filter
    local cols, rows = GRID_COLS, GRID_ROWS
    local vf = string.format("tile=%dx%d:padding=%d:margin=%d:color=black",
                             cols, rows, GRID_GAP, GRID_GAP)

    local args = {
        BIN_FFMPEG, "-hide_banner", "-loglevel", "error", "-y",
        "-framerate", "30",
        "-start_number", "1",
        "-i", join(tmp_dir, "cs_%03d.png"),
        "-frames:v", "1",
        "-vf", vf,
        out
    }
    local res = utils.subprocess({args=args, cancellable=false})
    if not (res and res.status == 0) then
        mp.osd_message("Export failed (tile)", 2.0)
        return false
    end

    mp.osd_message("Exported contact sheet: " .. basename(out), 2.0)
    return true
end

-- Export contact sheets for all pages
local function export_contact_sheet_all_pages()
    if not ensure_paths() then
        mp.osd_message("Video path unavailable", 2)
        return
    end
    refresh_files_list()
    local total = #state.files_all
    if total == 0 then
        mp.osd_message("No screenshots to export", 1.0)
        return
    end

    local page_size = GRID_COLS * GRID_ROWS
    local max_page = math.max(1, math.ceil(total / page_size))

    local original_page = state.page
    local ok_count = 0
    for p = 1, max_page do
        state.page = p
        if build_and_draw_page() then
            if export_contact_sheet_current_page() then
                ok_count = ok_count + 1
            end
        end
    end
    state.page = original_page
    build_and_draw_page()

    mp.osd_message(string.format("Contact sheet export: %d pages", ok_count), 2.5)
end

-- Export CSV (filename, timecode, seconds, frame)
local function export_csv()
    if not ensure_paths() then
        mp.osd_message("Video path unavailable", 2)
        return
    end
    refresh_files_list()
    if #state.files_all == 0 then
        mp.osd_message("No screenshots to export", 1.0)
        return
    end

    local fps = tonumber(get_fps()) or 0
    local out = join(state.exports_dir, string.format("%s_gallery.csv", state.video_name))
    local f, err = io.open(out, "wb")
    if not f then
        mp.osd_message("CSV: cannot write: " .. tostring(err), 2.5)
        return
    end

    f:write("filename,timecode,seconds,frame\n")
    for _, name in ipairs(state.files_all) do
        local frame = parse_frame_from_filename(name)
        local sec = ""
        local tc = ""
        if frame and fps > 0 then
            sec = string.format("%.3f", frame / fps)
            tc = seconds_to_hhmmss_floor(frame / fps)
        end
        f:write(string.format("%s,%s,%s,%s\n", name, tc, sec, tostring(frame or "")))
    end
    f:close()
    mp.osd_message("CSV export done: " .. basename(out), 2.0)
end

-- Export XLSX via Python/xlsxwriter
local function export_xlsx()
    if not ensure_paths() then
        mp.osd_message("Video path unavailable", 2)
        return
    end
    refresh_files_list()
    if #state.files_all == 0 then
        mp.osd_message("No screenshots to export", 1.0)
        return
    end

    local py = PYTHON_PATH or find_python()
    if not py then
        mp.osd_message("Python not found (set PYTHON_PATH or add python3/python to PATH)", 3.0)
        return
    end

    local script_path = XLSX_SCRIPT_PATH or join(SCRIPT_DIR, XLSX_SCRIPT_NAME)
    if not file_exists(script_path) then
        mp.osd_message("Missing Python script: " .. script_path, 3.5)
        return
    end

    local fps = tonumber(get_fps()) or 0
    local out = join(state.exports_dir, string.format("%s.xlsx", state.video_name))

    -- Build arguments for the Python exporter
    local args = {
        py,
        script_path,
        "--images-dir", state.images_dir,
        "--out", out,
        "--fps", tostring(fps),
        "--scale", tostring(XLSX_IMG_SCALE),
        "--video-name", state.video_name or "",
        "--resize", "physical",   -- try physical resize (requires Pillow)
        "--center", "1"           -- center image and text
    }

    local res = utils.subprocess({ args = args, cancellable = false })
    if not res or res.status ~= 0 then
        local err = (res and res.stderr) or "unknown error"
        msg.error("XLSX export failed: " .. tostring(err))
        mp.osd_message("XLSX export failed", 2.5)
        return
    end

    mp.osd_message("XLSX export done: " .. basename(out), 2.5)
end

-- ===================== Key bindings and toggle =====================

local function bind_gallery_keys()
    mp.add_forced_key_binding(KEY_CLICK,              "gallery-click",              on_click_normal)
    mp.add_forced_key_binding(KEY_MARK      ,         "gallery-ctrl-click",         on_click_ctrl)
    mp.add_forced_key_binding(KEY_SELECT_ALL,         "gallery-select-all",         function() select_all_visible(true) end)
    mp.add_forced_key_binding(KEY_UNSELECT_ALL,       "gallery-unselect-all",       function() select_all_visible(false) end)
    mp.add_forced_key_binding(KEY_DELETE_SELECTED,    "gallery-delete-selected",    delete_selected_now)
    mp.add_forced_key_binding(KEY_PAGE_NEXT,          "gallery-page-next",          next_page)
    mp.add_forced_key_binding(KEY_PAGE_PREV,          "gallery-page-prev",          prev_page)
    mp.add_forced_key_binding(KEY_EXPORT_CONTACT,     "gallery-export-contact",     export_contact_sheet_current_page)
    mp.add_forced_key_binding(KEY_EXPORT_CONTACT_ALL, "gallery-export-contact-all", export_contact_sheet_all_pages)
    mp.add_forced_key_binding(KEY_EXPORT_CSV,         "gallery-export-csv",         export_csv)
    mp.add_forced_key_binding(KEY_EXPORT_XLSX,        "gallery-export-xlsx",        export_xlsx)
end

local function unbind_gallery_keys()
    mp.remove_key_binding("gallery-click")
    mp.remove_key_binding("gallery-ctrl-click")
    mp.remove_key_binding("gallery-select-all")
    mp.remove_key_binding("gallery-unselect-all")
    mp.remove_key_binding("gallery-delete-selected")
    mp.remove_key_binding("gallery-page-next")
    mp.remove_key_binding("gallery-page-prev")
    mp.remove_key_binding("gallery-export-contact")
    mp.remove_key_binding("gallery-export-contact-all")
    mp.remove_key_binding("gallery-export-csv")
    mp.remove_key_binding("gallery-export-xlsx")
end

local function toggle_gallery()
    if not ensure_paths() then
        mp.osd_message("Video path unavailable", 2)
        return
    end

    if state.active then
        state.active = false
        clear_overlays()
        unbind_gallery_keys()
        mp.osd_message("[Gallery hidden]", 0.8)
    else
        refresh_files_list()
        -- restore last visited page for this video
        local last = state.last_page_by_video[state.video_name or ""]
        state.page = last or 1
        if build_and_draw_page() then
            state.active = true
            bind_gallery_keys()
            mp.osd_message("[Gallery] Goto: click | Mark: m | Select all: a | Unselect all: u | Delete: d | Prev. page: [ | Next page: ] | Export: c/e/x", 5)

        end
    end
end

-- ===================== Lifecycle and observers =====================

local function on_file_loaded()
    -- Reset per-video paths and state (keep last_page_by_video)
    state.basedir    = nil
    state.video_name = nil
    state.images_dir = nil
    state.bgra_dir   = nil
    state.exports_dir= nil

    state.files_all  = {}
    state.page       = 1
    state.rects      = {}
    state.overlays   = {}
    state.last_cell_w= nil
    state.last_cell_h= nil

    ensure_paths()

    if state.active then
        state.active = false
        clear_overlays()
        unbind_gallery_keys()
    end
end

mp.register_event("file-loaded", on_file_loaded)

-- Redraw gallery on OSD size change
mp.observe_property("osd-dimensions", "native", function()
    if state.active then
        build_and_draw_page()
    end
end)

-- Global keys
mp.add_forced_key_binding(KEY_SCREENSHOT,     "save-screenshot-custom", take_screenshot)
mp.add_forced_key_binding(KEY_TOGGLE_GALLERY, "toggle-gallery",         toggle_gallery)
