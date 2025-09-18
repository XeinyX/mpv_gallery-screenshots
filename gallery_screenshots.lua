-- gallery-screenshots.lua
-- Features and hotkeys:
--  s                  - save screenshot to images/<video_name>/f########.jpg
--  g                  - toggle gallery on/off
--    Click            - seek to the clicked thumbnail time
--    m                - toggle selection mark on a tile
--    [                - previous page
--    ]                - next page
--    a                - select/unselect all (toggle, current page)
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
-- Example (Linux)  : "/home/you/venv-mpv/bin/python"
-- Example (macOS)  : "/Users/you/venv-mpv/bin/python"
-- Leave as nil to auto-detect ("python3", "python", "py").
local PYTHON_PATH = nil

-- Scale factor for images in XLSX (0.15 - 0.35 recommended).
-- If Pillow is available, images will be physically resized (smaller XLSX).
-- Otherwise, a visual scale will be applied by Excel (larger XLSX).
local XLSX_IMG_SCALE = 0.2

-- Python exporter script name. The script is expected to sit next to this Lua script.
local XLSX_SCRIPT_NAME = "gallery_xlsx_export.py"

-- ===================== SETTINGS =====================

-- Gallery grid (per page)
local GRID_COLS = 7
local GRID_ROWS = 5
local GRID_GAP  = 12              -- px gap between tiles and edges

-- Safe margins as fractions of viewport (vw, vh)
-- TIP: ~0.09 ≈ 100 px on 1080p; ~0.14 ≈ 150 px na 1080p
local SAFE_MARGIN_TOP_REL    = 0.10    -- fraction of vh
local SAFE_MARGIN_BOTTOM_REL = 0.14    -- fraction of vh
local SAFE_MARGIN_LEFT_REL   = 0.02    -- fraction of vw
local SAFE_MARGIN_RIGHT_REL  = 0.02    -- fraction of vw

-- Keep pixel gap
local GRID_GAP = 12

-- Screenshot output sizing
local SCREENSHOT_SCALE      = 1.0       -- 0.1 .. 1.0, 0.5 means 50% of source
local SCREENSHOT_MIN_WIDTH  = 540       -- lower bound in px (0 disables)
local SCREENSHOT_MIN_HEIGHT = 540       -- lower bound in px (0 disables)

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
local LABEL_REL_SIZE     = 0.16

-- Simple delete mode: if nothing is selected, delete the tile under the mouse cursor
local ENABLE_SIMPLE_DELETE = true

-- Selection rendering (overdraw the tile)
-- Draw a red strip across the bottom (same height as the timecode bar) + optional filled square at bottom-left.
local SELECT_BOX_ALPHA   = 0.45

-- Key bindings
local KEY_SCREENSHOT         = "s"
local KEY_TOGGLE_GALLERY     = "g"
local KEY_CLICK              = "MBTN_LEFT"
local KEY_MARK               = "m"
local KEY_SELECT_ALL         = "a"
local KEY_DELETE_SELECTED    = "d"
local KEY_PAGE_NEXT          = "]"
local KEY_PAGE_PREV          = "["
local KEY_EXPORT_CONTACT     = "C"
local KEY_EXPORT_CONTACT_ALL = "c"
local KEY_EXPORT_CSV         = "e"
local KEY_EXPORT_XLSX        = "x"

-- ===========================================================================

-- Resolve absolute path of this Lua script directory (to locate the Python script).
local function get_script_dir()
    if type(mp.get_script_file) == "function" then
        local sf = mp.get_script_file()
        if sf and sf ~= "" then
            local d, _ = utils.split_path(sf)
            if d and d ~= "" then return d end
        end
    end
    local info = debug.getinfo(1, "S")
    local src = info and info.source or nil
    if src and src:sub(1, 1) == "@" then
        local path = src:sub(2)
        local d, _ = utils.split_path(path)
        if d and d ~= "" then return d end
    end
    return mp.get_property("working-directory") or "."
end

local SCRIPT_DIR = get_script_dir()
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
    last_page_by_video = {}, -- per-video memory of last visited page

    visible_total = 0,
    visible_missing = 0,
    visible_pending = {},
}

-- ===================== Forward declarations (to avoid nil globals) =====================

local ensure_bgra_thumbnail
local purge_bgra_all
local purge_bgra_for_source
local generate_bgra_auto_async
local generate_bgra_with_ffmpeg
local generate_bgra_with_mpv
local redraw_tile_by_name 
local start_prewarm_all_current_size
local stop_prewarm
local run_subprocess_async
local schedule_thumbnail_regen
local show_page_indicator

-- State for cooperative prewarm
local prewarm_queue = {}
local prewarm_running = false

-- Optional: debounce timer for regen
local regen_timer = nil
local regen_in_progress = false

-- ===================== Utilities =====================

local function is_windows()
    return package.config:sub(1,1) == "\\"
end

local function join(a, b, ...)
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

-- Ensure parent directory for a path exists
local function ensure_parent_dir(path)
    local dir, _ = utils.split_path(path)
    if not dir or dir == "" then return true end
    return ensure_dir(dir)
end

-- Normalize a filesystem path using mpv; falls back to original on failure
local function normpath(p)
    if not p or p == "" then return p end
    local ok, out = pcall(function()
        return mp.command_native({"normalize-path", p})
    end)
    if ok and type(out) == "string" and out ~= "" then
        return out
    end
    return p
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

local function safe_margins_px(vw, vh)
    local mt = math.floor(vh * SAFE_MARGIN_TOP_REL    + 0.5)
    local mb = math.floor(vh * SAFE_MARGIN_BOTTOM_REL + 0.5)
    local ml = math.floor(vw * SAFE_MARGIN_LEFT_REL   + 0.5)
    local mr = math.floor(vw * SAFE_MARGIN_RIGHT_REL  + 0.5)
    return mt, mb, ml, mr
end

-- Compute target screenshot size (keeps AR, never upscales)
local function screenshot_target_dims(src_w, src_h)
    if not src_w or not src_h or src_w <= 0 or src_h <= 0 then return nil end

    -- clamp scale to [0.01, 1.0]
    local base = math.max(0.01, math.min(1.0, SCREENSHOT_SCALE or 1.0))

    -- requirements to meet minima (0 means "no minimum")
    local need_w = (SCREENSHOT_MIN_WIDTH  or 0) / src_w
    local need_h = (SCREENSHOT_MIN_HEIGHT or 0) / src_h

    -- pick the largest requirement; but never upscale above 1.0
    local scale = math.max(base, need_w, need_h)
    scale = math.min(scale, 1.0)

    local tw = math.max(1, math.floor(src_w * scale + 0.5))
    local th = math.max(1, math.floor(src_h * scale + 0.5))
    return tw, th, scale
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
    local s = math.floor(sec)
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

local function find_python()
    local candidates = {"python3", "python", "py"}
    for _, n in ipairs(candidates) do
        local p = find_executable(n)
        if p then return p end
    end
    return nil
end

local function update_generation_status()
    if not state.active then return end
    local total = state.visible_total or 0
    local missing = state.visible_missing or 0
    if total <= 0 then return end
    if missing > 0 then
        local done = total - missing
        mp.osd_message(string.format("Generating gallery... %d/%d", done, total), 0.8)
    else
        show_page_indicator()
    end
end

local function mark_tile_done_if_visible(name)
    if not state.active then return end
    if state.visible_pending and state.visible_pending[name] then
        state.visible_pending[name] = nil
        if (state.visible_missing or 0) > 0 then
            state.visible_missing = state.visible_missing - 1
        end
        update_generation_status()
    end
end

-- Atomically replace target with tmp (best-effort on Windows)
local function atomic_replace(tmp, final)
    if file_exists(final) then pcall(os.remove, final) end
    local ok, err = os.rename(tmp, final)
    if not ok then
        -- As last resort, copy bytes (rarely needed)
        local rf = io.open(tmp, "rb")
        if not rf then return false, "rename failed: " .. tostring(err) end
        local wf = io.open(final, "wb")
        if not wf then rf:close(); return false, "rename+open failed: " .. tostring(err) end
        local data = rf:read("*all")
        wf:write(data or "")
        rf:close(); wf:close()
        pcall(os.remove, tmp)
        return true
    end
    return true
end

-- Remove any queued prewarm job for the given source name
local function cancel_prewarm_for(name)
    if not prewarm_queue or #prewarm_queue == 0 then return end
    local out = {}
    for _, it in ipairs(prewarm_queue) do
        if it.name ~= name then
            table.insert(out, it)
        end
    end
    prewarm_queue = out
end

-- ===================== Binaries =====================

local BIN_FFMPEG = nil
local BIN_MPV    = nil

local function init_binaries()
    if BIN_FFMPEG == nil then BIN_FFMPEG = find_executable("ffmpeg") end
    if BIN_MPV == nil then BIN_MPV = find_executable("mpv") end
end

local function ensure_binaries_ready()
    init_binaries()
end

-- ffmpeg presence check
local function has_ffmpeg()
    return BIN_FFMPEG ~= nil and BIN_FFMPEG ~= ""
end

-- Inform user that ffmpeg is required (or suggest Pillow)
local function inform_contact_sheet_unavailable()
    mp.osd_message("Contact sheet export requires ffmpeg. Not available. Install ffmpeg or enable Python/Pillow mode.", 3.0)
end

-- ===================== Per-video paths (LAZY: nothing is created) ======================

local function ensure_paths()
    local path = mp.get_property("path")
    if not path then return false end

    local dir, file = utils.split_path(path)
    if not dir or dir == "" then
        dir = mp.get_property("working-directory") or "."
    end

    local video_name = strip_ext(file or "video")

    -- Normalize all derived paths (Windows-safe, but harmless on other OSes)
    local basedir     = normpath(dir)
    local images_dir  = normpath(join(basedir, DIR_IMAGES_BASE, video_name))
    local bgra_dir    = normpath(join(images_dir, ".gallery_bgra"))
    local exports_dir = normpath(join(basedir, DIR_EXPORTS_NAME))

    state.basedir     = basedir
    state.video_name  = video_name
    state.images_dir  = images_dir
    state.bgra_dir    = bgra_dir
    state.exports_dir = exports_dir
    return true
end

-- ===================== BGRA generation and cache =====================

--- Expected BGRA size
local function expected_bgra_size(w, h)
    return w * h * 4
end

local function file_size(path)
    local info = utils.file_info(path)
    return info and info.size or nil
end

-- Common filter chain builder for ffmpeg/mpv
local function build_vf_chain(w, h, label_text, selected)
    local fontsize = math.max(10, math.floor(math.min(h, 1.5 * w) * LABEL_REL_SIZE))
    local box_h    = fontsize + 2 * LABEL_MARGIN_Y

    local parts = {
        string.format("scale=%d:%d:force_original_aspect_ratio=decrease", w, h),
        string.format("pad=%d:%d:(ow-iw)/2:(oh-ih)/2:color=black", w, h)
    }

    if ENABLE_TIME_LABELS and label_text and label_text ~= "" then
        local box = string.format("drawbox=x=0:y=ih-%d:w=iw:h=%d:color=black@%.2f:t=fill",
                                  box_h, box_h, LABEL_BOX_ALPHA)
        table.insert(parts, box)
        local text = string.format("drawtext=text=%s:x=w-tw-%d:y=h-%d+(%d-th)/2:fontcolor=white:fontsize=%d:shadowx=1:shadowy=1:borderw=1",
                                   label_text:gsub(":", "."), LABEL_MARGIN_X, box_h, box_h, fontsize)
        table.insert(parts, text)
    end

    if selected then
        -- same visual behavior as in your code (red bottom strip + optional square)
        local sel_box = string.format("drawbox=x=0:y=ih-%d:w=iw:h=%d:color=red@%.2f:t=fill",
                                      box_h, box_h, SELECT_BOX_ALPHA)
        table.insert(parts, sel_box)
    end

    return table.concat(parts, ","), box_h
end

-- ffmpeg driver
function generate_bgra_with_ffmpeg_driver(src, dst, w, h, label_text, selected)
    ensure_parent_dir(dst)
    local vf, _ = build_vf_chain(w, h, label_text, selected)
    local tmp = dst .. ".tmp"
    local args = {
        BIN_FFMPEG,
        "-hide_banner",
        "-loglevel",
        "error",
        "-nostdin",
        "-y",
        "-i", src,
        "-frames:v",
        "1",
        "-vf", vf,
        "-pix_fmt",
        "bgra",
        "-f",
        "rawvideo",
        tmp
    }
    local res = utils.subprocess({args=args, cancellable=false})
    if not (res and res.status == 0) then pcall(os.remove, tmp); return false end
    local ok = (file_size(tmp) == expected_bgra_size(w, h))
    if not ok then pcall(os.remove, tmp); return false end
    local rep_ok = atomic_replace(tmp, dst)
    if not rep_ok then pcall(os.remove, tmp); return false end
    return true
end

-- mpv driver
function generate_bgra_with_mpv_driver(src, dst, w, h, label_text, selected)
    ensure_parent_dir(dst)
    local vf, _ = build_vf_chain(w, h, label_text, selected)
    vf = vf .. ",format=bgra"
    local tmp = dst .. ".tmp"
    local args = {
        BIN_MPV,
        "--no-config",
        "--no-audio",
        "--hwdec=no",
        src,
        "--frames=1",
        "--vf=" .. vf,
        "--of=rawvideo",
        "--ovc=rawvideo",
        "-o", tmp
    }
    local res = utils.subprocess({ args = args, cancellable = false })
    if not (res and res.status == 0) then pcall(os.remove, tmp); return false end
    local ok = (file_size(tmp) == expected_bgra_size(w, h))
    if not ok then pcall(os.remove, tmp); return false end
    local rep_ok = atomic_replace(tmp, dst)
    if not rep_ok then pcall(os.remove, tmp); return false end
    return true
end

-- Public: automatically use ffmpeg when available, otherwise mpv; on failure, try fallback
local function generate_bgra_auto(src, dst, w, h, label_text, selected)
    ensure_binaries_ready()
    -- try ffmpeg -> mpv
    if BIN_FFMPEG and generate_bgra_with_ffmpeg_driver(src, dst, w, h, label_text, selected) then
        return true
    end
    if BIN_MPV and generate_bgra_with_mpv_driver(src, dst, w, h, label_text, selected) then
        return true
    end
    return false
end

local function bgra_name_for(src_name, selected)
    local base = strip_ext(src_name)
    return string.format("%s_%s.bgra", base, selected and "SEL" or "L0")
end

generate_bgra_with_ffmpeg = function(src, dst, w, h, label_text, selected)
    return generate_bgra_auto(src, dst, w, h, label_text, selected)
end

generate_bgra_with_mpv = function(src, dst, w, h, label_text, selected)
    return generate_bgra_auto(src, dst, w, h, label_text, selected)
end

-- Create or reuse a solid placeholder BGRA file to avoid blank grid while thumbnails are building
local function ensure_placeholder_bgra(w, h)
    ensure_dir(state.bgra_dir)
    local path = join(state.bgra_dir, string.format("placeholder_%dx%d.bgra", w, h))
    if file_exists(path) then return path end
    local px = string.char(0x20, 0x20, 0x20, 0x33) -- dark gray BGRA
    local row = px:rep(w)
    local data = row:rep(h)
    local f = io.open(path, "wb")
    if f then f:write(data) f:close() end
    return path
end

-- Asynchronous subprocess runner using mpv's 'subprocess' command.
-- This is the officially documented way to run external programs without blocking.
function run_subprocess_async(args, cb, opts)
    opts = opts or {}
    local cmd = {
        name = "subprocess",
        args = args,                 -- array, e.g. {"/usr/bin/ffmpeg", "-y", ...}
        playback_only = false,       -- don't tie process lifetime to playback state
        capture_stdout = opts.capture_stdout or false,
        capture_stderr = opts.capture_stderr or false,
        stdin_data = opts.stdin_data -- optional string
    }
    mp.command_native_async(cmd, function(success, res, err)
        -- Normalize result to something similar to utils.subprocess()
        -- res will contain: status (exit code), stdout/stderr if captured, error (string) on failure
        local out = {
            status = (success and res and res.status) or -1,
            stdout = (res and res.stdout) or "",
            stderr = (res and res.stderr) or "",
            error_string = (not success) and ( (res and res.error) or err or "subprocess failed" ) or nil,
        }
        cb(out)
    end)
end


-- Async BGRA generator: prefer ffmpeg, fallback to mpv, verify output size.
function generate_bgra_auto_async(src, dst, w, h, label_text, selected, cb)
    ensure_binaries_ready()
    ensure_parent_dir(dst)

    local vf, _box_h = build_vf_chain(w, h, label_text, selected)
    local expected = expected_bgra_size(w, h)
    local tmp = dst .. ".tmp"

    local function verify_and_publish()
        local ok = (file_size(tmp) == expected)
        if not ok then pcall(os.remove, tmp) end
        if ok then
            local rep_ok = atomic_replace(tmp, dst)
            if not rep_ok then ok = false end
        end
        -- yield tick pre overlay-read
        mp.add_timeout(0, function() cb(ok) end)
    end

    local function run_mpv_fallback()
        if not BIN_MPV then return cb(false) end
        local mpv_vf = vf .. ",format=bgra"
        local args2 = {
            BIN_MPV, "--no-config", "--no-audio", "--hwdec=no", "--msg-level=all=no",
            src, "--frames=1",
            "--vf=" .. mpv_vf,
            "--of=rawvideo", "--ovc=rawvideo",
            "-o", tmp
        }
        run_subprocess_async(args2, function(res2)
            if res2 and res2.status == 0 then return verify_and_publish() end
            cb(false)
        end)
    end

    if BIN_FFMPEG then
        local args = {
            BIN_FFMPEG, "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-i", src, "-frames:v", "1", "-vf", vf,
            "-pix_fmt", "bgra", "-f", "rawvideo", tmp
        }
        return run_subprocess_async(args, function(res)
            if res and res.status == 0 then return verify_and_publish() end
            run_mpv_fallback()
        end)
    end

    if BIN_MPV then
        return run_mpv_fallback()
    end

    cb(false)
end

-- Creates _SEL from an existing _L0: adds a red bottom strip (with alpha) + optional solid square
local function generate_sel_from_l0(l0_path, sel_path, w, h, box_h)
    local exp = expected_bgra_size(w, h)
    local sz = file_size(l0_path)
    if sz ~= exp then
        msg.warn(("SEL-from-L0: unexpected L0 size: got %s, want %d"):format(tostring(sz), exp))
        return false
    end

    local f = io.open(l0_path, "rb")
    if not f then
        msg.warn("SEL-from-L0: cannot open L0 for read: " .. tostring(l0_path))
        return false
    end
    local data = f:read("*all")
    f:close()
    if not data or #data ~= exp then
        msg.warn("SEL-from-L0: failed to read L0 data or size mismatch")
        return false
    end

    -- Colors (BGRA)
    local a_strip = math.max(0, math.min(255, math.floor(SELECT_BOX_ALPHA * 255 + 0.5)))
    local px_red_solid = string.char(0x00, 0x00, 0xFF, 0xFF)
    -- local px_red_alpha = string.char(0x00, 0x00, 0xFF, string.char(a_strip):byte())
    local px_red_alpha = string.char(0x00, 0x00, 0xFF, a_strip)

    local row_bytes = w * 4
    local top_h = math.max(0, h - box_h)

    local chunks = {}
    -- copy the top part unchanged
    if top_h > 0 then
        chunks[#chunks+1] = data:sub(1, top_h * row_bytes)
    end

    -- bottom rows: compose new ones
    local left_w = 0
    local right_w = w

    local left_part  = ""
    local right_part = (right_w > 0) and px_red_alpha:rep(right_w) or ""
    local bottom_row = left_part .. right_part

    local bottom_block = bottom_row:rep(box_h)
    chunks[#chunks+1] = bottom_block

    local out = table.concat(chunks)

    local wfh, err = io.open(sel_path, "wb")
    if not wfh then
        msg.warn("SEL-from-L0: cannot open SEL for write: " .. tostring(err))
        return false
    end
    wfh:write(out)
    wfh:close()

    local wsz = file_size(sel_path)
    if wsz ~= exp then
        msg.warn(("SEL-from-L0: unexpected SEL size: got %s, want %d"):format(tostring(wsz), exp))
        pcall(os.remove, sel_path)
        return false
    end

    return true
end


ensure_bgra_thumbnail = function(src_path, src_name, w, h, label_text, selected)
    ensure_binaries_ready()
    ensure_dir(state.bgra_dir)

    local l0_name  = bgra_name_for(src_name, false)
    local sel_name = bgra_name_for(src_name, true)
    local l0_path  = join(state.bgra_dir, l0_name)
    local sel_path = join(state.bgra_dir, sel_name)

    local fontsize = math.max(10, math.floor(math.min(h, 1.5 * w) * LABEL_REL_SIZE))
    local box_h    = math.floor(fontsize / 2)
    local exp      = expected_bgra_size(w, h)

    if not selected then
        if file_exists(l0_path) then
            return l0_path, true
        end
        local ok = generate_bgra_auto(src_path, l0_path, w, h, label_text, false)
        if not ok or file_size(l0_path) ~= exp then
            msg.error("Failed to generate BGRA L0: " .. tostring(src_name))
            pcall(os.remove, l0_path)
            return nil, false
        end
        return l0_path, false
    else
        if file_exists(sel_path) then
            return sel_path, true
        end

        -- fast path: from an existing L0
        if file_exists(l0_path) and file_size(l0_path) == exp then
            local ok = generate_sel_from_l0(l0_path, sel_path, w, h, box_h)
            if ok then return sel_path, false end
            pcall(os.remove, sel_path)
        end

        -- if L0 does not exist, generate it
        if not file_exists(l0_path) or file_size(l0_path) ~= exp then
            local ok = generate_bgra_auto(src_path, l0_path, w, h, label_text, false)
            if not ok or file_size(l0_path) ~= exp then
                msg.error("Failed to generate BGRA L0 (for SEL): " .. tostring(src_name))
                pcall(os.remove, l0_path)
                return nil, false
            end
        end

        -- try the quick SEL from L0 again
        local ok = generate_sel_from_l0(l0_path, sel_path, w, h, box_h)
        if ok then return sel_path, false end

        -- final fallback: generate SEL directly (slower)
        ok = generate_bgra_auto(src_path, sel_path, w, h, label_text, true)
        if not ok or file_size(sel_path) ~= exp then
            msg.error("Failed to generate BGRA SEL: " .. tostring(src_name))
            pcall(os.remove, sel_path)
            return nil, false
        end
        return sel_path, false
    end
end

purge_bgra_for_source = function(src_name)
    local base = strip_ext(src_name)
    local esc  = base:gsub("(%W)","%%%1") -- escape Lua pattern metacharacters
    for _, name in ipairs(readdir_files(state.bgra_dir)) do
        if name:match("^" .. esc .. "_L0%.bgra$") or name:match("^" .. esc .. "_SEL%.bgra$") then
            remove_file(join(state.bgra_dir, name))
        end
    end
end

purge_bgra_all = function()
    for _, name in ipairs(readdir_files(state.bgra_dir)) do
        if name:lower():match("%.bgra$") then
            remove_file(join(state.bgra_dir, name))
        end
    end
end

-- ===================== Layout helpers =====================

local function compute_cell_size(vw, vh)
    local SMT, SMB, SML, SMR = safe_margins_px(vw, vh)
    local avail_w = math.max(0, vw - SML - SMR)
    local avail_h = math.max(0, vh - SMT - SMB)
    local cols, rows, gap = GRID_COLS, GRID_ROWS, GRID_GAP
    local cell_w = math.max(2, math.floor((avail_w - (cols + 1) * gap) / cols))
    local cell_h = math.max(2, math.floor((avail_h - (rows + 1) * gap) / rows))
    return cell_w, cell_h
end

-- ===================== Screenshot ("s") =====================

local function save_screenshot()
    if not ensure_paths() then
        mp.osd_message("Video path unavailable", 2)
        return
    end

    -- Lazily create images/<video_name> only when actually saving
    if not ensure_dir(state.images_dir) then
        mp.osd_message("Cannot create images dir", 2.0)
        return
    end

    -- Source video dimensions (DAR-corrected if available)
    local vp = mp.get_property_native("video-params") or {}
    local iw = vp.dw or vp.w
    local ih = vp.dh or vp.h
    if not iw or not ih then
        mp.osd_message("Cannot determine video size", 1.5)
        return
    end

    -- Compute target size according to scale/min limits; no upscaling
    local tw, th, scale = screenshot_target_dims(iw, ih)
    if not tw then
        mp.osd_message("Resolution calc failed", 1.5)
        return
    end

    -- File naming based on the estimated frame number
    local frame = mp.get_property_native("estimated-frame-number") or 0
    local frame_str = string.format("%08d", frame)
    local filename = "f" .. frame_str .. ".jpg"
    local path = join(state.images_dir, filename)

    -- Always write to a temp file first (atomic replace afterwards)
    local tmp_capture = path .. ".cap.tmp"
    local tmp_scaled  = path .. ".sc.tmp"  -- reserved (not needed if ffmpeg writes directly)

    -- 1) Capture the frame via mpv
    local ok_cap, err_cap = pcall(function()
        mp.commandv("screenshot-to-file", tmp_capture, "video")
    end)
    if not ok_cap then
        msg.error("Failed to save screenshot: " .. tostring(err_cap))
        mp.osd_message("Save error: " .. filename, 3.0)
        pcall(os.remove, tmp_capture)
        return
    end

    -- 2) If downscale is not needed (~1.0) or ffmpeg is unavailable -> publish the capture as final
    ensure_binaries_ready()
    local need_resize = (scale < 0.999)
    if not need_resize or not BIN_FFMPEG then
        local rep_ok = atomic_replace(tmp_capture, path)
        if not rep_ok then
            pcall(os.remove, tmp_capture)
            mp.osd_message("Save error (atomic replace)", 2.0)
            return
        end
    else
        -- 3) Downscale via ffmpeg and publish atomically
        local vf = string.format("scale=%d:%d:flags=lanczos", tw, th)
        local args = {
            BIN_FFMPEG, "-hide_banner", "-loglevel", "error", "-nostdin", "-y",
            "-i", tmp_capture, "-vf", vf, path
        }
        local res = utils.subprocess({ args = args, cancellable = false })
        pcall(os.remove, tmp_capture)
        if not res or res.status ~= 0 then
            mp.osd_message("Save error (ffmpeg resize)", 2.0)
            return
        end
    end

    msg.info(("Saved screenshot: %s (%dx%d%s)"):format(
        path, tw, th, need_resize and " resized" or " original"
    ))
    mp.osd_message("Saved: " .. filename, 1.2)

    -- Eager BGRA generation for smooth gallery opening
    local _ml, _mt, vw, vh = get_osd_dims()
    local cell_w, cell_h = compute_cell_size(vw, vh)

    local fps = tonumber(get_fps()) or 0
    local label_text = nil
    if ENABLE_TIME_LABELS and fps > 0 and frame then
        label_text = seconds_to_hhmmss_floor(frame / fps)
    end

    local bg_ok_path, from_cache = ensure_bgra_thumbnail(path, filename, cell_w, cell_h, label_text, false)
    if bg_ok_path then
        msg.info("BGRA ready (" .. (from_cache and "cache" or "new") .. "): " .. bg_ok_path)
    else
        msg.warn("BGRA generation failed (post-screenshot).")
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

local function build_and_draw_page()
    clear_overlays()

    local ml, mt, vw, vh = get_osd_dims()
    local SMT, SMB, SML, SMR = safe_margins_px(vw, vh)
    local cell_w, cell_h = compute_cell_size(vw, vh)

    -- On cell size change: purge cache and rewrite
    if state.last_cell_w ~= cell_w or state.last_cell_h ~= cell_h then
        purge_bgra_all()
        state.last_cell_w = cell_w
        state.last_cell_h = cell_h
    end

    local page_size = GRID_COLS * GRID_ROWS
    local total = #state.files_all
    if total == 0 then
        mp.osd_message("No images in " .. DIR_IMAGES_BASE .. "/" .. (state.video_name or "?"), 2.0)
        return false
    end
    local max_page = math.max(1, math.ceil(total / page_size))
    if state.page > max_page then state.page = max_page end
    if state.page < 1 then state.page = 1 end

    state.last_page_by_video[state.video_name or ""] = state.page

    local start_idx = (state.page - 1) * page_size + 1
    local end_idx = math.min(start_idx + page_size - 1, total)

    -- Reset visible progress
    state.visible_total = (end_idx - start_idx + 1)
    state.visible_missing = 0
    state.visible_pending = {}



    -- 1) Draw placeholders or existing L0 immediately (no blocking)
    local fps = tonumber(get_fps()) or 0
    local placeholder = ensure_placeholder_bgra(cell_w, cell_h)
    state.rects = {}
    state.overlays = {}

    for i = start_idx, end_idx do
        local visible_idx = i - start_idx + 1
        local idx0 = visible_idx - 1
        local rrow = math.floor(idx0 / GRID_COLS)
        local ccol = idx0 % GRID_COLS

        local x_rel = SML + GRID_GAP + ccol * (cell_w + GRID_GAP)
        local y_rel = SMT + GRID_GAP + rrow * (cell_h + GRID_GAP)
        local x_abs = ml + x_rel
        local y_abs = mt + y_rel

        local name = state.files_all[i]
        local src_path = join(state.images_dir, name)
        local frame = parse_frame_from_filename(name)
        local label_text = nil
        if ENABLE_TIME_LABELS and fps > 0 and frame then
            label_text = seconds_to_hhmmss_floor(frame / fps)
        end

        local l0_path = join(state.bgra_dir, bgra_name_for(name, false))
        local exists = (file_size(l0_path) == expected_bgra_size(cell_w, cell_h))
        if not exists then
            state.visible_missing = state.visible_missing + 1
            state.visible_pending[name] = true
        end
        local use_path = exists and l0_path or placeholder
        local id = 1 + (visible_idx - 1)

        pcall(function()
            mp.command_native({"overlay-add", id, x_abs, y_abs, use_path, 0, "bgra", cell_w, cell_h, cell_w * 4})
        end)

        table.insert(state.overlays, id)
        table.insert(state.rects, {
            x_abs=x_abs, y_abs=y_abs, x_rel=x_rel, y_rel=y_rel,
            w=cell_w, h=cell_h,
            path=src_path, name=name,
            frame=frame, selected=false, id=id, label_text=label_text
        })
    end

    update_generation_status()

    -- 2) Kick off async build for visible tiles; redraw each one on completion
    for _, r in ipairs(state.rects) do
        local l0_path = join(state.bgra_dir, bgra_name_for(r.name, false))
        if file_size(l0_path) ~= expected_bgra_size(r.w, r.h) then
            generate_bgra_auto_async(r.path, l0_path, r.w, r.h, r.label_text, false, function(ok)
                if ok and state.active then
                    -- Update only if the same page is still visible
                    for _, rr in ipairs(state.rects or {}) do
                        if rr.name == r.name then
                            redraw_tile_by_name(r.name)
                            mark_tile_done_if_visible(r.name)  -- <- priebeÅ¾nÃ½ progres
                            break
                        end
                    end
                end
            end)
        end
    end

    msg.info(string.format("Gallery: page %d", state.page))
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
        mp.command_native({"overlay-add", r.id, r.x_abs, r.y_abs, bgra_path, 0, "bgra", r.w, r.h, r.w * 4})
    end)
end


-- Redraw a single tile by file name if it's still on the current page.
function redraw_tile_by_name(name)  -- NOTE: no 'local' here, it assigns to the forward-declared local
    for i, r in ipairs(state.rects or {}) do
        if r.name == name then
            local l0 = join(state.bgra_dir, bgra_name_for(name, false))
            if file_size(l0) == expected_bgra_size(r.w, r.h) then
                pcall(function()
                    mp.command_native({"overlay-add", r.id, r.x_abs, r.y_abs, l0, 0, "bgra", r.w, r.h, r.w * 4})
                end)
            end
            break
        end
    end
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

    state.active = false
    clear_overlays()
    mp.osd_message("Gallery hidden", 0.6)
    mp.remove_key_binding("gallery-click")
    mp.remove_key_binding("gallery-ctrl-click")
    mp.remove_key_binding("gallery-select-all-toggle")
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

local function mark_toggle()
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

-- Toggle: if not all selected, select all; otherwise unselect all
local function select_all_toggle()
    local total = #state.rects
    if total == 0 then
        mp.osd_message("No thumbnails on this page", 0.7)
        return
    end
    local selected = 0
    for i = 1, total do
        if state.rects[i].selected then selected = selected + 1 end
    end
    local flag = not (selected == total)  -- true -> select all, false -> unselect all
    for i = 1, total do
        state.rects[i].selected = flag
        redraw_tile(i)
    end
    mp.osd_message(flag and "Selected all" or "Unselected all", 0.7)
end


local function delete_selected_now()
    local to_delete_names = {}

    -- Collect selected tiles
    for _, r in ipairs(state.rects) do
        if r.selected then table.insert(to_delete_names, r.name) end
    end

    -- Simple-delete fallback: delete tile under cursor if nothing is selected
    if #to_delete_names == 0 and ENABLE_SIMPLE_DELETE then
        local pos = mp.get_property_native("mouse-pos")
        if pos and pos.x and pos.y then
            local _idx, r = rect_at(pos.x, pos.y)
            if r then
                table.insert(to_delete_names, r.name)
            end
        end
        if #to_delete_names == 0 then
            mp.osd_message("Nothing selected", 1.0)
            return
        end
    elseif #to_delete_names == 0 then
        mp.osd_message("Nothing selected", 1.0)
        return
    end

    -- Delete selected (or the single tile under cursor)
    for _, name in ipairs(to_delete_names) do
        cancel_prewarm_for(name)                    -- avoid generating thumbs for a file we are deleting
        purge_bgra_for_source(name)                 -- remove L0/SEL cache
        remove_file(join(state.images_dir, name))   -- remove screenshot
    end

    refresh_files_list()
    build_and_draw_page()
    local n = #to_delete_names
    mp.osd_message(string.format("Deleted %d screenshot%s", n, n == 1 and "" or "s"), 1.2)
end


-- Quick page toast shown for a short time
function show_page_indicator()
    if not state.active then return end
    local total_pages = math.max(1, math.ceil((#state.files_all) / GRID_COLS / GRID_ROWS))
    local text = string.format("[Gallery] Page %d/%d", state.page or 1, total_pages)
    mp.osd_message(text, 1) -- show for ~1s
end

local function next_page()
    if not state.active then return end
    local total_pages = math.max(1, math.ceil((#state.files_all) / GRID_COLS / GRID_ROWS))
    if state.page < total_pages then
        state.page = state.page + 1
        build_and_draw_page()
    end
    show_page_indicator()
end

local function prev_page()
    if not state.active then return end
    if state.page > 1 then
        state.page = state.page - 1
        build_and_draw_page()
    end
    show_page_indicator()
end

-- ===================== Exports =====================

local function export_contact_sheet_current_page()

    if #state.rects == 0 then
        mp.osd_message("No thumbnails on this page", 1.0)
        return false
    end

    ensure_binaries_ready()

    -- Guard: require ffmpeg
    if not has_ffmpeg() then
        inform_contact_sheet_unavailable()
        return false
    end

    if not ensure_dir(state.exports_dir) then
        mp.osd_message("Cannot create exports dir", 2.0)
        return false
    end

    local out = join(state.exports_dir, string.format("%s_contact_sheet_p%02d.png", state.video_name, state.page))

    local tmp_dir = join(state.bgra_dir, ".cs_tmp")
    ensure_dir(tmp_dir)
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
            mp.osd_message("Export failed (BGRA -> PNG)", 2.0)
            return false
        end
    end

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

local function export_contact_sheet_all_pages()
    if not ensure_paths() then
        mp.osd_message("Video path unavailable", 2)
        return
    end

    -- Guard: require ffmpeg
    ensure_binaries_ready()

    if not has_ffmpeg() then
        inform_contact_sheet_unavailable()
        return
    end

    refresh_files_list()
    local total = #state.files_all
    if total == 0 then
        mp.osd_message("No screenshots to export", 1.0)
        return
    end

    if not ensure_dir(state.exports_dir) then
        mp.osd_message("Cannot create exports dir", 2.0)
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

    if not ensure_dir(state.exports_dir) then
        mp.osd_message("Cannot create exports dir", 2.0)
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

    if not ensure_dir(state.exports_dir) then
        mp.osd_message("Cannot create exports dir", 2.0)
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

    local args = {
        py,
        script_path,
        "--images-dir", state.images_dir,
        "--out", out,
        "--fps", tostring(fps),
        "--scale", tostring(XLSX_IMG_SCALE),
        "--video-name", state.video_name or "",
        "--resize", "physical",
        "--center", "1"
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
    mp.add_forced_key_binding(KEY_MARK,               "gallery-ctrl-click",         mark_toggle)
    mp.add_forced_key_binding(KEY_SELECT_ALL,         "gallery-select-all-toggle",  select_all_toggle)
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
    mp.remove_key_binding("gallery-select-all-toggle")
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
        local last = state.last_page_by_video[state.video_name or ""]
        state.page = last or 1
        if build_and_draw_page() then
            state.active = true
            bind_gallery_keys()
            mp.osd_message("[Gallery]     goto: click     mark: m     (un)select all: a     delete: d     prev: [     next: ]     export: c/e/x", 5)
        end
    end
end


local function close_gallery()
    if not state.active then return end
    state.active = false
    clear_overlays()
    unbind_gallery_keys()
end



-- ===================== Lifecycle and observers =====================

local function on_file_loaded()

    -- New file opened: ensure gallery is closed to avoid showing stale thumbnails
    close_gallery()

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

    -- Stop any previous prewarm and schedule a fresh one
    if stop_prewarm then stop_prewarm() end
    schedule_thumbnail_regen(0.2)
end

-- Debounce timer for pre-generation

local function current_cell_size()
    local _ml, _mt, vw, vh = get_osd_dims()
    return compute_cell_size(vw, vh)
end

-- Cooperative prewarm: processes queue without freezing UI

local function prewarm_step()
    if not prewarm_running then return end
    local item = table.remove(prewarm_queue, 1)
    if not item then
        prewarm_running = false
        return
    end
    local src_path, name, w, h, label_text = item.src_path, item.name, item.w, item.h, item.label_text
    local l0_path = join(state.bgra_dir, bgra_name_for(name, false))
    if file_exists(l0_path) then
        return mp.add_timeout(0, prewarm_step)
    end

    generate_bgra_auto_async(src_path, l0_path, w, h, label_text, false, function(ok)
        if ok and state.active then
            -- Redraw only if the tile is still on the currently visible page
            for _, rr in ipairs(state.rects or {}) do
                if rr.name == name then
                    redraw_tile_by_name(name)
                    mark_tile_done_if_visible(name)
                    break
                end
            end
        end
        mp.add_timeout(0, prewarm_step) -- yield back to the event loop and start the next item
    end)
end


-- Start prewarming thumbnails for current cell size.
function start_prewarm_all_current_size()
    if prewarm_running then return end
    if not ensure_paths() then return end

    -- Always refresh file list before building the queue
    refresh_files_list()
    if #state.files_all == 0 then
        msg.info("[PreWarm] no files")
        return
    end

    local _ml, _mt, vw, vh = get_osd_dims()
    local w, h = compute_cell_size(vw, vh)

    -- Important: set cell size now, so gallery won't purge prewarmed L0 on open
    state.last_cell_w = w
    state.last_cell_h = h

    prewarm_queue = {}
    local fps = tonumber(get_fps()) or 0
    for _, name in ipairs(state.files_all) do
        local src_path = join(state.images_dir, name)
        local frame = parse_frame_from_filename(name)
        local label_text = nil
        if ENABLE_TIME_LABELS and fps > 0 and frame then
            label_text = seconds_to_hhmmss_floor(frame / fps)
        end
        table.insert(prewarm_queue, {src_path=src_path, name=name, w=w, h=h, label_text=label_text})
    end

    prewarm_running = true
    msg.info(string.format("[PreWarm] start, items=%d, cell=%dx%d", #prewarm_queue, w, h))
    mp.add_timeout(0, prewarm_step)
end

-- Stop/cancel any ongoing prewarm.
function stop_prewarm()
    prewarm_running = false
    prewarm_queue = {}
    msg.info("[PreWarm] stopped")
end


-- pre-generating all thumbnails for the current geometry (L0 and SEL)
local function prewarm_thumbnails_both_variants()
    if regen_in_progress then return end
    regen_in_progress = true

    if not ensure_paths() then
        regen_in_progress = false
        return
    end
    refresh_files_list()

    local cell_w, cell_h = current_cell_size()

    -- if the cell size changes, clear the old cache (same as during drawing)
    if state.last_cell_w ~= cell_w or state.last_cell_h ~= cell_h then
        purge_bgra_all()
        state.last_cell_w = cell_w
        state.last_cell_h = cell_h
    end

    local fps = tonumber(get_fps()) or 0
    for _, src_name in ipairs(state.files_all) do
        local src_path = join(state.images_dir, src_name)
        local frame = parse_frame_from_filename(src_name)
        local label_text = nil
        if ENABLE_TIME_LABELS and fps > 0 and frame then
            label_text = seconds_to_hhmmss_floor(frame / fps)
        end

        -- L0
        local l0_path, _c1 = ensure_bgra_thumbnail(src_path, src_name, cell_w, cell_h, label_text, false)
    end

    regen_in_progress = false
end


-- Debounced regen + prewarm
function schedule_thumbnail_regen(delay)
    if regen_timer then regen_timer:kill() regen_timer = nil end
    regen_timer = mp.add_timeout(delay or 1.0, function()
        regen_timer = nil
        if stop_prewarm then stop_prewarm() end
        if not ensure_paths() then return end

        local _ml, _mt, vw, vh = get_osd_dims()
        local w, h = compute_cell_size(vw, vh)
        local size_changed = (state.last_cell_w ~= w or state.last_cell_h ~= h)
        if size_changed then
            purge_bgra_all()
        end

        start_prewarm_all_current_size()
    end)
end


mp.register_event("file-loaded", on_file_loaded)

mp.observe_property("osd-dimensions", "native", function()
    -- if the gallery is open, close it (prevents recalculation during resize)
    if state.active then
        toggle_gallery()
    end
    -- schedule prewarm (debounce 1 s)
    schedule_thumbnail_regen(1.0)
end)


mp.add_forced_key_binding(KEY_SCREENSHOT,     "save-screenshot-custom", save_screenshot)
mp.add_forced_key_binding(KEY_TOGGLE_GALLERY, "toggle-gallery",         toggle_gallery)
