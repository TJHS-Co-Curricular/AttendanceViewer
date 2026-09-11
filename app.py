#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
出缺席记录查看网站 (Attendance Absence Record Viewer)

Reads every date-folder under Result/ (each one a Google Form "File responses"
export), parses every .xlsx uploaded by each club/society, and shows a
consolidated 姓名 / 班级 / 学号 / 缺席情况 / 备注 table in the browser.

Project layout:
    app.py                 routes + .xlsx parsing (this file)
    templates/              Jinja2 page templates (base/index/folder/grouped)
    static/style.css        all styling
    static/app.js            small shared front-end behaviour (copy button)
    static/favicon.svg       tab icon
    requirements.txt
    build_exe.bat            packages this into a standalone .exe (see README_app.md)
    Result/                  your data -- date-named folders of exported .xlsx files

Run from source:
    pip install -r requirements.txt
    python app.py
Then open http://127.0.0.1:5000 in your browser (it also opens automatically).

By default it looks for a "Result" folder next to this script (or next to
the .exe, when packaged with PyInstaller). Pass a different path as the
first command-line argument to point it elsewhere:
    python app.py "D:\path\to\Result"

To build a standalone .exe that needs no Python installed to run, see
build_exe.bat / README_app.md in the same folder.
"""
import os
import re
import sys
import socket
import datetime
import threading
import webbrowser
import urllib.parse
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from flask import Flask, render_template, abort, request, Response
import openpyxl

# Some source .xlsx files (typically ones with a dropdown / data-validation
# rule set up in Excel or Google Sheets) trigger this harmless openpyxl
# warning on read. We only ever read these files, never re-save them, so the
# unsupported feature being dropped has no effect on anything we show --
# it's just console noise (confusing in the packaged .exe's console window).
import warnings
warnings.filterwarnings(
    "ignore",
    message="Data Validation extension is not supported and will be removed",
    category=UserWarning,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# When packaged by PyInstaller (--onefile), the running script is extracted
# into a temp dir, so two separate "base" locations matter:
#   APP_DIR       -- where user data (Result/) lives: always beside the
#                    actual .exe (sys.executable), so it survives between runs.
#   RESOURCE_DIR  -- where bundled app assets (templates/, static/) live:
#                    the PyInstaller temp extraction dir (sys._MEIPASS) when
#                    frozen, otherwise the same as APP_DIR.
FROZEN = getattr(sys, "frozen", False)
if FROZEN:
    APP_DIR = Path(sys.executable).resolve().parent
    RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    RESOURCE_DIR = APP_DIR

# Command-line flags (e.g. "--dev") vs. the optional Result-folder path --
# any arg starting with "--" is a flag, the first remaining arg (if any) is
# the folder path override.
_argv_flags = {a for a in sys.argv[1:] if a.startswith("--")}
_argv_positional = [a for a in sys.argv[1:] if not a.startswith("--")]

# Developer mode: shows the full per-request access log (every page/CSV/
# favicon hit) in the console. Off by default so double-clicking the .exe
# gives a clean, non-technical-friendly window with just the startup banner.
# Turn it on with either `--dev` / `--verbose` on the command line, or by
# setting the ATTENDANCE_VIEWER_DEBUG=1 environment variable.
DEV_MODE = bool(
    _argv_flags & {"--dev", "--verbose"}
    or os.environ.get("ATTENDANCE_VIEWER_DEBUG") == "1"
)

BASE_DIR = Path(_argv_positional[0]).resolve() if _argv_positional else (APP_DIR / "Result")

app = Flask(
    __name__,
    template_folder=str(RESOURCE_DIR / "templates"),
    static_folder=str(RESOURCE_DIR / "static"),
)


@app.after_request
def _no_browser_cache(response):
    """Every page here is generated fresh from the .xlsx files on each
    request (the server-side cache below already re-reads any file that
    changed on disk), so make sure the *browser* never serves a stale
    cached copy of a folder page -- clicking a folder (or hitting back/
    forward) should always show current data, not a snapshot from before."""
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
    return response


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------

def _clean(v):
    """Normalize a cell value to a plain string (or None)."""
    if v is None:
        return None
    if isinstance(v, str):
        v = v.replace("\xa0", " ").strip()
        return v if v != "" else None
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float):
        # student IDs / serials come through as floats (e.g. 25323.0)
        if v == int(v):
            return str(int(v))
        return str(v)
    return str(v)


def _find_header_row(rows):
    """rows: list of already-_clean()ed row tuples/lists (0-indexed, one per
    sheet row). Returns (row_idx, col_map) where row_idx is 1-based (matching
    the sheet's own row numbers) and col_map maps field name -> 0-based
    column index."""
    for row_idx, cells in enumerate(rows, start=1):
        if "姓名" in cells and "缺席情况" in cells:
            col_map = {}
            for i, c in enumerate(cells):
                if c is None:
                    continue
                if c.startswith("序"):
                    col_map["序"] = i
                elif c.startswith("姓名"):
                    col_map["姓名"] = i
                elif c.startswith("班级"):
                    col_map["班级"] = i
                elif c.startswith("学号"):
                    col_map["学号"] = i
                elif c.startswith("缺席情况"):
                    col_map["缺席情况"] = i
                elif c.startswith("备注"):
                    col_map["备注"] = i
            if {"姓名", "缺席情况"} <= col_map.keys():
                return row_idx, col_map
    return None, None


def _find_meta(rows):
    """Best-effort extraction of club name / date / full-attendance flag from the
    top of the sheet (rows are usually: title, 学会名称, 日期, 全勤/有会员缺席).
    `rows` is the same pre-cleaned row list _find_header_row uses."""
    club_name = None
    session_date = None
    full_attendance = None
    for cells in rows[:8]:
        for i, c in enumerate(cells):
            if c is None:
                continue
            if c.startswith("学会名称") and club_name is None:
                for v in cells[i + 1:]:
                    if v:
                        club_name = v
                        break
            elif c.startswith("日期") and session_date is None:
                for v in cells[i + 1:]:
                    if v:
                        session_date = v
                        break
            elif c.startswith("全勤") and full_attendance is None:
                for v in cells[i + 1:]:
                    if v is not None:
                        full_attendance = v in ("True", "TRUE", "1", "Y", "是") or v is True
                        break
    return club_name, session_date, full_attendance


def _parse_xlsx_uncached(path: Path):
    """Parse one response .xlsx. Returns dict with meta + list of record dicts.
    Never raises -- parsing errors are captured in the 'error' key."""
    result = {
        "file": path.name,
        "club_name": None,
        "session_date": None,
        "full_attendance": None,
        "records": [],
        "error": None,
    }
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:  # noqa: BLE001
        result["error"] = f"无法打开文件：{e}"
        return result

    try:
        for ws in wb.worksheets:
            # Read the sheet's rows exactly once (read_only worksheets stream
            # from the underlying XML, so re-iterating means re-parsing --
            # doing it once here instead of 2-3x is the main per-file win).
            rows = [[_clean(c) for c in row] for row in ws.iter_rows(values_only=True)]

            club_name, session_date, full_attendance = _find_meta(rows)
            if club_name and not result["club_name"]:
                result["club_name"] = club_name
            if session_date and not result["session_date"]:
                result["session_date"] = session_date
            if full_attendance is not None and result["full_attendance"] is None:
                result["full_attendance"] = full_attendance

            header_row, col_map = _find_header_row(rows)
            if header_row is None:
                continue

            first_data_row = True
            for cells in rows[header_row:]:  # rows immediately after the header

                def get(field, cells=cells):
                    idx = col_map.get(field)
                    if idx is None or idx >= len(cells):
                        return None
                    return cells[idx]

                serial = get("序")
                name = get("姓名")

                # Skip the template's worked example row (序 == "例")
                if first_data_row and serial == "例":
                    first_data_row = False
                    continue
                first_data_row = False

                # Stop once we hit the trailing blank template rows (no name)
                if name is None:
                    # allow one stray blank line but then stop
                    break

                result["records"].append({
                    "序": serial,
                    "姓名": name,
                    "班级": get("班级"),
                    "学号": get("学号"),
                    "缺席情况": get("缺席情况"),
                    "备注": get("备注"),
                })
    except Exception as e:  # noqa: BLE001
        result["error"] = (result["error"] + "; " if result["error"] else "") + f"解析出错：{e}"
    finally:
        try:
            wb.close()
        except Exception:  # noqa: BLE001
            pass

    return result


# In-memory cache of parsed files, keyed by absolute path string and
# invalidated by (mtime, size). Since export .xlsx files never change once
# downloaded, this makes every page after the first one essentially free --
# and since parse_xlsx() is the single choke point, the CSV export, the
# grouped-by-type page and the main table all benefit automatically.
_PARSE_CACHE: dict = {}
_PARSE_CACHE_LOCK = threading.Lock()


def parse_xlsx(path: Path):
    cache_key = str(path)
    try:
        st = path.stat()
        stamp = (st.st_mtime, st.st_size)
    except OSError:
        stamp = None

    if stamp is not None:
        with _PARSE_CACHE_LOCK:
            cached = _PARSE_CACHE.get(cache_key)
        if cached is not None and cached[0] == stamp:
            return cached[1]

    result = _parse_xlsx_uncached(path)

    if stamp is not None:
        with _PARSE_CACHE_LOCK:
            _PARSE_CACHE[cache_key] = (stamp, result)

    return result


def parse_many(files):
    """Parse several files concurrently (helps when disk I/O -- e.g. a
    network drive -- dominates; falls back to sequential for 0-1 files)."""
    files = list(files)
    if len(files) <= 1:
        return [parse_xlsx(f) for f in files]
    with ThreadPoolExecutor(max_workers=min(8, len(files))) as pool:
        return list(pool.map(parse_xlsx, files))


# ---------------------------------------------------------------------------
# Folder discovery
# ---------------------------------------------------------------------------

DATE_PREFIX_RE = re.compile(r"^(\d{2})\.(\d{2})\.(\d{4})")
CLUB_CODE_RE = re.compile(r"^([A-Za-z]+)\s*0*([0-9]+)")


def _club_sort_key(parsed_entry):
    """Sort club sections by their code (A01, A02, ... B02, ... D28, ...),
    alphabetically by letter then numerically by number, so A01/A02/... sorts
    correctly ahead of B.../D... regardless of upload order."""
    name = (parsed_entry.get("club_name") or "").strip()
    m = CLUB_CODE_RE.match(name)
    if m:
        letters, digits = m.groups()
        return (0, letters.upper(), int(digits), name)
    fallback = name or parsed_entry.get("file") or ""
    return (1, fallback, 0, fallback)


def _club_code(parsed_entry) -> str:
    """Extract just the club/team code (e.g. "A01", "D19") from a parsed
    file's club name, falling back to the source filename if needed."""
    for candidate in (parsed_entry.get("club_name"), parsed_entry.get("file")):
        if not candidate:
            continue
        m = CLUB_CODE_RE.match(candidate.strip())
        if m:
            letters, digits = m.groups()
            return f"{letters.upper()}{digits.zfill(2)}"
    return "UNKNOWN"


# Canonical order matching the form template's own wording:
# "缺席情况 ：旷课、病假、事假、特别事假、公假等。"
TYPE_ORDER = ["旷课", "病假", "事假", "特别事假", "公假"]

# These 缺席情况 types are excluded from the grouped/batch export (they don't
# need to be submitted through that list) but still show normally in the
# main table and CSV export.
GROUPED_EXPORT_IGNORE_TYPES = {"内部公假", "迟到", "早退"}


def build_grouped_by_type(date_folder: Path, parsed=None) -> str:
    """Consolidate every absence record for a date into batches of at most
    10 student IDs per line, grouped first by 缺席情况 (leave type), then by
    学会/团体 (club) within that type. Format:

        "病假":
        A01_01: 12345,12346,...
        B05_01: 22222,...

        "事假":
        ...

    Pass an already-parsed list (see parse_many()) via `parsed` to avoid
    re-reading every .xlsx in the folder a second time -- the caller
    (view_folder) already has one.
    """
    if parsed is None:
        files = find_response_files(date_folder)
        parsed = parse_many(files)
    parsed = sorted(parsed, key=_club_sort_key)

    # type -> club_code -> [student_id, ...]  (dict preserves insertion order)
    groups: dict[str, dict[str, list[str]]] = {}
    for p in parsed:
        club_code = _club_code(p)
        for r in p["records"]:
            leave_type = (r["缺席情况"] or "未填写").strip() or "未填写"
            if leave_type in GROUPED_EXPORT_IGNORE_TYPES:
                continue
            sid = (r["学号"] or "").strip()
            if not sid:
                continue
            groups.setdefault(leave_type, {}).setdefault(club_code, []).append(sid)

    ordered_types = [t for t in TYPE_ORDER if t in groups]
    ordered_types += sorted(t for t in groups if t not in TYPE_ORDER)

    lines: list[str] = []
    for leave_type in ordered_types:
        if lines:
            lines.append("")  # blank line between type blocks
        lines.append(f'"{leave_type}":')
        club_map = groups[leave_type]
        for club_code in sorted(club_map.keys()):
            ids = club_map[club_code]
            for i in range(0, len(ids), 10):
                batch = ids[i:i + 10]
                seq = i // 10 + 1
                lines.append(f"{club_code}_{seq:02d}: " + ",".join(batch))

    if not lines:
        return "（此资料夹没有找到任何缺席记录）\n"

    return "\n".join(lines) + "\n"


def _folder_sort_key(name: str):
    m = DATE_PREFIX_RE.match(name)
    if m:
        d, mo, y = m.groups()
        return (0, y, mo, d)
    return (1, name)


def list_date_folders():
    if not BASE_DIR.exists():
        return []
    folders = [p for p in BASE_DIR.iterdir() if p.is_dir()]
    folders.sort(key=lambda p: _folder_sort_key(p.name))
    return folders


def find_response_files(date_folder: Path):
    """Each date folder contains one "<question> (File responses)" subfolder
    holding the actual uploaded .xlsx files. Search recursively to be safe."""
    return sorted(date_folder.rglob("*.xlsx"))


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/favicon.ico")
def favicon():
    # A real favicon is served via the <link rel="icon"> tag in base.html
    # (static/favicon.svg); this just avoids a 404 page for browsers/tools
    # that request /favicon.ico directly regardless of that tag.
    return Response(status=204)


@app.route("/")
def index():
    folders = list_date_folders()
    folder_data = [
        {"name": f.name, "count": len(find_response_files(f))}
        for f in folders
    ]
    return render_template(
        "index.html",
        title="出缺席记录查看网站",
        folders=folder_data,
        base_dir=str(BASE_DIR),
    )


@app.route("/folder/<path:folder_name>")
def view_folder(folder_name):
    date_folder = BASE_DIR / folder_name
    if not date_folder.exists() or not date_folder.is_dir():
        abort(404, f"找不到资料夹：{folder_name}")

    q = (request.args.get("q") or "").strip()

    files = find_response_files(date_folder)
    parsed = parse_many(files)
    parsed.sort(key=_club_sort_key)

    total_records = sum(len(p["records"]) for p in parsed)
    total_clubs = len(parsed)
    error_count = sum(1 for p in parsed if p["error"])
    reason_counts: dict = {}
    for p in parsed:
        for r in p["records"]:
            k = r["缺席情况"] or "未填写"
            reason_counts[k] = reason_counts.get(k, 0) + 1

    stats = {
        "total_clubs": total_clubs,
        "total_records": total_records,
        "error_count": error_count,
        "reason_counts": sorted(reason_counts.items(), key=lambda x: -x[1]),
    }

    sections = []
    for p in parsed:
        entry = {
            "club_label": p["club_name"] or p["file"],
            "session_date": str(p["session_date"]) if p["session_date"] else None,
            "file": p["file"],
            "error": p["error"],
            "full_attendance_empty": bool(p["full_attendance"] and not p["records"]),
            "records": [],
            "show_empty_message": False,
        }

        if p["error"]:
            sections.append(entry)
            continue

        records = p["records"]
        if q:
            ql = q.lower()
            records = [
                r for r in records
                if any(ql in str(v).lower() for v in r.values() if v)
            ]

        if not records:
            entry["show_empty_message"] = not p["full_attendance"]
            sections.append(entry)
            continue

        entry["records"] = records
        sections.append(entry)

    return render_template(
        "folder.html",
        title=folder_name,
        folder_name=folder_name,
        q=q,
        stats=stats,
        sections=sections,
    )


@app.route("/folder/<path:folder_name>/grouped")
def view_grouped(folder_name):
    """Standalone page (opened in a new tab) showing the 依缺席情况分组名单
    for one date -- kept separate from the main folder page so it's easy to
    keep open in its own tab/window while browsing other dates."""
    date_folder = BASE_DIR / folder_name
    if not date_folder.exists() or not date_folder.is_dir():
        abort(404, f"找不到资料夹：{folder_name}")

    grouped_text = build_grouped_by_type(date_folder)

    return render_template(
        "grouped.html",
        title=f"分组名单 - {folder_name}",
        folder_name=folder_name,
        grouped_text=grouped_text,
    )


@app.route("/folder/<path:folder_name>/grouped.txt")
def export_grouped(folder_name):
    date_folder = BASE_DIR / folder_name
    if not date_folder.exists() or not date_folder.is_dir():
        abort(404, f"找不到资料夹：{folder_name}")
    text = build_grouped_by_type(date_folder)
    return Response(text, mimetype="text/plain; charset=utf-8")


@app.route("/folder/<path:folder_name>/export.csv")
def export_csv(folder_name):
    import csv
    import io

    date_folder = BASE_DIR / folder_name
    if not date_folder.exists() or not date_folder.is_dir():
        abort(404)

    q = (request.args.get("q") or "").strip().lower()

    files = find_response_files(date_folder)
    buf = io.StringIO()
    buf.write("\ufeff")  # BOM so Excel opens UTF-8 correctly
    writer = csv.writer(buf)
    writer.writerow(["学会/团体", "日期", "序", "姓名", "班级", "学号", "缺席情况", "备注", "来源档案"])
    parsed = parse_many(files)
    parsed.sort(key=_club_sort_key)
    for p in parsed:
        for r in p["records"]:
            if q and not any(q in str(v).lower() for v in r.values() if v):
                continue
            writer.writerow([
                p["club_name"] or "", p["session_date"] or "",
                r["序"] or "", r["姓名"] or "", r["班级"] or "", r["学号"] or "",
                r["缺席情况"] or "", r["备注"] or "", p["file"],
            ])

    filename = f"{folder_name}.csv".replace("/", "-")
    quoted = urllib.parse.quote(filename)
    return Response(
        buf.getvalue(),
        mimetype="text/csv",
        headers={
            "Content-Disposition": f"attachment; filename=\"attendance.csv\"; filename*=UTF-8''{quoted}"
        },
    )


def _find_free_port(preferred=5000, tries=20):
    for port in range(preferred, preferred + tries):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", port))
                return port
            except OSError:
                continue
    return 0  # let the OS pick any free port


if __name__ == "__main__":
    # Quiet by default: hide Werkzeug's per-request access log line (one
    # printed for every page/CSV/favicon hit) so the console window stays
    # clean for a non-technical user. --dev / --verbose / the
    # ATTENDANCE_VIEWER_DEBUG=1 env var switch it back on for troubleshooting.
    import logging
    logging.getLogger("werkzeug").setLevel(logging.INFO if DEV_MODE else logging.WARNING)

    port = _find_free_port(5000)
    url = f"http://127.0.0.1:{port}"

    print("=" * 60)
    print(" 出缺席记录查看网站 / Attendance Viewer")
    print("=" * 60)
    print(f" 资料来源 (Result folder): {BASE_DIR}")
    print(f" 网址 (URL): {url}")
    if DEV_MODE:
        print(" 开发者模式：显示每次请求记录 (developer mode: request log shown)")
    else:
        print(" 如需显示每次请求记录，请加上 --dev 参数重新启动")
        print(" (to show the request log, restart with a --dev flag)")
    print(" 关闭这个窗口即可停止网站 (Close this window to stop the site)")
    print("=" * 60)

    # Open the default browser shortly after the server starts, so this
    # behaves like a normal double-click-to-run desktop app.
    def _open_browser():
        try:
            webbrowser.open(url)
        except Exception:  # noqa: BLE001
            pass

    threading.Timer(1.0, _open_browser).start()

    # debug/reloader must stay off: the reloader re-execs the process, which
    # breaks (or infinite-loops) inside a PyInstaller --onefile bundle.
    # threaded=True lets the browser's automatic favicon/prefetch requests
    # (or a second tab) proceed without queuing behind a slow folder parse.
    app.run(host="127.0.0.1", port=port, debug=False, use_reloader=False, threaded=True)
