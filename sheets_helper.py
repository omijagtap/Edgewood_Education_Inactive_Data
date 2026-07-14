import os
import re
import json
import gspread
from google.oauth2 import service_account
from gspread_formatting import *
from concurrent.futures import ThreadPoolExecutor

try:
    from dotenv import load_dotenv
    _env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(_env_path)
except ImportError:
    pass

SPREADSHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_CREDS_FILE = os.getenv("GOOGLE_CREDS_FILE", "linen-rex-436411-r4-9bba0db0c720.json")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = GOOGLE_CREDS_FILE if os.path.isabs(GOOGLE_CREDS_FILE) else os.path.join(SCRIPT_DIR, GOOGLE_CREDS_FILE)

# ── Palette ───────────────────────────────────────────────────────────────────
C_NAVY     = "1F3864"
C_WHITE    = "FFFFFF"
C_SLATE    = "475569"
C_GREEN_BG = "DCFCE7"; C_GREEN_FG = "166534"
C_BLUE_BG  = "E0F2FE"; C_BLUE_FG  = "075985"
C_AMB_BG   = "FEF3C7"; C_AMB_FG   = "92400E"
C_RED_BG   = "FEE2E2"; C_RED_FG   = "991B1B"
C_GREY_BG  = "F1F5F9"; C_GREY_FG  = "475569"
C_BORDER   = "D9D9D9"
C_ALT_ROW  = "F8FAFC"

CAT_STYLES = {
    "ACTIVE":                              {"bg": C_GREEN_BG, "fg": C_GREEN_FG},
    "LOW ACTIVITY - ASSIGNMENTS COMPLETED":{"bg": C_BLUE_BG,  "fg": C_BLUE_FG},
    "MISSED SUBMISSION":                   {"bg": C_AMB_BG,   "fg": C_AMB_FG},
    "NOT ACTIVE":                          {"bg": C_RED_BG,   "fg": C_RED_FG},
    "NO ACTIVITY":                         {"bg": C_GREY_BG,  "fg": C_GREY_FG},
}

WEEK_STYLES = {
    "MET":    {"bg": "E8F5E9", "fg": "2E7D32"},
    "NOT MET":{"bg": "FFEBEE", "fg": "C62828"},
    "-":      {"bg": C_WHITE,  "fg": "000000"},
}


def _color(hex_str):
    r = int(hex_str[0:2], 16) / 255.0
    g = int(hex_str[2:4], 16) / 255.0
    b = int(hex_str[4:6], 16) / 255.0
    return Color(r, g, b)

def _border(color_hex=C_BORDER):
    s = Border(style="SOLID", color=_color(color_hex))
    return Borders(top=s, bottom=s, left=s, right=s)

def _fmt(bg, fg=None, bold=False, size=9, halign="CENTER"):
    cf = CellFormat(
        backgroundColor=_color(bg),
        horizontalAlignment=halign,
        verticalAlignment="MIDDLE",
        borders=_border(),
        wrapStrategy="WRAP",
    )
    if fg:
        cf.textFormat = TextFormat(bold=bold, fontSize=size, foregroundColor=_color(fg))
    else:
        cf.textFormat = TextFormat(bold=bold, fontSize=size)
    return cf


def get_google_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    raw = os.getenv("GOOGLE_CREDS_JSON")
    if raw:
        creds = service_account.Credentials.from_service_account_info(json.loads(raw), scopes=scopes)
    elif os.path.exists(CREDENTIALS_PATH):
        creds = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scopes)
    else:
        raise FileNotFoundError(f"Credentials not found at {CREDENTIALS_PATH}")
    return gspread.Client(auth=creds)


def _hms(seconds):
    s = int(seconds or 0)
    return f"{s//3600:02d}:{(s%3600)//60:02d}:{s%60:02d}"


def _safe_title(sis_id):
    return re.sub(r'[:\\/\?\*\[\]]', '_', sis_id)[:31]


# ── Main push ─────────────────────────────────────────────────────────────────
def push_to_google_sheet(flat_rows, sis_id, course_name, duration_weeks=6):
    """Push FULL data (matching Excel) to a course-specific tab. No Consolidate tab."""
    print("  Connecting to Google Sheets...")
    gc  = get_google_client()
    sh  = gc.open_by_key(SPREADSHEET_ID)
    tab = _safe_title(sis_id)

    # Dynamic column count: base 15 + 2 per week
    total_cols = 15 + duration_weeks * 2

    # ── FIX: Create the course tab FIRST, then clean up stale tabs ────────────
    # We must never delete the last remaining sheet — Google Sheets API will error.
    # Strategy: add the new course tab first, THEN delete Consolidate + old course tab.

    # 1. Delete the old course tab if it exists (it won't be the last sheet
    #    because Consolidate is still there)
    try:
        old_ws = sh.worksheet(tab)
        print(f"  Worksheet '{tab}' exists. Deleting old version...")
        sh.del_worksheet(old_ws)
    except gspread.exceptions.WorksheetNotFound:
        pass

    # 2. Create fresh course tab
    print(f"  Creating worksheet '{tab}'...")
    ws = sh.add_worksheet(title=tab, rows=str(max(500, len(flat_rows)+10)), cols=str(total_cols + 2))

    # 3. NOW it is safe to remove stale Consolidate tab (course tab exists above)
    try:
        stale = sh.worksheet("Consolidate")
        sh.del_worksheet(stale)
        print("  Removed stale Consolidate tab.")
    except gspread.exceptions.WorksheetNotFound:
        pass  # already clean

    # ── Build header rows ────────────────────────────────────────────────────
    # Row 1: group headers
    grp_row1 = (
        ["Course Information"] * 2
        + ["Learner Information"] * 4
        + ["Activity Summary"] * 2
        + ["Assignment Tracking"] * 5
        + ["Engagement Classification"] * 2
    )
    for w in range(1, duration_weeks + 1):
        grp_row1 += [f"Week {w}"] * 2

    # Row 2: column headers
    hdr_row2 = [
        "Course ID", "Course Name",
        "Cohort", "Learner Name", "Official Email ID", "Enrollment Status",
        "Last Activity Timestamp", "Total Time Spent (HH:MM:SS)",
        "Total Assignments", "Submitted", "Missing", "Overdue", "On-Time Submissions",
        "Engagement Category", "Overall Activity"
    ]
    for w in range(1, duration_weeks + 1):
        hdr_row2 += [f"W{w} Duration (HH:MM:SS)", f"W{w} Status"]

    # ── Build data rows ──────────────────────────────────────────────────────
    rows = [grp_row1, hdr_row2]
    for r in flat_rows:
        last_act = r.get("last_activity_timestamp")
        last_str = last_act.strftime("%Y-%m-%d %H:%M:%S") if last_act else "N/A"
        row = [
            sis_id, course_name,
            r.get("cohort", "N/A"), r.get("name", "N/A"),
            r.get("email", "N/A"), r.get("status", "N/A"),
            last_str, _hms(r.get("total_engagement_seconds", 0)),
            r.get("total_assignments", 0), r.get("submitted_assignments", 0),
            r.get("missing_assignments", 0), r.get("overdue_assignments", 0),
            r.get("on_time_submissions", 0),
            r.get("category", "N/A"), r.get("overall_activity", "N/A"),
        ]
        wd = r.get("weekly_data", {})
        for w in range(1, duration_weeks + 1):
            wk = wd.get(w, {})
            t = wk.get("time_spent", "-")
            s = wk.get("status", "-")
            row.append(_hms(t) if isinstance(t, (int, float)) else str(t))
            row.append(str(s))
        rows.append(row)

    # Write all at once
    end_col_letter = _col_letter(total_cols)
    ws.update(range_name=f"A1:{end_col_letter}{len(rows)}", values=rows)

    # ── Format ───────────────────────────────────────────────────────────────
    print(f"  Formatting '{tab}'...")
    n_data = len(rows)

    with batch_updater(sh) as b:
        b.format_cell_range(ws, f"A1:{end_col_letter}1", _fmt(C_NAVY, C_WHITE, bold=True, size=10))
        b.format_cell_range(ws, f"A2:{end_col_letter}2", _fmt(C_SLATE, C_WHITE, bold=True, size=9))

        if n_data > 2:
            for i in range(3, n_data + 1):
                bg = C_ALT_ROW if i % 2 == 0 else C_WHITE
                b.format_cell_range(ws, f"A{i}:{end_col_letter}{i}",
                                    CellFormat(backgroundColor=_color(bg),
                                               borders=_border(),
                                               verticalAlignment="MIDDLE"))

            cat_col = "N"
            for i, r in enumerate(flat_rows, start=3):
                cat = r.get("category", "")
                style = CAT_STYLES.get(cat, {"bg": C_WHITE, "fg": C_SLATE})
                b.format_cell_range(ws, f"{cat_col}{i}",
                                    CellFormat(backgroundColor=_color(style["bg"]),
                                               textFormat=TextFormat(bold=True, foregroundColor=_color(style["fg"])),
                                               horizontalAlignment="CENTER",
                                               borders=_border()))

            for i, r in enumerate(flat_rows, start=3):
                oa = r.get("overall_activity", "")
                bg = C_GREEN_BG if oa == "Active" else C_RED_BG
                fg = C_GREEN_FG if oa == "Active" else C_RED_FG
                b.format_cell_range(ws, f"O{i}",
                                    CellFormat(backgroundColor=_color(bg),
                                               textFormat=TextFormat(bold=True, foregroundColor=_color(fg)),
                                               horizontalAlignment="CENTER",
                                               borders=_border()))

            for w in range(1, duration_weeks + 1):
                status_col_idx = 15 + (w - 1) * 2 + 2
                status_col = _col_letter(status_col_idx)
                for i, r in enumerate(flat_rows, start=3):
                    wk = r.get("weekly_data", {}).get(w, {})
                    st = wk.get("status", "-")
                    style = WEEK_STYLES.get(str(st), WEEK_STYLES["-"])
                    b.format_cell_range(ws, f"{status_col}{i}",
                                        CellFormat(backgroundColor=_color(style["bg"]),
                                                   textFormat=TextFormat(foregroundColor=_color(style["fg"]), bold=True),
                                                   horizontalAlignment="CENTER",
                                                   borders=_border()))

    ws.freeze(rows=2)
    ws.columns_auto_resize(0, min(total_cols - 1, 25))
    print(f"  Sheet tab '{tab}' updated successfully!")


def _update_consolidate(sh, duration_weeks=6):
    CON_TITLE = "Consolidate"
    total_cols = 15 + duration_weeks * 2
    end_col = _col_letter(total_cols)

    try:
        ws_con = sh.worksheet(CON_TITLE)
        ws_con.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws_con = sh.add_worksheet(title=CON_TITLE, rows="2000", cols=str(total_cols + 2))

    # Put Consolidate first
    all_ws = sh.worksheets()
    sh.reorder_worksheets([ws_con] + [w for w in all_ws if w.title != CON_TITLE])

    all_rows = []
    for ws in sh.worksheets():
        if ws.title == CON_TITLE:
            continue
        try:
            vals = ws.get_all_values()
            if len(vals) > 2:
                if not all_rows:
                    all_rows.extend(vals[:2])   # include headers from first sheet
                all_rows.extend(vals[2:])       # skip headers from subsequent sheets
        except Exception:
            pass

    if all_rows:
        ws_con.update(range_name=f"A1:{end_col}{len(all_rows)}", values=all_rows)
        with batch_updater(sh) as b:
            b.format_cell_range(ws_con, f"A1:{end_col}1", _fmt(C_NAVY, C_WHITE, bold=True, size=10))
            b.format_cell_range(ws_con, f"A2:{end_col}2", _fmt(C_SLATE, C_WHITE, bold=True, size=9))
        ws_con.freeze(rows=2)


# ── Column letter helper ───────────────────────────────────────────────────────
def _col_letter(n):
    """Convert 1-based column index to spreadsheet letter (A, B, … AA, AB…)."""
    result = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result


# ── Dashboard cache ───────────────────────────────────────────────────────────
_dashboard_cache = {"data": None, "fetched_at": 0}
CACHE_TTL = 3600  # Cache for 1 hour; cleared explicitly on new audit runs
CACHE_FILE = os.path.join(SCRIPT_DIR, "dashboard_cache.json")


def _fetch_ws(ws):
    try:
        return ws.title, ws.get_all_values()
    except Exception:
        return ws.title, None


def get_dashboard_data(force=False):
    import time as _t
    now = _t.time()

    # If not forced, try to return in-memory cache first
    if not force and _dashboard_cache["data"] and (now - _dashboard_cache["fetched_at"]) < CACHE_TTL:
        return _dashboard_cache["data"]

    # If not forced, try to load from disk cache
    if not force and os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                _dashboard_cache["data"] = data
                _dashboard_cache["fetched_at"] = now
                return data
        except Exception as ce:
            print(f"Error loading dashboard disk cache: {ce}")

    # Otherwise, fetch fresh data from Google Sheets API
    print("  Fetching fresh dashboard data from Google Sheets API...")
    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    # Read all course-specific tabs in parallel; skip any Consolidate tab
    all_data_rows = []
    course_tabs   = []
    with ThreadPoolExecutor(max_workers=10) as ex:
        for title, vals in ex.map(_fetch_ws, sh.worksheets()):
            if title == "Consolidate":
                continue          # ignore any leftover Consolidate tab
            if vals and len(vals) > 2:
                course_tabs.append(title)
                all_data_rows.extend(vals[2:])   # skip the 2 header rows per tab

    # ── Parse each row ───────────────────────────────────────────────────────
    total_courses_set = set()
    active_ct = 0; inactive_ct = 0
    total_missing = 0; total_submitted = 0; total_assignments = 0
    total_time_s = 0; valid_time = 0

    cat_counts   = {}   # category -> count
    course_map   = {}   # course_id -> {name, active, inactive, total, missing, submitted, total_asgn}
    cohort_map   = {}   # cohort -> {name, active, inactive, total, missing, total_time_s, valid_time}
    watchlist    = []   # inactive learners detail
    top_learners = []   # highest time spent

    for row in all_data_rows:
        if len(row) < 15:
            continue
        course_id    = row[0].strip()
        course_name  = row[1].strip()
        cohort       = row[2].strip() or "N/A"
        learner_name = row[3].strip()
        email        = row[4].strip()
        status       = row[5].strip()
        last_act     = row[6].strip()
        time_str     = row[7].strip()
        total_asgn   = _safe_int(row[8])
        submitted    = _safe_int(row[9])
        missing      = _safe_int(row[10])
        overdue      = _safe_int(row[11])
        on_time      = _safe_int(row[12])
        category     = row[13].strip()
        overall      = row[14].strip()

        if not learner_name:
            continue

        total_courses_set.add(course_id)
        total_missing    += missing
        total_submitted  += submitted
        total_assignments+= total_asgn

        ts = _parse_hms(time_str)
        if ts > 0:
            total_time_s += ts
            valid_time   += 1

        if overall == "Active":
            active_ct += 1
        else:
            inactive_ct += 1
            watchlist.append({
                "learner_name": learner_name,
                "email": email,
                "course_id": course_id,
                "course_name": course_name,
                "cohort": cohort,
                "category": category,
                "missing": missing,
                "overdue": overdue,
                "last_act": last_act,
                "time_hms": time_str,
            })

        cat_counts[category] = cat_counts.get(category, 0) + 1

        # Course breakdown
        if course_id not in course_map:
            course_map[course_id] = {"name": course_name, "active": 0, "inactive": 0, "total": 0,
                                      "missing": 0, "submitted": 0, "total_asgn": 0}
        cm = course_map[course_id]
        cm["total"] += 1
        cm["missing"]    += missing
        cm["submitted"]  += submitted
        cm["total_asgn"] += total_asgn
        if overall == "Active":
            cm["active"] += 1
        else:
            cm["inactive"] += 1

        # Cohort breakdown
        if cohort not in cohort_map:
            cohort_map[cohort] = {"name": cohort, "active": 0, "inactive": 0, "total": 0,
                                  "missing": 0, "total_time_s": 0, "valid_time": 0}
        coh = cohort_map[cohort]
        coh["total"] += 1
        coh["missing"] += missing
        coh["total_time_s"] += ts
        if ts > 0:
            coh["valid_time"] += 1
        if overall == "Active":
            coh["active"] += 1
        else:
            coh["inactive"] += 1

        top_learners.append({
            "learner_name": learner_name,
            "email": email,
            "course_name": course_name,
            "time_seconds": ts,
            "time_hms": time_str,
            "submitted": submitted,
            "total_asgn": total_asgn,
            "category": category,
            "overall": overall,
        })

    avg_time_h = round((total_time_s / valid_time) / 3600, 2) if valid_time else 0

    # Sort top learners by time spent
    top_learners.sort(key=lambda x: x["time_seconds"], reverse=True)

    # Process cohort map to calculate averages
    processed_cohorts = {}
    for cname, cinfo in cohort_map.items():
        avg_coh_t = round((cinfo["total_time_s"] / cinfo["valid_time"]) / 3600, 2) if cinfo["valid_time"] else 0
        processed_cohorts[cname] = {
            "name": cname,
            "total": cinfo["total"],
            "active": cinfo["active"],
            "inactive": cinfo["inactive"],
            "missing": cinfo["missing"],
            "avg_time_hours": avg_coh_t
        }

    payload = {
        "kpis": {
            "total_courses": len(total_courses_set),
            "total_learners": active_ct + inactive_ct,
            "active_learners": active_ct,
            "inactive_learners": inactive_ct,
            "total_missing": total_missing,
            "total_submitted": total_submitted,
            "avg_time_hours": avg_time_h,
        },
        "cat_counts": cat_counts,
        "course_tabs": sorted(course_tabs),
        "course_breakdown": course_map,
        "cohort_breakdown": processed_cohorts,
        "watchlist": watchlist[:100],
        "top_learners": top_learners[:50],
    }

    # Save to disk cache
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception as se:
        print(f"Error saving dashboard disk cache: {se}")

    _dashboard_cache["data"] = payload
    _dashboard_cache["fetched_at"] = now
    return payload


def _safe_int(v):
    try:
        return int(v)
    except Exception:
        return 0


def _parse_hms(hms_str):
    """Parse HH:MM:SS string to total seconds."""
    try:
        parts = hms_str.split(":")
        if len(parts) == 3:
            return int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
    except Exception:
        pass
    return 0
