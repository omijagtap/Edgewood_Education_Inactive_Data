import os
import re
import json
import datetime
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

SPREADSHEET_ID = os.getenv("GOOGLE_SHEET_ID", "14SKZH71x2hCgwfA2CCwifkWTD4VQht0n43MFpfDY6u0")
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
    "MISSED SUBMISSION":                   {"bg": C_AMB_BG,   "fg": C_AMB_FG},
    "INACTIVE":                            {"bg": C_RED_BG,   "fg": C_RED_FG},
    "INACTIVE - NOT LOGGED IN":            {"bg": "FECDD3",    "fg": "9F1239"},
    "In Active":                           {"bg": C_RED_BG,   "fg": C_RED_FG},
    "No Submission":                       {"bg": C_RED_BG,   "fg": C_RED_FG},
    "NOT ACTIVE":                          {"bg": C_RED_BG,   "fg": C_RED_FG},
    "NO ACTIVITY":                         {"bg": "FECDD3",    "fg": "9F1239"},
    "LOW ACTIVITY - ASSIGNMENTS COMPLETED":{"bg": C_AMB_BG,   "fg": C_AMB_FG},
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
    creds = None
    if raw:
        raw_clean = raw.strip()
        if (raw_clean.startswith("'") and raw_clean.endswith("'")) or (raw_clean.startswith('"') and raw_clean.endswith('"')):
            raw_clean = raw_clean[1:-1].strip()
        try:
            creds = service_account.Credentials.from_service_account_info(json.loads(raw_clean), scopes=scopes)
        except Exception as e:
            if os.path.exists(CREDENTIALS_PATH):
                creds = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scopes)
            else:
                raise ValueError(f"Failed to parse GOOGLE_CREDS_JSON: {e}")
    elif os.path.exists(CREDENTIALS_PATH):
        creds = service_account.Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=scopes)
    else:
        raise FileNotFoundError(f"Credentials not found at {CREDENTIALS_PATH} and GOOGLE_CREDS_JSON is empty.")
    
    if not SPREADSHEET_ID:
        raise ValueError("GOOGLE_SHEET_ID is missing from environment variables.")
        
    return gspread.Client(auth=creds)


def _hms(seconds):
    s = int(seconds or 0)
    return f"{s//3600:02d}:{(s%3600)//60:02d}:{s%60:02d}"


def _safe_title(sis_id):
    return re.sub(r'[:\\/\?\*\[\]]', '_', sis_id)[:31]


# ── Main push ─────────────────────────────────────────────────────────────────
def push_to_google_sheet(flat_rows, sis_id, course_name, duration_weeks=6, assignments=None):
    """Push FULL data (matching Excel) to a course-specific tab. No Consolidate tab."""
    print("  Connecting to Google Sheets...")
    gc  = get_google_client()
    sh  = gc.open_by_key(SPREADSHEET_ID)
    tab = _safe_title(sis_id)

    assignments = assignments or []
    num_asgn_cols = len(assignments)
    total_cols = 10 + num_asgn_cols + 2 + (duration_weeks * 2)

    try:
        old_ws = sh.worksheet(tab)
    except gspread.exceptions.WorksheetNotFound:
        old_ws = None

    if old_ws:
        all_ws = sh.worksheets()
        if len(all_ws) > 1:
            print(f"  Worksheet '{tab}' exists. Deleting old version...")
            sh.del_worksheet(old_ws)
            print(f"  Creating worksheet '{tab}'...")
            ws = sh.add_worksheet(title=tab, rows=str(max(500, len(flat_rows)+10)), cols=str(total_cols + 2))
        else:
            print(f"  Worksheet '{tab}' is the only tab. Updating via temp worksheet...")
            temp_title = f"{tab}_new"
            ws = sh.add_worksheet(title=temp_title, rows=str(max(500, len(flat_rows)+10)), cols=str(total_cols + 2))
            sh.del_worksheet(old_ws)
            ws.update_title(tab)
    else:
        print(f"  Creating worksheet '{tab}'...")
        ws = sh.add_worksheet(title=tab, rows=str(max(500, len(flat_rows)+10)), cols=str(total_cols + 2))

    try:
        stale = sh.worksheet("Consolidate")
        sh.del_worksheet(stale)
        print("  Removed stale Consolidate tab.")
    except gspread.exceptions.WorksheetNotFound:
        pass

    grp_row1 = (
        ["Course Information"] * 2
        + ["Learner Information"] * 4
        + ["Activity Summary"] * 2
        + ["Assignment Submissions"] * (2 + num_asgn_cols)
        + ["Engagement Classification"] * 2
    )
    for w in range(1, duration_weeks + 1):
        grp_row1 += [f"Week {w}"] * 2

    hdr_row2 = [
        "Course ID", "Course Name",
        "Cohort", "Learner Name", "Official Email ID", "Enrollment Status",
        "Last Activity Timestamp", "Total Time Spent (HH:MM:SS)",
        "Total Assignments", "Completed Assignments"
    ]

    for a in assignments:
        a_name = a.get("name", "Assignment")
        due_at = a.get("due_at")
        if due_at and not isinstance(due_at, datetime.datetime):
            try:
                from canvas_engagement_tracker import parse_iso_datetime
                due_at = parse_iso_datetime(due_at)
            except Exception:
                due_at = None
        from canvas_engagement_tracker import format_deadline_header
        hdr_row2.append(f"{a_name} {format_deadline_header(due_at)}")

    hdr_row2.extend(["Engagement Category", "Overall Activity"])
    for w in range(1, duration_weeks + 1):
        hdr_row2 += [f"W{w} Duration (HH:MM:SS)", f"W{w} Status"]

    rows = [grp_row1, hdr_row2]
    for r in flat_rows:
        last_act = r.get("last_activity_timestamp")
        last_str = last_act.strftime("%Y-%m-%d %H:%M:%S") if (last_act and hasattr(last_act, "strftime")) else (str(last_act) if last_act else "N/A")
        subs_map = r.get("assignment_submissions", {})
        completed_count = r.get("submitted_assignments")
        if completed_count is None:
            completed_count = sum(1 for a in assignments if subs_map.get(a["id"]) == "Yes")

        row = [
            sis_id, course_name,
            r.get("cohort", "N/A"), r.get("name", "N/A"),
            r.get("email", "N/A"), r.get("status", "N/A"),
            last_str, _hms(r.get("total_engagement_seconds", 0)),
            r.get("total_assignments", len(assignments)),
            completed_count
        ]

        for a in assignments:
            row.append(subs_map.get(a["id"], "No"))

        row.extend([
            r.get("category", "N/A"), r.get("overall_activity", "N/A")
        ])

        wd = r.get("weekly_data", {})
        for w in range(1, duration_weeks + 1):
            wk = wd.get(w, {})
            t = wk.get("time_spent", "-")
            s = wk.get("status", "-")
            row.append(_hms(t) if isinstance(t, (int, float)) else str(t))
            row.append(str(s))
        rows.append(row)

    end_col_letter = _col_letter(total_cols)
    ws.update(range_name=f"A1:{end_col_letter}{len(rows)}", values=rows)

    print(f"  Formatting '{tab}'...")
    n_data = len(rows)

    with batch_updater(sh) as b:
        b.format_cell_range(ws, f"A1:{end_col_letter}1", CellFormat(backgroundColor=_color(C_NAVY), textFormat=TextFormat(bold=True, foregroundColor=_color(C_WHITE)), horizontalAlignment="CENTER", verticalAlignment="MIDDLE"))
        b.format_cell_range(ws, f"A2:{end_col_letter}2", CellFormat(backgroundColor=_color(C_SLATE), textFormat=TextFormat(bold=True, foregroundColor=_color(C_WHITE)), horizontalAlignment="CENTER", verticalAlignment="MIDDLE", wrapStrategy="CLIP"))

        if n_data > 2:
            for i in range(3, n_data + 1):
                bg = C_ALT_ROW if i % 2 == 0 else C_WHITE
                # General row format
                b.format_cell_range(ws, f"A{i}:{end_col_letter}{i}",
                                    CellFormat(backgroundColor=_color(bg),
                                               borders=_border(),
                                               horizontalAlignment="CENTER",
                                               verticalAlignment="MIDDLE"))
                # Left align Cohort, Name, Email (cols 3, 4, 5 -> C, D, E)
                b.format_cell_range(ws, f"C{i}:E{i}",
                                    CellFormat(backgroundColor=_color(bg),
                                               borders=_border(),
                                               horizontalAlignment="LEFT",
                                               verticalAlignment="MIDDLE"))

            for idx_a in range(num_asgn_cols):
                col_let = _col_letter(11 + idx_a)
                for i, r in enumerate(flat_rows, start=3):
                    a_id = assignments[idx_a]["id"]
                    sub_val = r.get("assignment_submissions", {}).get(a_id, "No")
                    bg_color = C_GREEN_BG if sub_val == "Yes" else C_RED_BG
                    fg_color = C_GREEN_FG if sub_val == "Yes" else C_RED_FG
                    b.format_cell_range(ws, f"{col_let}{i}",
                                        CellFormat(backgroundColor=_color(bg_color),
                                                   textFormat=TextFormat(bold=True, foregroundColor=_color(fg_color)),
                                                   horizontalAlignment="CENTER",
                                                   verticalAlignment="MIDDLE",
                                                   borders=_border()))

            cat_col_idx = 11 + num_asgn_cols
            cat_col = _col_letter(cat_col_idx)
            for i, r in enumerate(flat_rows, start=3):
                cat = r.get("category", "")
                style = CAT_STYLES.get(cat, {"bg": C_WHITE, "fg": C_SLATE})
                b.format_cell_range(ws, f"{cat_col}{i}",
                                    CellFormat(backgroundColor=_color(style["bg"]),
                                               textFormat=TextFormat(bold=True, foregroundColor=_color(style["fg"])),
                                               horizontalAlignment="CENTER",
                                               verticalAlignment="MIDDLE",
                                               borders=_border()))

            oa_col = _col_letter(cat_col_idx + 1)
            for i, r in enumerate(flat_rows, start=3):
                oa = r.get("overall_activity", "")
                bg = C_GREEN_BG if oa == "Active" else C_RED_BG
                fg = C_GREEN_FG if oa == "Active" else C_RED_FG
                b.format_cell_range(ws, f"{oa_col}{i}",
                                    CellFormat(backgroundColor=_color(bg),
                                               textFormat=TextFormat(bold=True, foregroundColor=_color(fg)),
                                               horizontalAlignment="CENTER",
                                               verticalAlignment="MIDDLE",
                                               borders=_border()))

            for w in range(1, duration_weeks + 1):
                status_col_idx = cat_col_idx + 1 + (w - 1) * 2 + 2
                status_col = _col_letter(status_col_idx)
                for i, r in enumerate(flat_rows, start=3):
                    wk = r.get("weekly_data", {}).get(w, {})
                    st = wk.get("status", "-")
                    style = WEEK_STYLES.get(str(st), WEEK_STYLES["-"])
                    b.format_cell_range(ws, f"{status_col}{i}",
                                        CellFormat(backgroundColor=_color(style["bg"]),
                                                   textFormat=TextFormat(foregroundColor=_color(style["fg"]), bold=True),
                                                   horizontalAlignment="CENTER",
                                                   verticalAlignment="MIDDLE",
                                                   borders=_border()))

    ws.freeze(rows=2)
    try:
        dim_reqs = []
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 2}, "properties": {"pixelSize": 180}, "fields": "pixelSize"}})
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 5}, "properties": {"pixelSize": 220}, "fields": "pixelSize"}})
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 5, "endIndex": 10}, "properties": {"pixelSize": 190}, "fields": "pixelSize"}})
        if num_asgn_cols > 0:
            dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 10, "endIndex": 10 + num_asgn_cols}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}})
        c_st = 10 + num_asgn_cols
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": c_st, "endIndex": c_st + 2}, "properties": {"pixelSize": 260}, "fields": "pixelSize"}})
        w_st = c_st + 2
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": w_st, "endIndex": w_st + (duration_weeks * 2)}, "properties": {"pixelSize": 190}, "fields": "pixelSize"}})
        sh.batch_update({"requests": dim_reqs})
    except Exception as dim_err:
        print(f"  Note: Column width adjustment: {dim_err}")

    try:
        stale_dash = sh.worksheet("Dashboard")
        sh.del_worksheet(stale_dash)
        print("  Removed stale Dashboard tab.")
    except Exception:
        pass

    try:
        stale_con = sh.worksheet("Consolidate")
        sh.del_worksheet(stale_con)
        print("  Removed stale Consolidate tab.")
    except Exception:
        pass

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


def _format_hms(s):
    s = int(s or 0)
    h = s // 3600
    m = (s % 3600) // 60
    sec = s % 60
    return f"{h:02d}:{m:02d}:{sec:02d}"


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
CACHE_TTL = 30  # 30-second memory cache TTL for instant sub-millisecond responses
CACHE_FILE = os.path.join(SCRIPT_DIR, "dashboard_cache.json")


def _fetch_ws(ws):
    try:
        return ws.title, ws.get_all_values()
    except Exception:
        return ws.title, None


def get_dashboard_data(force=False):
    import time as _t
    now = _t.time()

    # 1. Return in-memory cache if fresh and not forced
    if not force and _dashboard_cache["data"] and (now - _dashboard_cache["fetched_at"]) < CACHE_TTL:
        return _dashboard_cache["data"]

    # 2. Return disk cache if available and not forced
    if not force and os.path.exists(CACHE_FILE):
        try:
            mtime = os.path.getmtime(CACHE_FILE)
            if (now - mtime) < 120:  # Fresh disk cache within 2 minutes
                with open(CACHE_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    _dashboard_cache["data"] = data
                    _dashboard_cache["fetched_at"] = now
                    return data
        except Exception:
            pass

    # 3. Otherwise, fetch fresh live data from Google Sheets API
    print("  Fetching fresh dashboard data from Google Sheets API...")
    gc = get_google_client()
    sh = gc.open_by_key(SPREADSHEET_ID)

    # Read all course-specific tabs in parallel; skip any Consolidate tab
    all_data_rows = []   # list of (row_list, header_row) tuples
    course_tabs   = []
    with ThreadPoolExecutor(max_workers=10) as ex:
        for title, vals in ex.map(_fetch_ws, sh.worksheets()):
            if title == "Consolidate" or "Assignments" in title:
                continue          # ignore any leftover Consolidate tab
            if vals and len(vals) > 2:
                course_tabs.append(title)
                header_row = vals[1]  # row 2 (0-indexed: index 1) has column names
                for data_row in vals[2:]:  # data starts at row 3
                    all_data_rows.append((data_row, header_row))

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
    learner_profiles = {} # email/name -> profile detail
    weekly_agg   = {}   # week_num -> {met: N, not_met: N, total: N}
    course_weekly = {}  # course_id -> {week_num -> {met: N, not_met: N, total: N}}

    for row, header_row in all_data_rows:
        if len(row) < 9:
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

        if not learner_name:
            continue

        # ── Use header to find exact column positions ──────────────────────────
        # Header row contains: ..."Engagement Category", "Overall Activity", ...
        hdr_lower = [h.strip().lower() for h in header_row]
        try:
            overall_idx  = hdr_lower.index("overall activity")
            category_idx = hdr_lower.index("engagement category")
        except ValueError:
            # Fallback: dynamic calculation based on total_asgn
            # Columns: [0..8 fixed] + [9=completed] + [10..10+N-1 per-asgn Yes/No] + [cat] + [overall]
            category_idx = 10 + total_asgn
            overall_idx  = category_idx + 1

        overall  = row[overall_idx].strip()  if len(row) > overall_idx  else "N/A"
        category = row[category_idx].strip() if len(row) > category_idx else "N/A"

        # Count submissions and extract exact missed assignment module names
        asgn_start = 10
        asgn_end   = asgn_start + total_asgn
        asgn_cells = row[asgn_start:asgn_end] if len(row) >= asgn_end else row[asgn_start:]
        submitted  = sum(1 for c in asgn_cells if c.strip().lower() == "yes")
        missing    = sum(1 for c in asgn_cells if c.strip().lower() == "no")

        missed_assignments = []
        for idx_col, cell_val in enumerate(asgn_cells):
            if cell_val.strip().lower() == "no":
                col_header_idx = asgn_start + idx_col
                title = header_row[col_header_idx].strip() if len(header_row) > col_header_idx else f"Assignment {idx_col+1}"
                missed_assignments.append(title)

        # ── Read Weekly W1-W6 Status columns ──────────────────────────
        learner_weekly = {}
        for wi in range(1, 7):
            wk_status_hdr = f"w{wi} status"
            if wk_status_hdr in hdr_lower:
                wk_idx = hdr_lower.index(wk_status_hdr)
                wk_val = row[wk_idx].strip().upper() if len(row) > wk_idx else "-"
            else:
                wk_val = "-"
            learner_weekly[wi] = wk_val

            if wk_val in ("MET", "NOT MET"):
                if wi not in weekly_agg:
                    weekly_agg[wi] = {"met": 0, "not_met": 0, "total": 0}
                weekly_agg[wi]["total"] += 1
                if wk_val == "MET":
                    weekly_agg[wi]["met"] += 1
                else:
                    weekly_agg[wi]["not_met"] += 1

                # Per-course weekly aggregation
                if course_id not in course_weekly:
                    course_weekly[course_id] = {}
                if wi not in course_weekly[course_id]:
                    course_weekly[course_id][wi] = {"met": 0, "not_met": 0, "total": 0}
                course_weekly[course_id][wi]["total"] += 1
                if wk_val == "MET":
                    course_weekly[course_id][wi]["met"] += 1
                else:
                    course_weekly[course_id][wi]["not_met"] += 1

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
                "overdue": missing,  # overdue == missing for per-assignment format
                "last_act": last_act,
                "time_hms": time_str,
                "missed_assignments": missed_assignments,
            })


        cat_counts[category] = cat_counts.get(category, 0) + 1

        # Course breakdown
        if course_id not in course_map:
            course_map[course_id] = {
                "course_id": course_id,
                "name": course_name,
                "active": 0,
                "inactive": 0,
                "missed_submissions": 0,
                "total": 0,
                "missing": 0,
                "submitted": 0,
                "total_asgn": 0,
                "health_pct": 0
            }
        cm = course_map[course_id]
        cm["total"] += 1
        cm["missing"]    += missing
        cm["submitted"]  += submitted
        cm["total_asgn"] += total_asgn
        if overall == "Active":
            cm["active"] += 1
        else:
            cm["inactive"] += 1
        if category == "MISSED SUBMISSION":
            cm["missed_submissions"] += 1

        # Cohort breakdown
        if cohort not in cohort_map:
            cohort_map[cohort] = {
                "name": cohort,
                "active": 0,
                "inactive": 0,
                "missed_submissions": 0,
                "total": 0,
                "missing": 0,
                "total_time_s": 0,
                "valid_time": 0
            }
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
        if category == "MISSED SUBMISSION":
            coh["missed_submissions"] += 1
        # Populate learner profile data for search modal
        l_key = (email or learner_name).strip().lower()
        if l_key:
            if l_key not in learner_profiles:
                learner_profiles[l_key] = {
                    "learner_name": learner_name,
                    "email": email,
                    "courses": [],
                    "total_time_s": 0,
                    "submitted": 0,
                    "total_asgn": 0,
                    "overall": overall,
                    "category": category,
                    "missed_assignments": []
                }
            lp = learner_profiles[l_key]
            if course_name and course_name not in lp["courses"]:
                lp["courses"].append(course_name)
            lp["total_time_s"] += ts
            lp["submitted"] += submitted
            lp["total_asgn"] += total_asgn
            if overall == "Active":
                lp["overall"] = "Active"
            for ma in missed_assignments:
                if ma not in lp["missed_assignments"]:
                    lp["missed_assignments"].append(ma)

        # Only include learners with > 10 minutes (600s) time spent for content consumption leaderboard
        if ts >= 600:
            top_learners.append({
                "learner_name": learner_name,
                "email": email,
                "course_id": course_id,
                "course_name": course_name,
                "cohort": cohort,
                "time_seconds": ts,
                "time_hms": time_str,
                "submitted": submitted,
                "total_asgn": total_asgn,
                "category": category,
                "overall": overall,
                "last_act": last_act,
                "missed_assignments": missed_assignments,
            })

    avg_time_h = round((total_time_s / valid_time) / 3600, 2) if valid_time else 0

    # Sort top learners by time spent
    top_learners.sort(key=lambda x: x["time_seconds"], reverse=True)

    # Format learner profiles HMS strings
    for lk, lp in learner_profiles.items():
        lp["total_time_hms"] = _format_hms(lp["total_time_s"])
        avg_s = int(lp["total_time_s"] / max(1, len(lp["courses"])))
        lp["avg_time_hms"] = _format_hms(avg_s)

    # Process course health percentage
    for cid, cinfo in course_map.items():
        max_possible = cinfo["total"] * (cinfo["total_asgn"] / max(1, cinfo["total"]))
        if max_possible > 0:
            cinfo["health_pct"] = round(min(100.0, (cinfo["submitted"] / max(1, max_possible)) * 100.0), 1)
        else:
            cinfo["health_pct"] = 0.0

    # Process cohort map to calculate averages
    processed_cohorts = {}
    for cname, cinfo in cohort_map.items():
        avg_coh_t = round((cinfo["total_time_s"] / cinfo["valid_time"]) / 3600, 2) if cinfo["valid_time"] else 0
        processed_cohorts[cname] = {
            "name": cname,
            "total": cinfo["total"],
            "active": cinfo["active"],
            "inactive": cinfo["inactive"],
            "missed_submissions": cinfo["missed_submissions"],
            "missing": cinfo["missing"],
            "avg_time_hours": avg_coh_t
        }

    total_learners_cnt = active_ct + inactive_ct
    missed_submissions_cnt = cat_counts.get("MISSED SUBMISSION", 0)
    inactive_only_cnt = cat_counts.get("INACTIVE", 0) + cat_counts.get("INACTIVE - NOT LOGGED IN", 0)
    eng_ratio = round((active_ct / max(1, total_learners_cnt)) * 100.0, 1) if total_learners_cnt else 0.0

    # Build weekly summary list sorted by week number
    weekly_summary = []
    for wi in sorted(weekly_agg.keys()):
        wa = weekly_agg[wi]
        pct = round((wa["met"] / max(1, wa["total"])) * 100, 1)
        weekly_summary.append({
            "week": wi,
            "met": wa["met"],
            "not_met": wa["not_met"],
            "total": wa["total"],
            "met_pct": pct
        })

    # Attach per-course weekly data to course_map
    for cid, wk_data in course_weekly.items():
        if cid in course_map:
            cw_list = []
            for wi in sorted(wk_data.keys()):
                wd = wk_data[wi]
                pct = round((wd["met"] / max(1, wd["total"])) * 100, 1)
                cw_list.append({
                    "week": wi,
                    "met": wd["met"],
                    "not_met": wd["not_met"],
                    "total": wd["total"],
                    "met_pct": pct
                })
            course_map[cid]["weekly"] = cw_list

    payload = {
        "kpis": {
            "total_courses": len(total_courses_set),
            "total_learners": total_learners_cnt,
            "active_learners": active_ct,
            "missed_submission_learners": missed_submissions_cnt,
            "inactive_learners": inactive_ct,
            "total_missing": total_missing,
            "total_submitted": total_submitted,
            "avg_time_hours": avg_time_h,
            "engagement_ratio": eng_ratio
        },
        "cat_counts": cat_counts,
        "course_tabs": sorted(course_tabs),
        "course_breakdown": course_map,
        "cohort_breakdown": processed_cohorts,
        "weekly_summary": weekly_summary,
        "watchlist": watchlist[:100],
        "top_learners": top_learners[:100],
        "learner_profiles": learner_profiles
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
