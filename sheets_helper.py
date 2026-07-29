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
    "NOT ACTIVE":                          {"bg": C_RED_BG,   "fg": C_RED_FG},
    "In Active":                           {"bg": C_BLUE_BG,  "fg": C_BLUE_FG},
    "No Submission":                       {"bg": C_GREY_BG,  "fg": C_GREY_FG},
    "NO ACTIVITY":                         {"bg": C_BLUE_BG,  "fg": C_BLUE_FG},
    "LOW ACTIVITY - ASSIGNMENTS COMPLETED":{"bg": C_GREY_BG,  "fg": C_GREY_FG},
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
    total_cols = 9 + num_asgn_cols + 2 + (duration_weeks * 2)

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
        + ["Assignment Submissions"] * (1 + num_asgn_cols)
        + ["Engagement Classification"] * 2
    )
    for w in range(1, duration_weeks + 1):
        grp_row1 += [f"Week {w}"] * 2

    hdr_row2 = [
        "Course ID", "Course Name",
        "Cohort", "Learner Name", "Official Email ID", "Enrollment Status",
        "Last Activity Timestamp", "Total Time Spent (HH:MM:SS)",
        "Total Assignments"
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
        row = [
            sis_id, course_name,
            r.get("cohort", "N/A"), r.get("name", "N/A"),
            r.get("email", "N/A"), r.get("status", "N/A"),
            last_str, _hms(r.get("total_engagement_seconds", 0)),
            r.get("total_assignments", len(assignments))
        ]

        subs_map = r.get("assignment_submissions", {})
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
                b.format_cell_range(ws, f"A{i}:{end_col_letter}{i}",
                                    CellFormat(backgroundColor=_color(bg),
                                               borders=_border(),
                                               verticalAlignment="MIDDLE"))

            for idx_a in range(num_asgn_cols):
                col_let = _col_letter(10 + idx_a)
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

            cat_col_idx = 10 + num_asgn_cols
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
    ws.columns_auto_resize(0, total_cols)
    try:
        dim_reqs = []
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 0, "endIndex": 2}, "properties": {"pixelSize": 180}, "fields": "pixelSize"}})
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 2, "endIndex": 6}, "properties": {"pixelSize": 220}, "fields": "pixelSize"}})
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 6, "endIndex": 9}, "properties": {"pixelSize": 180}, "fields": "pixelSize"}})
        if num_asgn_cols > 0:
            dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": 9, "endIndex": 9 + num_asgn_cols}, "properties": {"pixelSize": 280}, "fields": "pixelSize"}})
        c_st = 9 + num_asgn_cols
        dim_reqs.append({"updateDimensionProperties": {"range": {"sheetId": ws.id, "dimension": "COLUMNS", "startIndex": c_st, "endIndex": c_st + 2}, "properties": {"pixelSize": 280}, "fields": "pixelSize"}})
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
            if title == "Consolidate" or "Assignments" in title:
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

        if len(row) > 9 and row[9].strip().isdigit():
            # Legacy format with numeric submission counts
            submitted = _safe_int(row[9])
            missing   = _safe_int(row[10])
            overdue   = _safe_int(row[11])
            category  = row[13].strip() if len(row) > 13 else "N/A"
            overall   = row[14].strip() if len(row) > 14 else "N/A"
        else:
            # New format with per-assignment Yes/No columns
            asgn_cells = row[9 : 9 + total_asgn] if len(row) >= 9 + total_asgn else row[9:]
            submitted  = sum(1 for c in asgn_cells if c.strip().lower() == "yes")
            missing    = sum(1 for c in asgn_cells if c.strip().lower() == "no")
            overdue    = missing
            cat_idx    = 9 + total_asgn
            category   = row[cat_idx].strip() if len(row) > cat_idx else "N/A"
            overall    = row[cat_idx + 1].strip() if len(row) > cat_idx + 1 else "N/A"

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
