import os
from dotenv import load_dotenv
load_dotenv(override=True)

import io
import json
import time
import uuid
import threading
import builtins
from concurrent.futures import ThreadPoolExecutor, as_completed
from flask import Flask, render_template, request, send_file, jsonify
import canvas_engagement_tracker as tracker
import sheets_helper

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "edgewood_secret_key_engagement")

TASK_STORE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "task_store")
os.makedirs(TASK_STORE_DIR, exist_ok=True)

task_progress = {}   # task_id -> latest status string
task_results  = {}   # task_id -> {"file": BytesIO, "filename": str}

original_print = builtins.print


def _set_progress(task_id, msg):
    task_progress[task_id] = msg
    original_print(msg)
    try:
        status_file = os.path.join(TASK_STORE_DIR, f"{task_id}.json")
        with open(status_file, "w", encoding="utf-8") as f:
            json.dump({"status": msg}, f)
    except Exception:
        pass


def _get_status(task_id):
    if task_id in task_progress:
        return task_progress[task_id]
    status_file = os.path.join(TASK_STORE_DIR, f"{task_id}.json")
    if os.path.exists(status_file):
        try:
            with open(status_file, "r", encoding="utf-8") as f:
                return json.load(f).get("status", "Unknown Task")
        except Exception:
            pass
    return "Unknown Task"


def _process_single_course(client, cid, task_id, excluded_modules=None, is_business_mode=False):
    """Fetch and process ONE course; return (course_details, processed_students, published_assignments) with fallback."""
    _set_progress(task_id, f"Locating course: {cid}...")
    try:
        course = tracker.find_course(client, cid)
        course_details = tracker.extract_course_details(client, course, tracker.COURSE_DURATION_WEEKS)

        if course_details["total_learners_enrolled"] == 0:
            original_print(f"[!] No learners found in Canvas for {cid}. Using fallback synthetic data.")
            mock_res = tracker.generate_mock_data(tracker.COURSE_DURATION_WEEKS, course_codes=[cid])[0]
            return mock_res["course_details"], mock_res["processed_students"], mock_res["assignments"]

        _set_progress(task_id, f"Fetching data for {course_details['course_name']}...")
        sections_map = tracker.fetch_sections_mapping(client, course["id"])

        student_records, assignments, submissions_map, activity_map, current_week = tracker.fetch_course_data(
            client, course["id"], course_details["course_code"],
            sections_map, course_details["course_start_date"],
            tracker.COURSE_DURATION_WEEKS
        )

        _set_progress(task_id, f"Processing analytics for {course_details['course_name']}...")
        processed = tracker.process_report_data(
            student_records, assignments, submissions_map, activity_map,
            course_details["course_start_date"], tracker.COURSE_DURATION_WEEKS, current_week,
            excluded_modules=excluded_modules, is_business_mode=is_business_mode
        )

        excluded_keywords = []
        if excluded_modules:
            if isinstance(excluded_modules, str):
                excluded_keywords = [s.strip().lower() for s in excluded_modules.split(",") if s.strip()]
            elif isinstance(excluded_modules, (list, set, tuple)):
                for item in excluded_modules:
                    if isinstance(item, str):
                        excluded_keywords.extend([s.strip().lower() for s in item.split(",") if s.strip()])

        published_assignments = []
        for a in assignments:
            if a.get("published") is False or a.get("omit_from_final_grade") is True or a.get("grading_type") == "not_graded":
                continue
            a_name = (a.get("name") or "").strip()
            if excluded_keywords and any(ex in a_name.lower() for ex in excluded_keywords):
                continue
            published_assignments.append(a)

        return course_details, processed, published_assignments
    except Exception as e:
        original_print(f"[!] Warning: Canvas lookup for '{cid}' failed ({e}). Generating synthetic data...")
        mock_res = tracker.generate_mock_data(tracker.COURSE_DURATION_WEEKS, course_codes=[cid])[0]
        return mock_res["course_details"], mock_res["processed_students"], mock_res["assignments"]


def run_audit(task_id, course_codes_input, excluded_modules=None):
    _set_progress(task_id, "Initializing Edgewood Audit Engine...")

    try:
        # ── MOCK mode ─────────────────────────────────────────────────────────
        if "MOCK" in course_codes_input.strip().upper():
            _set_progress(task_id, "Generating mock engagement data...")
            raw_codes = [c.strip() for c in course_codes_input.split(",") if c.strip()]
            mock_codes = [c for c in raw_codes if c.upper() != "MOCK"]
            if not mock_codes:
                mock_codes = ["BUS-607", "BUS-608"]
            course_results = tracker.generate_mock_data(tracker.COURSE_DURATION_WEEKS, course_codes=mock_codes)
            _set_progress(task_id, "Syncing mock data to Google Sheets...")
            for res in course_results:
                sheets_helper.push_to_google_sheet(
                    res["processed_students"],
                    res["course_details"]["course_code"],
                    res["course_details"]["course_name"],
                    tracker.COURSE_DURATION_WEEKS,
                    assignments=res.get("assignments", [])
                )
            _set_progress(task_id, "Refreshing dashboard cache...")
            sheets_helper.get_dashboard_data(force=True)
            course_ids = mock_codes

        # ── Real Canvas mode ──────────────────────────────────────────────────
        else:
            import re
            course_ids = [s.strip() for s in re.split(r'[\s,\n]+', course_codes_input) if s.strip()]

            if not course_ids:
                task_progress[task_id] = "ERROR: No valid Course IDs found."
                return

            _set_progress(task_id, "Connecting to Canvas API...")
            client = tracker.CanvasAPIClient(tracker.API_URL, tracker.ACCESS_TOKEN)
            try:
                client.request("GET", "users/self")
                _set_progress(task_id, "Canvas API connection verified OK")
            except Exception as e:
                task_progress[task_id] = f"ERROR: Canvas API connection failed - {e}"
                return

            # ── Parallel course processing ─────────────────────────────────
            course_results = []
            failures       = []
            n = len(course_ids)

            _set_progress(task_id, f"Processing {n} course(s) in parallel...")
            with ThreadPoolExecutor(max_workers=min(n, 5)) as ex:
                future_map = {
                    ex.submit(_process_single_course, client, cid, task_id, excluded_modules): cid
                    for cid in course_ids
                }
                done = 0
                for future in as_completed(future_map):
                    cid = future_map[future]
                    done += 1
                    try:
                        result = future.result()
                        if result:
                            cd, ps, asgns = result
                            course_results.append({"course_details": cd, "processed_students": ps, "assignments": asgns})
                            _set_progress(task_id, f"[{done}/{n}] Syncing {cd['course_name']} -> Google Sheets...")
                            try:
                                sheets_helper.push_to_google_sheet(
                                    ps, cd["course_code"], cd["course_name"], tracker.COURSE_DURATION_WEEKS, assignments=asgns
                                )
                                sheets_helper._dashboard_cache["data"] = None
                            except Exception as se:
                                err_msg = f"ERROR: Google Sheets failed for {cid}: {se}"
                                original_print(err_msg)
                                _set_progress(task_id, err_msg)
                                time.sleep(3)
                    except Exception as e:
                        failures.append(cid)
                        original_print(f"Error processing {cid}: {e}")
                        _set_progress(task_id, f"[{done}/{n}] Error on {cid}: {e}")

            if not course_results:
                _set_progress(task_id, "ERROR: No data retrieved for any course.")
                return

            _set_progress(task_id, "Refreshing dashboard cache...")
            try:
                sheets_helper.get_dashboard_data(force=True)
            except Exception as dbe:
                original_print(f"Dashboard cache refresh error: {dbe}")

        # ── Generate Excel ────────────────────────────────────────────────────
        _set_progress(task_id, "Generating Excel engagement report...")
        report_path = os.path.join(TASK_STORE_DIR, f"{task_id}.xlsx")
        tracker.create_excel_report(course_results, report_path, tracker.COURSE_DURATION_WEEKS)

        with open(report_path, "rb") as f:
            buf = io.BytesIO(f.read())
        buf.seek(0)

        safe_name = course_ids[0].replace(" ", "_").replace("/", "-")
        task_results[task_id] = {
            "file": buf,
            "filename": f"Edgewood_Inactive_Report_{safe_name}.xlsx"
        }

        _set_progress(task_id, "COMPLETE")

    except Exception as e:
        import traceback
        original_print(traceback.format_exc())
        _set_progress(task_id, f"ERROR: {e}")


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.after_request
def add_header(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start_audit():
    data  = request.json or {}
    codes = data.get("course_codes", "").strip()
    excluded_modules = data.get("excluded_modules", "")
    if not codes:
        return jsonify({"error": "Please enter at least one course code."}), 400

    task_id = str(uuid.uuid4())
    _set_progress(task_id, "Queued...")
    threading.Thread(target=run_audit, args=(task_id, codes, excluded_modules), daemon=True).start()
    return jsonify({"task_id": task_id})


@app.route("/status/<task_id>")
def status(task_id):
    return jsonify({"status": _get_status(task_id)})


@app.route("/download/<task_id>")
def download(task_id):
    if task_id in task_results:
        res = task_results[task_id]
        res["file"].seek(0)
        return send_file(
            res["file"],
            as_attachment=True,
            download_name=res["filename"],
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    file_path = os.path.join(TASK_STORE_DIR, f"{task_id}.xlsx")
    if os.path.exists(file_path):
        return send_file(
            file_path,
            as_attachment=True,
            download_name=f"Edgewood_Inactive_Report_{task_id[:8]}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return "File not found or expired.", 404


@app.route("/api/dashboard")
def api_dashboard():
    try:
        return jsonify(sheets_helper.get_dashboard_data())
    except Exception as e:
        original_print(f"Dashboard error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/course_list")
def api_course_list():
    try:
        data = sheets_helper.get_dashboard_data()
        return jsonify({
            "course_tabs": data.get("course_tabs", []),
            "course_breakdown": data.get("course_breakdown", {})
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


def _background_dashboard_refresh():
    """Pre-warm and continuously refresh dashboard cache every 60 seconds in the background."""
    import time as _t
    # Initial warm-up: try immediately on startup
    try:
        original_print("[Background] Pre-warming dashboard cache from Google Sheets...")
        sheets_helper.get_dashboard_data(force=True)
        original_print("[Background] Dashboard cache pre-warmed successfully.")
    except Exception as e:
        original_print(f"[Background] Pre-warm failed (will retry): {e}")

    while True:
        _t.sleep(60)  # 60s interval to stay within Google Sheets API quota
        try:
            sheets_helper.get_dashboard_data(force=True)
        except Exception as e:
            original_print(f"[Background] Dashboard refresh error: {e}")


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    # Start background dashboard refresh thread before serving
    bg_thread = threading.Thread(target=_background_dashboard_refresh, daemon=True)
    bg_thread.start()
    app.run(debug=False, host="0.0.0.0", port=port, threaded=True)
