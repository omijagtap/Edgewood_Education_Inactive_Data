import os
from dotenv import load_dotenv
load_dotenv(override=True)

import io
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

task_progress = {}   # task_id -> latest status string
task_results  = {}   # task_id -> {"file": BytesIO, "filename": str}

original_print = builtins.print


def _set_progress(task_id, msg):
    task_progress[task_id] = msg
    original_print(msg)


def _process_single_course(client, cid, task_id):
    """Fetch and process ONE course; return (course_details, processed_students) or raise."""
    _set_progress(task_id, f"Locating course: {cid}...")
    course = tracker.find_course(client, cid)
    course_details = tracker.extract_course_details(client, course, tracker.COURSE_DURATION_WEEKS)

    if course_details["total_learners_enrolled"] == 0:
        _set_progress(task_id, f"[Skip] No learners found in {cid}.")
        return None

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
        course_details["course_start_date"], tracker.COURSE_DURATION_WEEKS, current_week
    )

    return course_details, processed


def run_audit(task_id, course_codes_input):
    _set_progress(task_id, "Initializing Edgewood Audit Engine...")

    try:
        # ── MOCK mode ─────────────────────────────────────────────────────────
        if course_codes_input.strip().upper() == "MOCK":
            _set_progress(task_id, "Generating mock engagement data...")
            course_results = tracker.generate_mock_data(tracker.COURSE_DURATION_WEEKS)
            _set_progress(task_id, "Syncing mock data to Google Sheets...")
            for res in course_results:
                sheets_helper.push_to_google_sheet(
                    res["processed_students"],
                    res["course_details"]["course_code"],
                    res["course_details"]["course_name"],
                    tracker.COURSE_DURATION_WEEKS
                )
            _set_progress(task_id, "Refreshing dashboard cache...")
            sheets_helper.get_dashboard_data(force=True)
            course_ids = ["MOCK"]

        # ── Real Canvas mode ──────────────────────────────────────────────────
        else:
            # Parse input (comma-separated or space-separated)
            if "," in course_codes_input:
                course_ids = [s.strip() for s in course_codes_input.split(",") if s.strip()]
            else:
                course_ids = [s.strip() for s in course_codes_input.split() if s.strip()]

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
                    ex.submit(_process_single_course, client, cid, task_id): cid
                    for cid in course_ids
                }
                done = 0
                for future in as_completed(future_map):
                    cid = future_map[future]
                    done += 1
                    try:
                        result = future.result()
                        if result:
                            cd, ps = result
                            course_results.append({"course_details": cd, "processed_students": ps})
                            _set_progress(task_id, f"[{done}/{n}] Syncing {cd['course_name']} -> Google Sheets...")
                            try:
                                sheets_helper.push_to_google_sheet(
                                    ps, cid, cd["course_name"], tracker.COURSE_DURATION_WEEKS
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
                task_progress[task_id] = "ERROR: No data retrieved for any course."
                return

            _set_progress(task_id, "Refreshing dashboard cache...")
            try:
                sheets_helper.get_dashboard_data(force=True)
            except Exception as dbe:
                original_print(f"Dashboard cache refresh error: {dbe}")

        # ── Generate Excel ────────────────────────────────────────────────────
        _set_progress(task_id, "Generating Excel engagement report...")
        tmp = f"_report_{task_id}.xlsx"
        tracker.create_excel_report(course_results, tmp, tracker.COURSE_DURATION_WEEKS)

        with open(tmp, "rb") as f:
            buf = io.BytesIO(f.read())
        os.remove(tmp)
        buf.seek(0)

        safe_name = course_ids[0].replace(" ", "_").replace("/", "-")
        task_results[task_id] = {
            "file": buf,
            "filename": f"Edgewood_Inactive_Report_{safe_name}.xlsx"
        }

        task_progress[task_id] = "COMPLETE"

    except Exception as e:
        import traceback
        original_print(traceback.format_exc())
        task_progress[task_id] = f"ERROR: {e}"


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/start", methods=["POST"])
def start_audit():
    data  = request.json or {}
    codes = data.get("course_codes", "").strip()
    if not codes:
        return jsonify({"error": "Please enter at least one course code."}), 400

    task_id = str(uuid.uuid4())
    task_progress[task_id] = "Queued..."
    threading.Thread(target=run_audit, args=(task_id, codes), daemon=True).start()
    return jsonify({"task_id": task_id})


@app.route("/status/<task_id>")
def status(task_id):
    return jsonify({"status": task_progress.get(task_id, "Unknown Task")})


@app.route("/download/<task_id>")
def download(task_id):
    if task_id in task_results:
        res = task_results[task_id]
        return send_file(
            res["file"],
            as_attachment=True,
            download_name=res["filename"],
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


if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port, threaded=True)
