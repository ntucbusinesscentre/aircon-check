"""
Aircon Reconciliation Web App
Flask backend that accepts file uploads, runs the original aircon_check_final.py
reconciliation, and serves results.
"""

import os
import re
import sys
import uuid
import shutil
from pathlib import Path

import io

from flask import Flask, request, jsonify, send_file, render_template
from openpyxl import load_workbook

# Lazy import of reconcile() — deferred to first use so app can start even if deps have issues
_reconcile = None

def get_reconcile():
    global _reconcile
    if _reconcile is None:
        try:
            from aircon_check_final import reconcile
            _reconcile = reconcile
        except ImportError:
            PARENT_DIR = str(Path(__file__).resolve().parent.parent)
            if PARENT_DIR not in sys.path:
                sys.path.insert(0, PARENT_DIR)
            from aircon_check_final import reconcile
            _reconcile = reconcile
    return _reconcile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, "templates"),
    static_folder=os.path.join(BASE_DIR, "static"),
)

# Store temp session data: session_id -> {folder_path, output_path, output_filename}
SESSIONS = {}

# Base temp directory — use system temp on Railway, local subfolder otherwise
import tempfile
TEMP_BASE = Path(tempfile.gettempdir()) / "aircon_reconciliation"
TEMP_BASE.mkdir(exist_ok=True)


def parse_summary_text(summary_text):
    """Parse the text summary returned by reconcile() into a structured dict."""
    def extract_int(pattern, text, default=0):
        m = re.search(pattern, text)
        return int(m.group(1)) if m else default

    def extract_float(pattern, text, default=0.0):
        m = re.search(pattern, text)
        return float(m.group(1)) if m else default

    return {
        "matched": extract_int(r"Matched cleanly:\s*(\d+)", summary_text),
        "zero_charge": extract_int(r"Zero-charge \(review\):\s*(\d+)", summary_text),
        "unclear": extract_int(r"Unclear \(review\):\s*(\d+)", summary_text),
        "mismatch": extract_int(r"Amount mismatch:\s*(\d+)", summary_text),
        "missing": extract_int(r"Missing from PDFs:\s*(\d+)", summary_text),
        "total_landlord_rows": extract_int(r"Landlord rows:\s*(\d+)", summary_text),
        "unbilled_pdfs": extract_int(r"PDF bookings not billed:\s*(\d+)", summary_text),
    }


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/health")
def health():
    """Debug endpoint to check if all imports work."""
    status = {"status": "ok", "imports": {}}
    for mod_name in ["flask", "pdfplumber", "openpyxl", "aircon_check_final"]:
        try:
            __import__(mod_name)
            status["imports"][mod_name] = "ok"
        except Exception as e:
            status["imports"][mod_name] = str(e)
            status["status"] = "error"
    return jsonify(status)


@app.route("/upload", methods=["POST"])
def upload():
    """Accept uploaded files, run reconciliation, return summary JSON."""
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded."}), 400

    # Create a unique session folder
    session_id = uuid.uuid4().hex[:12]
    session_folder = TEMP_BASE / session_id
    session_folder.mkdir(parents=True, exist_ok=True)

    try:
        # Save uploaded files
        saved_count = {"pdf": 0, "xlsx": 0}
        log_lines = []

        for f in files:
            if not f.filename:
                continue
            # Get just the filename (strip any directory path from folder uploads)
            filename = Path(f.filename).name
            # Skip temp Excel files and already-reconciled outputs
            if filename.startswith("~$") or filename.startswith("Reconciled_"):
                continue
            ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
            if ext == "pdf":
                saved_count["pdf"] += 1
            elif ext == "xlsx":
                saved_count["xlsx"] += 1
            else:
                continue  # Skip non-relevant files

            save_path = session_folder / filename
            f.save(str(save_path))
            log_lines.append(f"  Uploaded: {filename}")

        if saved_count["pdf"] == 0:
            shutil.rmtree(session_folder, ignore_errors=True)
            return jsonify({"error": "No PDF files found in your upload. Please include the Form F PDFs."}), 400
        if saved_count["xlsx"] == 0:
            shutil.rmtree(session_folder, ignore_errors=True)
            return jsonify({"error": "No Excel (.xlsx) file found. Please include the landlord's monthly summary."}), 400

        # Run the original reconcile() function, capturing its stdout
        captured = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = captured
        try:
            summary_text, output_path = get_reconcile()(str(session_folder))
        finally:
            sys.stdout = old_stdout
        captured_output = captured.getvalue()

        # Parse the text summary into structured data
        result = parse_summary_text(summary_text)

        # Extract output filename
        output_filename = os.path.basename(str(output_path))
        result["output_filename"] = output_filename

        # Build log from captured stdout + summary
        log_lines.append("")
        log_lines.append(captured_output)
        log_lines.append(summary_text)
        result["log_lines"] = log_lines

        # Read the generated output Excel's Summary sheet for accurate values
        try:
            wb_out = load_workbook(str(output_path), data_only=True)
            ws_summary = wb_out["Summary"]
            for row in ws_summary.iter_rows(min_row=2, max_col=2, values_only=True):
                label, value = row
                if label and "Total PDF bookings parsed" in str(label):
                    result["total_pdf_bookings"] = int(value) if value else 0
                elif label and "Total billed amount" in str(label):
                    result["total_billed"] = round(float(value), 2) if value else 0.0
            wb_out.close()
        except Exception:
            # Fallback if we can't read the output Excel
            result.setdefault("total_pdf_bookings", 0)
            result.setdefault("total_billed", 0.0)

        # Store session info for download
        SESSIONS[session_id] = {
            "folder_path": str(session_folder),
            "output_path": str(output_path),
            "output_filename": output_filename,
        }

        result["session_id"] = session_id
        result["download_url"] = f"/download/{session_id}"

        return jsonify(result)

    except Exception as e:
        # Clean up on error
        shutil.rmtree(session_folder, ignore_errors=True)
        return jsonify({"error": str(e)}), 500


@app.route("/download/<session_id>")
def download(session_id):
    """Serve the reconciled Excel file for download."""
    session = SESSIONS.get(session_id)
    if not session:
        return jsonify({"error": "Session expired or not found. Please re-upload your files."}), 404

    output_path = session["output_path"]
    if not os.path.exists(output_path):
        return jsonify({"error": "Output file no longer exists. Please re-upload your files."}), 404

    return send_file(
        output_path,
        as_attachment=True,
        download_name=session["output_filename"],
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/cleanup/<session_id>", methods=["POST"])
def cleanup(session_id):
    """Clean up temp files for a session."""
    session = SESSIONS.pop(session_id, None)
    if session:
        shutil.rmtree(session["folder_path"], ignore_errors=True)
    return jsonify({"ok": True})


if __name__ == "__main__":
    # Clean up any leftover temp files from previous runs
    if TEMP_BASE.exists():
        for child in TEMP_BASE.iterdir():
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=True)
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("RAILWAY_ENVIRONMENT") is None  # debug only locally
    print(f"\n  Aircon Reconciliation Web App")
    print(f"  Open http://localhost:{port} in your browser\n")
    app.run(debug=debug, host="0.0.0.0", port=port)
