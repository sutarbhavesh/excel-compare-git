from flask import Flask, render_template, request, abort, flash, redirect, url_for
import os
import time
import logging
import traceback
import uuid
import shutil 
from werkzeug.utils import secure_filename

from excel_diff.excel_parser import ExcelParser
from excel_diff.diff_engine import DiffEngine
from excel_diff.git_reader import GitReader 

app = Flask(__name__)
app.secret_key = "super_secret_key_for_flash_messages" 

current_dir = os.path.dirname(os.path.abspath(__file__))
if os.path.basename(current_dir).lower() == "internal":
    PROJECT_ROOT = os.path.dirname(current_dir)
else:
    PROJECT_ROOT = current_dir

# --- LOGGING SETUP ---
LOG_FOLDER = os.path.join(PROJECT_ROOT, "logs")
os.makedirs(LOG_FOLDER, exist_ok=True)

log_format = (
    "\n" + "#"*80 + "\n"
    "TIMESTAMP: %(asctime)s\n"
    "ERROR TYPE: %(levelname)s\n"
    "MESSAGE: %(message)s\n"
    "#"*80 + "\n"
)

logging.basicConfig(
    filename=os.path.join(LOG_FOLDER, 'app_error.log'),
    level=logging.ERROR,
    format=log_format
)

# --- UPLOAD SETUP ---
BASE_UPLOAD_FOLDER = os.path.join(PROJECT_ROOT, "uploads", "temp")
os.makedirs(BASE_UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = os.path.abspath(BASE_UPLOAD_FOLDER)

# SECURITY: Limit maximum upload size to 16MB
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024

@app.errorhandler(413)
def request_entity_too_large(error):
    return "File is too large! Please upload a file smaller than 16MB.", 413

@app.route("/", methods=["GET", "POST"])
def excel_diff():
    if request.method == "GET":
        return render_template("excel_diff.html")

    # 1. Create a unique folder for THIS specific request/tab
    request_id = str(uuid.uuid4())
    request_folder = os.path.join(app.config["UPLOAD_FOLDER"], request_id)
    os.makedirs(request_folder, exist_ok=True)

    excel_a_path = None
    excel_b_path = None

    try:
        source_a = request.form.get("source_a", "pc")
        source_b = request.form.get("source_b", "pc")

        # Resolve Excel A
        if source_a == "pc":
            file_a = request.files.get("file_a")
            if not file_a or file_a.filename == "":
                flash("Excel A file missing", "error")
                return redirect(url_for('excel_diff'))
            
            display_name_a = file_a.filename
            filename = secure_filename(file_a.filename)
            excel_a_path = os.path.join(request_folder, f"a_{filename}")
            file_a.save(excel_a_path)
        else:  
            branch_a = request.form.get("branch_a", "N/A")
            path_a = request.form.get("path_a", "")
            git_filename_a = os.path.basename(path_a)
            display_name_a = f"[Git: {branch_a}] {git_filename_a}"
            
            excel_a_path = GitReader.fetch_excel(
                branch=branch_a,
                path=path_a,
                url=request.form.get("url_a"), 
                target_dir=request_folder
            )

        # Resolve Excel B
        if source_b == "pc":
            file_b = request.files.get("file_b")
            if not file_b or file_b.filename == "":
                flash("Excel B file missing", "error")
                return redirect(url_for('excel_diff'))
            
            display_name_b = file_b.filename
            filename = secure_filename(file_b.filename)
            excel_b_path = os.path.join(request_folder, f"b_{filename}")
            file_b.save(excel_b_path)
        else:  
            branch_b = request.form.get("branch_b", "N/A")
            path_b = request.form.get("path_b", "")
            git_filename_b = os.path.basename(path_b)
            display_name_b = f"[Git: {branch_b}] {git_filename_b}"
            
            excel_b_path = GitReader.fetch_excel(
                branch=branch_b,
                path=path_b,
                url=request.form.get("url_b"),
                target_dir=request_folder
            )

        # Parse Excel files (Handles .xls and .xlsx via your updated parser)
        parser_a = ExcelParser(excel_a_path)
        parser_b = ExcelParser(excel_b_path)
        data_a = parser_a.parse()
        data_b = parser_b.parse()

        # Diff Engine Logic
        diff_engine = DiffEngine(data_a, data_b)
        diff_result = diff_engine.compare()

        # Calculate total row-based changes
        total_row_changes = 0
        for sheet in diff_result:
            if sheet.get('data') and 'rows' in sheet['data']:
                for row in sheet['data']['rows']:
                    if any(cell.get('status') != 'equal' and (cell.get('a') or cell.get('b')) for cell in row.get('cells', [])):
                        total_row_changes += 1

        return render_template(
            "excel_diff_result.html",
            diff=diff_result,
            total_diffs=total_row_changes,
            excel_a_name=display_name_a,
            excel_b_name=display_name_b,
        )

    except Exception as e:
        error_info = traceback.format_exc()
        app.logger.error(f"DIFF_ERROR: {str(e)}\n{error_info}")
        
        if "git" in str(e).lower():
            friendly_msg = "Could not access the Git repository. Please check your URL, Branch, and Path."
        elif "permission" in str(e).lower():
            friendly_msg = "The system could not access the file. It might be open in another program."
        else:
            friendly_msg = "An unexpected error occurred during comparison. Please verify your inputs."
            
        flash(friendly_msg, "error")
        return redirect(url_for('excel_diff'))

    finally:
        try:
            if request_folder and os.path.exists(request_folder):
                shutil.rmtree(request_folder)
        except Exception as e:
            app.logger.error(f"Cleanup error for {request_folder}: {e}")

if __name__ == "__main__":
    app.run(debug=True)