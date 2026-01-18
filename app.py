from flask import Flask, render_template, request, abort, flash, redirect, url_for, jsonify, session
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

# Security: Generate random secret key (not hardcoded)
import secrets
app.secret_key = secrets.token_hex(32)

# Storage for temporary comparison results
comparison_results = {} 

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

        # Prepare commit history metadata (ensure no None values)
        commit_metadata = {
            'source_a': source_a,
            'branch_a': request.form.get("branch_a", "") or "",
            'path_a': request.form.get("path_a", "") or "",
            'url_a': request.form.get("url_a", "") or "",
        }

        return render_template(
            "excel_diff_result.html",
            diff=diff_result,
            total_diffs=total_row_changes,
            excel_a_name=display_name_a,
            excel_b_name=display_name_b,
            commit_metadata=commit_metadata,
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

@app.route("/api/commit-history", methods=["POST"])
def get_commit_history():
    """API endpoint to fetch commit history for a Git file."""
    try:
        data = request.get_json()
        branch = data.get("branch", "main")
        path = data.get("path", "")
        url = data.get("url")
        limit = min(int(data.get("limit", 20)), 100)  # Cap at 100
        
        if not path:
            return jsonify({"error": "Path is required"}), 400

        # Validate file extension - only Excel files supported
        _, ext = os.path.splitext(path or "")
        if ext.lower() not in ('.xlsx', '.xls'):
            return jsonify({"error": "Only .xlsx and .xls files are supported. Please provide the exact file path (including extension)."}), 400
        
        commits = GitReader.fetch_commit_history(branch, path, url, limit)
        
        # If no commits found, provide debugging info
        if not commits:
            return jsonify({
                "commits": [],
                "warning": f"No commits found. Try checking the branch name and file path. Branch: {branch}, Path: {path}"
            })
        
        return jsonify({"commits": commits})
    except Exception as e:
        error_msg = str(e)
        app.logger.error(f"COMMIT_HISTORY_ERROR: {error_msg}")
        return jsonify({"error": "Unable to load commit history", "details": "Check branch name (case-sensitive) and file path"}), 500

@app.route("/api/compare-with-commit", methods=["POST"])
def compare_with_commit():
    """API endpoint to compare two commits from Git history."""
    try:
        data = request.get_json()
        commit_hash_a = data.get("commit_hash_a")
        commit_hash_b = data.get("commit_hash_b")
        branch = data.get("branch", "main")
        path = data.get("path", "")
        url = data.get("url")
        
        if not commit_hash_a or not commit_hash_b or not path:
            return jsonify({"error": "Missing required parameters"}), 400
        
        # Create temp directory for fetching commit versions
        request_id = str(uuid.uuid4())
        request_folder = os.path.join(app.config["UPLOAD_FOLDER"], request_id)
        os.makedirs(request_folder, exist_ok=True)
        
        try:
            # Fetch both commit versions
            excel_a_path = GitReader.fetch_excel_by_commit(commit_hash_a, path, request_folder, url)
            excel_b_path = GitReader.fetch_excel_by_commit(commit_hash_b, path, request_folder, url)
            
            # Parse and compare
            parser_a = ExcelParser(excel_a_path)
            parser_b = ExcelParser(excel_b_path)
            data_a = parser_a.parse()
            data_b = parser_b.parse()
            
            diff_engine = DiffEngine(data_a, data_b)
            diff_result = diff_engine.compare()
            
            # Calculate total changes
            total_row_changes = 0
            for sheet in diff_result:
                if sheet.get('data') and 'rows' in sheet['data']:
                    for row in sheet['data']['rows']:
                        if any(cell.get('status') != 'equal' and (cell.get('a') or cell.get('b')) for cell in row.get('cells', [])):
                            total_row_changes += 1
            
            return jsonify({
                "diff": diff_result,
                "total_diffs": total_row_changes,
                "commit_info": {
                    "hash_a": commit_hash_a[:7],
                    "hash_b": commit_hash_b[:7],
                    "branch": branch
                }
            })
        finally:
            try:
                if request_folder and os.path.exists(request_folder):
                    shutil.rmtree(request_folder)
            except Exception as e:
                app.logger.error(f"Cleanup error: {e}")
    
    except Exception as e:
        app.logger.error(f"COMPARE_COMMIT_ERROR: {str(e)}\n{traceback.format_exc()}")
        return jsonify({"error": str(e)}), 500


@app.route('/compare-commits', methods=['GET'])
def compare_commits_page():
    """Render a full comparison page for two commits (opens in new tab)."""
    try:
        commit_a = request.args.get('commit_a')
        commit_b = request.args.get('commit_b')
        branch = request.args.get('branch', 'main')
        path = request.args.get('path', '')
        url = request.args.get('url')
        # Get File B metadata (defaults match File A if not provided)
        branch_b = request.args.get('branch_b', branch)
        path_b = request.args.get('path_b', path)
        url_b = request.args.get('url_b', url)

        if not commit_a or not commit_b or not path:
            flash('Missing parameters for commit comparison.', 'error')
            return redirect(url_for('excel_diff'))

        # Validate file extension early to avoid fetching non-excel blobs
        _, ext = os.path.splitext(path or "")
        if ext.lower() not in ('.xlsx', '.xls'):
            flash('Invalid file path: only .xlsx and .xls files are supported. Please provide the full file path including extension.', 'error')
            return redirect(url_for('excel_diff'))

        # Create temp dir
        request_id = str(uuid.uuid4())
        request_folder = os.path.join(app.config['UPLOAD_FOLDER'], request_id)
        os.makedirs(request_folder, exist_ok=True)

        try:
            excel_a_path = GitReader.fetch_excel_by_commit(commit_a, path, request_folder, url)
            excel_b_path = GitReader.fetch_excel_by_commit(commit_b, path, request_folder, url)

            parser_a = ExcelParser(excel_a_path)
            parser_b = ExcelParser(excel_b_path)
            data_a = parser_a.parse()
            data_b = parser_b.parse()

            diff_engine = DiffEngine(data_a, data_b)
            diff_result = diff_engine.compare()

            # Calculate total row-based changes
            total_row_changes = 0
            for sheet in diff_result:
                if sheet.get('data') and 'rows' in sheet['data']:
                    for row in sheet['data']['rows']:
                        if any(cell.get('status') != 'equal' and (cell.get('a') or cell.get('b')) for cell in row.get('cells', [])):
                            total_row_changes += 1

            # Ensure no None values in metadata
            commit_metadata = {
                'branch_a': branch or '',
                'path_a': path or '',
                'url_a': url or '',
                'branch_b': branch_b or '',
                'path_b': path_b or '',
                'url_b': url_b or ''
            }
            return render_template(
                'excel_diff_result.html',
                diff=diff_result,
                total_diffs=total_row_changes,
                excel_a_name=f'[Commit: {commit_a[:7]}]',
                excel_b_name=f'[Commit: {commit_b[:7]}]',
                commit_metadata=commit_metadata
            )
        finally:
            try:
                if request_folder and os.path.exists(request_folder):
                    shutil.rmtree(request_folder)
            except Exception as e:
                app.logger.error(f'Cleanup error: {e}')

    except Exception as e:
        app.logger.error(f'COMPARE_PAGE_ERROR: {e}\n{traceback.format_exc()}')
        flash('Error generating commit comparison. Check parameters and try again.', 'error')
        return redirect(url_for('excel_diff'))

@app.route('/api/compare-commit-local-unified', methods=['POST'])
def compare_commit_local_unified():
    """API endpoint: Compare Git commit with local uploaded file (unified)."""
    try:
        # Input validation with length limits
        MAX_HASH_LEN = 40
        MAX_BRANCH_LEN = 255
        MAX_PATH_LEN = 512
        MAX_URL_LEN = 1024
        
        hash_val = request.form.get('hash', '').strip()
        branch = request.form.get('branch', '').strip()
        path = request.form.get('path', '').strip()
        url = request.form.get('url', '').strip() or None
        hash_side = request.form.get('hash_side', 'a').strip()  # 'a' or 'b'
        local_file = request.files.get('local_file')
        # Get File B metadata
        branch_b = request.form.get('branch_b', '').strip()
        path_b = request.form.get('path_b', '').strip()
        url_b = request.form.get('url_b', '').strip() or None
        upload_file_b_name = request.form.get('upload_file_b_name', '').strip()

        # Validation
        if not hash_val or not local_file:
            return jsonify({"error": "Missing commit hash or local file"}), 400
        
        # Length validation
        if len(hash_val) > MAX_HASH_LEN or len(path) > MAX_PATH_LEN:
            return jsonify({"error": "Input parameters invalid"}), 400
        if len(branch) > MAX_BRANCH_LEN or (url and len(url) > MAX_URL_LEN):
            return jsonify({"error": "Input parameters too long"}), 400

        if not local_file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({"error": "Uploaded file must be .xlsx or .xls format"}), 400

        # Create temp directory
        request_id = str(uuid.uuid4())
        request_folder = os.path.join(app.config['UPLOAD_FOLDER'], request_id)
        os.makedirs(request_folder, exist_ok=True)

        try:
            # Fetch commit file and save local file
            commit_file_path = GitReader.fetch_excel_by_commit(hash_val, path, request_folder, url)
            local_file_path = os.path.join(request_folder, secure_filename(local_file.filename))
            local_file.save(local_file_path)

            # Parse both files
            parser_commit = ExcelParser(commit_file_path)
            parser_local = ExcelParser(local_file_path)
            data_commit = parser_commit.parse()
            data_local = parser_local.parse()

            # Run diff - commit on side A, local on side B (unless hash_side is 'b')
            if hash_side == 'a':
                diff_engine = DiffEngine(data_commit, data_local)
                excel_a_name = f'[Commit: {hash_val[:7]}]'
                excel_b_name = f'[Local: {local_file.filename}]'
            else:
                diff_engine = DiffEngine(data_local, data_commit)
                excel_a_name = f'[Local: {local_file.filename}]'
                excel_b_name = f'[Commit: {hash_val[:7]}]'

            diff_result = diff_engine.compare()

            # Calculate total changes
            total_row_changes = 0
            for sheet in diff_result:
                if sheet.get('data') and 'rows' in sheet['data']:
                    for row in sheet['data']['rows']:
                        if any(cell.get('status') != 'equal' and (cell.get('a') or cell.get('b')) for cell in row.get('cells', [])):
                            total_row_changes += 1

            # Store result and return a redirect URL
            result_id = str(uuid.uuid4())
            comparison_results[result_id] = {
                "diff": diff_result,
                "total_diffs": total_row_changes,
                "excel_a_name": excel_a_name,
                "excel_b_name": excel_b_name,
                "commit_metadata": {'branch_a': branch, 'path_a': path, 'url_a': url, 'branch_b': branch_b, 'path_b': path_b, 'url_b': url_b}
            }
            
            return jsonify({
                "success": True,
                "redirect_url": f"/view-result/{result_id}"
            }), 200

        except Exception as e:
            app.logger.error(f'Unified compare error: {e}\n{traceback.format_exc()}')
            return jsonify({"error": "Comparison failed"}), 500

        finally:
            # Cleanup temp folder
            try:
                shutil.rmtree(request_folder)
            except:
                pass

    except Exception as e:
        app.logger.error(f'API_COMPARE_UNIFIED_ERROR: {e}\n{traceback.format_exc()}')
        return jsonify({"error": "Server error"}), 500
@app.route('/view-result/<result_id>', methods=['GET'])
def view_comparison_result(result_id):
    """Render a stored comparison result."""
    try:
        if result_id not in comparison_results:
            flash('Comparison result not found or expired.', 'error')
            return redirect(url_for('excel_diff'))
        
        result = comparison_results[result_id]
        
        # Clean up after viewing (optional - comment out if you want to keep results longer)
        # del comparison_results[result_id]
        
        return render_template(
            'excel_diff_result.html',
            diff=result.get('diff', []),
            total_diffs=result.get('total_diffs', 0),
            excel_a_name=result.get('excel_a_name', 'File A'),
            excel_b_name=result.get('excel_b_name', 'File B'),
            commit_metadata=result.get('commit_metadata', {})
        )
    except Exception as e:
        app.logger.error(f'VIEW_RESULT_ERROR: {e}\n{traceback.format_exc()}')
        flash('Error loading comparison result.', 'error')
        return redirect(url_for('excel_diff'))

@app.route('/comparison-result', methods=['POST'])
def comparison_result_page():
    """Render a full comparison page from JSON data."""
    try:
        data = request.get_json()
        return render_template(
            'excel_diff_result.html',
            diff=data.get('diff', []),
            total_diffs=data.get('total_diffs', 0),
            excel_a_name=data.get('excel_a_name', 'File A'),
            excel_b_name=data.get('excel_b_name', 'File B'),
            commit_metadata=data.get('commit_metadata', {})
        )
    except Exception as e:
        app.logger.error(f'COMPARISON_RESULT_ERROR: {e}\n{traceback.format_exc()}')
        return jsonify({"error": "Failed to render comparison result"}), 500

if __name__ == "__main__":
    app.run(debug=True)