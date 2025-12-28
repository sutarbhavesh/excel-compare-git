from flask import Flask, render_template, request
import os
from werkzeug.utils import secure_filename

from excel_diff.excel_parser import ExcelParser
from excel_diff.diff_engine import DiffEngine
from excel_diff.git_reader import GitReader 

app = Flask(__name__)

UPLOAD_FOLDER = "uploads/temp"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


@app.route("/excel-diff", methods=["GET", "POST"])
def excel_diff():
    if request.method == "GET":
        return render_template("excel_diff.html")

    source_a = request.form.get("source_a", "pc")
    source_b = request.form.get("source_b", "pc")

    excel_a_path = None
    excel_b_path = None

    # Resolve Excel A
    if source_a == "pc":
        file_a = request.files.get("file_a")
        if not file_a or file_a.filename == "":
            return "Excel A file missing", 400

        filename = secure_filename(file_a.filename)
        excel_a_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file_a.save(excel_a_path)

    else:  
        branch = request.form.get("branch_a")
        path = request.form.get("path_a")
        url = request.form.get("url_a")

        excel_a_path = GitReader.fetch_excel(
        branch=branch,
        path=path,
        url=url, 
        target_dir=UPLOAD_FOLDER
        )


    # Resolve Excel B
    if source_b == "pc":
        file_b = request.files.get("file_b")
        if not file_b or file_b.filename == "":
            return "Excel B file missing", 400

        filename = secure_filename(file_b.filename)
        excel_b_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file_b.save(excel_b_path)

    else:  
        branch = request.form.get("branch_b")
        path = request.form.get("path_b")
        url = request.form.get("url_b")

        excel_b_path = GitReader.fetch_excel(
            branch=branch,
            path=path,
            url=url,
            target_dir=UPLOAD_FOLDER
        )


    # Parse Excel files using your shared classes
    parser_a = ExcelParser(excel_a_path)
    parser_b = ExcelParser(excel_b_path)
    data_a = parser_a.parse()
    data_b = parser_b.parse()

    # Clean up files immediately to keep server light
    if os.path.exists(excel_a_path): os.remove(excel_a_path)
    if os.path.exists(excel_b_path): os.remove(excel_b_path)

    diff_engine = DiffEngine(data_a, data_b)
    diff_result = diff_engine.compare()

    stats = {"modified": 0, "added": 0, "deleted": 0}
    for sheet_name, sheet_data in diff_result.items():
        for row in sheet_data.get('rows', []):
            statuses = [c['status'] for c in row['cells']]
            if "modified" in statuses: stats["modified"] += 1
            elif "added" in statuses: stats["added"] += 1
            elif "deleted" in statuses: stats["deleted"] += 1

    return render_template(
        "excel_diff_result.html",
        diff=diff_result,
        diff_stats=stats,
        excel_a_name=os.path.basename(excel_a_path),
        excel_b_name=os.path.basename(excel_b_path),
    )

@app.after_request
def cleanup_temp_files(response):
    try:
        folder = app.config["UPLOAD_FOLDER"]
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            # Only delete files older than 5 minutes to avoid deleting 
            # files while another user is currently parsing them
            import time
            if os.path.isfile(file_path) and time.time() - os.path.getmtime(file_path) > 300:
                os.remove(file_path)
    except Exception as e:
        print(f"Error during cleanup: {e}")
    return response
    
if __name__ == "__main__":
    app.run(debug=True)
