from flask import Flask, render_template, request
import subprocess
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
REPORT_FOLDER = "reports"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/compare", methods=["POST"])
def compare():

    file1 = request.files["file1"]
    file2 = request.files["file2"]
    output = request.form["output"]

    # Ensure .xlsx extension
    if not output.endswith(".xlsx"):
        output += ".xlsx"

    path1 = os.path.join(
        UPLOAD_FOLDER,
        file1.filename
    )

    path2 = os.path.join(
        UPLOAD_FOLDER,
        file2.filename
    )

    output_path = os.path.join(
        REPORT_FOLDER,
        output
    )

    file1.save(path1)
    file2.save(path2)

    result = subprocess.run(
        [
            "python",
            "compare_excel.py",
            path1,
            path2,
            output_path
        ],
        capture_output=True,
        text=True
    )

    if result.returncode != 0:
        return f"""
        <h2>Error</h2>
        <pre>{result.stderr}</pre>
        """

    return f"""
    <!DOCTYPE html>
    <html>
    <head>
        <title>Comparison Complete</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                text-align: center;
                margin-top: 50px;
            }}

            .btn {{
                display: inline-block;
                padding: 12px 20px;
                margin: 10px;
                text-decoration: none;
                color: white;
                border-radius: 5px;
                font-weight: bold;
            }}

            .download {{
                background-color: #28a745;
            }}

            .back {{
                background-color: #007bff;
            }}
        </style>
    </head>
    <body>

        <h2>✅ Comparison Completed</h2>

        <p>Report Generated:</p>
        <p><b>{output_path}</b></p>

        <a class="btn download" href="/reports/{output}">
            Download Report
        </a>

        <a class="btn back" href="/">
            Back to Home
        </a>

    </body>
    </html>
    """


@app.route("/reports/<filename>")
def download_report(filename):
    from flask import send_from_directory

    return send_from_directory(
        REPORT_FOLDER,
        filename,
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(debug=True)