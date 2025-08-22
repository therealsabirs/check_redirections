import os
import requests
import pandas as pd
from flask import Flask, request, render_template, send_from_directory
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ✅ Function to check URL
def check_url(url):
    try:
        response = requests.get(url, timeout=10, allow_redirects=True)
        final_url = response.url if response.history else url
        return response.status_code, final_url
    except Exception:
        return "Error", url

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return "No file uploaded", 400
    
    file = request.files["file"]
    if file.filename == "":
        return "No file selected", 400
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)

    # Read Excel
    df = pd.read_excel(filepath)

    # Ensure at least 7 columns exist (A–G)
    while df.shape[1] < 7:
        df[df.shape[1]] = ""

    results = []
    for i in range(len(df)):
        actual_url = str(df.iloc[i, 1]).strip()   # Column B (Actual URL)
        expected_url = str(df.iloc[i, 2]).strip() # Column C (Expected URL)

        actual_status, redirect_url = check_url(actual_url)
        expected_status, _ = check_url(expected_url)

        df.iloc[i, 3] = actual_status   # Column D → Status Actual
        df.iloc[i, 4] = expected_status # Column E → Status Expected
        df.iloc[i, 5] = redirect_url    # Column F → Redirected To

        status_msg = "Success" if redirect_url == expected_url else "Failed"
        df.iloc[i, 6] = status_msg      # Column G → Final Result

        results.append({
            "actual": actual_url,
            "expected": expected_url,
            "status_actual": actual_status,
            "status_expected": expected_status,
            "redirect_to": redirect_url,
            "final_status": status_msg
        })

    output_path = os.path.join(UPLOAD_FOLDER, "output_" + filename)
    df.to_excel(output_path, index=False, header=[
        "Col A", "Actual URL", "Expected URL", "Status Actual",
        "Status Expected", "Redirected To", "Result"
    ])

    return render_template("results.html", results=results, download_file="output_" + filename)


@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(UPLOAD_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
