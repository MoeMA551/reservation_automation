from flask import Flask, render_template, request, send_file
import pandas as pd 
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/")
def render():
    return render_template("index.html")

@app.route("/process", methods=["POST"])
def process_file():

    if "file" not in request.files:
        return "No file uploaded."

    file = request.files["file"]

    if file.filename == "":
        return "No file selected."

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    df = pd.read_excel(input_path)

    if "reservation site" not in df.columns:
        return "Error: 'reservation site' column was not found."
    output_path = os.path.join(
        OUTPUT_FOLDER,
        "modified_reservation.xlsx"
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        for site, site_data in df.groupby("reservation site"):

            sheet_name = re.sub(
                r'[:\\/?*\[\]]',
                '',
                str(site)
            )

            sheet_name = sheet_name[:31]

            site_data.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )
        
        return send_file(
        output_path,
        as_attachment=True,
        download_name="modified_reservation.xlsx"
    )

if __name__== "__main__":
    app.run(debug=True)