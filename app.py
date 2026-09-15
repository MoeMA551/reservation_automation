from flask import Flask, render_template, request, send_file
import pandas as pd 
import os
import re
from openpyxl.utils import get_column_letter

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

#implement autofit_column feature
def autofit_column(worksheet, data):
    for i,col in enumerate(data.columns, start=1):
        longest = max(data[col].astype(str).str.len().max(), #longest value in column
                      len(str(col)))#header length
        worksheet.column_dimensions[get_column_letter(i)].width = longest + 2

#implement autosum feature
def auto_sum(worksheet, data, column="amount"):
    colum_letter = get_column_letter(data.columns.get_loc(column) + 1)
    last_row = len(data) + 1
    worksheet[f"{colum_letter}{last_row + 1}"] = f"=SUM({colum_letter}2:{colum_letter}{last_row})"
    print(f"{colum_letter}{last_row + 1}")

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

    #Reading data from uploaded excel file
    df = pd.read_excel(input_path)

    if "reservation site" not in df.columns:
        return "Error: 'reservation site' column was not found."
    
    output_path = os.path.join(
        OUTPUT_FOLDER,
        "modified_reservation.xlsx"
    )

    #Writing output excel file
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        #Groupby reservaton site
        for site, site_data in df.groupby("reservation site"):

            #remove unallowed characters
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
            autofit_column(writer.sheets[sheet_name],site_data)
            auto_sum(writer.sheets[sheet_name],site_data)

    #autodownloading the file
    return send_file(
    output_path,
    as_attachment=True,
    download_name="modified_reservation.xlsx"
    )

if __name__== "__main__":
    app.run(debug=True)