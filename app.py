from flask import Flask, render_template, request, send_file
import pandas as pd 
import os
import re
from datetime import datetime
from openpyxl.utils import get_column_letter

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

remove_sheetnames = ["クレジットカード", "現金"]

#implement autofit_column feature
def autofit_column(worksheet, data):
    for i,col in enumerate(data.columns, start=1):
        longest = max(data[col].astype(str).str.len().max(), #longest value in column
                      len(str(col)))#header length
        worksheet.column_dimensions[get_column_letter(i)].width = longest + 10

#implement autosum feature
def auto_sum(worksheet, data, column):
    column_letter = get_column_letter(data.columns.get_loc(column) + 1)
    print(column_letter)
    last_row = len(data) + 1
    print(last_row)
    worksheet[f"{column_letter}{last_row + 1}"] = f"=SUM({column_letter}2:{column_letter}{last_row})"
    print(f"{column_letter}{last_row + 1}")

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
    try:
        df = pd.read_csv(input_path, encoding="utf-8-sig")
    except UnicodeDecodeError:
        df = pd.read_csv(input_path, encoding="cp932")

    if "予約サイト名" in df.columns:
        sheet_type = "TL"
        group_column = "予約サイト名"
        
    elif "科目名" in df.columns:
        sheet_type = "SB"
        group_column = "科目名"
        
    else:
        return "Error: 'reservation site' column was not found."

    if sheet_type == "TL":
        filename = f"{datetime.now():%m%d} TL.xlsx"
    elif sheet_type == "SB":
        filename = f"{datetime.now():%m%d} SB.xlsx"
    
    output_path = os.path.join(
        OUTPUT_FOLDER,
        filename
    )

    #Writing output excel file
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:

        #Groupby reservaton site
        for site, site_data in df.groupby(group_column):

            #remove unallowed characters
            sheet_name = re.sub(
                r'[:\\/?*\[\]]',
                '',
                str(site)
            )

            sheet_name = sheet_name[:31]
            if sheet_name in remove_sheetnames:
                continue
            site_data.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )
            autofit_column(writer.sheets[sheet_name],site_data)
            if sheet_type == "TL":
                auto_sum(writer.sheets[sheet_name],site_data,"合計料金")
                auto_sum(writer.sheets[sheet_name],site_data,"請求料金")
                auto_sum(writer.sheets[sheet_name],site_data,"ポイント")
            elif sheet_type == "SB":
                auto_sum(writer.sheets[sheet_name],site_data,"支払額")

    #autodownloading the file
    return send_file(
    output_path,
    as_attachment=True,
    download_name= filename
    )

if __name__== "__main__":
    app.run(debug=True)