# Hostel Reservation Arranger

A small Flask web app that takes a reservation file (`.xlsx` or `.csv`), splits it into one sheet per booking site, and returns an Excel workbook with a total at the bottom of each sheet. (just for presentation only)

## What it does

1. You upload an `.xlsx` or `.csv` file in the browser.
2. Rows are grouped by the `reservation site` column.
3. Each site gets its own worksheet, named after the site.
4. Column widths are adjusted to fit the longest value in each column.
5. A `SUM` formula is added below the last row of the `amount` column.
6. The workbook downloads automatically as `MMDD TL.xlsx` (for example `0917 TL.xlsx` on 17 September).

## Requirements

- Python 3.9 or later
- Flask, pandas and openpyxl (listed in `requirements.txt`)

## Installation

```bash
git clone https://github.com/MoeMA551/reservation_automation.git
cd reservation_automation
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Usage

```bash
python app.py
```

Open http://127.0.0.1:5000, choose your file and click **Arrange reservations**. The processed workbook downloads immediately.

## Input format

| File type | Notes |
| --------- | ----- |
| `.xlsx`   | Only the first sheet is read. |
| `.csv`    | UTF-8 (with or without BOM) or Windows-1252, the format English Excel uses when saving as plain CSV. |

The header must be in row 1.

| Column             | Required | Purpose                                    |
| ------------------ | -------- | ------------------------------------------ |
| `reservation site` | Yes      | Rows are grouped into sheets by this value |
| `amount`           | Yes      | Totalled at the bottom of each sheet       |
| Any other column   | No       | Copied to the output unchanged             |

Column names must match exactly, including lowercase letters and spaces. Values in `amount` must be plain numbers (`1000`, not `1,000` or `¥1000`), otherwise Excel treats them as text and the total will be wrong.

### Example

Input (`uploads/reservation.csv`):

```csv
reservation site,guest name,amount
agoda,David,1000
booking.com,Tom,2000
agoda,Chloe,3000
```

Output, sheet `agoda`:

| reservation site | guest name | amount               |
| ---------------- | ---------- | -------------------- |
| agoda            | David      | 1000                 |
| agoda            | Chloe      | 3000                 |
|                  |            | `=SUM(C2:C3)` → 4000 |

Sheet names follow Excel's rules: the characters `: \ / ? * [ ]` are removed and names are cut to 31 characters.

## How it works

All logic is in `app.py`.

### Routes

| Route      | Method | Description |
| ---------- | ------ | ----------- |
| `/`        | GET    | Shows the upload page (`templates/index.html`) |
| `/process` | POST   | Takes the multipart field `file` and returns the workbook, or a plain-text error with status 400 |

### Processing steps in `/process`

1. Checks that a file was sent and that its extension is `.xlsx` or `.csv`.
2. Saves the upload to `uploads/` with a timestamp name, for example `uploads/20260917_101500.csv`. The original filename is not used, so a crafted name can't write outside the folder.
3. Reads the file with `read_file()`.
4. Checks that the `reservation site` and `amount` columns exist.
5. Writes one sheet per site to `outputs/MMDD TL.xlsx`, then sends that file as the download. Running it again on the same day overwrites the file.

### Functions

| Function | Description |
| -------- | ----------- |
| `read_file(path, extension)` | Reads `.xlsx` with `pd.read_excel`. Reads `.csv` as UTF-8, falling back to Windows-1252. |
| `clean_sheet_name(name)` | Removes characters Excel doesn't allow in sheet names and cuts the name to 31 characters. |
| `autofit_columns(worksheet, data)` | Sets each column's width to its longest value (or header) plus 2 characters. Empty cells count as zero length. |
| `add_sum_row(worksheet, data, column)` | Writes a `=SUM(...)` formula in the row directly below the data. |

### Settings

The constants at the top of `app.py` control the app's behaviour:

| Constant | Default | Purpose |
| -------- | ------- | ------- |
| `UPLOAD_FOLDER` | `"uploads"` | Where uploaded files are saved |
| `OUTPUT_FOLDER` | `"outputs"` | Where generated workbooks are saved |
| `ALLOWED_EXTENSIONS` | `{".xlsx", ".csv"}` | File types the app accepts |
| `GROUP_COLUMN` | `"reservation site"` | Column used to split rows into sheets |
| `SUM_COLUMN` | `"amount"` | Column that gets a total |
| `MAX_SHEET_NAME_LENGTH` | `31` | Excel's sheet name limit |
| `INVALID_SHEET_CHARS` | `[:\\/?*\[\]]` | Characters removed from sheet names |
| `COLUMN_PADDING` | `2` | Extra width added to each column |

To use different column names, change `GROUP_COLUMN` and `SUM_COLUMN`.

Both folders are created automatically when the app starts.

## Project structure

```
reservation_automation/
├── app.py               # Flask app: routes and file processing
├── requirements.txt     # Python dependencies
├── templates/
│   └── index.html       # Upload page
├── static/
│   └── style.css        # Page styles
├── uploads/             # Uploaded files (created automatically)
└── outputs/             # Generated workbooks (created automatically)
```

## Troubleshooting

| Message or symptom | Cause | Fix |
| ------------------ | ----- | --- |
| `Error: no file selected.` | The form was submitted without a file | Choose a file before clicking the button |
| `Error: please upload a .xlsx or .csv file.` | Wrong file type, e.g. `.xls` or `.txt` | Save the file as `.xlsx` or `.csv` |
| `Error: missing column(s): ...` | A required header is missing or spelled differently | Rename the header to exactly `reservation site` or `amount` |
| Internal Server Error | The file is damaged or isn't really the format its extension says | Open the file in Excel and save it again |
| Garbled text in the output | The CSV uses an encoding other than UTF-8 or Windows-1252, e.g. saved from Japanese Excel | Save it as **CSV UTF-8** in Excel |
| Total shows 0 or is wrong | `amount` contains text such as `1,000` or currency symbols | Use plain numbers in `amount` |
