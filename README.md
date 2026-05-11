# Score Report Generator

Flask app for converting reader-machine Excel data into score reports and an APP paste-in worksheet.

## Features

- Upload `.xlsx` or `.xlsm` reader-machine files.
- Read student name from column F and multiple-choice score from column J.
- Add students manually, including leave/absence records.
- Calculate total score with `選擇 / 25 * 85 + 非選 / 6 * 15`.
- Generate the formatted official score report.
- Generate an APP paste-in Excel file using a pasted name order.

## Run Locally

```powershell
pip install -r requirements.txt
python index.py
```

Then open `http://127.0.0.1:5000`.

## Deploy

This repo includes `vercel.json` for Vercel's Python runtime. The app entry point is `index.py`.
