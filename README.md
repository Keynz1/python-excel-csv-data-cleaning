# Python Excel/CSV Data Cleaning & Validation

Self-created demonstration project for recurring spreadsheet cleanup work.

## What it does
- Reads mixed CSV and XLSX files with inconsistent header casing/spacing.
- Normalizes text, email, amount, and status fields.
- Merges records deterministically.
- Removes duplicates by normalized ID using a documented stable winner rule.
- Flags missing/invalid email, required fields, and invalid numeric amounts.
- Exports matching clean CSV/XLSX files plus a source-row issues report.

## Run
```powershell
python generate_sample_inputs.py
python run_demo.py
python -m pytest -v
```

## Demo result
8 input rows в†’ 7 unique rows; 1 duplicate removed; 4 validation issues detected (invalid email, invalid amount, missing email, missing name). CSV/XLSX outputs matched exactly and 5 automated tests passed.

All sample data is generated for this self-created demonstration project. No private or client data is included.

