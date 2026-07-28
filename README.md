# AmEx Receipt Checker

AmEx Receipt Checker is a Python utility that audits an American Express transaction workbook against receipt files in the same folder tree.

It helps answer one practical question: **which charges already have receipts, and which ones still need follow-up?**

## What It Does

The program:

- opens an AmEx Excel transaction workbook
- finds the correct transaction sheet automatically
- scans the receipt folder and subfolders
- matches receipts using amount, date, and vendor text
- allows a date window for late-posted AmEx charges
- writes a new audited Excel workbook
- color-codes each row by status
- creates CSV review reports
- optionally renames matched receipts
- optionally moves matched receipts into a `DONE` folder
- optionally uses OCR on image receipts
- creates one merged PDF packet of all acquired receipts with a cover page and table of contents
- can process one workbook, a whole folder batch, or a simple file-picker GUI

The original workbook is not overwritten unless you explicitly choose the same output path.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

OCR is optional. To use `--ocr`, install the Tesseract app for Windows in addition to the Python requirements.

## Basic Run

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx"
```

By default, receipts are searched in the same folder as the workbook.

If receipts are in a different folder:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx" --receipts-dir "C:\Path\To\Receipts"
```

## Finished-Folder Workflow

To match the completed manual workflow, rename matched receipts and move them into `DONE`:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx" --receipts-dir "C:\Path\To\Month Folder" --rename-receipts --move-to-done
```

Renamed receipts use this format:

```text
Vendor - Invoice - Date.ext
```

Example:

```text
Example Cloud - INV-1001 - 2026-07-01.pdf
```

## Outputs

The program creates an audited workbook:

```text
AMEX Transaction Details_receipt_audit.xlsx
```

It adds these columns:

- `Receipt Status`
- `Matched Receipt`
- `Receipt Match Notes`

Rows are marked:

- `Acquired` when a matching receipt is found
- `Needs Review` when the program found a possible receipt but the match is not strong enough to trust automatically
- `Missing` when no strong receipt match is found

If the workbook already has a `Receipt` column, acquired rows are marked `yes` and missing rows are left blank.

The program also creates these files:

```text
AMEX Transaction Details_audit_report.csv
AMEX Transaction Details_audit_report_unmatched_receipts.csv
```

The audit report lists every charge, its status, score, matched receipt, and match notes. The unmatched receipt report lists receipt files that were not used by any acquired row.

The program also creates a merged PDF packet of acquired receipts:

```text
AMEX Transaction Details_merged_receipts.pdf
```

The PDF starts with a cover page and table of contents, then appends the acquired receipt files.

To skip the merged PDF:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx" --no-merge-pdf
```

## Matching Rules

The matcher checks:

- exact charge amount when an amount is available
- transaction date, with a default 3-day date window
- vendor or description text

The date window helps when the receipt date and AmEx posted date are a few days apart.

Change the date window with:

```powershell
--date-window-days 5
```

Possible but uncertain matches can be controlled with:

```powershell
--review-threshold 0.40
```

Duplicate receipt candidates are called out in the match notes so those rows can be reviewed before filing.

## Batch Run

To process every transaction workbook under a folder:

```powershell
.\.venv\Scripts\python receipt_audit.py --batch-root "C:\Path\To\American Express"
```

The batch mode looks for workbook names containing `Transaction Details` and creates outputs beside each workbook.

## Simple GUI

To run with file pickers:

```powershell
.\.venv\Scripts\python receipt_audit.py --gui
```

The GUI asks for the workbook and receipt folder, then runs the standard rename, `DONE`, report, and merged PDF workflow.

## Example

This repository includes one sanitized before/after example:

```text
examples/before/example_charges_before.xlsx
examples/after/example_charges_after.xlsx
```

The example also includes:

```text
examples/before/receipts/
examples/after/DONE/
examples/after/example_acquired_receipts.pdf
```

The example uses fake data only.

## Tests

```powershell
python -m unittest discover -s tests -v
```
