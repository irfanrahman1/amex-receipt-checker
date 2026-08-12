# AmEx Receipt Checker

Simple Python tool that checks an American Express Excel workbook against receipt files in a folder.

It marks each charge as:

- `Acquired` when a receipt is found
- `Needs Review` when there is a possible match
- `Missing` when no receipt is found

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

## Test Run

This repo includes one test Excel file:

```text
examples/demo/test_charges.xlsx
```

Run the demo:

```powershell
.\.venv\Scripts\python receipt_audit.py "examples\demo\test_charges.xlsx" --receipts-dir "examples\demo\receipts"
```

The program creates:

```text
examples/demo/test_charges_receipt_audit.xlsx
examples/demo/test_charges_audit_report.csv
examples/demo/test_charges_audit_report_unmatched_receipts.csv
examples/demo/test_charges_merged_receipts.pdf
```

## Use Your Own Files

Put the Excel workbook and receipts in the same folder, then run:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx"
```

If the receipts are in a different folder:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx" --receipts-dir "C:\Path\To\Receipts"
```

## Helpful Options

```powershell
--date-window-days 5
--rename-receipts
--move-to-done
--batch-root "C:\Path\To\American Express"
--gui
```

Receipt renaming uses:

```text
Vendor - Invoice - Date.ext
```

## Tests

```powershell
python -m unittest discover -s tests -v
```
