# AmEx Receipt Checker

This is a small Python utility for auditing an American Express charge workbook against receipt files in the same folder.

It opens the Excel file, looks for charge rows, checks receipt filenames and text-based PDFs in the receipt folder, and writes a copy of the workbook with these columns:

- `Receipt Status`: `Acquired` or `Missing`
- `Matched Receipt`: the receipt file that matched
- `Receipt Match Notes`: the score and signals used for the match

The original workbook is not overwritten unless you explicitly pass it as `--output`.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

## Run

If the workbook is inside the same folder as the receipts:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AP Interns 2026\Margaret Jing\American Express\charges.xlsx"
```

If receipts are in a different folder:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\charges.xlsx" --receipts-dir "C:\Path\To\Receipts"
```

The output defaults to a new file beside the workbook, such as:

```text
charges_receipt_audit.xlsx
```

## Matching Notes

The checker looks for:

- charge amount
- charge date
- merchant or description text

It scans common receipt formats including PDF, PNG, JPG, JPEG, HEIC, WEBP, TIF, and TIFF. For PDFs, it can read embedded text when the PDF contains text. Scanned image-only PDFs and photos are matched by filename only.

If the workbook headers are unusual, rename them to common labels such as `Date`, `Description`, and `Amount`.

The workbook does not need to open directly on the transaction sheet. The checker searches worksheets for transaction-style headers and prefers a sheet named `Transaction Details` when present.

## Tests

```powershell
python -m unittest discover -s tests -v
```
