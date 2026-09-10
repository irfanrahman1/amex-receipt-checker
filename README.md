# AmEx Receipt Checker

Simple Python tool that checks an American Express Excel workbook against receipt files in a folder.

It marks each charge as:

- `Acquired` when a receipt is found
- `Needs Review` when there is a possible match
- `Missing` when no receipt is found

It also looks up the **Vendor**, **Account Code**, and **Account Description** for each charge automatically -- see [Vendor & Account Code Lookup](#vendor--account-code-lookup) below.

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
```

## Quick Start (recommended)

The easiest way to run this is with the file-picker window -- no paths to type:

```powershell
.\.venv\Scripts\python receipt_audit.py --gui
```

A window will open asking you to choose the AmEx workbook, then the receipts folder. Everything else happens automatically.

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

The simplest way is `--gui` (see Quick Start above). If you'd rather type the path directly:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx"
```

If the receipts are in a different folder:

```powershell
.\.venv\Scripts\python receipt_audit.py "C:\Path\To\AMEX Transaction Details.xlsx" --receipts-dir "C:\Path\To\Receipts"
```

## Vendor & Account Code Lookup

Every charge automatically gets a **Vendor**, **Account Code**, and **Account Description** filled in, using three sources in order:

1. **`vendor_lookup_reference.xlsx`** -- the official, manually-curated vendor list. Checked first.
2. **`historical_vendor_database.xlsx`** -- built from past AmEx files. If we've bought from a vendor before, it's used here automatically, no review needed.
3. **`category_suggestions.xlsx`** -- for brand-new vendors not found anywhere above, a best guess is made using AmEx's own transaction Category, but only for categories that were highly consistent historically. These rows are flagged in the **Coding Note** column ("Best guess (new vendor) - please review") so they can be double-checked.

All three files live in this folder alongside `receipt_audit.py` and grow over time as new vendors are added.

### Keeping the historical database updated

Once you've reviewed a month's audited file and the flagged best-guess rows look correct, run:

```powershell
.\.venv\Scripts\python update_historical_database.py --gui
```

(or pass a file path directly instead of `--gui`)

This finds every flagged best-guess row in that file and adds it to `historical_vendor_database.xlsx`. From then on, that vendor is recognized automatically -- no more guessing, no manual copy-pasting.

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
