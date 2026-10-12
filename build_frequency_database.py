#!/usr/bin/env python3
"""Build frequency_database.xlsx from historical AmEx statement files.

This is the one-time (or "re-run when you get more history") build step for
the Frequency feature Thomas asked for -- same idea as how
historical_vendor_database.xlsx was originally built from ~17 months of past
statements, but this one keeps a row per transaction (Vendor, Date, Amount)
instead of collapsing to one row per vendor, because frequency classification
needs the actual date spacing between charges.

Point this at the same folder of monthly "Transaction Details" statement
files already used to build the vendor lookup database -- it reuses
receipt_audit.py's own worksheet/header detection so it finds the same sheet
and columns that tool already knows how to read.

Usage:
    python build_frequency_database.py --statements-root "path\\to\\AmEx Statements"

Re-run any time you have more history (a new month closes out) to fold it
in -- existing rows for a vendor/date/amount combination already in the
database are skipped, so re-running on the same files is safe.
"""

import argparse
import datetime as dt
import sys
from pathlib import Path
from typing import List, Tuple

from openpyxl import Workbook, load_workbook

from receipt_audit import (
    find_charge_sheet,
    infer_target_year,
    iter_charges,
)
from vendor_lookup import extract_frequency_key


def find_statement_workbooks(root: Path) -> List[Path]:
    """Find monthly statement workbooks under root. Deliberately more lenient
    than receipt_audit.py's own candidate_workbooks(): the real files use
    underscores ("..._Transaction_Details_TR.xlsx"), not the literal phrase
    "transaction details" with a space, so match on the two words appearing
    with any separator instead."""
    workbooks = []
    for path in root.rglob("*.xlsx"):
        name = path.name.lower()
        if name.startswith("~$") or "receipt_audit" in name:
            continue
        normalized = name.replace("_", " ").replace("-", " ")
        if "transaction" in normalized and "details" in normalized:
            workbooks.append(path)
    return sorted(workbooks)

FREQUENCY_SHEET = "Frequency Database"
OUTPUT_PATH = Path(__file__).resolve().parent / "frequency_database.xlsx"


def collect_charges_from_statement(path: Path):
    """Read one monthly statement workbook and return (vendor, date, amount)
    rows for every real charge line, using the same sheet/header detection
    receipt_audit.py already relies on. date is a real datetime.date object,
    not a string -- frequency_lookup.py needs actual dates to compute gaps,
    and storing real dates keeps the column sortable/readable in Excel too.

    Uses extract_frequency_key(), not extract_vendor() -- for most vendors
    they return the same thing, but for a vendor like Microsoft that mixes
    several different recurring products under one generic coding name,
    this keeps the extra product detail so each one gets its own clean
    gap pattern instead of all of them being averaged into noise."""
    wb = load_workbook(path, data_only=True)
    ws, header_row, columns = find_charge_sheet(wb, infer_target_year(path))
    rows = []
    for charge in iter_charges(ws, header_row, columns):
        if not charge.merchant or charge.date is None:
            continue
        vendor = extract_frequency_key(charge.merchant)
        rows.append((vendor, charge.date, str(charge.amount or "")))
    return rows


def load_existing(path: Path):
    rows = []
    seen = set()
    if path.exists():
        wb = load_workbook(path, data_only=True)
        if FREQUENCY_SHEET in wb.sheetnames:
            ws = wb[FREQUENCY_SHEET]
            for r in range(2, ws.max_row + 1):
                vendor = ws.cell(row=r, column=1).value
                date = ws.cell(row=r, column=2).value
                amount = ws.cell(row=r, column=3).value
                if not vendor or not date:
                    continue
                if isinstance(date, dt.datetime):
                    date = date.date()
                date_key = date.isoformat() if isinstance(date, dt.date) else str(date)
                rows.append([str(vendor), date, str(amount or "")])
                seen.add((str(vendor).strip().upper(), date_key))
    return rows, seen


def save(path: Path, rows):
    wb = Workbook()
    ws = wb.active
    ws.title = FREQUENCY_SHEET
    ws.append(["Vendor", "Date", "Amount"])
    for row in rows:
        ws.append(row)
    wb.save(path)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--statements-root",
        type=Path,
        help="Folder containing the historical monthly AmEx statement files "
        "(the same ones used to build historical_vendor_database.xlsx). "
        "Not needed if you use --gui instead.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=OUTPUT_PATH,
        help="Where to save frequency_database.xlsx (defaults to the copy next to this script).",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Open a folder-picker window instead of typing a path.",
    )
    args = parser.parse_args(argv)

    if args.gui:
        try:
            import tkinter as tk
            from tkinter import filedialog, messagebox
        except Exception as exc:
            print(f"GUI is unavailable: {exc}", file=sys.stderr)
            return 2
        root_window = tk.Tk()
        root_window.withdraw()
        picked = filedialog.askdirectory(
            title="Choose the folder with your historical AmEx statement files"
        )
        if not picked:
            return 1
        args.statements_root = Path(picked)

    if not args.statements_root:
        print("A statements folder is required: use --statements-root or --gui.", file=sys.stderr)
        return 2

    root = args.statements_root.expanduser().resolve()
    if not root.exists():
        print(f"Statements folder not found: {root}", file=sys.stderr)
        return 2

    workbooks = find_statement_workbooks(root)
    if not workbooks:
        print(f"No 'Transaction Details' workbooks found under: {root}", file=sys.stderr)
        return 2

    existing_rows, seen = load_existing(args.output)
    added = 0
    skipped_files = []

    for workbook in workbooks:
        print(f"Reading: {workbook}")
        try:
            charges = collect_charges_from_statement(workbook)
        except ValueError as exc:
            print(f"  Skipped ({exc})")
            skipped_files.append(workbook)
            continue
        for vendor, charge_date, amount in charges:
            date_key = charge_date.isoformat() if isinstance(charge_date, dt.date) else str(charge_date)
            key = (vendor.strip().upper(), date_key)
            if key in seen:
                continue
            existing_rows.append([vendor, charge_date, amount])
            seen.add(key)
            added += 1

    save(args.output, existing_rows)

    print()
    print(f"Added {added} new charge row(s) to {args.output.name}.")
    print(f"Total rows now: {len(existing_rows)}")
    if skipped_files:
        print(f"Skipped {len(skipped_files)} file(s) that didn't look like a statement workbook:")
        for path in skipped_files:
            print(f"  {path}")

    if args.gui:
        message = (
            f"Frequency database built.\n\n"
            f"Files read: {len(workbooks)}\n"
            f"New charge rows added: {added}\n"
            f"Total rows now: {len(existing_rows)}"
        )
        if skipped_files:
            message += f"\nSkipped {len(skipped_files)} file(s) -- see the terminal for details."
        messagebox.showinfo("Frequency Database Builder", message)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
