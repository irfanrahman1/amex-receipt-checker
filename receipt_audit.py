#!/usr/bin/env python3
"""Audit an AmEx charge workbook against receipt files in the same folder."""

import argparse
import datetime as dt
import logging
import re
import sys
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Set, Tuple

from openpyxl import load_workbook

try:
    from pypdf import PdfReader
except Exception:  # pragma: no cover - optional dependency at runtime
    PdfReader = None

logging.getLogger("pypdf").setLevel(logging.ERROR)


RECEIPT_EXTENSIONS = {
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".heic",
    ".webp",
    ".tif",
    ".tiff",
}

HEADER_ALIASES = {
    "date": ("date", "transaction date", "posted date", "charge date"),
    "merchant": (
        "merchant",
        "description",
        "vendor",
        "supplier",
        "payee",
        "name",
        "doing business as",
    ),
    "amount": ("amount", "charge", "charge amount", "debit", "total", "extended amount"),
}

STATUS_HEADER = "Receipt Status"
MATCH_HEADER = "Matched Receipt"
NOTES_HEADER = "Receipt Match Notes"


@dataclass
class Charge:
    row: int
    date: Optional[dt.date]
    merchant: str
    amount: Optional[Decimal]


@dataclass
class Receipt:
    path: Path
    searchable_text: str


@dataclass
class Match:
    receipt: Optional[Receipt]
    score: float
    notes: str


def clean_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).lower().replace(",", "")
    text = re.sub(r"[^a-z0-9.]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_amount(value: object) -> Optional[Decimal]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value)).copy_abs().quantize(Decimal("0.01"))

    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    negative = "(" in text and ")" in text
    text = re.sub(r"[^0-9.\-]", "", text)
    if not text or text in {"-", "."}:
        return None
    try:
        amount = Decimal(text)
    except InvalidOperation:
        return None
    if negative:
        amount *= Decimal("-1")
    return amount.copy_abs().quantize(Decimal("0.01"))


def parse_date(value: object) -> Optional[dt.date]:
    if value is None or value == "":
        return None
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value

    text = str(value).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%m-%d-%Y", "%m.%d.%Y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def date_tokens(value: Optional[dt.date]) -> List[str]:
    if value is None:
        return []
    return [
        value.strftime("%m%d%Y"),
        value.strftime("%m%d%y"),
        value.strftime("%Y%m%d"),
        value.strftime("%Y %m%d"),
        value.strftime("%Y %m %d"),
        value.strftime("%m %d %Y").lower(),
        value.strftime("%b %d %Y").lower(),
        value.strftime("%B %d %Y").lower(),
    ]


def extract_dates(text: str) -> List[dt.date]:
    dates: Set[dt.date] = set()
    for year, month, day in re.findall(r"\b(20\d{2})\s*(\d{2})(\d{2})\b", text):
        try:
            dates.add(dt.date(int(year), int(month), int(day)))
        except ValueError:
            pass
    for month, day, year in re.findall(r"\b(\d{1,2})\s+(\d{1,2})\s+(20\d{2})\b", text):
        try:
            dates.add(dt.date(int(year), int(month), int(day)))
        except ValueError:
            pass
    return sorted(dates)


def amount_tokens(value: Optional[Decimal]) -> List[str]:
    if value is None:
        return []
    plain = f"{value:.2f}"
    dollars = plain.split(".")[0]
    trimmed = plain.rstrip("0").rstrip(".")
    tokens = [plain, plain.replace(".", ""), f"${plain}", dollars]
    if trimmed != plain:
        tokens.extend([trimmed, f"${trimmed}"])
    return tokens


def merchant_filename_text(merchant: str) -> str:
    text = re.sub(r"[^A-Za-z0-9]+", " ", str(merchant or ""))
    words = [word for word in text.split() if len(word) > 1]
    return " ".join(words[:5]) or "Receipt"


def safe_filename_part(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*]+', "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value[:80] or "Receipt"


def invoice_text(receipt: Receipt) -> str:
    generic_values = {"INVOICE", "NUMBER", "NO", "DATE", "AMOUNT", "TOTAL", "RECEIPT"}
    stem_match = re.search(r"#([a-zA-Z0-9-]{3,})", receipt.path.stem)
    if stem_match:
        value = stem_match.group(1).upper()
        if value not in generic_values:
            return value

    text = receipt.searchable_text
    patterns = (
        r"\binvoice\s*(?:number|no|#)?\s*([a-z0-9-]{3,})\b",
        r"\binv\s*(?:number|no|#)?\s*([a-z0-9-]{3,})\b",
        r"#\s*([a-z0-9-]{3,})\b",
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            value = match.group(1).upper()
            if value not in generic_values:
                return value
    return "NoInvoice"


def receipt_filename(charge: Charge, receipt: Receipt) -> str:
    date_part = charge.date.strftime("%Y-%m-%d") if charge.date else "NoDate"
    merchant_part = safe_filename_part(merchant_filename_text(charge.merchant))
    invoice_part = safe_filename_part(invoice_text(receipt))
    return f"{merchant_part} - {invoice_part} - {date_part}{receipt.path.suffix.lower()}"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    for index in range(2, 1000):
        candidate = path.with_name(f"{stem}_{index}{suffix}")
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Could not find an available filename for {path}")


def rename_receipt(charge: Charge, receipt: Receipt) -> Path:
    new_path = unique_path(receipt.path.with_name(receipt_filename(charge, receipt)))
    if new_path == receipt.path:
        return receipt.path
    receipt.path.rename(new_path)
    receipt.path = new_path
    return new_path


def merchant_score(merchant: str, text: str) -> float:
    merchant_clean = clean_text(merchant)
    if not merchant_clean:
        return 0.0
    if merchant_clean in text:
        return 1.0

    merchant_words = [word for word in merchant_clean.split() if len(word) > 2]
    if not merchant_words:
        return 0.0

    text_words = set(text.split())
    word_hits = sum(1 for word in merchant_words if word in text_words)
    word_score = word_hits / len(merchant_words)
    ratio = SequenceMatcher(None, merchant_clean, text[: max(80, len(merchant_clean) * 4)]).ratio()
    return max(word_score, ratio)


def extract_pdf_text(path: Path) -> str:
    if PdfReader is None:
        return ""
    try:
        reader = PdfReader(str(path))
        pages = reader.pages[:3]
        return " ".join(page.extract_text() or "" for page in pages)
    except Exception:
        return ""


def load_receipts(folder: Path, excel_path: Path, recursive: bool) -> List[Receipt]:
    pattern = "**/*" if recursive else "*"
    receipts = []
    for path in folder.glob(pattern):
        if not path.is_file():
            continue
        if path.resolve() == excel_path.resolve():
            continue
        if path.suffix.lower() not in RECEIPT_EXTENSIONS:
            continue
        text = path.stem
        if path.suffix.lower() == ".pdf":
            text = f"{text} {extract_pdf_text(path)}"
        receipts.append(Receipt(path=path, searchable_text=clean_text(text)))
    return receipts


def find_header_row_and_columns(ws) -> Tuple[int, Dict[str, int]]:
    for row in range(1, min(ws.max_row, 15) + 1):
        seen = {}
        for col in range(1, ws.max_column + 1):
            header = clean_text(ws.cell(row=row, column=col).value)
            if not header:
                continue
            for key, aliases in HEADER_ALIASES.items():
                if header in aliases and key not in seen:
                    seen[key] = col
        if "amount" in seen and ("merchant" in seen or "date" in seen):
            return row, seen
    raise ValueError(
        "Could not identify the workbook headers. Expected columns like Date, "
        "Description/Merchant, and Amount/Charge."
    )


def find_charge_sheet(wb):
    errors = []
    preferred_names = ("transaction details", "master_use me", "master")
    worksheets = sorted(
        wb.worksheets,
        key=lambda sheet: (
            clean_text(sheet.title) not in preferred_names,
            wb.worksheets.index(sheet),
        ),
    )
    for ws in worksheets:
        try:
            header_row, columns = find_header_row_and_columns(ws)
            return ws, header_row, columns
        except ValueError as exc:
            errors.append(f"{ws.title}: {exc}")
    raise ValueError("Could not find a transaction worksheet. " + " | ".join(errors))


def ensure_output_columns(ws, header_row: int) -> Dict[str, int]:
    headers = {
        clean_text(ws.cell(row=header_row, column=col).value): col
        for col in range(1, ws.max_column + 1)
    }
    output = {}
    for label in (STATUS_HEADER, MATCH_HEADER, NOTES_HEADER):
        key = clean_text(label)
        if key in headers:
            output[label] = headers[key]
        else:
            col = ws.max_column + 1
            ws.cell(row=header_row, column=col).value = label
            output[label] = col
            headers[key] = col
    return output


def iter_charges(ws, header_row: int, columns: Dict[str, int]) -> Iterable[Charge]:
    for row in range(header_row + 1, ws.max_row + 1):
        amount = parse_amount(ws.cell(row=row, column=columns["amount"]).value)
        merchant = ""
        if "merchant" in columns:
            merchant = str(ws.cell(row=row, column=columns["merchant"]).value or "")
        date = None
        if "date" in columns:
            date = parse_date(ws.cell(row=row, column=columns["date"]).value)
        if amount is None and not merchant and date is None:
            continue
        yield Charge(row=row, date=date, merchant=merchant, amount=amount)


def score_receipt(charge: Charge, receipt: Receipt, date_window_days: int) -> Tuple[float, List[str]]:
    text = receipt.searchable_text
    score = 0.0
    notes = []

    amount_hit = any(token and token in text for token in amount_tokens(charge.amount))
    if amount_hit:
        score += 0.45
        notes.append("amount")

    date_hit = any(token and clean_text(token) in text for token in date_tokens(charge.date))
    if date_hit:
        score += 0.20
        notes.append("date")
    elif charge.date and date_window_days > 0:
        receipt_dates = extract_dates(text)
        if receipt_dates:
            closest_days = min(abs((receipt_date - charge.date).days) for receipt_date in receipt_dates)
            if closest_days <= date_window_days:
                score += 0.14
                notes.append(f"date within {closest_days} day(s)")

    merch = merchant_score(charge.merchant, text)
    if merch >= 0.55:
        score += min(0.35, merch * 0.35)
        notes.append(f"merchant {merch:.0%}")

    return score, notes


def best_match(
    charge: Charge,
    receipts: List[Receipt],
    threshold: float,
    date_window_days: int,
    used_receipts: Set[Path],
) -> Match:
    best = Match(receipt=None, score=0.0, notes="")
    for receipt in receipts:
        if receipt.path.resolve() in used_receipts:
            continue
        score, notes = score_receipt(charge, receipt, date_window_days)
        if score > best.score:
            best = Match(receipt=receipt, score=score, notes=", ".join(notes))

    if best.receipt is not None and best.score >= threshold:
        return best
    return Match(receipt=None, score=best.score, notes=best.notes or "no strong match")


def audit_workbook(
    excel_path: Path,
    receipts_dir: Path,
    output_path: Path,
    recursive: bool,
    threshold: float,
    date_window_days: int,
    rename_receipts: bool,
) -> Tuple[int, int, int]:
    wb = load_workbook(excel_path)
    ws, header_row, columns = find_charge_sheet(wb)
    output_columns = ensure_output_columns(ws, header_row)
    receipts = load_receipts(receipts_dir, excel_path, recursive)

    acquired = 0
    missing = 0
    renamed = 0
    used_receipts: Set[Path] = set()
    for charge in iter_charges(ws, header_row, columns):
        match = best_match(charge, receipts, threshold, date_window_days, used_receipts)
        if match.receipt:
            acquired += 1
            old_path = match.receipt.path
            if rename_receipts:
                rename_receipt(charge, match.receipt)
                if match.receipt.path != old_path:
                    renamed += 1
            used_receipts.add(match.receipt.path.resolve())
            ws.cell(row=charge.row, column=output_columns[STATUS_HEADER]).value = "Acquired"
            ws.cell(row=charge.row, column=output_columns[MATCH_HEADER]).value = str(
                match.receipt.path.relative_to(receipts_dir)
            )
        else:
            missing += 1
            ws.cell(row=charge.row, column=output_columns[STATUS_HEADER]).value = "Missing"
            ws.cell(row=charge.row, column=output_columns[MATCH_HEADER]).value = ""
        ws.cell(row=charge.row, column=output_columns[NOTES_HEADER]).value = (
            f"score {match.score:.2f}; {match.notes}"
        )

    wb.save(output_path)
    return acquired, missing, renamed


def default_output_path(excel_path: Path) -> Path:
    return excel_path.with_name(f"{excel_path.stem}_receipt_audit{excel_path.suffix}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Mark AmEx workbook charges as Acquired or Missing based on receipt files."
    )
    parser.add_argument("excel", type=Path, help="Path to the AmEx Excel workbook.")
    parser.add_argument(
        "--receipts-dir",
        type=Path,
        help="Folder containing receipts. Defaults to the workbook's folder.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Where to save the audited workbook. Defaults to *_receipt_audit.xlsx.",
    )
    parser.add_argument(
        "--no-recursive",
        action="store_true",
        help="Only check files directly in the receipt folder, not subfolders.",
    )
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.55,
        help="Match threshold from 0.0 to 1.0. Default: 0.55.",
    )
    parser.add_argument(
        "--date-window-days",
        type=int,
        default=3,
        help="Allow receipt dates this many days before/after the workbook date. Default: 3.",
    )
    parser.add_argument(
        "--rename-receipts",
        action="store_true",
        help="Rename matched receipt files to Vendor - Invoice - Date.ext.",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    excel_path = args.excel.expanduser().resolve()
    receipts_dir = (args.receipts_dir or excel_path.parent).expanduser().resolve()
    output_path = (args.output or default_output_path(excel_path)).expanduser().resolve()

    if not excel_path.exists():
        print(f"Excel file not found: {excel_path}", file=sys.stderr)
        return 2
    if not receipts_dir.exists():
        print(f"Receipt folder not found: {receipts_dir}", file=sys.stderr)
        return 2

    acquired, missing, renamed = audit_workbook(
        excel_path=excel_path,
        receipts_dir=receipts_dir,
        output_path=output_path,
        recursive=not args.no_recursive,
        threshold=args.threshold,
        date_window_days=args.date_window_days,
        rename_receipts=args.rename_receipts,
    )
    print(f"Saved audited workbook: {output_path}")
    print(f"Acquired: {acquired}")
    print(f"Missing: {missing}")
    if args.rename_receipts:
        print(f"Renamed receipts: {renamed}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
