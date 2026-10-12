"""Charge-frequency classification: One-time / Monthly / Annually.

Implements the feature Thomas asked for after the vendor-coding work: a
Frequency column that feeds Shilpa's 2025-vs-2026 spend comparison, which
Camille currently builds by hand from pivot tables.

Design (a hybrid of the two approaches Thomas floated on the call):
  - Pure "have we seen this vendor in the last N months" doesn't handle a
    subscription that only started 2 months ago -- with just 2 charges it
    already looks "monthly" even though it hasn't hit any fixed N-month
    threshold yet.
  - Pure occurrence-counting ("12 charges = monthly") doesn't distinguish a
    canceled subscription (charged 3x in a row, then stopped) from one still
    active, and doesn't naturally separate annual charges from one-offs.
  - So instead: for each vendor, look at every known charge date (from
    frequency_database.xlsx, built from ~16 months of historical statements)
    plus the current charge being audited, and compute the *typical gap*
    between consecutive charges for that vendor. That single number does
    both jobs at once:
      - median gap ~30 days (with tolerance)      -> Monthly
      - median gap ~330-400 days (with tolerance) -> Annually
      - only one charge ever, or no clean pattern -> One-time
  - Quarterly is intentionally NOT a bucket yet -- Thomas said to defer it;
    a charge every ~90 days currently falls out as One-time, same as any
    other charge without a clean monthly/annual pattern.

Per Thomas: this needs the vendor name, the date, and (for a later step he
flagged, not built yet) the amount, so frequency_database.xlsx stores all
three even though amount isn't used in classify() yet. He wants this to
catch cases like "looks annual but the amount is way off -- might be a
different charge" -- that's a follow-on refinement, not in scope for the
first working version.

This module intentionally mirrors vendor_lookup.py's shape: a small
dataclass result, a lookup class that loads a workbook once, and a database
file that lives next to this file and keeps growing over time.
"""

import datetime as dt
import statistics
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

from openpyxl import load_workbook

FREQUENCY_WORKBOOK = Path(__file__).resolve().parent / "frequency_database.xlsx"
FREQUENCY_SHEET = "Frequency Database"
FREQ_VENDOR_COL = 1
FREQ_DATE_COL = 2
FREQ_AMOUNT_COL = 3
FREQ_START_ROW = 2

MONTHLY = "Monthly"
ANNUALLY = "Annually"
ONE_TIME = "One-time"

# Tolerance bands around the "ideal" gap. Real-world billing dates drift by a
# few days (weekends, billing-cycle quirks), and an annual charge can land
# up to a month off from last year without being a different thing.
MONTHLY_MIN_DAYS = 20
MONTHLY_MAX_DAYS = 40
ANNUAL_MIN_DAYS = 330
# Widened 400 -> 430 after the Sept 28 check-in: Thomas flagged GoDaddy (a
# true single annual renewal -- not a vendor-aliasing case like Apple) as
# showing One-time when it should be Annually. With ~16-18 months of
# history, a renewal that lands even a couple weeks later than last year
# pushes the gap past the old 400-day ceiling and falls through to
# One-time. 430 days (~14.1 months) gives that drift more room without
# reaching into territory a real semi-annual pattern would occupy.
ANNUAL_MAX_DAYS = 430


def _normalize(vendor: str) -> str:
    return (vendor or "").strip().upper()


@dataclass
class FrequencyResult:
    vendor_input: str
    frequency: str  # "Monthly", "Annually", or "One-time"
    charge_count: int  # total known charges for this vendor, including this one
    typical_gap_days: Optional[float]  # median gap between charges, if 2+


class FrequencyLookup:
    """Classifies a charge's frequency using a vendor's full known charge
    history (frequency_database.xlsx) plus the charge being audited right now.
    """

    def __init__(self, database_path: Path = FREQUENCY_WORKBOOK):
        self.database_path = database_path
        # vendor (normalized) -> sorted list of known charge dates
        self._history: Dict[str, List[dt.date]] = {}
        self._load()

    def _load(self):
        if not self.database_path.exists():
            return
        wb = load_workbook(self.database_path, data_only=True)
        if FREQUENCY_SHEET not in wb.sheetnames:
            return
        ws = wb[FREQUENCY_SHEET]
        for row in range(FREQ_START_ROW, ws.max_row + 1):
            vendor = ws.cell(row=row, column=FREQ_VENDOR_COL).value
            date_val = ws.cell(row=row, column=FREQ_DATE_COL).value
            if not vendor or date_val is None:
                continue
            date = date_val.date() if isinstance(date_val, dt.datetime) else date_val
            if not isinstance(date, dt.date):
                continue
            key = _normalize(str(vendor))
            self._history.setdefault(key, []).append(date)
        for key in self._history:
            self._history[key].sort()

    def classify(self, vendor: str, charge_date: Optional[dt.date]) -> FrequencyResult:
        """Classify this charge's frequency, folding in the vendor's known
        history plus the date of the charge being audited right now (so a
        brand-new month's charge is included even before it's promoted into
        frequency_database.xlsx)."""
        if not vendor or charge_date is None:
            return FrequencyResult(vendor, ONE_TIME, 0, None)

        key = _normalize(vendor)
        dates = list(self._history.get(key, []))
        if charge_date not in dates:
            dates.append(charge_date)
        dates.sort()

        if len(dates) < 2:
            return FrequencyResult(vendor, ONE_TIME, len(dates), None)

        gaps = [(dates[i + 1] - dates[i]).days for i in range(len(dates) - 1)]
        median_gap = statistics.median(gaps)

        if MONTHLY_MIN_DAYS <= median_gap <= MONTHLY_MAX_DAYS:
            return FrequencyResult(vendor, MONTHLY, len(dates), median_gap)
        if ANNUAL_MIN_DAYS <= median_gap <= ANNUAL_MAX_DAYS:
            return FrequencyResult(vendor, ANNUALLY, len(dates), median_gap)
        return FrequencyResult(vendor, ONE_TIME, len(dates), median_gap)
