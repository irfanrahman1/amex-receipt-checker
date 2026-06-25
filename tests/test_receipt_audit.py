import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from receipt_audit import audit_workbook


class ReceiptAuditTests(unittest.TestCase):
    def test_marks_acquired_and_missing_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            workbook_path = folder / "charges.xlsx"
            output_path = folder / "audited.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.append(["Date", "Description", "Amount"])
            ws.append(["06/01/2026", "Hotel Example", "125.40"])
            ws.append(["06/02/2026", "No Receipt Vendor", "9.99"])
            wb.save(workbook_path)

            (folder / "20260601 Hotel Example 125.40.jpg").write_text("receipt", encoding="utf-8")

            acquired, missing = audit_workbook(
                excel_path=workbook_path,
                receipts_dir=folder,
                output_path=output_path,
                recursive=True,
                threshold=0.55,
            )

            self.assertEqual(acquired, 1)
            self.assertEqual(missing, 1)

            audited = load_workbook(output_path)
            ws = audited.active
            self.assertEqual(ws["D2"].value, "Acquired")
            self.assertEqual(ws["D3"].value, "Missing")
            self.assertIn("Hotel Example", ws["E2"].value)


if __name__ == "__main__":
    unittest.main()
