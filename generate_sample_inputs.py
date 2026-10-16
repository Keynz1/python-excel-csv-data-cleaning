from pathlib import Path
import csv
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent
INPUT = ROOT / "input"
INPUT.mkdir(parents=True, exist_ok=True)

csv_rows = [
    [" ID ", "NAME", "Email", "Amount", "Status"],
    ["1001", "alice morgan", "ALICE.MORGAN@example.com", "120", "paid"],
    ["1002", "bob lee", "bob.lee@example.com", "75", "new"],
    ["1003", "carol king", "bad-email", "199.99", "pending"],
    ["1004", "", "missing.name@example.com", "30", "paid"],
]
with (INPUT / "source_a.csv").open("w", encoding="utf-8-sig", newline="") as f:
    csv.writer(f).writerows(csv_rows)

xlsx_rows = [
    ["id", " name ", "EMAIL", "amount", "status"],
    ["1002", "Bob Duplicate", "duplicate@example.com", "80", "paid"],
    ["1005", "David Stone", "david.stone@example.com", "not-a-number", "new"],
    ["1006", "Ethan Hall", "ethan.hall@example.com", "-15", "pending"],
    ["1007", "Maya Chen", "", "310.25", "new"],
]
wb = Workbook()
ws = wb.active
for row in xlsx_rows:
    ws.append(row)
wb.save(INPUT / "source_b.xlsx")
print(f"Created sample inputs in {INPUT}")
