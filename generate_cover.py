from pathlib import Path
import csv
from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent
clean = list(csv.DictReader((ROOT / "output" / "clean_merged.csv").open("r", encoding="utf-8-sig")))
issues = list(csv.DictReader((ROOT / "output" / "issues_report.csv").open("r", encoding="utf-8-sig")))
report = (ROOT / "output" / "processing_report.txt").read_text(encoding="utf-8")
input_rows = int(next(line.split(":", 1)[1] for line in report.splitlines() if line.startswith("Input rows:")))
duplicates = int(next(line.split(":", 1)[1] for line in report.splitlines() if line.startswith("Duplicates removed:")))


def f(size, bold=False):
    p = Path(r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf")
    return ImageFont.truetype(str(p), size) if p.exists() else ImageFont.load_default()

img = Image.new("RGB", (1600, 1200), "#0b1426")
d = ImageDraw.Draw(img)
white, muted, cyan, green, orange, panel, line = "#f4f7fb", "#a8b4c8", "#39c6ff", "#55e39b", "#ffbd59", "#152239", "#2c4162"
d.text((70, 55), "Excel/CSV Merge & Validation", font=f(62, True), fill=white)
d.text((72, 140), "Self-created Python automation demo", font=f(30), fill=muted)
steps = ["READ CSV", "READ XLSX", "NORMALIZE", "DEDUPLICATE", "VALIDATE", "EXPORT"]
x, y = 70, 260
for i, label in enumerate(steps):
    d.rounded_rectangle((x, y, x+210, y+150), radius=22, fill=panel, outline=cyan if i < 4 else green, width=3)
    d.text((x+20, y+48), label, font=f(22, True), fill=white)
    if i < len(steps)-1:
        d.line((x+224, y+75, x+265, y+75), fill=muted, width=5)
        d.polygon([(x+265, y+75), (x+248, y+65), (x+248, y+85)], fill=muted)
    x += 245
d.rounded_rectangle((70, 490, 1530, 910), radius=28, fill=panel, outline=line, width=3)
d.text((110, 535), "Verified processing result", font=f(38, True), fill=white)
metrics = [(str(input_rows), "input rows", cyan), (str(len(clean)), "unique rows", green), (str(duplicates), "duplicate removed", orange), (str(len(issues)), "issues flagged", orange)]
mx = 120
for value, label, color in metrics:
    d.text((mx, 645), value, font=f(58, True), fill=color)
    d.text((mx, 725), label, font=f(26), fill=muted)
    mx += 350
d.text((105, 815), "Detected: invalid email  •  invalid amount  •  missing email  •  missing name", font=f(27), fill=white)
d.text((80, 1010), "Python  •  Excel  •  CSV  •  openpyxl  •  Validation  •  pytest", font=f(29), fill=white)
d.text((80, 1065), "Generated sample data only — no private or client information", font=f(25), fill=muted)
path = ROOT / "portfolio_cover_1600x1200.png"
img.save(path)
print(path)
