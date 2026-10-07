from pathlib import Path
import json
import re
from openpyxl import load_workbook
from docx import Document

ROOT = Path(r"C:\Codex\Renshe_Derby\transcript")

def xlsx_rows(path):
    wb = load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            vals = [str(v).strip() for v in row if v not in (None, "")]
            if vals:
                out.append({"loc": f"{ws.title}!{i}", "text": " | ".join(vals)})
    wb.close()
    return out

def docx_rows(path):
    doc = Document(path)
    return [{"loc": f"p{i}", "text": p.text.strip()} for i, p in enumerate(doc.paragraphs) if p.text.strip()]

def classify(name):
    u = name.upper()
    if "EGD" in u:
        return "EGD"
    if "PO" in u or "R5" in u or "R2_" in u:
        return "PO"
    return "UNKNOWN"

files = []
for path in sorted(ROOT.iterdir()):
    if path.suffix.lower() == ".xlsx":
        rows = xlsx_rows(path)
    elif path.suffix.lower() == ".docx":
        rows = docx_rows(path)
    else:
        continue
    files.append({
        "name": path.name,
        "kind": classify(path.name),
        "rows": len(rows),
        "chars": sum(len(r["text"]) for r in rows),
        "content": rows,
    })

out = Path(r"C:\Codex\Renshe_Derby\transcript_audit.json")
out.write_text(json.dumps(files, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps([{k:v for k,v in f.items() if k != "content"} for f in files], ensure_ascii=False, indent=2))

# Structured coverage register for the final workbook. For tagged EGD spreadsheets,
# every topic label on the checked transcript sheet is retained. RM2 is a DOCX and
# therefore uses analyst-defined paragraph bands spanning the complete discussion.
coverage = []
group_ids = {"CHECKED_DM3": "DM3", "CHECKED_DM6": "DM6", "CHECKED_RM5": "RM5"}
for path in sorted(ROOT.glob("*.xlsx")):
    up = path.name.upper()
    gid = next((v for k, v in group_ids.items() if up.startswith(k)), None)
    if not gid:
        continue
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    stats = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        vals = [str(v).strip() for v in row if v not in (None, "")]
        if not vals:
            continue
        tag = vals[-1]
        rec = stats.setdefault(tag, {"count": 0, "min": i, "max": i})
        rec["count"] += 1
        rec["min"] = min(rec["min"], i)
        rec["max"] = max(rec["max"], i)
    wb.close()
    for tag, rec in stats.items():
        if rec["count"] >= 3 and len(tag) <= 80:
            coverage.append({"group": gid, "section": tag, "location": f"rows {rec['min']}–{rec['max']}", "records": rec["count"], "included": "Yes", "source": path.name})

rm2_path = next(ROOT.glob("RM2_EGD_*.docx"))
rm2_bands = [
    ("Introduction and life context", "paragraphs 0–360"),
    ("Dreams, success and challenges", "paragraphs 361–660"),
    ("Entertainment, media and shopping", "paragraphs 661–922"),
    ("Brand meaning and status", "paragraphs 923–1016"),
    ("Category associations and smoker meanings", "paragraphs 1017–1269"),
    ("Consumption moments, routines and needs", "paragraphs 1270–1450"),
    ("Purchase pattern and Derby journey", "paragraphs 1451–1680"),
    ("Occasional brands, aspiration and social use", "paragraphs 1681–1801"),
    ("Product improvement, fallback and rejection", "paragraphs 1802–2006"),
    ("Derby image, persona, name, pack and stick", "paragraphs 2007–2333"),
    ("Price, product experience and ratings", "paragraphs 2334–2428"),
    ("Awareness, promotion, trial and best quality", "paragraphs 2429–2634"),
]
for section, location in rm2_bands:
    coverage.append({"group": "RM2", "section": section, "location": location, "records": None, "included": "Yes", "source": rm2_path.name})

Path(r"C:\Codex\Renshe_Derby\egd_section_coverage.json").write_text(
    json.dumps(coverage, ensure_ascii=False, indent=2), encoding="utf-8"
)

# Recover the already checked brand-preference evidence and PO comparison from
# the preceding workbook so the expanded version can retain those citations.
prior_path = Path(r"C:\Codex\Renshe_Derby\Derby_EGD_Primary_Cigarette_Brand_Preference_Content_Analysis.xlsx")
if prior_path.exists():
    prior = load_workbook(prior_path, read_only=True, data_only=True)
    ev_ws = prior["EGD Evidence Log"]
    evidence_rows = [[ev_ws.cell(r, c).value for c in range(1, 8)] for r in range(5, ev_ws.max_row + 1)]
    po_ws = prior["PO Comparison"]
    po_rows = [[po_ws.cell(r, c).value for c in range(1, 7)] for r in range(5, po_ws.max_row + 1)]
    prior.close()
    Path(r"C:\Codex\Renshe_Derby\prior_analysis_extract.json").write_text(
        json.dumps({"evidence": evidence_rows, "po": po_rows}, ensure_ascii=False, indent=2), encoding="utf-8"
    )

print("\nPOSSIBLE SECTION MARKERS")
keywords = re.compile(r"(section|category|brand|journey|consum|purchase|pack|price|promotion|aspirat|alternative|reject|strength|weakness|intention|quality|lifestyle|life|future|personality|media|smok|সেকশন|ব্র্যান্ড|সিগারেট|প্যাক|দাম|প্রমো)", re.I)
for f in files:
    if f["kind"] != "EGD":
        continue
    print("\n###", f["name"])
    shown = 0
    for r in f["content"]:
        t = r["text"]
        if len(t) <= 180 and keywords.search(t):
            print(r["loc"], t.replace("\n", " "))
            shown += 1
            if shown >= 160:
                break
