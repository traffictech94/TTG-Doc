from __future__ import annotations

import json
import re
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    print("The Python package 'openpyxl' is required to refresh the tracker.")
    print("Run: py -m pip install openpyxl")
    raise


ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "TTG worksheet for wir and mir.xlsx"
OUTPUT = ROOT / "home.html"
STATUS_OPTIONS = ["Just Created", "Not Submitted", "ITP Need", "With Consultant", "No Updates", "Process", "Code-A", "Code-B", "with Doc Controller"]


def clean(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace("\ufffd", "-").strip()


def value_at(ws, row, col):
    return clean(ws.cell(row, col).value)


def find_header(ws, required):
    for row in range(1, min(ws.max_row, 8) + 1):
        vals = [clean(c.value).lower() for c in ws[row]]
        if all(any(needle in v for v in vals) for needle in required):
            return row
    return None


def status_choice(value, fallback="Process"):
    v = clean(value).lower()
    if v == "pending":
        return "Not Submitted"
    return next((s for s in STATUS_OPTIONS if s.lower() == v), fallback)


def attachment_index():
    found = {"MIR": [], "WIR": [], "Delivery": [], "CRS": []}
    folders = (("MIR", ROOT / "TTG MIR", "*.pdf"),
               ("WIR", ROOT / "TTG WIR", "*.pdf"),
               ("Delivery", ROOT / "Delivery Notes", "*.pdf"),
               ("CRS", ROOT / "CRS", "*"))
    for kind, folder, pattern in folders:
        if folder.exists():
            for path in sorted(folder.rglob(pattern), key=lambda p: p.name.lower()):
                if path.is_file():
                    found[kind].append({"name": path.name, "path": path.relative_to(ROOT).as_posix()})
    return found


def record(sheet, row, kind, number="", description="", mir="", wir="", mir_status="Process", wir_status="Process", status="Process", extra=None, source_status=""):
    key = f"{sheet}:{row}"
    return {
        "key": key,
        "sheet": sheet,
        "row": row,
        "kind": kind,
        "number": number,
        "description": description,
        "mir": mir,
        "wir": wir,
        "mirStatus": mir_status,
        "wirStatus": wir_status,
        "status": status,
        "sourceStatus": source_status,
        "extra": extra or [],
        "item": "",
        "delivery": "",
        "comments": "",
        "deliveryFile": None,
        "commentFile": None,
    }


def extract_records(workbook):
    out = []

    # Only the main combined status log is in scope. Include every populated row,
    # including new records and repeated Excel numbers.
    ws = workbook["Sheet1"]
    hdr = find_header(ws, ["sr. no", "material", "mir", "wir"])
    if hdr:
        for r in range(hdr + 1, ws.max_row + 1):
            sr, desc, mir, ms, wir, item, wst, delivery, comments = (
                value_at(ws, r, c) for c in (1, 2, 11, 12, 13, 14, 15, 16, 17)
            )
            if not any((sr, desc, mir, wir, item, delivery, comments)):
                continue
            rec = record(ws.title, r, "MIR & WIR", sr, desc, mir, wir,
                         status_choice(ms), status_choice(wst), status_choice(ms),
                         [("MIR Status (workbook)", ms), ("WIR Status (workbook)", wst)])
            rec["item"] = item
            rec["delivery"] = delivery
            rec["comments"] = comments
            out.append(rec)
    return out


def attach_matches(records, attachments):
    for rec in records:
        rec["files"] = []
        searches = [("MIR", rec.get("mir", "")), ("WIR", rec.get("wir", ""))]
        for kind, number in searches:
            if not number:
                continue
            token = number.casefold()
            for file in attachments[kind]:
                if token in file["name"].casefold():
                    rec["files"].append({**file, "kind": kind})
        delivery_digits = "".join(ch for ch in rec.get("delivery", "") if ch.isdigit())
        if delivery_digits:
            matches = [f for f in attachments["Delivery"]
                       if "".join(ch for ch in Path(f["name"]).stem if ch.isdigit()) == delivery_digits]
            if len(matches) == 1:
                rec["deliveryFile"] = matches[0]
        comment_key = re.sub(r"[^a-z0-9]", "", rec.get("comments", "").casefold())
        if comment_key:
            matches = [f for f in attachments["CRS"]
                       if re.sub(r"[^a-z0-9]", "", Path(f["name"]).stem.casefold()) == comment_key]
            if len(matches) == 1:
                rec["commentFile"] = matches[0]


def main():
    if "--initialize-missing-wir" in sys.argv:
        wb = load_workbook(WORKBOOK, data_only=False)
        ws = wb["Sheet1"]
        wir_files = [p.name.casefold() for p in (ROOT / "TTG WIR").rglob("*.pdf")]
        changed = []
        header = find_header(ws, ["sr. no", "material", "mir", "wir"])
        for row in range((header or 4) + 1, ws.max_row + 1):
            excel_no = ws.cell(row, 1).value
            wir_no = clean(ws.cell(row, 13).value)
            has_pdf = bool(wir_no and any(wir_no.casefold() in name for name in wir_files))
            if not has_pdf and ws.cell(row, 15).value != "Process":
                ws.cell(row, 15).value = "Process"
                changed.append((excel_no, row))
        wb.save(WORKBOOK)
        wb.close()
        print(f"Set missing WIR PDF statuses to Process in {len(changed)} workbook rows.")
        return
    if not WORKBOOK.exists():
        sys.exit(f"Workbook not found: {WORKBOOK.name}")
    wb = load_workbook(WORKBOOK, data_only=True, read_only=False)
    data = {"records": extract_records(wb), "attachments": attachment_index(), "workbook": WORKBOOK.name, "logo": "Logo/TR.jpg"}
    attach_matches(data["records"], data["attachments"])
    json_data = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    template = ROOT / "tracker_template.html"
    if not template.exists():
        sys.exit("tracker_template.html is missing.")
    html = template.read_text(encoding="utf-8").replace("__TRACKER_DATA__", json_data)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"Updated {OUTPUT.name}: {len(data['records'])} records, "
          f"{len(data['attachments']['MIR'])} MIR PDFs, {len(data['attachments']['WIR'])} WIR PDFs.")


if __name__ == "__main__":
    main()
