from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

from openpyxl import load_workbook

import update_tracker

ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "TTG worksheet for wir and mir.xlsx"
TEMPLATE = ROOT / "github_static_viewer.html"


def build(output: Path) -> None:
    output = output.resolve()
    if output != (ROOT / "_site").resolve():
        raise SystemExit("The static build output must be the project's _site folder.")
    if not WORKBOOK.is_file():
        raise SystemExit(f"Workbook not found: {WORKBOOK.name}")
    if output.exists():
        shutil.rmtree(output)
    output.mkdir(parents=True)

    workbook = load_workbook(WORKBOOK, data_only=True, read_only=False)
    try:
        if "Sheet1" not in workbook.sheetnames:
            raise SystemExit("The tracker workbook must contain Sheet1.")
        records = update_tracker.extract_records(workbook)
    finally:
        workbook.close()
    attachments = update_tracker.attachment_index()
    update_tracker.attach_matches(records, attachments)

    copied = {"MIR": [], "WIR": [], "Delivery": [], "CRS": []}
    for kind, files in attachments.items():
        for info in files:
            source = ROOT / Path(info["path"])
            if not source.is_file() or source.suffix.lower() != ".pdf":
                continue
            destination = output / Path(info["path"])
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
            copied[kind].append(info)

    records_json = json.loads(json.dumps(records, ensure_ascii=False))
    valid_paths = {info["path"] for files in copied.values() for info in files}
    for rec in records_json:
        rec["files"] = [f for f in rec.get("files", []) if f["path"] in valid_paths]
        if rec.get("deliveryFile") and rec["deliveryFile"].get("path") not in valid_paths:
            rec["deliveryFile"] = None
        if rec.get("commentFile") and rec["commentFile"].get("path") not in valid_paths:
            rec["commentFile"] = None

    data = {
        "records": records_json,
        "attachments": copied,
        "workbook": WORKBOOK.name,
    }
    (output / "data.json").write_text(json.dumps(data, ensure_ascii=False).replace("</", "<\\/"), encoding="utf-8")
    (output / "index.html").write_text(TEMPLATE.read_text(encoding="utf-8"), encoding="utf-8")
    logo = ROOT / "Logo" / "TR.jpg"
    if logo.is_file():
        (output / "Logo").mkdir(exist_ok=True)
        shutil.copy2(logo, output / "Logo" / "TR.jpg")
    print(f"Built static viewer: {len(records_json)} records, {len(copied['MIR'])} MIR PDFs, {len(copied['WIR'])} WIR PDFs.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build the read-only GitHub Pages tracker snapshot.")
    parser.add_argument("--output", default=str(ROOT / "_site"), help="Output folder (must be _site in the project).")
    args = parser.parse_args()
    build(Path(args.output))
