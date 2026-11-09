from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

from openpyxl import load_workbook

import update_tracker

ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "TTG worksheet for wir and mir.xlsx"
USERS_WORKBOOK = ROOT / "users.xlsx"
TEMPLATE = ROOT / "github_static_viewer.html"
LOGIN_TEMPLATE = ROOT / "index.html"


def load_public_users() -> list[dict[str, object]]:
    """Read the account list for the requested public, client-side login gate."""
    if not USERS_WORKBOOK.is_file():
        raise SystemExit("users.xlsx is required to build the GitHub Pages sign-in page.")

    workbook = load_workbook(USERS_WORKBOOK, data_only=True, read_only=True)
    try:
        for sheet in workbook.worksheets:
            for header_row in range(1, min(sheet.max_row, 10) + 1):
                headers = {
                    str(cell.value).strip().casefold(): cell.column
                    for cell in sheet[header_row]
                    if cell.value is not None
                }
                if "username" not in headers or "password" not in headers:
                    continue

                accounts = []
                for row in range(header_row + 1, sheet.max_row + 1):
                    username = sheet.cell(row, headers["username"]).value
                    password = sheet.cell(row, headers["password"]).value
                    if username is None or password is None:
                        continue
                    level_column = (
                        headers.get("level")
                        or headers.get("user level")
                        or headers.get("user_level")
                        or headers.get("role")
                    )
                    try:
                        level = int(sheet.cell(row, level_column).value) if level_column else 3
                    except (TypeError, ValueError):
                        level = 3
                    if level not in (1, 2, 3):
                        level = 3
                    first_name_column = (
                        headers.get("first name")
                        or headers.get("firstname")
                        or headers.get("first_name")
                    )
                    last_name_column = (
                        headers.get("last name")
                        or headers.get("lastname")
                        or headers.get("last_name")
                    )
                    first_name = sheet.cell(row, first_name_column).value if first_name_column else ""
                    last_name = sheet.cell(row, last_name_column).value if last_name_column else ""
                    accounts.append({
                        "username": str(username).strip(),
                        "password": str(password),
                        "level": level,
                        "firstName": str(first_name or "").strip(),
                        "lastName": str(last_name or "").strip(),
                    })
                if accounts:
                    return accounts
        raise SystemExit("Could not find username and password columns in users.xlsx.")
    finally:
        workbook.close()


def build(output: Path) -> None:
    output = output.resolve()
    if output != (ROOT / "_site").resolve():
        raise SystemExit("The static build output must be the project's _site folder.")
    if not WORKBOOK.is_file():
        raise SystemExit(f"Workbook not found: {WORKBOOK.name}")
    users = load_public_users()
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
    crs_file_types = {".pdf", ".xlsx", ".xls", ".docx", ".doc", ".png", ".jpg", ".jpeg"}
    for kind, files in attachments.items():
        for info in files:
            source = ROOT / Path(info["path"])
            allowed = source.suffix.lower() == ".pdf" or (
                kind == "CRS" and source.suffix.lower() in crs_file_types
            )
            if not source.is_file() or not allowed:
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
    (output / "users.json").write_text(json.dumps(users, ensure_ascii=False).replace("</", "<\\/"), encoding="utf-8")
    (output / "index.html").write_text(LOGIN_TEMPLATE.read_text(encoding="utf-8"), encoding="utf-8")
    (output / "home.html").write_text(TEMPLATE.read_text(encoding="utf-8"), encoding="utf-8")
    logo = ROOT / "Logo" / "TR.jpg"
    if logo.is_file():
        (output / "Logo").mkdir(exist_ok=True)
        shutil.copy2(logo, output / "Logo" / "TR.jpg")
    print(f"Built Pages site: {len(users)} login accounts, {len(records_json)} records, {len(copied['MIR'])} MIR PDFs, {len(copied['WIR'])} WIR PDFs, {len(copied['CRS'])} CRS attachments.")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Build the read-only GitHub Pages tracker snapshot.")
    parser.add_argument("--output", default=str(ROOT / "_site"), help="Output folder (must be _site in the project).")
    args = parser.parse_args()
    build(Path(args.output))
