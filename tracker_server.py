from __future__ import annotations

import json
import hmac
import re
import secrets
import sys
import threading
import time
from copy import copy
from http.cookies import SimpleCookie
from datetime import datetime
from email import policy
from email.parser import BytesParser
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlsplit

from openpyxl import load_workbook

import update_tracker


ROOT = Path(__file__).resolve().parent
WORKBOOK = ROOT / "TTG worksheet for wir and mir.xlsx"
USERS_WORKBOOK = ROOT / "users.xlsx"
HOST = "127.0.0.1"
PORT = 8767
SESSION_COOKIE = "TTGSESSION"
SESSION_TTL = 8 * 60 * 60
SESSIONS = {}
SESSION_LOCK = threading.Lock()
USERS_LOCK = threading.Lock()
ALWAYS_BLOCKED = {
    "users.xlsx",
    "Form.xlsx",
    WORKBOOK.name.casefold(),
    "tracker_server.py",
    "tracker_template.html",
    "update_tracker.py",
    "update ttg tracker.bat",
    "tracker_server.log",
}
VALID_STATUSES = {"Just Created", "Not Submitted", "ITP Need", "With Consultant", "No Updates", "Process", "Code-A", "Code-B", "with Doc Controller"}
KEY_RE = re.compile(r"^Sheet1:(\d+)$")


def log(message):
    with (ROOT / "tracker_server.log").open("a", encoding="utf-8") as stream:
        stream.write(f"{datetime.now().isoformat(timespec='seconds')} {message}\n")


class TrackerHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(ROOT), **kwargs)

    def log_message(self, fmt, *args):
        # Keep the background helper quiet.
        return

    def send_json(self, code, value):
        body = json.dumps(value, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def _session_token(self):
        cookie = SimpleCookie()
        try:
            cookie.load(self.headers.get("Cookie", ""))
        except Exception:
            return ""
        morsel = cookie.get(SESSION_COOKIE)
        return morsel.value if morsel else ""

    def _session_identity(self):
        session = self._session_details()
        return session.get("username") if session else None

    def _session_details(self):
        token = self._session_token()
        if not token:
            return None
        now = time.time()
        with SESSION_LOCK:
            session = SESSIONS.get(token)
            if isinstance(session, dict) and session.get("expires", 0) > now:
                return dict(session)
            SESSIONS.pop(token, None)
        return None

    def _is_authenticated(self):
        return self._session_identity() is not None

    def _redirect_to_login(self):
        self.send_response(303)
        self.send_header("Location", "/")
        self.end_headers()

    def _origin_allowed(self):
        origin = self.headers.get("Origin", "")
        return not origin or origin in (f"http://{HOST}:{PORT}", f"http://localhost:{PORT}")

    def _serve_page_or_file(self, head=False):
        request_path = unquote(urlsplit(self.path).path)
        if request_path in ("", "/", "/index.html"):
            self.path = "/index.html"
            return super().do_HEAD() if head else super().do_GET()
        if request_path == "/Logo/TR.jpg":
            self.path = request_path
            return super().do_HEAD() if head else super().do_GET()
        if not self._is_authenticated():
            return self._redirect_to_login()
        target = Path(self.translate_path(self.path)).resolve()
        try:
            relative = target.relative_to(ROOT).as_posix()
        except ValueError:
            return self.send_error(404)
        if target.name.casefold() in ALWAYS_BLOCKED or any(part.startswith(".") or part.casefold() == "__pycache__" for part in target.relative_to(ROOT).parts):
            return self.send_error(404)
        return super().do_HEAD() if head else super().do_GET()

    def do_GET(self):
        if urlsplit(self.path).path == "/session":
            session = self._session_details()
            if session:
                return self.send_json(200, {
                    "ok": True,
                    "username": session.get("username"),
                    "firstName": session.get("firstName", ""),
                    "lastName": session.get("lastName", ""),
                    "level": session.get("level", 3),
                })
            return self.send_json(401, {"ok": False})
        self._serve_page_or_file()

    def do_HEAD(self):
        self._serve_page_or_file(head=True)

    def _find_credentials(self, username, password):
        if not USERS_WORKBOOK.is_file():
            raise FileNotFoundError("users.xlsx is missing")
        workbook = load_workbook(USERS_WORKBOOK, data_only=True, read_only=True)
        try:
            sheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
            header_row = None
            username_col = password_col = level_col = first_name_col = last_name_col = None
            for row in range(1, min(sheet.max_row, 10) + 1):
                headers = {str(cell.value).strip().casefold(): cell.column for cell in sheet[row] if cell.value is not None}
                if "username" in headers and "password" in headers:
                    header_row = row
                    username_col = headers["username"]
                    password_col = headers["password"]
                    level_col = headers.get("level") or headers.get("user level") or headers.get("user_level") or headers.get("role")
                    first_name_col = headers.get("first name") or headers.get("firstname") or headers.get("first_name")
                    last_name_col = headers.get("last name") or headers.get("lastname") or headers.get("last_name")
                    break
            if not header_row:
                raise ValueError("username/password columns were not found")
            for row in range(header_row + 1, sheet.max_row + 1):
                stored_username = sheet.cell(row, username_col).value
                stored_password = sheet.cell(row, password_col).value
                if stored_username is None or stored_password is None:
                    continue
                stored_username = str(stored_username).strip()
                stored_password = str(stored_password)
                if (hmac.compare_digest(stored_username.casefold().encode("utf-8"), username.casefold().encode("utf-8"))
                        and hmac.compare_digest(stored_password.encode("utf-8"), password.encode("utf-8"))):
                    try:
                        level = int(sheet.cell(row, level_col).value) if level_col else 3
                    except (TypeError, ValueError):
                        level = 3
                    if level not in (1, 2, 3):
                        level = 3
                    first_name = sheet.cell(row, first_name_col).value if first_name_col else ""
                    last_name = sheet.cell(row, last_name_col).value if last_name_col else ""
                    return {
                        "username": stored_username,
                        "row": row,
                        "level": level,
                        "firstName": str(first_name or "").strip(),
                        "lastName": str(last_name or "").strip(),
                    }
            return None
        finally:
            workbook.close()

    def _write_password(self, username, current_password, new_password):
        with USERS_LOCK:
            workbook = load_workbook(USERS_WORKBOOK, data_only=False)
            try:
                sheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
                header_row = None
                username_col = password_col = None
                for row in range(1, min(sheet.max_row, 10) + 1):
                    headers = {str(cell.value).strip().casefold(): cell.column for cell in sheet[row] if cell.value is not None}
                    if "username" in headers and "password" in headers:
                        header_row = row
                        username_col = headers["username"]
                        password_col = headers["password"]
                        break
                if not header_row:
                    raise ValueError("The username and password columns were not found in users.xlsx.")
                for row in range(header_row + 1, sheet.max_row + 1):
                    stored_username = sheet.cell(row, username_col).value
                    stored_password = sheet.cell(row, password_col).value
                    if stored_username is None or stored_password is None:
                        continue
                    stored_username = str(stored_username).strip()
                    stored_password = str(stored_password)
                    if (hmac.compare_digest(stored_username.casefold().encode("utf-8"), username.casefold().encode("utf-8"))
                            and hmac.compare_digest(stored_password.encode("utf-8"), current_password.encode("utf-8"))):
                        sheet.cell(row, password_col).value = new_password
                        workbook.save(USERS_WORKBOOK)
                        return stored_username
                raise ValueError("The current password does not match this account.")
            finally:
                workbook.close()

    def _handle_login(self):
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "message": "This sign-in request was rejected."})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 1 or length > 16_384:
                raise ValueError("Invalid request")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            username = payload.get("username", "")
            password = payload.get("password", "")
            if not isinstance(username, str) or not isinstance(password, str) or not username or not password:
                return self.send_json(401, {"ok": False, "message": "Username or password is incorrect."})
            account = self._find_credentials(username, password)
            if not account:
                log("Login rejected.")
                return self.send_json(401, {"ok": False, "message": "Username or password is incorrect."})
            token = secrets.token_urlsafe(32)
            with SESSION_LOCK:
                SESSIONS[token] = {
                    "expires": time.time() + SESSION_TTL,
                    "username": account["username"],
                    "firstName": account.get("firstName", ""),
                    "lastName": account.get("lastName", ""),
                    "level": account["level"],
                }
            log("Login accepted.")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Set-Cookie", f"{SESSION_COOKIE}={token}; Path=/; Max-Age={SESSION_TTL}; HttpOnly; SameSite=Lax")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')
        except (ValueError, json.JSONDecodeError):
            self.send_json(400, {"ok": False, "message": "The sign-in form could not be read."})
        except Exception as exc:
            log(f"Login service error: {type(exc).__name__}")
            self.send_json(503, {"ok": False, "message": "The sign-in service is unavailable. Check users.xlsx and try again."})

    def _handle_logout(self):
        token = self._session_token()
        if token:
            with SESSION_LOCK:
                SESSIONS.pop(token, None)
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Set-Cookie", f"{SESSION_COOKIE}=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax")
        self.end_headers()
        self.wfile.write(b'{"ok":true}')

    def _handle_change_password(self):
        if not self._is_authenticated():
            return self.send_json(401, {"ok": False, "message": "Please sign in again before changing your password."})
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "message": "This password update request was rejected."})
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 1 or length > 16_384:
                raise ValueError("The password form could not be read.")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            current_password = payload.get("currentPassword", "")
            new_password = payload.get("newPassword", "")
            confirm_password = payload.get("confirmPassword", "")
            if not all(isinstance(value, str) for value in (current_password, new_password, confirm_password)):
                raise ValueError("Enter valid password values.")
            if not current_password or not new_password:
                raise ValueError("Complete all password fields.")
            if new_password != confirm_password:
                raise ValueError("The new passwords do not match.")
            if len(new_password) < 8:
                raise ValueError("Choose a new password with at least 8 characters.")
            username = self._session_identity()
            saved_username = self._write_password(username, current_password, new_password)
            token = self._session_token()
            with SESSION_LOCK:
                for other_token, session in list(SESSIONS.items()):
                    if other_token != token and isinstance(session, dict) and str(session.get("username", "")).casefold() == saved_username.casefold():
                        SESSIONS.pop(other_token, None)
            log("Password changed.")
            self.send_json(200, {"ok": True, "message": "Password changed successfully."})
        except ValueError as exc:
            self.send_json(400, {"ok": False, "message": str(exc)})
        except PermissionError:
            self.send_json(409, {"ok": False, "message": "Close users.xlsx in Excel, then try changing the password again."})
        except Exception as exc:
            log(f"Password update error: {type(exc).__name__}")
            self.send_json(500, {"ok": False, "message": "The password could not be saved to users.xlsx."})

    def _handle_add_record(self):
        session = self._session_details()
        if not session:
            return self.send_json(401, {"ok": False, "error": "Please sign in before adding a record."})
        if session.get("level") != 1:
            return self.send_json(403, {"ok": False, "error": "Only a Level 1 super admin can add records."})
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "error": "This add-record request was rejected."})
        created_files = []
        workbook = None
        try:
            length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            if length < 1 or length > 220 * 1024 * 1024 or not content_type.lower().startswith("multipart/form-data"):
                raise ValueError("The record form or attachments are too large or invalid.")
            body = self.rfile.read(length)
            message = BytesParser(policy=policy.default).parsebytes(
                f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("ascii") + body
            )
            fields, uploads = {}, {}
            for part in message.iter_parts():
                name = part.get_param("name", header="content-disposition")
                if not name:
                    continue
                filename = part.get_filename()
                payload = part.get_payload(decode=True) or b""
                if filename:
                    uploads[name] = (Path(filename).name, part.get_content_type(), payload)
                else:
                    fields[name] = payload.decode(part.get_content_charset() or "utf-8", errors="replace").strip()

            status_fields = {"mirStatus": fields.get("mirStatus", "Process"), "wirStatus": fields.get("wirStatus", "Process")}
            if any(value not in VALID_STATUSES for value in status_fields.values()):
                raise ValueError("Choose a valid MIR and WIR status.")
            if not fields.get("material"):
                raise ValueError("Material & Location is required.")

            saved_names = {}
            folder_map = {"mirFile": "TTG MIR", "wirFile": "TTG WIR", "deliveryFile": "Delivery Notes", "commentFile": "CRS"}
            for field, (filename, mime, payload) in uploads.items():
                if field not in folder_map or not filename:
                    continue
                if len(payload) > 50 * 1024 * 1024 or (not filename.lower().endswith(".pdf") and mime != "application/pdf"):
                    raise ValueError(f"{filename} must be a PDF no larger than 50 MB.")
                safe_name = re.sub(r"[^A-Za-z0-9._-]", "_", filename)
                folder = ROOT / folder_map[field]
                folder.mkdir(parents=True, exist_ok=True)
                target = folder / safe_name
                if target.exists():
                    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
                    target = folder / f"{target.stem}_{stamp}{target.suffix}"
                target.write_bytes(payload)
                created_files.append(target)
                saved_names[field] = target.name

            workbook = load_workbook(WORKBOOK, data_only=False)
            sheet = workbook["Sheet1"]
            header = update_tracker.find_header(sheet, ["sr. no", "material", "mir", "wir"])
            if not header:
                raise ValueError("Could not locate the MIR/WIR table in Sheet1.")
            used = [r for r in range(header + 1, sheet.max_row + 1)
                    if any(sheet.cell(r, c).value is not None for c in (1, 2, 11, 13, 14, 16, 17))]
            row = max(used, default=header) + 1
            source_row = max(used, default=row - 1)
            if source_row <= sheet.max_row:
                for col in range(1, sheet.max_column + 1):
                    src, dst = sheet.cell(source_row, col), sheet.cell(row, col)
                    if src.has_style:
                        dst._style = copy(src._style)
                    if src.number_format:
                        dst.number_format = src.number_format
                    if src.alignment:
                        dst.alignment = copy(src.alignment)
            numbers = []
            for old_row in used:
                try:
                    numbers.append(int(str(sheet.cell(old_row, 1).value).strip()))
                except (ValueError, TypeError):
                    pass
            sr_no = fields.get("srNo") or str(max(numbers, default=0) + 1)
            mir = fields.get("mirNumber") or saved_names.get("mirFile", "")
            wir = fields.get("wirNumber") or saved_names.get("wirFile", "")
            delivery = fields.get("deliveryNotes") or saved_names.get("deliveryFile", "")
            comments = fields.get("comments") or saved_names.get("commentFile", "")
            values = {1: sr_no, 2: fields["material"], 11: mir, 12: status_fields["mirStatus"],
                      13: wir, 14: fields.get("items", ""), 15: status_fields["wirStatus"],
                      16: delivery, 17: comments}
            for col, value in values.items():
                sheet.cell(row, col).value = value
            workbook.save(WORKBOOK)
            workbook.close()
            workbook = None
            update_tracker.main()
            return self.send_json(200, {"ok": True, "row": row})
        except ValueError as exc:
            for path in created_files:
                path.unlink(missing_ok=True)
            return self.send_json(400, {"ok": False, "error": str(exc)})
        except PermissionError:
            for path in created_files:
                path.unlink(missing_ok=True)
            return self.send_json(409, {"ok": False, "error": "Close the tracker workbook in Excel, then try again."})
        except Exception as exc:
            for path in created_files:
                path.unlink(missing_ok=True)
            log(f"Add record error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The record could not be saved to Excel."})
        finally:
            if workbook is not None:
                workbook.close()

    def do_POST(self):
        route = urlsplit(self.path).path
        if route == "/login":
            return self._handle_login()
        if route == "/logout":
            return self._handle_logout()
        if route == "/change-password":
            return self._handle_change_password()
        if route == "/add-record":
            return self._handle_add_record()
        if route != "/update":
            self.send_json(404, {"error": "Unknown endpoint."})
            return
        if not self._is_authenticated():
            self.send_json(401, {"error": "Please sign in before updating Excel."})
            return
        if (self._session_details() or {}).get("level") != 1:
            self.send_json(403, {"error": "Only a Level 1 user can update MIR/WIR statuses."})
            return
        if not self._origin_allowed():
            self.send_json(403, {"error": "Requests must come from the local tracker page."})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 1 or length > 100_000:
                raise ValueError("The status update is empty or too large.")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            changes = payload.get("changes")
            if not isinstance(changes, list) or not changes:
                raise ValueError("There are no status changes to save.")
            log("Received status update: " + repr([(c.get("key"), c.get("field"), c.get("status")) for c in changes if isinstance(c, dict)]))

            validated = []
            for change in changes:
                if not isinstance(change, dict):
                    raise ValueError("Invalid status change.")
                match = KEY_RE.fullmatch(str(change.get("key", "")))
                field = change.get("field")
                status = change.get("status")
                if not match or field not in ("mirStatus", "wirStatus") or status not in VALID_STATUSES:
                    raise ValueError("A status change has an invalid record or status.")
                row = int(match.group(1))
                validated.append((row, field, status))

            workbook = load_workbook(WORKBOOK, data_only=False)
            sheet = workbook["Sheet1"]
            header = update_tracker.find_header(sheet, ["sr. no", "material", "mir", "wir"])
            if not header:
                raise ValueError("Could not locate the MIR/WIR table in Sheet1.")
            for row, field, status in validated:
                if row <= header or row > sheet.max_row:
                    raise ValueError("A status change refers to a row outside the current workbook data.")
                if not any(sheet.cell(row, col).value is not None for col in (1, 2, 11, 13, 14, 16, 17)):
                    raise ValueError("The selected workbook row is empty. Refresh the tracker and try again.")
                target_col = 12 if field == "mirStatus" else 15
                sheet.cell(row, target_col).value = status
            try:
                workbook.save(WORKBOOK)
            except PermissionError as exc:
                raise RuntimeError("Close the workbook in Excel, then press Update Excel again.") from exc
            finally:
                workbook.close()

            update_tracker.main()
            self.send_json(200, {"saved": len(validated)})
        except (ValueError, KeyError, json.JSONDecodeError) as exc:
            log(f"Rejected status update: {exc}")
            self.send_json(400, {"error": str(exc)})
        except RuntimeError as exc:
            log(f"Workbook update unavailable: {exc}")
            self.send_json(409, {"error": str(exc)})
        except PermissionError:
            log("Workbook update failed because the workbook is locked.")
            self.send_json(409, {"error": "The workbook is open or locked. Close it in Excel and try again."})
        except Exception as exc:
            log(f"Unexpected workbook update error: {exc!r}")
            self.send_json(500, {"error": f"Could not save the workbook: {exc}"})


class SingleTrackerServer(ThreadingHTTPServer):
    # Prevent multiple versions from sharing the same port on Windows.
    allow_reuse_address = False


def main():
    try:
        server = SingleTrackerServer((HOST, PORT), TrackerHandler)
    except OSError as exc:
        # A previous launch may already be serving the tracker.
        log(f"Could not bind to {HOST}:{PORT}: {exc}")
        return
    log(f"Serving tracker at http://{HOST}:{PORT}/")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        log("Tracker helper stopped.")


if __name__ == "__main__":
    main()
