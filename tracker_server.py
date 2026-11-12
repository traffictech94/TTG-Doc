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
TRACKER_LOCK = threading.Lock()
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


def locate_users_header(sheet):
    aliases = {
        "no": {"no", "number", "sno", "srno"},
        "username": {"username", "user name"},
        "password": {"password", "pass"},
        "firstName": {"firstname", "first name", "first_name"},
        "lastName": {"lastname", "last name", "last_name"},
        "email": {"email", "e-mail", "email address"},
        "level": {"level", "user level", "user_level", "role"},
    }
    for row_idx in range(1, min(sheet.max_row, 10) + 1):
        columns = {}
        for cell in sheet[row_idx]:
            if cell.value is None:
                continue
            normalized = re.sub(r"[\s_-]+", "", str(cell.value).strip().casefold())
            for key, options in aliases.items():
                normalized_options = {re.sub(r"[\s_-]+", "", item.casefold()) for item in options}
                if normalized in normalized_options:
                    columns[key] = cell.column
        required = {"username", "password", "firstName", "lastName", "email", "level"}
        if required.issubset(columns):
            return row_idx, columns
    raise ValueError("Could not find the required user columns in users.xlsx.")


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
        if urlsplit(self.path).path == "/users":
            return self._handle_list_users()
        if urlsplit(self.path).path == "/records":
            return self._handle_list_records()
        guarded_admin_pages = {"/register.html", "/add_record.html", "/userlist.html", "/recordlist.html"}
        if urlsplit(self.path).path in guarded_admin_pages:
            session = self._session_details()
            if not session:
                return self._redirect_to_login()
            if session.get("level") != 1:
                self.send_response(303)
                self.send_header("Location", "/home.html")
                self.end_headers()
                return
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
            crs_extensions = {".pdf", ".xlsx", ".xls", ".docx", ".doc", ".png", ".jpg", ".jpeg"}
            for field, (filename, mime, payload) in uploads.items():
                if field not in folder_map or not filename:
                    continue
                suffix = Path(filename).suffix.casefold()
                is_crs = field == "commentFile"
                allowed_crs_type = suffix in crs_extensions
                allowed_pdf_type = suffix == ".pdf" or mime == "application/pdf"
                if len(payload) > 50 * 1024 * 1024 or not (allowed_crs_type if is_crs else allowed_pdf_type):
                    allowed_text = "a supported CRS document (PDF, Excel, Word, PNG, or JPG)" if is_crs else "a PDF"
                    raise ValueError(f"{filename} must be {allowed_text} no larger than 50 MB.")
                safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename).strip(" .")
                if not safe_name:
                    raise ValueError("Attachment filename is not valid.")
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
            try:
                sr_no = int(fields["srNo"]) if fields.get("srNo") else max(numbers, default=0) + 1
            except ValueError:
                raise ValueError("Serial number must be a whole number.")
            mir = Path(saved_names["mirFile"]).stem if saved_names.get("mirFile") else ""
            wir = Path(saved_names["wirFile"]).stem if saved_names.get("wirFile") else ""
            delivery = Path(saved_names["deliveryFile"]).stem if saved_names.get("deliveryFile") else ""
            comments = Path(saved_names["commentFile"]).stem if saved_names.get("commentFile") else ""
            values = {1: sr_no, 2: fields["material"], 11: mir, 12: status_fields["mirStatus"],
                      13: wir, 14: fields.get("items", ""), 15: status_fields["wirStatus"],
                      16: delivery, 17: comments}
            for col, value in values.items():
                sheet.cell(row, col).value = value
            workbook.save(WORKBOOK)
            workbook.close()
            workbook = None
            update_tracker.main()
            return self.send_json(200, {"ok": True, "row": row, "number": sr_no})
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

    def _handle_add_user(self):
        session = self._session_details()
        if not session:
            return self.send_json(401, {"ok": False, "error": "Please sign in before registering a user."})
        if session.get("level") != 1:
            return self.send_json(403, {"ok": False, "error": "Only a Level 1 user can register users."})
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "error": "This registration request was rejected."})
        workbook = None
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 1 or length > 16_384:
                raise ValueError("The registration form could not be read.")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            fields = {key: payload.get(key, "") for key in ("username", "password", "firstName", "lastName", "email", "level")}
            if any(not isinstance(value, str) for value in fields.values()):
                raise ValueError("Enter valid values in all user fields.")
            fields = {key: value.strip() for key, value in fields.items()}
            if not all(fields[key] for key in ("username", "password", "firstName", "lastName", "level")):
                raise ValueError("Complete username, password, first name, last name, and level.")
            if len(fields["username"]) > 100 or len(fields["password"]) > 200:
                raise ValueError("Username or password is too long.")
            if fields["email"] and (len(fields["email"]) > 254 or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", fields["email"])):
                raise ValueError("Enter a valid email address, or leave it blank.")
            if fields["level"] not in ("1", "2", "3"):
                raise ValueError("Choose user level 1, 2, or 3.")

            aliases = {
                "no": {"no", "number", "sno", "srno"},
                "username": {"username", "user name"},
                "password": {"password", "pass"},
                "firstName": {"firstname", "first name", "first_name"},
                "lastName": {"lastname", "last name", "last_name"},
                "email": {"email", "e-mail", "email address"},
                "level": {"level", "user level", "user_level", "role"},
            }
            with USERS_LOCK:
                if not USERS_WORKBOOK.is_file():
                    raise ValueError("users.xlsx was not found in the project folder.")
                workbook = load_workbook(USERS_WORKBOOK, data_only=False)
                sheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
                header_row, columns = None, {}
                for row_idx in range(1, min(sheet.max_row, 10) + 1):
                    values = {}
                    for cell in sheet[row_idx]:
                        if cell.value is None:
                            continue
                        normalized = re.sub(r"[\s_-]+", "", str(cell.value).strip().casefold())
                        for key, options in aliases.items():
                            normalized_options = {re.sub(r"[\s_-]+", "", item.casefold()) for item in options}
                            if normalized in normalized_options:
                                values[key] = cell.column
                    if all(key in values for key in ("username", "password", "firstName", "lastName", "email", "level")):
                        header_row, columns = row_idx, values
                        break
                if not header_row:
                    raise ValueError("Could not find the required user columns in users.xlsx.")
                username_col = columns["username"]
                for row_idx in range(header_row + 1, sheet.max_row + 1):
                    existing = sheet.cell(row_idx, username_col).value
                    if existing is not None and str(existing).strip().casefold() == fields["username"].casefold():
                        raise ValueError("That username already exists. Choose another username.")
                no_col = columns.get("no")
                next_no = 1
                if no_col:
                    existing_numbers = []
                    for row_idx in range(header_row + 1, sheet.max_row + 1):
                        try:
                            existing_numbers.append(int(sheet.cell(row_idx, no_col).value))
                        except (TypeError, ValueError):
                            pass
                    next_no = max(existing_numbers, default=0) + 1
                target_row = max(header_row + 1, sheet.max_row + 1)
                for col in range(1, sheet.max_column + 1):
                    source = sheet.cell(target_row - 1, col)
                    target = sheet.cell(target_row, col)
                    if source.has_style:
                        target._style = copy(source._style)
                    if source.number_format:
                        target.number_format = source.number_format
                row_values = {"no": next_no, **fields}
                for key, value in row_values.items():
                    if key in columns:
                        sheet.cell(target_row, columns[key]).value = value
                workbook.save(USERS_WORKBOOK)
                workbook.close()
                workbook = None
            log("User registered by Level 1 administrator.")
            return self.send_json(200, {"ok": True, "number": next_no, "username": fields["username"]})
        except ValueError as exc:
            return self.send_json(400, {"ok": False, "error": str(exc)})
        except PermissionError:
            return self.send_json(409, {"ok": False, "error": "Close users.xlsx in Excel, then try registering the user again."})
        except Exception as exc:
            log(f"User registration error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The user could not be saved to users.xlsx."})
        finally:
            if workbook is not None:
                workbook.close()

    def _admin_session(self):
        session = self._session_details()
        if not session:
            self.send_json(401, {"ok": False, "error": "Please sign in again."})
            return None
        if session.get("level") != 1:
            self.send_json(403, {"ok": False, "error": "Only a Level 1 user can manage the user list."})
            return None
        return session

    def _handle_list_users(self):
        if not self._admin_session():
            return
        workbook = None
        try:
            with USERS_LOCK:
                workbook = load_workbook(USERS_WORKBOOK, data_only=True, read_only=True)
                sheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
                header_row, columns = locate_users_header(sheet)
                rows = []
                next_number = 1
                for row_idx in range(header_row + 1, sheet.max_row + 1):
                    username = sheet.cell(row_idx, columns["username"]).value
                    password = sheet.cell(row_idx, columns["password"]).value
                    if username is None or not str(username).strip():
                        continue
                    number = sheet.cell(row_idx, columns["no"]).value if "no" in columns else None
                    if number in (None, ""):
                        number = next_number
                    try:
                        next_number = max(next_number, int(number) + 1)
                    except (TypeError, ValueError):
                        pass
                    try:
                        level = int(sheet.cell(row_idx, columns["level"]).value or 3)
                    except (TypeError, ValueError):
                        level = 3
                    rows.append({
                        "row": row_idx,
                        "number": number,
                        "username": str(username).strip(),
                        "firstName": str(sheet.cell(row_idx, columns["firstName"]).value or "").strip(),
                        "lastName": str(sheet.cell(row_idx, columns["lastName"]).value or "").strip(),
                        "email": str(sheet.cell(row_idx, columns["email"]).value or "").strip(),
                        "level": level if level in (1, 2, 3) else 3,
                        "passwordSet": password is not None and str(password) != "",
                    })
                workbook.close()
                workbook = None
            return self.send_json(200, {"ok": True, "users": rows})
        except Exception as exc:
            log(f"User list error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The user list could not be read from users.xlsx."})
        finally:
            if workbook is not None:
                workbook.close()

    def _handle_edit_user(self):
        admin = self._admin_session()
        if not admin:
            return
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "error": "This user update request was rejected."})
        workbook = None
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 1 or length > 16_384:
                raise ValueError("The user update form could not be read.")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            try:
                row_idx = int(payload.get("row"))
            except (TypeError, ValueError):
                raise ValueError("Select a valid user to edit.")
            keys = ("username", "firstName", "lastName", "email", "level", "password")
            fields = {key: payload.get(key, "") for key in keys}
            if any(not isinstance(value, str) for value in fields.values()):
                raise ValueError("Enter valid values in the user fields.")
            fields = {key: value.strip() for key, value in fields.items()}
            if not all(fields[key] for key in ("username", "firstName", "lastName", "level")):
                raise ValueError("Username, first name, last name, and level are required.")
            if fields["level"] not in ("1", "2", "3"):
                raise ValueError("Choose user level 1, 2, or 3.")
            if fields["email"] and (len(fields["email"]) > 254 or not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", fields["email"])):
                raise ValueError("Enter a valid email address, or leave it blank.")
            if len(fields["username"]) > 100 or len(fields["password"]) > 200:
                raise ValueError("Username or password is too long.")

            old_username = ""
            with USERS_LOCK:
                workbook = load_workbook(USERS_WORKBOOK, data_only=False)
                sheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
                header_row, columns = locate_users_header(sheet)
                if row_idx <= header_row or row_idx > sheet.max_row:
                    raise ValueError("The selected user row no longer exists. Refresh the list and try again.")
                old_username = str(sheet.cell(row_idx, columns["username"]).value or "").strip()
                if not old_username:
                    raise ValueError("The selected user row is empty. Refresh the list and try again.")
                try:
                    old_level = int(sheet.cell(row_idx, columns["level"]).value or 3)
                except (TypeError, ValueError):
                    old_level = 3
                if old_username.casefold() == str(admin.get("username", "")).casefold() and fields["level"] != "1":
                    raise ValueError("Use another Level 1 account to change your own access level.")
                if old_level == 1 and fields["level"] != "1":
                    level_one_count = 0
                    for check_row in range(header_row + 1, sheet.max_row + 1):
                        name = sheet.cell(check_row, columns["username"]).value
                        try:
                            level = int(sheet.cell(check_row, columns["level"]).value or 3)
                        except (TypeError, ValueError):
                            level = 3
                        if name is not None and str(name).strip() and level == 1:
                            level_one_count += 1
                    if level_one_count <= 1:
                        raise ValueError("You cannot downgrade the last Level 1 account.")
                for check_row in range(header_row + 1, sheet.max_row + 1):
                    other = str(sheet.cell(check_row, columns["username"]).value or "").strip()
                    if check_row != row_idx and other.casefold() == fields["username"].casefold():
                        raise ValueError("That username already exists. Choose another username.")
                for key in ("username", "firstName", "lastName", "email"):
                    sheet.cell(row_idx, columns[key]).value = fields[key]
                sheet.cell(row_idx, columns["level"]).value = int(fields["level"])
                if fields["password"]:
                    sheet.cell(row_idx, columns["password"]).value = fields["password"]
                workbook.save(USERS_WORKBOOK)
                workbook.close()
                workbook = None

            token = self._session_token()
            with SESSION_LOCK:
                for other_token, active in list(SESSIONS.items()):
                    if not isinstance(active, dict) or str(active.get("username", "")).casefold() != old_username.casefold():
                        continue
                    if old_username.casefold() == str(admin.get("username", "")).casefold() and other_token == token:
                        active.update({"username": fields["username"], "firstName": fields["firstName"],
                                       "lastName": fields["lastName"], "level": int(fields["level"])})
                    else:
                        SESSIONS.pop(other_token, None)
            log("User account updated by Level 1 administrator.")
            return self.send_json(200, {"ok": True})
        except ValueError as exc:
            return self.send_json(400, {"ok": False, "error": str(exc)})
        except PermissionError:
            return self.send_json(409, {"ok": False, "error": "Close users.xlsx in Excel, then try saving the user changes."})
        except Exception as exc:
            log(f"User update error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The user changes could not be saved."})
        finally:
            if workbook is not None:
                workbook.close()

    def _handle_delete_user(self):
        admin = self._admin_session()
        if not admin:
            return
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "error": "This user deletion request was rejected."})
        workbook = None
        target_username = ""
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 1 or length > 4_096:
                raise ValueError("The user deletion request could not be read.")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            try:
                row_idx = int(payload.get("row"))
            except (TypeError, ValueError):
                raise ValueError("Select a valid user to delete.")
            with USERS_LOCK:
                workbook = load_workbook(USERS_WORKBOOK, data_only=False)
                sheet = workbook["Sheet1"] if "Sheet1" in workbook.sheetnames else workbook[workbook.sheetnames[0]]
                header_row, columns = locate_users_header(sheet)
                if row_idx <= header_row or row_idx > sheet.max_row:
                    raise ValueError("The selected user row no longer exists. Refresh the list and try again.")
                target_username = str(sheet.cell(row_idx, columns["username"]).value or "").strip()
                if not target_username:
                    raise ValueError("The selected user row is empty. Refresh the list and try again.")
                if target_username.casefold() == str(admin.get("username", "")).casefold():
                    raise ValueError("You cannot delete the account you are currently using.")
                level_one_count = 0
                for check_row in range(header_row + 1, sheet.max_row + 1):
                    name = sheet.cell(check_row, columns["username"]).value
                    try:
                        level = int(sheet.cell(check_row, columns["level"]).value or 3)
                    except (TypeError, ValueError):
                        level = 3
                    if name is not None and str(name).strip() and level == 1:
                        level_one_count += 1
                try:
                    target_level = int(sheet.cell(row_idx, columns["level"]).value or 3)
                except (TypeError, ValueError):
                    target_level = 3
                if target_level == 1 and level_one_count <= 1:
                    raise ValueError("You cannot delete the last Level 1 account.")
                sheet.delete_rows(row_idx, 1)
                workbook.save(USERS_WORKBOOK)
                workbook.close()
                workbook = None
            with SESSION_LOCK:
                for other_token, active in list(SESSIONS.items()):
                    if isinstance(active, dict) and str(active.get("username", "")).casefold() == target_username.casefold():
                        SESSIONS.pop(other_token, None)
            log("User account deleted by Level 1 administrator.")
            return self.send_json(200, {"ok": True})
        except ValueError as exc:
            return self.send_json(400, {"ok": False, "error": str(exc)})
        except PermissionError:
            return self.send_json(409, {"ok": False, "error": "Close users.xlsx in Excel, then try deleting the user again."})
        except Exception as exc:
            log(f"User deletion error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The user could not be deleted from users.xlsx."})
        finally:
            if workbook is not None:
                workbook.close()

    def _handle_list_records(self):
        if not self._admin_session():
            return
        workbook = None
        try:
            with TRACKER_LOCK:
                workbook = load_workbook(WORKBOOK, data_only=True, read_only=True)
                records = update_tracker.extract_records(workbook)
                workbook.close()
                workbook = None
                attachments = update_tracker.attachment_index()
                update_tracker.attach_matches(records, attachments)
            result = [{
                "row": int(rec["row"]),
                "number": rec.get("number", ""),
                "description": rec.get("description", ""),
                "mir": rec.get("mir", ""),
                "mirStatus": rec.get("mirStatus", "Process"),
                "wir": rec.get("wir", ""),
                "items": rec.get("item", ""),
                "wirStatus": rec.get("wirStatus", "Process"),
                "delivery": rec.get("delivery", ""),
                "comments": rec.get("comments", ""),
                "files": rec.get("files", []),
                "deliveryFile": rec.get("deliveryFile"),
                "commentFile": rec.get("commentFile"),
            } for rec in records]
            return self.send_json(200, {"ok": True, "records": result})
        except Exception as exc:
            log(f"Record list error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The record list could not be read from the tracker workbook."})
        finally:
            if workbook is not None:
                workbook.close()

    def _handle_edit_record(self):
        admin = self._admin_session()
        if not admin:
            return
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "error": "This record update request was rejected."})
        workbook = None
        created_files = []
        workbook_saved = False
        try:
            length = int(self.headers.get("Content-Length", "0"))
            content_type = self.headers.get("Content-Type", "")
            if length < 1 or length > 220 * 1024 * 1024 or not content_type.lower().startswith("multipart/form-data"):
                raise ValueError("The record update form could not be read.")
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
                data = part.get_payload(decode=True) or b""
                if filename:
                    uploads[name] = (Path(filename).name, part.get_content_type(), data)
                else:
                    fields[name] = data.decode(part.get_content_charset() or "utf-8", errors="replace").strip()
            try:
                row = int(fields.get("row", ""))
            except (TypeError, ValueError):
                raise ValueError("Select a valid record to edit.")
            field_limits = {"description": 10_000, "mir": 500, "wir": 500, "items": 10_000,
                            "delivery": 1_000, "comments": 3_000}
            for key in (*field_limits, "mirStatus", "wirStatus"):
                fields[key] = str(fields.get(key, "")).strip()
            if not fields["description"]:
                raise ValueError("Material & Location is required.")
            for key, max_length in field_limits.items():
                if len(fields[key]) > max_length:
                    raise ValueError(f"{key.capitalize()} is too long.")
            if fields["mirStatus"] not in VALID_STATUSES or fields["wirStatus"] not in VALID_STATUSES:
                raise ValueError("Choose valid MIR and WIR statuses.")

            folder_map = {"mirFile": "TTG MIR", "wirFile": "TTG WIR", "deliveryFile": "Delivery Notes", "commentFile": "CRS"}
            crs_extensions = {".pdf", ".xlsx", ".xls", ".docx", ".doc", ".png", ".jpg", ".jpeg"}
            for file_field, (filename, mime, data) in uploads.items():
                if file_field not in folder_map:
                    continue
                suffix = Path(filename).suffix.casefold()
                is_crs = file_field == "commentFile"
                valid_type = suffix in crs_extensions if is_crs else (suffix == ".pdf" or mime == "application/pdf")
                if len(data) > 50 * 1024 * 1024 or not valid_type:
                    allowed_text = "a supported CRS document (PDF, Excel, Word, PNG, or JPG)" if is_crs else "a PDF"
                    raise ValueError(f"{filename} must be {allowed_text} no larger than 50 MB.")

            for file_field, (filename, mime, data) in uploads.items():
                if file_field not in folder_map:
                    continue
                safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename).strip(" .")
                if not safe_name:
                    raise ValueError("Attachment filename is not valid.")
                if Path(safe_name).suffix == "":
                    safe_name += ".pdf"
                folder = ROOT / folder_map[file_field]
                folder.mkdir(parents=True, exist_ok=True)
                target = folder / safe_name
                if target.exists():
                    stamp = datetime.now().strftime("%Y%m%d%H%M%S")
                    target = folder / f"{target.stem}_{stamp}{target.suffix}"
                target.write_bytes(data)
                created_files.append(target)
                ref_field = {"mirFile": "mir", "wirFile": "wir", "deliveryFile": "delivery", "commentFile": "comments"}[file_field]
                fields[ref_field] = target.stem

            with TRACKER_LOCK:
                workbook = load_workbook(WORKBOOK, data_only=False)
                sheet = workbook["Sheet1"]
                header = update_tracker.find_header(sheet, ["sr. no", "material", "mir", "wir"])
                if not header:
                    raise ValueError("Could not locate the MIR/WIR table in Sheet1.")
                current_records = update_tracker.extract_records(workbook)
                if row <= header or not any(int(rec["row"]) == row for rec in current_records):
                    raise ValueError("The selected record no longer exists. Refresh the list and try again.")
                values = {2: fields["description"], 11: fields["mir"], 12: fields["mirStatus"],
                          13: fields["wir"], 14: fields["items"], 15: fields["wirStatus"],
                          16: fields["delivery"], 17: fields["comments"]}
                for col, value in values.items():
                    sheet.cell(row, col).value = value
                workbook.save(WORKBOOK)
                workbook.close()
                workbook = None
                workbook_saved = True
            update_tracker.main()
            log("Tracker record updated by Level 1 administrator.")
            return self.send_json(200, {"ok": True, "references": {key: fields[key] for key in ("mir", "wir", "delivery", "comments")}})
        except ValueError as exc:
            if not workbook_saved:
                for path in created_files:
                    path.unlink(missing_ok=True)
            return self.send_json(400, {"ok": False, "error": str(exc)})
        except PermissionError:
            if not workbook_saved:
                for path in created_files:
                    path.unlink(missing_ok=True)
            return self.send_json(409, {"ok": False, "error": "Close the tracker workbook in Excel, then save the record again."})
        except Exception as exc:
            if not workbook_saved:
                for path in created_files:
                    path.unlink(missing_ok=True)
            log(f"Record update error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The record changes could not be saved."})
        finally:
            if workbook is not None:
                workbook.close()

    def _handle_delete_record(self):
        admin = self._admin_session()
        if not admin:
            return
        if not self._origin_allowed():
            return self.send_json(403, {"ok": False, "error": "This record deletion request was rejected."})
        workbook = None
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length < 1 or length > 4_096:
                raise ValueError("The record deletion request could not be read.")
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
            try:
                row = int(payload.get("row"))
            except (TypeError, ValueError):
                raise ValueError("Select a valid record to delete.")
            with TRACKER_LOCK:
                workbook = load_workbook(WORKBOOK, data_only=False)
                sheet = workbook["Sheet1"]
                header = update_tracker.find_header(sheet, ["sr. no", "material", "mir", "wir"])
                if not header:
                    raise ValueError("Could not locate the MIR/WIR table in Sheet1.")
                current_records = update_tracker.extract_records(workbook)
                if row <= header or not any(int(rec["row"]) == row for rec in current_records):
                    raise ValueError("The selected record no longer exists. Refresh the list and try again.")
                sheet.delete_rows(row, 1)
                workbook.save(WORKBOOK)
                workbook.close()
                workbook = None
            update_tracker.main()
            log("Tracker record deleted by Level 1 administrator.")
            return self.send_json(200, {"ok": True})
        except ValueError as exc:
            return self.send_json(400, {"ok": False, "error": str(exc)})
        except PermissionError:
            return self.send_json(409, {"ok": False, "error": "Close the tracker workbook in Excel, then try deleting the record."})
        except Exception as exc:
            log(f"Record deletion error: {type(exc).__name__}")
            return self.send_json(500, {"ok": False, "error": "The record could not be deleted from the tracker workbook."})
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
        if route == "/edit-record":
            return self._handle_edit_record()
        if route == "/delete-record":
            return self._handle_delete_record()
        if route == "/add-user":
            return self._handle_add_user()
        if route == "/edit-user":
            return self._handle_edit_user()
        if route == "/delete-user":
            return self._handle_delete_user()
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
