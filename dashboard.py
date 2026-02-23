#!/usr/bin/env python3
"""
Admin To-Do Dashboard Server
Scans .gdoc / .docx files in its own directory (Admin folder) for to-do items,
serves an interactive HTML dashboard at http://localhost:5678

Google Docs integration:
  - Place credentials.json (OAuth Desktop app) in the same folder.
  - First run opens a browser for one-click sign-in; token.json is saved for future runs.
  - Without credentials.json the scanner falls back to filename-based detection.
"""

import hashlib
import json
import os
import re
import threading
import webbrowser
from datetime import datetime, timedelta
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from urllib.parse import urlparse

from docx import Document

# ── Google API (optional) ──────────────────────────────────────────────────────
try:
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    GOOGLE_LIBS_AVAILABLE = True
except ImportError:
    GOOGLE_LIBS_AVAILABLE = False

# ── Anthropic SDK (optional) ───────────────────────────────────────────────────
try:
    import anthropic as _anthropic_sdk
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

# ── Load .env from Admin folder ────────────────────────────────────────────────
def _load_env():
    env_file = Path(__file__).parent / ".env"
    if env_file.exists():
        with open(env_file, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, val = line.partition("=")
                    os.environ.setdefault(key.strip(), val.strip())
_load_env()

# ── Config ─────────────────────────────────────────────────────────────────────
ADMIN_DIR        = Path(__file__).parent.resolve()
TASKS_FILE       = ADMIN_DIR / "dashboard_tasks.json"
HTML_FILE        = ADMIN_DIR / "dashboard.html"
CREDENTIALS_FILE = ADMIN_DIR / "credentials.json"
TOKEN_FILE       = ADMIN_DIR / "token.json"
PORT             = 5678
SHEET_ID         = os.environ.get("SHEET_ID", "").strip()
SHEET_NAME       = "Sheet1"

SCOPES = [
    "https://www.googleapis.com/auth/documents.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/calendar.events",
    "https://www.googleapis.com/auth/spreadsheets",
]

# Action verbs for heuristic detection
ACTION_VERBS = {
    "call", "email", "send", "review", "check", "fix", "update", "create",
    "write", "schedule", "book", "prepare", "contact", "follow", "submit",
    "complete", "finish", "respond", "reply", "read", "edit", "organise",
    "organize", "clean", "buy", "order", "pay", "file", "sign", "approve",
    "confirm", "arrange", "discuss", "draft", "research", "investigate",
    "resolve", "handle", "process", "upload", "download", "share", "print",
    "archive", "delete", "cancel", "renew", "register", "setup", "install",
    "test", "deploy", "backup", "migrate", "report", "interview", "hire",
    "onboard", "invoice", "chase", "collect", "transfer",
}

# ── Google Auth ────────────────────────────────────────────────────────────────

_docs_service     = None
_drive_service    = None
_calendar_service = None
_sheets_service   = None

def _get_creds():
    """Load, refresh, or obtain OAuth credentials. Handles stale scope automatically."""
    creds = None
    if TOKEN_FILE.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
            # Force re-auth if scopes are missing or don't cover all required scopes
            if not creds.scopes or not all(s in creds.scopes for s in SCOPES):
                print("  ℹ  New permissions needed — re-authenticating…")
                TOKEN_FILE.unlink()
                creds = None
        except Exception:
            TOKEN_FILE.unlink()
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                TOKEN_FILE.unlink()
                creds = None
        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            print("\n" + "─" * 60)
            print("  Google sign-in required.")
            print("  If no browser opens, copy the URL below into your browser.")
            print("─" * 60)
            creds = flow.run_local_server(port=0, open_browser=True)
        with open(TOKEN_FILE, "w", encoding="utf-8") as fh:
            fh.write(creds.to_json())

    return creds


def _init_services():
    global _docs_service, _drive_service, _calendar_service, _sheets_service
    if _docs_service is not None:
        return _docs_service, _drive_service, _calendar_service, _sheets_service
    if not GOOGLE_LIBS_AVAILABLE:
        print("  ℹ  Google API libraries not installed — pip3 install -r requirements.txt")
        return None, None, None, None
    if not CREDENTIALS_FILE.exists():
        print("  ℹ  No credentials.json — running without Google integration")
        return None, None, None, None

    creds             = _get_creds()
    _docs_service     = build("docs",     "v1", credentials=creds)
    _drive_service    = build("drive",    "v3", credentials=creds)
    _calendar_service = build("calendar", "v3", credentials=creds)
    _sheets_service   = build("sheets",   "v4", credentials=creds)
    print("  ✓  Google Docs, Drive, Calendar + Sheets API authenticated")
    return _docs_service, _drive_service, _calendar_service, _sheets_service


def get_docs_service():
    return _init_services()[0]


def get_drive_service():
    return _init_services()[1]


def get_calendar_service():
    return _init_services()[2]


def get_sheets_service():
    return _init_services()[3]


# ── Google Sheets helpers ──────────────────────────────────────────────────────

SHEET_HEADERS = ["id", "task", "category", "source", "completed",
                 "removed", "added_date", "updated_at"]
_COL = {h: i for i, h in enumerate(SHEET_HEADERS)}   # "id"→0, "task"→1 …


def sheets_init():
    """Write the header row if the sheet is blank. Safe to call every startup."""
    svc = get_sheets_service()
    if not svc or not SHEET_ID:
        return
    try:
        res = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_ID, range=f"{SHEET_NAME}!A1:A1"
        ).execute()
        if res.get("values"):
            return   # headers already present
        svc.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            body={"values": [SHEET_HEADERS]},
        ).execute()
        print("  ✓  Google Sheet initialised with headers")
    except Exception as e:
        print(f"  ⚠  sheets_init: {e}")


def sheets_load_all() -> list:
    """Return every data row as a list of dicts (skips header row)."""
    svc = get_sheets_service()
    if not svc or not SHEET_ID:
        return []
    try:
        res  = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_ID, range=SHEET_NAME
        ).execute()
        rows = res.get("values", [])
        if len(rows) < 2:
            return []
        header = rows[0]
        return [
            {header[i]: (row[i] if i < len(row) else "")
             for i in range(len(header))}
            for row in rows[1:]
        ]
    except Exception as e:
        print(f"  ⚠  sheets_load_all: {e}")
        return []


def _item_to_row(item: dict) -> list:
    """Convert a task dict to a Sheets row aligned with SHEET_HEADERS."""
    return [
        item.get("id",         ""),
        item.get("text",       ""),
        item.get("category",   "Other"),
        item.get("source",     "manual"),
        "TRUE"  if item.get("completed") else "FALSE",
        "TRUE"  if item.get("removed") or item.get("deleted") else "FALSE",
        item.get("added_date", datetime.now().strftime("%Y-%m-%d")),
        item.get("updated_at", datetime.now().isoformat()),
    ]


def _row_to_item(row: dict) -> dict:
    """Convert a Sheets row dict back to a task dict (for local cache rebuild)."""
    removed = row.get("removed", "FALSE").upper() == "TRUE"
    return {
        "id":              row.get("id",       ""),
        "text":            row.get("task",     ""),
        "category":        row.get("category", "Other"),
        "source":          row.get("source",   "manual"),
        "completed":       row.get("completed", "FALSE").upper() == "TRUE",
        "removed":         removed,
        "deleted":         removed,   # dashboard uses 'deleted' to filter
        "added_date":      row.get("added_date", ""),
        "updated_at":      row.get("updated_at", ""),
        # These fields aren't in Sheets; preserved from local JSON merge below
        "source_file":     None,
        "source_folder":   None,
        "file_name":       None,
        "pattern_type":    row.get("source", "manual"),
        "doc_id":          None,
        "doc_start_index": None,
        "doc_end_index":   None,
        "source_url":      None,
    }


def _sheets_find_row(task_id: str) -> int:
    """Return the 1-based Sheets row number for task_id, or -1 if not found."""
    svc = get_sheets_service()
    if not svc or not SHEET_ID:
        return -1
    try:
        res = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_ID, range=f"{SHEET_NAME}!A:A"
        ).execute()
        ids = [r[0] if r else "" for r in res.get("values", [])]
        try:
            return ids.index(task_id) + 1   # 1-based
        except ValueError:
            return -1
    except Exception as e:
        print(f"  ⚠  _sheets_find_row: {e}")
        return -1


def sheets_upsert(item: dict):
    """Append a new row for item. No-op if a row with item['id'] already exists.
    This enforces the additive-only rule: scan never overwrites existing data."""
    svc = get_sheets_service()
    if not svc or not SHEET_ID:
        return
    if _sheets_find_row(item["id"]) > 0:
        return   # already in sheet — do not overwrite
    try:
        svc.spreadsheets().values().append(
            spreadsheetId=SHEET_ID,
            range=f"{SHEET_NAME}!A1",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [_item_to_row(item)]},
        ).execute()
        print(f"  ➕ Sheet: added '{item['text'][:50]}'")
    except Exception as e:
        print(f"  ⚠  sheets_upsert: {e}")


def sheets_update_row(task_id: str, **item_fields):
    """Update specific fields on the Sheets row matching task_id.
    Keyword keys match item dict keys (e.g. text=, category=, completed=, removed=)."""
    # Map item dict key names → Sheets column names
    KEY_MAP = {"text": "task"}
    svc = get_sheets_service()
    if not svc or not SHEET_ID:
        return False
    row_num = _sheets_find_row(task_id)
    if row_num < 0:
        print(f"  ⚠  sheets_update_row: id '{task_id}' not found in sheet")
        return False
    try:
        col_last = chr(ord("A") + len(SHEET_HEADERS) - 1)
        res = svc.spreadsheets().values().get(
            spreadsheetId=SHEET_ID,
            range=f"{SHEET_NAME}!A{row_num}:{col_last}{row_num}",
        ).execute()
        existing = list(res.get("values", [[]])[0])
        while len(existing) < len(SHEET_HEADERS):
            existing.append("")
        for item_key, value in item_fields.items():
            col_name = KEY_MAP.get(item_key, item_key)
            if col_name in _COL:
                existing[_COL[col_name]] = (
                    ("TRUE" if value else "FALSE") if isinstance(value, bool)
                    else str(value)
                )
        existing[_COL["updated_at"]] = datetime.now().isoformat()
        svc.spreadsheets().values().update(
            spreadsheetId=SHEET_ID,
            range=f"{SHEET_NAME}!A{row_num}",
            valueInputOption="RAW",
            body={"values": [existing]},
        ).execute()
        return True
    except Exception as e:
        print(f"  ⚠  sheets_update_row: {e}")
        return False


def sheets_sync_to_local():
    """Load all rows from Sheets and rebuild local JSON cache.
    Called on startup + after each scan so the dashboard reflects Sheets truth."""
    rows = sheets_load_all()
    if not rows:
        return
    items_from_sheet = [_row_to_item(r) for r in rows]
    # Preserve metadata fields (source_file, source_url, etc.) from existing local cache
    local      = load_tasks()
    local_by_id = {i["id"]: i for i in local.get("items", [])}
    for item in items_from_sheet:
        existing = local_by_id.get(item["id"], {})
        for field in ("source_file", "source_folder", "file_name",
                      "source_url", "doc_id", "doc_start_index", "doc_end_index"):
            item[field] = existing.get(field, item.get(field))
    local["items"] = items_from_sheet
    save_tasks(local)
    active = sum(1 for i in items_from_sheet if not i.get("removed"))
    print(f"  ✓  Synced {len(items_from_sheet)} row(s) from Sheets ({active} active)")


def sheets_migrate_from_json():
    """One-time: push all existing local JSON items into Sheets (skips existing IDs)."""
    data  = load_tasks()
    items = data.get("items", [])
    if not items:
        return
    print(f"  🔄 Migrating {len(items)} item(s) from JSON → Sheets…")
    for item in items:
        item.setdefault("removed", item.get("deleted", False))
        item.setdefault("updated_at", datetime.now().isoformat())
        sheets_upsert(item)
    print("  ✓  Migration complete")


# ── Persistence ────────────────────────────────────────────────────────────────

def load_tasks():
    if TASKS_FILE.exists():
        with open(TASKS_FILE, encoding="utf-8") as f:
            return json.load(f)
    return {"items": [], "last_scan": None}


def save_tasks(data):
    with open(TASKS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


# ── Helpers ────────────────────────────────────────────────────────────────────

def make_id(key: str, text: str) -> str:
    return hashlib.md5(f"{key}::{text}".encode()).hexdigest()[:12]


def is_action_like(text: str) -> bool:
    words = text.strip().split()
    if not words or len(words) > 15:
        return False
    return words[0].lower().rstrip(",:;") in ACTION_VERBS


def rel_folder(file_path: Path) -> str:
    try:
        return str(file_path.parent.relative_to(ADMIN_DIR))
    except ValueError:
        return str(file_path.parent)


COMPLETED_RE = re.compile(r"^COMPLETED[\s_-]*", re.IGNORECASE)
DONE_RE      = re.compile(r"^DONE[\s_-]*",      re.IGNORECASE)


def _task_base(file_path: Path, task_text: str, pattern_type: str,
               completed: bool, doc_id=None, start_index=None,
               end_index=None, source_url=None) -> dict:
    return {
        "id":              make_id(f"{file_path}:{start_index}", task_text),
        "text":            task_text,
        "source":          "scanned",
        "source_file":     str(file_path),
        "source_folder":   rel_folder(file_path),
        "file_name":       file_path.name,
        "pattern_type":    pattern_type,
        "completed":       completed,
        "deleted":         False,
        "added_date":      datetime.now().strftime("%Y-%m-%d"),
        "category":        "Other",
        "doc_id":          doc_id,
        "doc_start_index": start_index,
        "doc_end_index":   end_index,
        "source_url":      source_url,
    }


# ── .gdoc metadata reader ──────────────────────────────────────────────────────

def _search_drive_for_doc(name: str):
    """Search Google Drive for a Google Doc with the given name.
    Returns (webViewLink, doc_id) or (None, None)."""
    drive = get_drive_service()
    if not drive:
        return None, None
    try:
        safe = name.replace("\\", "\\\\").replace("'", "\\'")
        res  = drive.files().list(
            q=(f"name='{safe}' "
               f"and mimeType='application/vnd.google-apps.document' "
               f"and trashed=false"),
            fields="files(id, name, webViewLink)",
            pageSize=5,
        ).execute()
        files = res.get("files", [])
        if files:
            f = files[0]
            return f.get("webViewLink"), f.get("id")
    except Exception as e:
        print(f"  ⚠  Drive search for '{name}': {e}")
    return None, None


def read_gdoc_metadata(file_path: Path):
    """Get the Google Doc URL and ID for a .gdoc file.

    Strategy:
    1. Try reading the .gdoc file directly (works if Google Drive materialises it).
    2. Fall back to searching Google Drive by filename via the Drive API.
    """
    # ── Try reading the local file ─────────────────────────────────────────────
    try:
        with open(str(file_path), "r", encoding="utf-8") as f:
            data = json.load(f)
        url = data.get("url", "")
        m   = re.search(r"/document/d/([a-zA-Z0-9_-]+)", url)
        if m:
            return url, m.group(1)
    except Exception:
        pass

    # ── Fall back: Drive API search by filename ────────────────────────────────
    # Strip COMPLETED/DONE prefix before searching
    stem = file_path.stem.strip()
    stem = COMPLETED_RE.sub("", stem).strip()
    stem = DONE_RE.sub("", stem).strip()
    return _search_drive_for_doc(stem)


# ── Anthropic LLM extraction ───────────────────────────────────────────────────

_anthropic_client = None

def get_anthropic_client():
    global _anthropic_client
    if _anthropic_client is not None:
        return _anthropic_client
    if not ANTHROPIC_AVAILABLE:
        print("  ℹ  anthropic not installed — pip3 install -r requirements.txt")
        return None
    api_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if not api_key:
        print("  ℹ  No ANTHROPIC_API_KEY — falling back to rule-based parsing")
        return None
    _anthropic_client = _anthropic_sdk.Anthropic(api_key=api_key)
    print("  ✓  Claude (Haiku) ready for task extraction")
    return _anthropic_client


TASK_PROMPT = """\
Extract every action item or to-do task from Alex's Google Doc.

Tasks appear as bullet or numbered list items, checklist items, or action-oriented prose.
Treat each sub-task as a separate item; never merge sub-tasks.
Ignore headings, notes, and non-actionable descriptions.

For each task return:
  "p"        : paragraph number (int)
  "task"     : clean, imperative task text
  "category" : one of "Holiday", "Work", "Finances", "Other"
               Holiday  = travel, holidays, leisure bookings
               Work     = professional tasks, meetings, clients, colleagues;
                          also assign Work to any item containing "(work)"
               Finances = payments, invoices, bills, banking, tax, insurance
               Default to "Other" if unclear

Return ONLY a valid JSON array — no explanation, no markdown. Example:
[{{"p": 3, "task": "Call Sarah about the contract", "category": "Work"}}]

If no tasks: []

Paragraphs:
{paragraphs}"""


def extract_tasks_with_llm(all_paras: list):
    """Call Claude Haiku to identify task paragraphs.
    Returns list of {p, task, category} dicts, or None on failure."""
    client = get_anthropic_client()
    if not client:
        return None

    numbered = "\n".join(f"[{p['n']}] {p['text']}" for p in all_paras)
    prompt   = TASK_PROMPT.format(paragraphs=numbered)

    try:
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=2048,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = msg.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$",          "", raw)
        parsed = json.loads(raw)
        print("  ── Haiku output ──────────────────────────────────")
        for r in parsed:
            cat  = r.get("category", "Other")
            task = r.get("task", "")[:80]
            print(f"  ○ [{cat:8}] p{r.get('p','?'):>3}  {task}")
        print("  ─────────────────────────────────────────────────")
        return parsed
    except Exception as e:
        print(f"  ⚠  LLM extraction failed: {e}")
        return None


# ── Google Doc content parser ──────────────────────────────────────────────────

def parse_google_doc(doc: dict, doc_id: str, file_path: Path) -> list:
    """
    Extract task items from a Google Docs API response.
    Uses Claude Haiku when available; falls back to rule-based parsing.
    The doc is treated as read-only input — all tasks enter as incomplete.
    """
    source_url = f"https://docs.google.com/document/d/{doc_id}/edit"
    lists_info = doc.get("lists", {})

    # ── Step 1: Build flat paragraph list ─────────────────────────────────────
    all_paras = []
    n = 0
    for element in doc.get("body", {}).get("content", []):
        para = element.get("paragraph")
        if not para:
            continue

        elem_start = element.get("startIndex", 0)
        elem_end   = element.get("endIndex",   0)

        text_parts = []
        for pe in para.get("elements", []):
            content = pe.get("textRun", {}).get("content", "").rstrip("\n")
            if content:
                text_parts.append(content)

        full_text = "".join(text_parts).strip()
        if not full_text or len(full_text) < 3:
            continue

        n += 1
        all_paras.append({
            "n":     n,
            "text":  full_text,
            "start": elem_start,
            "end":   max(elem_start, elem_end - 1),
            "bullet": para.get("bullet"),
        })

    if not all_paras:
        return []

    # ── Step 2: LLM extraction ─────────────────────────────────────────────────
    llm_results = extract_tasks_with_llm(all_paras)

    if llm_results is not None:
        para_by_n = {p["n"]: p for p in all_paras}
        items     = []
        for r in llm_results:
            para = para_by_n.get(r.get("p"))
            if not para:
                continue
            task_text = (r.get("task") or "").strip() or para["text"]
            category  = r.get("category", "Other")
            if category not in {"Holiday", "Work", "Finances", "Other"}:
                category = "Other"
            item = _task_base(
                file_path, task_text, "llm", False,
                doc_id      = doc_id,
                start_index = para["start"],
                end_index   = para["end"],
                source_url  = source_url,
            )
            item["category"] = category
            items.append(item)
        print(f"    ✨ Claude identified {len(items)} task(s)")
        return items

    # ── Step 3: Rule-based fallback ────────────────────────────────────────────
    print("  ℹ  Using rule-based extraction")
    items = []
    for para in all_paras:
        full_text    = para["text"]
        pattern_type = None
        task_text    = full_text
        bullet       = para.get("bullet")

        if bullet:
            list_id        = bullet.get("listId", "")
            nesting_level  = bullet.get("nestingLevel", 0)
            nesting_levels = (
                lists_info.get(list_id, {})
                          .get("listProperties", {})
                          .get("nestingLevels", [])
            )
            glyph_type = ""
            if nesting_level < len(nesting_levels):
                glyph_type = nesting_levels[nesting_level].get("glyphType", "")
            pattern_type = "numbered" if glyph_type in (
                "DECIMAL", "ZERO_DECIMAL", "UPPER_ALPHA", "ALPHA", "UPPER_ROMAN", "ROMAN"
            ) else "bullet"

        elif re.match(r"^(TODO|TASK|ACTION|FIXME)[\s:]+", full_text, re.IGNORECASE):
            task_text    = re.sub(r"^(TODO|TASK|ACTION|FIXME)[\s:]+", "",
                                  full_text, flags=re.IGNORECASE).strip()
            pattern_type = "todo_prefix"

        elif is_action_like(full_text):
            pattern_type = "action"

        if pattern_type and task_text:
            items.append(_task_base(
                file_path, task_text, pattern_type, False,
                doc_id      = doc_id,
                start_index = para["start"],
                end_index   = para["end"],
                source_url  = source_url,
            ))
    return items


def create_calendar_event(title: str, date_str: str, time_str: str,
                          duration_minutes: int):
    """Insert an event into the user's primary Google Calendar.

    date_str:          "YYYY-MM-DD"
    time_str:          "HH:MM" (24-hour) — pass "" or duration_minutes=0 for all-day
    duration_minutes:  0 → all-day; positive → timed event
    Returns the HTML link to the new event, or None on failure.
    """
    service = get_calendar_service()
    if not service:
        return None
    try:
        if not time_str or duration_minutes == 0:
            event_body = {
                "summary": title,
                "start":   {"date": date_str},
                "end":     {"date": date_str},
            }
        else:
            start_dt = datetime.strptime(f"{date_str}T{time_str}", "%Y-%m-%dT%H:%M")
            end_dt   = start_dt + timedelta(minutes=duration_minutes)
            tz       = "Europe/London"
            event_body = {
                "summary": title,
                "start":   {"dateTime": start_dt.isoformat(), "timeZone": tz},
                "end":     {"dateTime": end_dt.isoformat(),   "timeZone": tz},
            }
        result = service.events().insert(
            calendarId="primary", body=event_body,
        ).execute()
        print(f"  ✓  Calendar event created: {result.get('htmlLink')}")
        return result.get("htmlLink")
    except Exception as e:
        print(f"  ⚠  Could not create calendar event: {e}")
        return None


# ── Scanners ───────────────────────────────────────────────────────────────────

def get_doc_modified_time(doc_id: str):
    """Lightweight Drive metadata call — returns modifiedTime string or None."""
    drive = get_drive_service()
    if not drive:
        return None
    try:
        meta = drive.files().get(fileId=doc_id, fields="modifiedTime").execute()
        return meta.get("modifiedTime")
    except Exception as e:
        print(f"  ⚠  Could not get modifiedTime for {doc_id}: {e}")
        return None


def extract_from_gdoc(file_path: Path, mtime_cache: dict = None) -> tuple:
    """
    Returns (items, doc_id, new_mtime, used_cache).
    - mtime_cache: {doc_id: mtime} from a previous scan run.
    - If doc is unchanged (live modifiedTime matches cache): returns ([], doc_id, mtime, True)
      so the caller knows to preserve existing cached tasks for this doc.
    - If changed or new: fetches doc, runs LLM, returns fresh items.
    - Falls back to filename-as-task if API unavailable.
    """
    url, doc_id = read_gdoc_metadata(file_path)

    if doc_id:
        # ── Check modifiedTime before doing any heavy work ─────────────────────
        mtime        = get_doc_modified_time(doc_id)
        cached_mtime = (mtime_cache or {}).get(doc_id)
        if mtime and mtime == cached_mtime:
            print(f"  ↩  {file_path.name}: unchanged — skipping LLM")
            return [], doc_id, mtime, True   # used_cache = True

        service = get_docs_service()
        if service:
            try:
                doc   = service.documents().get(documentId=doc_id).execute()
                items = parse_google_doc(doc, doc_id, file_path)
                if items:
                    return items, doc_id, mtime, False
                # Doc exists but has no detectable tasks — fall through to filename
            except Exception as e:
                print(f"  ⚠  Docs API error for {file_path.name}: {e}")
        else:
            mtime = None
    else:
        mtime = None

    # ── Filename fallback ──────────────────────────────────────────────────────
    stem = file_path.stem.strip()
    if not stem:
        return [], doc_id, mtime, False

    completed = False
    task_text = stem
    if COMPLETED_RE.match(stem):
        completed = True
        task_text = COMPLETED_RE.sub("", stem).strip()
    elif DONE_RE.match(stem):
        completed = True
        task_text = DONE_RE.sub("", stem).strip()

    if not task_text:
        return [], doc_id, mtime, False

    return [_task_base(
        file_path, task_text, "gdoc", completed,
        doc_id     = doc_id,
        source_url = url,
    )], doc_id, mtime, False


def extract_from_docx(file_path: Path) -> list:
    items = []
    try:
        doc = Document(str(file_path))
        for para in doc.paragraphs:
            text = para.text.strip()
            if not text or len(text) < 3:
                continue

            pattern_type = None
            task_text    = text
            completed    = False

            m = re.match(r"^\[([xX ])\]\s+(.+)", text)
            if m:
                completed    = m.group(1).lower() == "x"
                task_text    = m.group(2).strip()
                pattern_type = "checkbox"
            elif re.match(r"^(TODO|TASK|ACTION|FIXME)[\s:]+", text, re.IGNORECASE):
                task_text    = re.sub(r"^(TODO|TASK|ACTION|FIXME)[\s:]+", "",
                                      text, flags=re.IGNORECASE).strip()
                pattern_type = "todo_prefix"
            elif "list" in para.style.name.lower():
                pattern_type = "bullet"
            elif re.match(r"^[-•*]\s+(.+)", text):
                task_text    = re.sub(r"^[-•*]\s+", "", text).strip()
                pattern_type = "bullet"
            elif re.match(r"^\d+[.)]\s+(.+)", text):
                task_text    = re.sub(r"^\d+[.)]\s+", "", text).strip()
                pattern_type = "numbered"
            elif is_action_like(text):
                pattern_type = "action"

            if pattern_type and task_text:
                items.append(_task_base(file_path, task_text, pattern_type, completed))
    except Exception as e:
        print(f"  ⚠  Could not read {file_path.name}: {e}")
    return items


def scan_all(existing_data: dict) -> tuple:
    """
    Scan all .gdoc and .docx files.
    Returns (new_items, new_mtime_cache, cached_doc_ids).
    - new_items:       freshly extracted items (excludes unchanged docs)
    - new_mtime_cache: {doc_id: mtime} for all gdocs with a known modifiedTime
    - cached_doc_ids:  set of doc_ids whose content is unchanged → caller re-uses existing tasks
    """
    mtime_cache     = existing_data.get("doc_mtime_cache", {})
    new_items       = []
    new_mtime_cache = {}
    cached_doc_ids  = set()

    for gdoc_file in sorted(ADMIN_DIR.rglob("*.gdoc")):
        items, doc_id, mtime, used_cache = extract_from_gdoc(gdoc_file, mtime_cache)
        if doc_id and mtime:
            new_mtime_cache[doc_id] = mtime
        if used_cache and doc_id:
            cached_doc_ids.add(doc_id)
        elif items:
            print(f"  {gdoc_file.relative_to(ADMIN_DIR)}: {len(items)} item(s)")
        new_items.extend(items)

    for docx_file in sorted(ADMIN_DIR.rglob("*.docx")):
        if docx_file.name.startswith("~$"):
            continue
        found = extract_from_docx(docx_file)
        if found:
            print(f"  {docx_file.relative_to(ADMIN_DIR)}: {len(found)} item(s)")
        new_items.extend(found)

    return new_items, new_mtime_cache, cached_doc_ids


def run_scan() -> dict:
    """Scan Google Docs, upsert new tasks into Sheets, then sync Sheets → local cache."""
    print("Scanning files…")
    data = load_tasks()
    new_scanned, new_mtime_cache, _ = scan_all(data)

    # Upsert only genuinely new items — never overwrite existing Sheets rows
    for item in new_scanned:
        item.setdefault("removed",    False)
        item.setdefault("updated_at", datetime.now().isoformat())
        sheets_upsert(item)

    # Rebuild local cache from Sheets (Sheets is now source of truth)
    sheets_sync_to_local()

    # Persist scan metadata (mtime cache + timestamp) to local JSON
    data = load_tasks()
    data["doc_mtime_cache"] = new_mtime_cache
    data["last_scan"]       = datetime.now().isoformat()
    save_tasks(data)

    active = sum(1 for i in data["items"] if not i.get("removed") and not i.get("deleted"))
    print(f"Done — {active} active task(s)")
    return data


# ── HTTP Server ────────────────────────────────────────────────────────────────

class Handler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        pass

    # ── helpers ───────────────────────────────────────────────────────────────

    def send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path, mime="text/html"):
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def read_body(self) -> dict:
        length = int(self.headers.get("Content-Length", 0))
        if length:
            return json.loads(self.rfile.read(length))
        return {}

    def task_id_from_path(self):
        """Extract task id from paths like /api/tasks/<id>/toggle"""
        parts = self.path.split("/")
        if len(parts) >= 4 and parts[2] == "tasks":
            return parts[3]
        return None

    # ── OPTIONS ───────────────────────────────────────────────────────────────

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET,POST,DELETE,OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    # ── GET ───────────────────────────────────────────────────────────────────

    def do_GET(self):
        path = urlparse(self.path).path

        if path in ("/", "/dashboard.html"):
            self.send_file(HTML_FILE)
        elif path == "/api/tasks":
            self.send_json(load_tasks())
        else:
            self.send_response(404); self.end_headers()

    # ── POST ──────────────────────────────────────────────────────────────────

    def do_POST(self):
        path = urlparse(self.path).path

        # Re-scan
        if path == "/api/scan":
            data = run_scan()
            self.send_json(data)
            return

        # Toggle complete / incomplete → writes to Sheets
        if re.match(r"^/api/tasks/[^/]+/toggle$", path):
            task_id = self.task_id_from_path()
            data    = load_tasks()
            for item in data["items"]:
                if item["id"] == task_id:
                    item["completed"]   = not item["completed"]
                    item["updated_at"]  = datetime.now().isoformat()
                    sheets_update_row(task_id, completed=item["completed"])
                    save_tasks(data)
                    self.send_json({"ok": True, "completed": item["completed"]})
                    return
            self.send_json({"error": "not found"}, 404)
            return

        # Rename task text → writes to Sheets
        if re.match(r"^/api/tasks/[^/]+/edit$", path):
            task_id  = self.task_id_from_path()
            body     = self.read_body()
            new_text = body.get("text", "").strip()
            if not new_text:
                self.send_json({"error": "empty text"}, 400)
                return
            data = load_tasks()
            for item in data["items"]:
                if item["id"] == task_id:
                    item["text"]       = new_text
                    item["updated_at"] = datetime.now().isoformat()
                    sheets_update_row(task_id, text=new_text)
                    save_tasks(data)
                    self.send_json({"ok": True})
                    return
            self.send_json({"error": "not found"}, 404)
            return

        # Update task category → writes to Sheets
        if re.match(r"^/api/tasks/[^/]+/category$", path):
            task_id  = self.task_id_from_path()
            body     = self.read_body()
            category = body.get("category", "Other").strip()
            if category not in {"Holiday", "Work", "Finances", "Other"}:
                self.send_json({"error": "invalid category"}, 400)
                return
            data = load_tasks()
            for item in data["items"]:
                if item["id"] == task_id:
                    item["category"]   = category
                    item["updated_at"] = datetime.now().isoformat()
                    sheets_update_row(task_id, category=category)
                    save_tasks(data)
                    self.send_json({"ok": True})
                    return
            self.send_json({"error": "not found"}, 404)
            return

        # Add event to Google Calendar
        if re.match(r"^/api/tasks/[^/]+/calendar$", path):
            body     = self.read_body()
            title    = body.get("title",    "").strip()
            date_s   = body.get("date",     "").strip()
            time_s   = body.get("time",     "").strip()
            duration = int(body.get("duration", 60))
            if not title or not date_s:
                self.send_json({"error": "title and date are required"}, 400)
                return
            link = create_calendar_event(title, date_s, time_s, duration)
            if link:
                self.send_json({"ok": True, "event_url": link})
            else:
                self.send_json({"error": "Could not create event — check Calendar API access"}, 500)
            return

        # Add manual item → appends to Sheets then local cache
        if path == "/api/tasks/add":
            body = self.read_body()
            text = body.get("text", "").strip()
            if not text:
                self.send_json({"error": "empty text"}, 400)
                return
            now      = datetime.now()
            new_item = {
                "id":              make_id("manual", text + now.isoformat()),
                "text":            text,
                "source":          "manual",
                "source_file":     None,
                "source_folder":   None,
                "file_name":       None,
                "pattern_type":    "manual",
                "completed":       False,
                "deleted":         False,
                "removed":         False,
                "added_date":      now.strftime("%Y-%m-%d"),
                "updated_at":      now.isoformat(),
                "category":        "Other",
                "doc_id":          None,
                "doc_start_index": None,
                "doc_end_index":   None,
                "source_url":      None,
            }
            sheets_upsert(new_item)
            data = load_tasks()
            data["items"].append(new_item)
            save_tasks(data)
            self.send_json(new_item)
            return

        self.send_response(404); self.end_headers()

    # ── DELETE ────────────────────────────────────────────────────────────────

    def do_DELETE(self):
        path    = urlparse(self.path).path
        task_id = self.task_id_from_path()

        if re.match(r"^/api/tasks/[^/]+$", path) and task_id:
            data   = load_tasks()
            target = next((i for i in data["items"] if i["id"] == task_id), None)
            if target is None:
                self.send_json({"error": "not found"}, 404)
                return

            # Mark removed in Sheets (row is never physically deleted — audit trail)
            target["removed"]    = True
            target["deleted"]    = True   # keeps dashboard filter working
            target["updated_at"] = datetime.now().isoformat()
            sheets_update_row(task_id, removed=True)
            save_tasks(data)
            self.send_json({"ok": True})
            return

        self.send_response(404); self.end_headers()


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    # Initialise Google auth (opens browser if token.json missing/stale)
    get_docs_service()

    # Set up Sheets: write headers if blank, then migrate existing JSON items
    sheets_init()
    sheets_migrate_from_json()

    run_scan()

    server = HTTPServer(("localhost", PORT), Handler)
    url    = f"http://localhost:{PORT}"

    def open_browser():
        import time; time.sleep(0.8)
        webbrowser.open(url)

    threading.Thread(target=open_browser, daemon=True).start()
    print(f"\nDashboard → {url}")
    print("Press Ctrl+C to stop\n")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
