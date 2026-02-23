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

SCOPES = [
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/calendar.events",
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
    global _docs_service, _drive_service, _calendar_service
    if _docs_service is not None:
        return _docs_service, _drive_service, _calendar_service
    if not GOOGLE_LIBS_AVAILABLE:
        print("  ℹ  Google API libraries not installed — pip3 install -r requirements.txt")
        return None, None, None
    if not CREDENTIALS_FILE.exists():
        print("  ℹ  No credentials.json — running without Google integration")
        return None, None, None

    creds             = _get_creds()
    _docs_service     = build("docs",     "v1", credentials=creds)
    _drive_service    = build("drive",    "v3", credentials=creds)
    _calendar_service = build("calendar", "v3", credentials=creds)
    print("  ✓  Google Docs, Drive + Calendar API authenticated")
    return _docs_service, _drive_service, _calendar_service


def get_docs_service():
    return _init_services()[0]


def get_drive_service():
    return _init_services()[1]


def get_calendar_service():
    return _init_services()[2]


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
Extract every action item or to-do task from Alex's Google Doc (paragraphs below).

Doc structure:
- Active to-dos appear at the top of the document.
- Items under "── COMPLETED ──" or "── REMOVED ──" headings are already handled — skip them.
- Tasks appear as bullet/numbered list items, checklist items, or action-oriented prose.
- Treat each sub-task as a separate item; never merge sub-tasks.
- Ignore headings, notes, and non-actionable descriptions.

For each task return:
  "p"        : paragraph number (int)
  "task"     : clean, imperative task text
  "done"     : true only if context strongly implies already complete; false otherwise
  "category" : one of "Holiday", "Work", "Finances", "Other"
               Holiday  = travel, holidays, leisure bookings
               Work     = professional tasks, meetings, clients, colleagues;
                          also assign Work to any item containing "(work)" in the text
               Finances = payments, invoices, bills, banking, tax, insurance
               Default to "Other" if the task doesn't clearly fit the above.

Return ONLY a valid JSON array — no explanation, no markdown. Example:
[{{"p": 3, "task": "Call Sarah about the contract", "done": false, "category": "Work"}}]

If no tasks: []

Paragraphs:
{paragraphs}"""


def extract_tasks_with_llm(all_paras: list):
    """Call Claude Haiku to identify task paragraphs.
    Returns list of {p, task, done} dicts, or None on failure."""
    client = get_anthropic_client()
    if not client:
        return None

    def _para_label(p):
        if p.get("is_checkbox"):
            prefix = "[☑ checked]" if p.get("completed") else "[☐ unchecked]"
            return f"[{p['n']}] {prefix} {p['text']}"
        return f"[{p['n']}] {p['text']}"

    numbered = "\n".join(_para_label(p) for p in all_paras)
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
            status = "✓" if r.get("done") else "○"
            cat    = r.get("category", "Other")
            task   = r.get("task", "")[:80]
            print(f"  {status} [{cat:8}] p{r.get('p','?'):>3}  {task}")
        print("  ─────────────────────────────────────────────────")
        return parsed
    except Exception as e:
        print(f"  ⚠  LLM extraction failed: {e}")
        return None


# ── Google Doc content parser ──────────────────────────────────────────────────

def parse_google_doc(doc: dict, doc_id: str, file_path: Path) -> list:
    """
    Extract task items from a Google Docs API response.
    Uses Claude Haiku when an API key is available; falls back to rule-based parsing.
    """
    source_url = f"https://docs.google.com/document/d/{doc_id}/edit"
    lists_info = doc.get("lists", {})

    # ── Step 1: Build flat paragraph list with position tracking ──────────────
    all_paras = []
    n = 0
    for element in doc.get("body", {}).get("content", []):
        para = element.get("paragraph")
        if not para:
            continue

        elem_start = element.get("startIndex", 0)
        elem_end   = element.get("endIndex",   0)

        text_parts        = []
        all_strikethrough = True
        any_strikethrough = False
        has_content       = False

        for pe in para.get("elements", []):
            tr      = pe.get("textRun", {})
            content = tr.get("content", "").rstrip("\n")
            if not content:
                continue
            has_content = True
            text_parts.append(content)
            if tr.get("textStyle", {}).get("strikethrough", False):
                any_strikethrough = True
            else:
                all_strikethrough = False

        full_text = "".join(text_parts).strip()
        if not full_text or len(full_text) < 3:
            continue

        # Detect if this paragraph is a native Google Docs checkbox (checklist) item
        is_checkbox_item = False
        if para.get("bullet"):
            list_id_val = para["bullet"].get("listId", "")
            nesting_lvl = para["bullet"].get("nestingLevel", 0)
            nl_list = (
                lists_info.get(list_id_val, {})
                          .get("listProperties", {})
                          .get("nestingLevels", [])
            )
            if nesting_lvl < len(nl_list):
                is_checkbox_item = nl_list[nesting_lvl].get("glyphType", "") == "CHECKBOX"

        # Completed if: all text runs have strikethrough (manual strike-through),
        # OR it's a checklist item where Google Docs applied strikethrough to any run
        # (Google Docs marks checked checkboxes by applying strikethrough to the text).
        completed = has_content and (
            all_strikethrough or (is_checkbox_item and any_strikethrough)
        )

        n += 1
        all_paras.append({
            "n":           n,
            "text":        full_text,
            "start":       elem_start,
            "end":         max(elem_start, elem_end - 1),
            "completed":   completed,
            "is_checkbox": is_checkbox_item,
            "bullet":      para.get("bullet"),
            "list_id":     para.get("bullet", {}).get("listId", "") if para.get("bullet") else "",
        })

    if not all_paras:
        return []

    # ── Step 2: LLM extraction ─────────────────────────────────────────────────
    llm_results = extract_tasks_with_llm(all_paras)

    if llm_results is not None:
        para_by_n = {p["n"]: p for p in all_paras}
        items     = []
        for r in llm_results:
            pn   = r.get("p")
            para = para_by_n.get(pn)
            if not para:
                continue
            task_text = (r.get("task") or "").strip() or para["text"]
            # Strikethrough in the doc always wins; LLM adds context-based completion
            completed = para["completed"] or bool(r.get("done", False))
            category  = r.get("category", "Other")
            if category not in {"Holiday", "Work", "Finances", "Other"}:
                category = "Other"
            item = _task_base(
                file_path, task_text, "llm", completed,
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
            list_id        = para.get("list_id", "")
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
                file_path, task_text, pattern_type, para["completed"],
                doc_id      = doc_id,
                start_index = para["start"],
                end_index   = para["end"],
                source_url  = source_url,
            ))
    return items


# ── Google Doc writer (feature 3) ──────────────────────────────────────────────

def sync_doc_strikethrough(doc_id: str, start_index: int, end_index: int,
                            strikethrough: bool) -> bool:
    """Apply or remove strikethrough on a paragraph range in a Google Doc."""
    service = get_docs_service()
    if not service or start_index is None or end_index is None:
        return False
    try:
        service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": [{
                "updateTextStyle": {
                    "range": {"startIndex": start_index, "endIndex": end_index},
                    "textStyle": {"strikethrough": strikethrough},
                    "fields": "strikethrough",
                }
            }]},
        ).execute()
        return True
    except Exception as e:
        print(f"  ⚠  Could not update Google Doc: {e}")
        return False


def sync_doc_text(doc_id: str, start_index: int, end_index: int,
                  new_text: str, is_completed: bool = False):
    """Replace the text of a task paragraph in a Google Doc.

    Deletes the existing content in [start_index, end_index) and inserts new_text.
    If is_completed, also reapplies strikethrough to the new text so the doc stays
    consistent with the completed state shown in the dashboard.

    Returns the new end_index (start_index + len(new_text)) on success, or None on failure.
    """
    service = get_docs_service()
    if not service or start_index is None or end_index is None:
        return None
    try:
        requests = [
            {
                "deleteContentRange": {
                    "range": {"startIndex": start_index, "endIndex": end_index}
                }
            },
            {
                "insertText": {
                    "location": {"index": start_index},
                    "text": new_text,
                }
            },
        ]
        new_end = start_index + len(new_text)
        if is_completed:
            # Reapply strikethrough so the doc reflects completed state
            requests.append({
                "updateTextStyle": {
                    "range": {"startIndex": start_index, "endIndex": new_end},
                    "textStyle": {"strikethrough": True},
                    "fields": "strikethrough",
                }
            })
        service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": requests},
        ).execute()
        return new_end
    except Exception as e:
        print(f"  ⚠  Could not update text in Google Doc: {e}")
        return None


_SECTION_DIVIDER_RE = re.compile(r"^──.*──$")


def move_doc_item_to_section(doc_id: str, start_index: int, end_index: int,
                             item_text: str, section_name: str) -> bool:
    """Move a task paragraph to a named section (COMPLETED / REMOVED) at the bottom
    of the Google Doc, creating the section heading if it doesn't exist yet.

    Key behaviours:
    - Re-discovers the item's CURRENT position by scanning the fresh doc for a
      paragraph whose text matches item_text (closest to stored start_index).
      This prevents corruption when stored indices are stale from earlier edits.
    - Appends the item AFTER the last existing entry in the section (not right
      after the heading) so multiple entries stack cleanly.
    - Prefixes items with '• ' for a tidy log appearance.
    - Falls back to stored indices if the text cannot be found in the doc.
    """
    service = get_docs_service()
    if not service:
        return False
    heading = f"── {section_name} ──"
    try:
        doc     = service.documents().get(documentId=doc_id).execute()
        content = doc.get("body", {}).get("content", [])

        doc_end          = 1
        heading_end      = None   # endIndex of the target heading paragraph
        section_last_end = None   # endIndex of the last paragraph IN the section
        in_target        = False  # are we currently inside the target section?

        actual_start = None       # current startIndex of the item to move
        actual_end   = None       # current endIndex-1 of the item to move
        best_dist    = float("inf")

        for elem in content:
            elem_end = elem.get("endIndex", 0)
            doc_end  = max(doc_end, elem_end)
            para     = elem.get("paragraph")
            if not para:
                in_target = False
                continue

            raw = "".join(
                pe.get("textRun", {}).get("content", "")
                for pe in para.get("elements", [])
            ).strip()

            # ── Track target section boundaries ──────────────────────────────
            if raw == heading:
                heading_end      = elem_end
                section_last_end = elem_end
                in_target        = True
            elif in_target:
                if _SECTION_DIVIDER_RE.match(raw):
                    in_target = False           # hit the next section
                else:
                    section_last_end = elem_end # extend section to this para

            # ── Find the current position of the item to delete ──────────────
            # Match against both the stored task text and the raw para text
            # (LLM may have slightly reworded, so check raw too)
            clean = raw.lstrip("•–- ").strip()
            if raw == item_text or clean == item_text:
                elem_s = elem.get("startIndex", 0)
                dist   = abs(elem_s - (start_index or 0))
                if dist < best_dist:
                    best_dist    = dist
                    actual_start = elem_s
                    actual_end   = max(elem_s, elem_end - 1)

        # Fall back to stored indices if the item text was not found in the doc
        if actual_start is None:
            if start_index is None or end_index is None:
                print(f"  ⚠  Could not locate '{item_text[:50]}' in doc; skipping move")
                return False
            print(f"  ℹ  Text not found in doc — using stored indices as fallback")
            actual_start = start_index
            actual_end   = end_index

        # ── Decide where to insert ───────────────────────────────────────────
        if section_last_end is not None and section_last_end > actual_end:
            # Section exists and its tail is below the item → append to section
            insert_at   = section_last_end
            insert_text = f"• {item_text}\n"
        else:
            # Section missing (or entirely above the item) → append at doc end
            insert_at   = doc_end - 1   # just before the final sentinel \n
            if heading_end is not None:
                # Heading exists but is above the item — append item only
                insert_text = f"• {item_text}\n"
            else:
                # No heading yet — create it with the first item
                insert_text = f"\n{heading}\n• {item_text}\n"

        if insert_at <= actual_end:
            print(f"  ⚠  Cannot safely move item: insert_at={insert_at} <= actual_end={actual_end}")
            return False

        # Insert first (higher index), then delete original (lower index) — safe ordering
        service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": [
                {"insertText": {
                    "location": {"index": insert_at},
                    "text": insert_text,
                }},
                {"deleteContentRange": {
                    "range": {
                        "startIndex": actual_start,
                        "endIndex":   actual_end + 1,  # +1 to include trailing \n
                    },
                }},
            ]},
        ).execute()
        return True
    except Exception as e:
        print(f"  ⚠  Could not move item to '{section_name}': {e}")
        return False


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


def merge(existing_data: dict, new_scanned: list, cached_doc_ids: set = None) -> list:
    existing       = {item["id"]: item for item in existing_data.get("items", [])}
    cached_doc_ids = cached_doc_ids or set()
    merged         = []
    seen           = set()

    # Add fresh scanned items, preserving completed/deleted state from existing
    for item in new_scanned:
        iid = item["id"]
        if iid in existing:
            old = existing[iid]
            item["completed"]  = old["completed"]
            item["deleted"]    = old["deleted"]
            item["added_date"] = old["added_date"]
            # Never overwrite a category that's already been set (by LLM or manually).
            # The LLM only gets to assign it on the item's first appearance.
            item["category"]   = old.get("category", item.get("category", "Other"))
        merged.append(item)
        seen.add(iid)

    # Re-add existing scanned items from unchanged (cached) docs
    for item in existing_data.get("items", []):
        if item["id"] in seen:
            continue
        if item.get("source") == "scanned" and item.get("doc_id") in cached_doc_ids:
            merged.append(item)
            seen.add(item["id"])

    # Always keep manual items
    for item in existing_data.get("items", []):
        if item.get("source") == "manual" and item["id"] not in seen:
            merged.append(item)

    return merged


def run_scan() -> dict:
    print("Scanning files…")
    data                                    = load_tasks()
    new_scanned, new_mtime_cache, cached_ids = scan_all(data)
    data["items"]                           = merge(data, new_scanned, cached_ids)
    data["doc_mtime_cache"]                 = new_mtime_cache
    data["last_scan"]                       = datetime.now().isoformat()
    save_tasks(data)
    scanned = sum(1 for i in data["items"] if i["source"] == "scanned" and not i["deleted"])
    manual  = sum(1 for i in data["items"] if i["source"] == "manual"  and not i["deleted"])
    cached  = len(cached_ids)
    print(f"Done — {scanned} scanned, {manual} manual ({cached} doc(s) served from cache)")
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

        # Toggle complete / incomplete  ── also syncs to Google Doc
        if re.match(r"^/api/tasks/[^/]+/toggle$", path):
            task_id = self.task_id_from_path()
            data    = load_tasks()
            for item in data["items"]:
                if item["id"] == task_id:
                    item["completed"] = not item["completed"]

                    synced = False
                    if item["completed"]:
                        # Marking DONE → move item to COMPLETED section in the doc
                        doc_id = item.get("doc_id")
                        s_idx  = item.get("doc_start_index")
                        e_idx  = item.get("doc_end_index")
                        if doc_id and s_idx is not None and e_idx is not None:
                            synced = move_doc_item_to_section(
                                doc_id, s_idx, e_idx, item["text"], "COMPLETED"
                            )
                            if synced:
                                # Item no longer at its original position in the doc
                                item["doc_start_index"] = None
                                item["doc_end_index"]   = None
                    # Toggling back to undone: one-way log — no doc change

                    save_tasks(data)
                    self.send_json({
                        "ok":        True,
                        "completed": item["completed"],
                        "doc_synced": synced,
                    })
                    return
            self.send_json({"error": "not found"}, 404)
            return

        # Edit task text  ── also syncs new text to Google Doc
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
                    item["text"] = new_text
                    doc_id = item.get("doc_id")
                    s_idx  = item.get("doc_start_index")
                    e_idx  = item.get("doc_end_index")
                    synced = False
                    if doc_id and s_idx is not None and e_idx is not None:
                        new_end = sync_doc_text(
                            doc_id, s_idx, e_idx, new_text,
                            is_completed=item.get("completed", False),
                        )
                        if new_end is not None:
                            item["doc_end_index"] = new_end
                            synced = True
                    save_tasks(data)
                    self.send_json({"ok": True, "doc_synced": synced})
                    return
            self.send_json({"error": "not found"}, 404)
            return

        # Update task category
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
                    item["category"] = category
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

        # Add manual item
        if path == "/api/tasks/add":
            body = self.read_body()
            text = body.get("text", "").strip()
            if not text:
                self.send_json({"error": "empty text"}, 400)
                return
            data     = load_tasks()
            new_item = {
                "id":              make_id("manual", text + datetime.now().isoformat()),
                "text":            text,
                "source":          "manual",
                "source_file":     None,
                "source_folder":   None,
                "file_name":       None,
                "pattern_type":    "manual",
                "completed":       False,
                "deleted":         False,
                "added_date":      datetime.now().strftime("%Y-%m-%d"),
                "category":        "Other",
                "doc_id":          None,
                "doc_start_index": None,
                "doc_end_index":   None,
                "source_url":      None,
            }
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

            # Move to REMOVED section in the Google Doc (scanned items only)
            synced = False
            doc_id = target.get("doc_id")
            s_idx  = target.get("doc_start_index")
            e_idx  = target.get("doc_end_index")
            if doc_id and s_idx is not None and e_idx is not None:
                synced = move_doc_item_to_section(
                    doc_id, s_idx, e_idx, target["text"], "REMOVED"
                )

            data["items"] = [i for i in data["items"] if i["id"] != task_id]
            save_tasks(data)
            self.send_json({"ok": True, "doc_synced": synced})
            return

        self.send_response(404); self.end_headers()


# ── Entry point ────────────────────────────────────────────────────────────────

def main():
    # Initialise Google auth before scan so auth prompt happens first
    get_docs_service()

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
