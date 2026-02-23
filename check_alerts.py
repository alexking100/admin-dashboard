#!/usr/bin/env python3
"""check_alerts.py — send ntfy notifications for tasks overdue by 7+ days.

Run daily via cron, e.g.:
    0 9 * * * /usr/bin/python3 /path/to/admin-dashboard/check_alerts.py

Configure via .env in the same directory:
    NTFY_TOPIC=your-topic-name
    NTFY_URL=https://ntfy.sh          # optional, defaults to ntfy.sh
"""

import json
import os
import sys
import urllib.request
import urllib.error
from datetime import date
from pathlib import Path

SCRIPT_DIR  = Path(__file__).parent.resolve()
TASKS_FILE  = SCRIPT_DIR / "dashboard_tasks.json"
OVERDUE_DAYS = 7

# ── Load .env ──────────────────────────────────────────────────────────────────

def _load_env():
    env_file = SCRIPT_DIR / ".env"
    if env_file.exists():
        with open(env_file, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, _, val = line.partition("=")
                    os.environ.setdefault(key.strip(), val.strip())

_load_env()

NTFY_URL   = os.environ.get("NTFY_URL", "https://ntfy.sh").rstrip("/")
NTFY_TOPIC = os.environ.get("NTFY_TOPIC", "").strip()

# ── Helpers ────────────────────────────────────────────────────────────────────

def load_tasks():
    if not TASKS_FILE.exists():
        print(f"No tasks file found at {TASKS_FILE}", file=sys.stderr)
        return []
    with open(TASKS_FILE, encoding="utf-8") as f:
        data = json.load(f)
    return data.get("items", [])


def days_overdue(task):
    raw = task.get("added_date", "")
    if not raw:
        return 0
    try:
        added = date.fromisoformat(raw)
    except ValueError:
        return 0
    return (date.today() - added).days


def send_ntfy(task, overdue_days):
    text     = task.get("text", "(no text)").strip()
    category = task.get("category", "Other")

    tag_map = {
        "Holiday":  "beach_with_umbrella",
        "Work":     "briefcase",
        "Finances": "moneybag",
        "Other":    "pushpin",
    }
    tag = tag_map.get(category, "pushpin")

    # Escalate priority after 2 weeks
    priority = "high" if overdue_days >= 14 else "default"
    title    = f"Overdue by {overdue_days}d [{category}]"

    req = urllib.request.Request(
        f"{NTFY_URL}/{NTFY_TOPIC}",
        data=text.encode("utf-8"),
        method="POST",
    )
    req.add_header("Title",    title)
    req.add_header("Tags",     f"warning,{tag}")
    req.add_header("Priority", priority)

    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            return resp.status == 200
    except urllib.error.URLError as e:
        print(f"  ntfy error for '{text[:50]}': {e}", file=sys.stderr)
        return False


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    if not NTFY_TOPIC:
        print(
            "ERROR: NTFY_TOPIC is not set.\n"
            "Add  NTFY_TOPIC=your-topic  to the .env file next to dashboard.py",
            file=sys.stderr,
        )
        sys.exit(1)

    tasks = load_tasks()

    overdue = [
        (t, days_overdue(t))
        for t in tasks
        if not t.get("completed")
        and not t.get("deleted")
        and days_overdue(t) >= OVERDUE_DAYS
    ]

    if not overdue:
        print(f"No tasks overdue by {OVERDUE_DAYS}+ days.")
        return

    print(f"Sending {len(overdue)} overdue task(s) to ntfy topic '{NTFY_TOPIC}'")
    ok = 0
    for task, d in overdue:
        snippet = task.get("text", "")[:60]
        if send_ntfy(task, d):
            ok += 1
            print(f"  ✓  {d}d overdue: {snippet}")
        else:
            print(f"  ✗  failed:     {snippet}")

    print(f"Done: {ok}/{len(overdue)} sent.")


if __name__ == "__main__":
    main()
