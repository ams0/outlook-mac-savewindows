#!/usr/bin/env python3
"""
save_windows.py — Snapshot open Outlook email window titles to a JSON state file.

Runs every 2 minutes via launchd. Only acts when Microsoft Outlook is running.
Writes atomically to avoid corruption on abrupt shutdown.
"""

import json
import os
import subprocess
import sys
import tempfile
from datetime import datetime

STATE_FILE = os.path.expanduser("~/.outlook-windows-state.json")
MAX_SNAPSHOTS = 5

# Titles to skip — main app window and known non-email windows
SKIP_TITLES = {
    "Microsoft Outlook",
    "",
}

# Heuristic: skip very short titles (likely dialogs/status windows)
MIN_TITLE_LENGTH = 4


def is_outlook_running() -> bool:
    result = subprocess.run(
        ["pgrep", "-x", "Microsoft Outlook"],
        capture_output=True,
    )
    return result.returncode == 0


def get_outlook_window_titles() -> list[str]:
    script = """\
tell application "System Events"
    if exists process "Microsoft Outlook" then
        tell process "Microsoft Outlook"
            return name of every window
        end tell
    end if
    return {}
end tell
"""
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        print(f"osascript error: {result.stderr.strip()}", file=sys.stderr)
        return []

    raw = result.stdout.strip()
    if not raw:
        return []

    # osascript returns a comma-separated list for AppleScript lists
    titles = [t.strip() for t in raw.split(",")]
    return titles


def filter_titles(titles: list[str]) -> list[str]:
    filtered = []
    for title in titles:
        if title in SKIP_TITLES:
            continue
        if len(title) < MIN_TITLE_LENGTH:
            continue
        filtered.append(title)
    return filtered


def load_existing_state() -> dict:
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def write_state_atomic(state: dict) -> None:
    dir_ = os.path.dirname(STATE_FILE)
    with tempfile.NamedTemporaryFile(
        mode="w",
        dir=dir_,
        suffix=".tmp",
        delete=False,
    ) as tf:
        json.dump(state, tf, indent=2)
        tmp_path = tf.name
    os.replace(tmp_path, STATE_FILE)


def main() -> None:
    if not is_outlook_running():
        # Don't overwrite good state when Outlook isn't open
        sys.exit(0)

    titles = get_outlook_window_titles()
    titles = filter_titles(titles)

    now = datetime.now().isoformat(timespec="seconds")

    existing = load_existing_state()
    snapshots = existing.get("snapshots", [])

    snapshot = {
        "saved_at": now,
        "windows": [{"title": t, "saved_at": now} for t in titles],
    }

    snapshots.append(snapshot)
    # Keep only the last MAX_SNAPSHOTS
    snapshots = snapshots[-MAX_SNAPSHOTS:]

    state = {
        "saved_at": now,
        "windows": snapshot["windows"],
        "snapshots": snapshots,
    }

    write_state_atomic(state)


if __name__ == "__main__":
    main()
