#!/usr/bin/env python3
"""
restore_windows.py — Restore previously open Outlook email windows after a crash/reboot.

Tries three strategies per window:
  A) Outlook AppleScript message search (best effort — may not work on all versions)
  B) UI automation via System Events search shortcut (fallback)
  C) Summary dialog listing all saved titles (always shown at the end)
"""

import json
import os
import subprocess
import sys
import time

STATE_FILE = os.path.expanduser("~/.outlook-windows-state.json")


def load_state() -> dict:
    try:
        with open(STATE_FILE) as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"State file not found: {STATE_FILE}", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"State file is corrupt: {e}", file=sys.stderr)
        sys.exit(1)


def run_applescript(script: str) -> tuple[bool, str]:
    result = subprocess.run(
        ["osascript", "-e", script],
        capture_output=True,
        text=True,
    )
    return result.returncode == 0, result.stderr.strip()


def ensure_outlook_open() -> None:
    script = """\
tell application "Microsoft Outlook"
    activate
end tell
"""
    run_applescript(script)
    time.sleep(2)


def strategy_a_applescript_search(subject: str) -> bool:
    """Try to find and open the message via Outlook AppleScript."""
    # Escape double quotes in subject
    safe_subject = subject.replace('"', '\\"')
    script = f"""\
tell application "Microsoft Outlook"
    activate
    set foundMessages to {{}}
    repeat with anAccount in every account
        repeat with aFolder in every mail folder of anAccount
            try
                set msgs to (messages of aFolder whose subject contains "{safe_subject}")
                if (count of msgs) > 0 then
                    set foundMessages to foundMessages & msgs
                end if
            end try
        end repeat
    end repeat
    if (count of foundMessages) > 0 then
        open item 1 of foundMessages
        return "ok"
    else
        return "not_found"
    end if
end tell
"""
    success, err = run_applescript(script)
    if not success:
        return False
    return True


def strategy_b_ui_search(subject: str) -> bool:
    """Trigger Outlook's search UI and type the subject."""
    # Escape special characters for AppleScript string
    safe_subject = subject.replace('"', '\\"')
    script = f"""\
tell application "Microsoft Outlook"
    activate
end tell
delay 0.5
tell application "System Events"
    tell process "Microsoft Outlook"
        -- Try Cmd+E (search shortcut)
        keystroke "e" using command down
        delay 1
        keystroke "{safe_subject}"
        delay 0.5
        key code 36  -- Return
    end tell
end tell
"""
    success, _ = run_applescript(script)
    return success


def strategy_c_summary_dialog(windows: list[dict], saved_at: str) -> None:
    """Show a native dialog listing all saved window titles."""
    if not windows:
        msg = "No email windows were saved in the last snapshot."
    else:
        lines = [f"Saved at: {saved_at}", "", "Open email windows at time of snapshot:"]
        for i, w in enumerate(windows, 1):
            lines.append(f"  {i}. {w['title']}")
        lines.extend([
            "",
            "These windows could not all be restored automatically.",
            "Use Outlook's search (Cmd+E) to find them manually.",
        ])
        msg = "\n".join(lines)

    # Escape for AppleScript
    safe_msg = msg.replace('"', '\\"').replace("\\", "\\\\")
    script = f"""\
tell application "Microsoft Outlook"
    activate
end tell
tell application "System Events"
    display dialog "{safe_msg}" with title "Outlook Window Restore" buttons {{"OK"}} default button "OK"
end tell
"""
    run_applescript(script)


def restore_window(window: dict) -> bool:
    title = window["title"]
    print(f"  Restoring: {title!r}")

    # Strategy A
    try:
        if strategy_a_applescript_search(title):
            print(f"    Strategy A succeeded.")
            return True
    except Exception as e:
        print(f"    Strategy A error: {e}", file=sys.stderr)

    # Strategy B
    try:
        if strategy_b_ui_search(title):
            print(f"    Strategy B (UI search) triggered.")
            time.sleep(1.5)
            return True
    except Exception as e:
        print(f"    Strategy B error: {e}", file=sys.stderr)

    print(f"    Could not auto-restore — will appear in summary.")
    return False


def main() -> None:
    state = load_state()
    windows = state.get("windows", [])
    saved_at = state.get("saved_at", "unknown")

    if not windows:
        print("No windows in state file. Nothing to restore.")
        strategy_c_summary_dialog([], saved_at)
        return

    print(f"Restoring {len(windows)} window(s) from snapshot at {saved_at}")
    ensure_outlook_open()

    failed = []
    for window in windows:
        ok = restore_window(window)
        if not ok:
            failed.append(window)
        time.sleep(0.5)

    print(f"\nDone. {len(windows) - len(failed)}/{len(windows)} window(s) restored automatically.")

    # Always show the summary dialog
    strategy_c_summary_dialog(windows, saved_at)


if __name__ == "__main__":
    main()
