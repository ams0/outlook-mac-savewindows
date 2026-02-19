# outlook-mac-savewindows

Automatically saves open Outlook email windows and restores them after a crash or reboot.

---

## Motivation

If you use Microsoft Outlook for Mac and your machine ever reboots unexpectedly — kernel panic, power cut, forced update — you lose track of every email window you had open. Unlike Safari or Terminal, Outlook does not restore its window state. You're left trying to remember what you were working on.

This tool runs silently in the background, snapshots your open email windows every two minutes, and gives you a restore script to get back to where you were.

---

## How It Works

**Saving** — A launchd agent runs `save_windows.py` every two minutes. It uses macOS System Events (via `osascript`) to read the title of every open Outlook window, filters out non-email windows (the main app window, dialogs), and writes the list atomically to `~/.outlook-windows-state.json`. If Outlook isn't running, the script exits immediately and leaves the last good snapshot untouched.

**Restoring** — Run `restore_windows.py` after a crash or reboot. It reads the snapshot and attempts to reopen each email via two strategies:

1. **AppleScript message search** — queries Outlook's mail folders directly for a message whose subject matches the saved title.
2. **UI automation fallback** — triggers Outlook's search shortcut (Cmd+E), types the subject, and presses Enter.

Regardless of how many windows were restored automatically, a native macOS dialog always appears listing every saved title so you have a reference to find anything that slipped through.

---

## Why System Events Instead of Outlook's AppleScript Dictionary?

New Outlook for Mac (the Microsoft 365 version) has a severely limited AppleScript dictionary — it exposes almost no window or message objects. Querying window titles through System Events works at the OS level and is unaffected by Outlook's AppleScript limitations.

---

## Requirements

- macOS (tested on macOS Sequoia / 15.x)
- Microsoft Outlook for Mac (any recent version)
- Python 3 — ships with macOS at `/usr/bin/python3`
- No third-party dependencies

---

## Installation

```bash
git clone https://github.com/YOUR_USERNAME/outlook-mac-savewindows.git
cd outlook-mac-savewindows
bash install.sh
```

The installer:
1. Copies the launchd plist to `~/Library/LaunchAgents/`
2. Loads the agent immediately — saving begins right away
3. Prints a confirmation with log and restore paths

To verify the agent is running:

```bash
launchctl list | grep outlook-savewindows
```

---

## Usage

### Check the saved state

After opening a few email windows and waiting up to two minutes:

```bash
cat ~/.outlook-windows-state.json
```

You should see something like:

```json
{
  "saved_at": "2026-02-19T10:30:00",
  "windows": [
    {"title": "Re: Q4 Budget Review", "saved_at": "2026-02-19T10:30:00"},
    {"title": "Fwd: Meeting Tomorrow", "saved_at": "2026-02-19T10:30:00"}
  ],
  "snapshots": [...]
}
```

### Restore after a crash or reboot

```bash
python3 restore_windows.py
```

Run this once Outlook has finished launching. The script will attempt to reopen each saved window and then show a summary dialog listing all titles.

### Check logs

```bash
tail -f /tmp/outlook-savewindows.log
```

On a normal run the log is silent. Errors from `osascript` or unexpected conditions are written here.

---

## Uninstall

```bash
bash uninstall.sh
```

This unloads the launchd agent and removes the plist from `~/Library/LaunchAgents/`. The state file at `~/.outlook-windows-state.json` is left in place — delete it manually if you no longer need it:

```bash
rm ~/.outlook-windows-state.json
```

---

## File Overview

```
outlook-mac-savewindows/
├── save_windows.py        # Snapshot agent — run by launchd every 2 minutes
├── restore_windows.py     # Restore script — run manually after a crash/reboot
├── install.sh             # Installs and loads the launchd agent
├── uninstall.sh           # Unloads and removes the launchd agent
└── launchd/
    └── com.user.outlook-savewindows.plist   # launchd agent definition (template)
```

**State file**: `~/.outlook-windows-state.json`
Stores the latest snapshot plus a rolling history of the last five snapshots.

---

## Limitations

- **Subject-based matching only** — window titles are email subjects. If two emails share the same subject, the restore script opens whichever Outlook finds first.
- **AppleScript search coverage** — Strategy A iterates all mail folders in all accounts, but very large mailboxes or certain folder types (e.g. shared mailboxes, archives) may not be fully traversed depending on your Outlook version.
- **UI automation permissions** — Strategy B (UI search fallback) requires Accessibility access for Terminal (or whichever app you run the script from). macOS will prompt for this on first use. Grant it in **System Settings → Privacy & Security → Accessibility**.
- **New Outlook only** — tested against the current Microsoft 365 / New Outlook build. The legacy Outlook 2019 build has a richer AppleScript dictionary and may not need this tool at all.

---

## Permissions

macOS will ask for permission the first time each capability is used:

| Capability | Required for | Where to grant |
|---|---|---|
| Automation → System Events | Reading window titles (save) | System Settings → Privacy & Security → Automation |
| Automation → Microsoft Outlook | Opening messages (restore) | System Settings → Privacy & Security → Automation |
| Accessibility | UI search fallback (restore) | System Settings → Privacy & Security → Accessibility |

Grant these for Terminal (or whichever app you use to run the scripts).

---

## License

MIT
