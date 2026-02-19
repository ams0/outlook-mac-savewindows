#!/usr/bin/env bash
set -euo pipefail

LABEL="com.user.outlook-savewindows"
PLIST_NAME="${LABEL}.plist"
LAUNCH_AGENTS_DIR="${HOME}/Library/LaunchAgents"
DEST="${LAUNCH_AGENTS_DIR}/${PLIST_NAME}"

echo "Uninstalling Outlook window saver..."

if [ -f "${DEST}" ]; then
    launchctl unload "${DEST}" 2>/dev/null || true
    rm -f "${DEST}"
    echo "Removed: ${DEST}"
else
    echo "Plist not found at ${DEST} — already uninstalled?"
fi

echo ""
echo "Uninstalled. The state file ~/.outlook-windows-state.json was not removed."
echo "Delete it manually if you no longer need it:"
echo "  rm ~/.outlook-windows-state.json"
