#!/usr/bin/env bash
set -euo pipefail

LABEL="com.user.outlook-savewindows"
PLIST_NAME="${LABEL}.plist"
LAUNCH_AGENTS_DIR="${HOME}/Library/LaunchAgents"
TEMPLATE="$(cd "$(dirname "$0")" && pwd)/launchd/${PLIST_NAME}"
INSTALL_DIR="$(cd "$(dirname "$0")" && pwd)"
DEST="${LAUNCH_AGENTS_DIR}/${PLIST_NAME}"

echo "Installing Outlook window saver..."
echo "  Script directory : ${INSTALL_DIR}"
echo "  Plist destination: ${DEST}"

# Substitute INSTALL_DIR placeholder in plist template
mkdir -p "${LAUNCH_AGENTS_DIR}"
sed "s|INSTALL_DIR|${INSTALL_DIR}|g" "${TEMPLATE}" > "${DEST}"

# Unload first if already loaded (ignore errors)
launchctl unload "${DEST}" 2>/dev/null || true

# Load the agent
launchctl load "${DEST}"

echo ""
echo "Installed and started successfully."
echo ""
echo "The agent will run every 2 minutes while Microsoft Outlook is open."
echo "State is saved to: ~/.outlook-windows-state.json"
echo ""
echo "To restore windows after a crash or reboot, run:"
echo "  python3 ${INSTALL_DIR}/restore_windows.py"
echo ""
echo "To check logs:"
echo "  tail -f /tmp/outlook-savewindows.log"
echo ""
echo "To uninstall:"
echo "  bash ${INSTALL_DIR}/uninstall.sh"
