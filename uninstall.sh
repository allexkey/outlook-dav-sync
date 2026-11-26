#!/bin/bash
# Remove outlook-dav-sync for the current macOS user. Asks before deleting anything
# beyond the background job. Server data is never touched.
set -uo pipefail

LABEL="local.outlook-dav-sync"
APP_HOME="$HOME/Library/Application Support/outlook-dav-sync"
BIN="$APP_HOME/bin"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
CFG="$APP_HOME/config.json"

ask() { read -r -p "$1 [y/N] " a; [[ "$a" =~ ^[Yy] ]]; }
cfgval() { /usr/bin/plutil -extract "$1" raw -o - "$CFG" 2>/dev/null; }

echo "Stopping background sync"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
rm -f "$PLIST"

if [ -f "$CFG" ]; then
  CAL="$(cfgval outlook_calendar)"; AB="$(cfgval outlook_addressbook)"; USERNAME="$(cfgval username)"
  if [ -f "$BIN/outlook.js" ] && ask "Delete the Outlook folders \"$CAL\" and \"$AB\" (only the local copies)?"; then
    osascript -l JavaScript "$BIN/outlook.js" removeFolder calendar "$CAL"
    osascript -l JavaScript "$BIN/outlook.js" removeFolder addressbook "$AB"
    echo
  fi
  if [ -n "$USERNAME" ] && ask "Delete the saved password for $USERNAME from Keychain?"; then
    security delete-generic-password -s outlook-dav-sync -a "$USERNAME" >/dev/null && echo "Password removed."
  fi
fi

if ask "Delete settings, sync state and backups in $APP_HOME?"; then
  rm -rf "$APP_HOME"
  echo "Removed $APP_HOME"
else
  rm -rf "$BIN"
  echo "Program removed; settings and backups kept in $APP_HOME"
fi
echo "Uninstalled."
