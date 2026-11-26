#!/bin/bash
# Install outlook-dav-sync for the current macOS user.
#   ./install.sh                 install or upgrade (asks for settings on first install)
#   ./install.sh --reconfigure   ask for server / account / collections again
#   ./install.sh --interval 300  sync every 300 s instead of 180 s
set -euo pipefail

LABEL="local.outlook-dav-sync"
APP_HOME="$HOME/Library/Application Support/outlook-dav-sync"
BIN="$APP_HOME/bin"
PLIST="$HOME/Library/LaunchAgents/$LABEL.plist"
SRC="$(cd "$(dirname "$0")" && pwd)"
INTERVAL=180
RECONFIGURE=0
YES=0

while [ $# -gt 0 ]; do
  case "$1" in
    --interval) INTERVAL="$2"; shift 2 ;;
    --reconfigure) RECONFIGURE=1; shift ;;
    -y|--yes) YES=1; shift ;;
    -h|--help) sed -n '2,6p' "$0"; exit 0 ;;
    *) echo "Unknown option: $1"; exit 2 ;;
  esac
done

say()  { printf '\n\033[1m==> %s\033[0m\n' "$*"; }
die()  { printf '\n\033[31mError:\033[0m %s\n' "$*" >&2; exit 1; }
confirm() { [ "$YES" = 1 ] && return 0; read -r -p "$1 [Y/n] " a; [[ -z "$a" || "$a" =~ ^[Yy] ]]; }

say "Checking requirements"
[ "$(uname)" = "Darwin" ] || die "This tool only runs on macOS."
[ -d "/Applications/Microsoft Outlook.app" ] || die "Microsoft Outlook is not installed in /Applications."
if [ "$(defaults read com.microsoft.Outlook IsRunningNewOutlook 2>/dev/null || echo 0)" = "1" ]; then
  die "Outlook is in 'New Outlook' mode, which cannot be scripted. In Outlook choose Outlook > Legacy Outlook, then run this again."
fi
PY="$(command -v python3 || true)"
[ -n "$PY" ] && "$PY" -c 'import sys; sys.exit(sys.version_info < (3, 9))' 2>/dev/null \
  || die "Python 3.9+ is needed. Install it from https://www.python.org/downloads/macos/ (or run: xcode-select --install)."
echo "Python: $PY ($("$PY" -V 2>&1))"

say "Installing files to $BIN"
mkdir -p "$BIN"
cp "$SRC/davsync.py" "$SRC/outlook.js" "$BIN/"
chmod +x "$BIN/davsync.py"

if [ "$RECONFIGURE" = 1 ] || [ ! -f "$APP_HOME/config.json" ]; then
  say "Account setup"
  "$PY" "$BIN/davsync.py" setup
else
  echo "Keeping existing settings ($APP_HOME/config.json). Use --reconfigure to change them."
fi

say "Starting Outlook"
open -g -a "Microsoft Outlook"
for _ in $(seq 1 30); do
  osascript -l JavaScript "$BIN/outlook.js" running 2>/dev/null | grep -q '"result":true' && break
  sleep 2
done
echo "If macOS asks to allow control of \"Microsoft Outlook\", click OK."

say "Dry run (nothing is changed)"
"$PY" "$BIN/davsync.py" run --dry-run | grep -v "would " || true
confirm "Run the first real sync now?" || { echo "Stopped before syncing. Run ./install.sh again when ready."; exit 0; }

say "First sync"
"$PY" "$BIN/davsync.py" run || die "First sync failed; see ~/Library/Logs/outlook-dav-sync.log"

say "Scheduling background sync every ${INTERVAL}s"
mkdir -p "$HOME/Library/LaunchAgents"
launchctl bootout "gui/$(id -u)/$LABEL" 2>/dev/null || true
cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>$LABEL</string>
  <key>ProgramArguments</key>
  <array>
    <string>$PY</string>
    <string>$BIN/davsync.py</string>
    <string>run</string>
    <string>--quiet</string>
  </array>
  <key>StartInterval</key><integer>$INTERVAL</integer>
  <key>RunAtLoad</key><true/>
  <key>ProcessType</key><string>Background</string>
  <key>LowPriorityIO</key><true/>
  <key>StandardOutPath</key><string>/dev/null</string>
  <key>StandardErrorPath</key><string>$HOME/Library/Logs/outlook-dav-sync.launchd.log</string>
</dict>
</plist>
EOF
plutil -lint "$PLIST" >/dev/null
launchctl bootstrap "gui/$(id -u)" "$PLIST"

say "Done"
cat <<EOF
Sync runs every ${INTERVAL}s while Outlook is open.
  Log:        ~/Library/Logs/outlook-dav-sync.log
  Settings:   $APP_HOME/config.json
  Sync now:   "$PY" "$BIN/davsync.py" run
  Uninstall:  ./uninstall.sh
EOF
