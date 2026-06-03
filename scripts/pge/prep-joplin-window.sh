#!/usr/bin/env bash
# Idempotent: prepare Joplin's window for evaluator screenshot capture.
#
# Three things go wrong if we screenshot Joplin without prep:
#   1. Joplin window is small / not maximized → editor pane is narrower
#      than the Univer canvas needs, so Univer renders fewer columns
#      AND any whole-page screenshot crops the canvas.
#   2. Sidebar (notebooks) and note list are visible → they consume
#      ~500-700px of horizontal real estate that Univer could
#      otherwise use to render more columns. The user only cares
#      about the editor in evaluator mode.
#   3. DevTools window is open as a separate Electron window → the
#      Joplin renderer pane shrinks accordingly, AND
#      `eval-screenshot.js`'s CDP page picker can land on the
#      DevTools page if the scoring is unlucky.
#
# What this script does (in order):
#   - Bring Joplin to the front via AppleScript so subsequent menu/
#     keyboard actions go to the right app.
#   - Invoke "Window > Fill" to size Joplin to the display it's on.
#     macOS's window manager handles the multi-display case correctly.
#     Idempotent: if already at fill bounds, the menu item is a no-op.
#   - Hand off to `prep-joplin-panes.js` (Playwright/CDP) which closes
#     any DevTools page and toggles sidebar + note list off if visible.
#     The JS half handles state-aware toggling because Joplin's
#     "Toggle sidebar" / "Toggle note list" menu items don't expose
#     AXMenuItemMarked we could read from AppleScript — the renderer
#     DOM (`element.style.display === 'none'`) is the source of truth.
#
# All actions are safe to invoke repeatedly. AppleScript needs
# Accessibility permission for the terminal app in
# System Settings → Privacy & Security → Accessibility.
#
# Usage:
#   bash scripts/pge/prep-joplin-window.sh
#
# Env:
#   PGE_CDP_PORT  CDP port (default 8315; must match launch-joplin.sh).
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"

# 1. Bail out early if Joplin isn't running. Caller (launch-joplin.sh)
# is responsible for booting it.
if ! pgrep -f 'Joplin.app/Contents/MacOS/Joplin' >/dev/null 2>&1; then
    echo "prep-joplin-window: Joplin is not running. Run launch-joplin.sh first." >&2
    exit 1
fi

# 2. Activate Joplin and fill its window to the current display.
osascript >&2 <<'APPLESCRIPT'
tell application "Joplin" to activate
delay 0.2
tell application "System Events"
    tell process "Joplin"
        try
            click menu item "Fill" of menu "Window" of menu bar 1
            return "prep-joplin-window: invoked Window > Fill"
        on error errMsg
            return "prep-joplin-window: Window > Fill failed: " & errMsg
        end try
    end tell
end tell
APPLESCRIPT

# 3. Close DevTools, hide sidebar + note list (state-aware).
NODE_PATH="$REPO_ROOT/node_modules" node "$REPO_ROOT/scripts/pge/prep-joplin-panes.js"

echo "prep-joplin-window: done." >&2
