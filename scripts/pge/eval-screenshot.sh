#!/usr/bin/env bash
# Evaluator's authoritative screenshot capture.
#
# Strategy: launch the dev-profile Joplin via launch-joplin.sh (which
# pins --env dev), open the feature's test note via the joplin://
# URL scheme, wait for Univer to render, then capture the Joplin
# window with macOS `screencapture -l <window-id>`.
#
# We previously tried Playwright's _electron.launch — it adds
# `--inspect=0` to every Electron process, which Joplin rejects with
# a fatal "Unknown flag" modal. There's no clean way to suppress it.
# `screencapture -l` against a process the harness controls gives
# us a real PNG without that fragility.
#
# Usage:
#   bash eval-screenshot.sh <feature-id>
#
# Output: prints the saved screenshot path on stdout.
set -euo pipefail

FEATURE_ID="${1:-}"
if [ -z "$FEATURE_ID" ]; then
    echo "usage: eval-screenshot.sh <feature-id>" >&2
    exit 2
fi

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
SHOT_DIR="$REPO_ROOT/screenshots/$FEATURE_ID"
mkdir -p "$SHOT_DIR"
UTC=$(date -u '+%Y-%m-%dT%H-%M-%SZ')
OUT="$SHOT_DIR/eval-$UTC.png"

# 1. Ensure dev Joplin is up (pins --env dev internally).
"$REPO_ROOT/scripts/pge/launch-joplin.sh" >&2

# 2. Discover API + token.
HOSTPORT=$("$REPO_ROOT/scripts/pge/discover-api-port.sh")
TOKEN=""
if [ -n "${JOPLIN_TOKEN:-}" ]; then
    TOKEN="$JOPLIN_TOKEN"
elif [ -f "$REPO_ROOT/.claude/joplin-token.local" ]; then
    TOKEN=$(tr -d '[:space:]' < "$REPO_ROOT/.claude/joplin-token.local")
fi
if [ -z "$TOKEN" ]; then
    echo "eval-screenshot: no Joplin token (set JOPLIN_TOKEN or write to .claude/joplin-token.local)" >&2
    exit 1
fi

# 3. Per-feature note title prefix. The smoke uses "PGE smoke note ".
case "$FEATURE_ID" in
    feature-1-smoke-red-cell) PREFIX="PGE smoke note " ;;
    *)
        echo "eval-screenshot: unknown feature '$FEATURE_ID' — add a title prefix to the case statement in this script" >&2
        exit 1
        ;;
esac

# 4. Find the most recent matching note by title.
NOTE_ID=$(curl -sf "http://$HOSTPORT/notes?token=$TOKEN&fields=id,title,updated_time&order_by=updated_time&order_dir=DESC" \
    | python3 -c "
import json, sys
items = json.load(sys.stdin).get('items', [])
prefix = sys.argv[1]
for n in items:
    if (n.get('title') or '').startswith(prefix):
        print(n['id'])
        break
" "$PREFIX")

if [ -z "$NOTE_ID" ]; then
    echo "eval-screenshot: no Joplin note with title prefix '$PREFIX'. Generator should have created one before invoking the evaluator." >&2
    exit 1
fi
echo "eval-screenshot: opening note id $NOTE_ID" >&2

# 5. Open the note via joplin:// URL scheme. Joplin registers the
# protocol handler at install time; macOS routes to the running
# instance.
open "joplin://x-callback-url/openNote?id=$NOTE_ID"

# 6. Give Univer time to hydrate.
sleep 5

# 7. Find Joplin's window id and capture it.
WINDOW_ID=$(osascript <<'APPLESCRIPT'
tell application "System Events"
    tell process "Joplin"
        try
            set frontWindow to first window
            return id of frontWindow as string
        on error
            return "0"
        end try
    end tell
end tell
APPLESCRIPT
)

if [ "$WINDOW_ID" = "0" ] || [ -z "$WINDOW_ID" ]; then
    echo "eval-screenshot: could not find Joplin window via AppleScript. Falling back to full-screen capture." >&2
    screencapture -x "$OUT"
else
    # Bring Joplin to front first so `screencapture -l` captures the
    # right window pixels (it captures by window id, but z-order
    # affects what's visible if windows overlap).
    osascript -e 'tell application "Joplin" to activate' >/dev/null 2>&1 || true
    sleep 1
    screencapture -l "$WINDOW_ID" -x "$OUT"
fi

if [ ! -s "$OUT" ]; then
    echo "eval-screenshot: screencapture produced no/empty file" >&2
    exit 1
fi

echo "eval-screenshot: saved $OUT" >&2
echo "$OUT"
