#!/usr/bin/env bash
# Evaluator's authoritative screenshot capture — thin wrapper around
# eval-screenshot.js.
#
# Strategy: Joplin is launched by `launch-joplin.sh` with
# `--remote-debugging-port=$PGE_CDP_PORT` (default 8315) which exposes
# the renderer's Chrome DevTools Protocol. eval-screenshot.js uses
# Playwright's `chromium.connectOverCDP(...)` to attach, opens the
# feature's test note via the joplin:// URL scheme, waits for Univer
# to render, and saves a PNG via Playwright's `page.screenshot()`.
#
# We previously tried Playwright's `_electron.launch` — it injects
# `--inspect=0` into every Electron child, which Joplin rejects with
# a fatal "Unknown flag" modal. Then we tried macOS `screencapture`
# as a workaround, but it gave us only OS-level pixels (no DOM
# access, no waitForSelector). CDP attach gets us both real pixels
# and the renderer's DOM.
#
# Usage:
#   bash eval-screenshot.sh <feature-id>
#
# Output: prints the saved screenshot path on stdout.
#
# Env (all optional, forwarded to the JS):
#   PGE_CDP_PORT       CDP port (default 8315; must match launch-joplin.sh).
#   PGE_NOTE_ID        Skip note lookup; open this id directly.
#   PGE_OUT            Override output path (verification mode).
#   PGE_DEBUG_CONTEXTS Print every CDP page's title/URL to stderr.
set -euo pipefail

FEATURE_ID="${1:-}"
if [ -z "$FEATURE_ID" ]; then
    echo "usage: eval-screenshot.sh <feature-id>" >&2
    exit 2
fi

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"

# 1. Ensure Joplin is up (Web Clipper API + CDP both responding).
"$REPO_ROOT/scripts/pge/launch-joplin.sh" >&2

# 2. Hand off to the Node driver. It prints the saved PNG path on
# stdout; we pass that through unchanged so the evaluator agent can
# capture it from `$(bash eval-screenshot.sh ...)`.
exec node "$REPO_ROOT/scripts/pge/eval-screenshot.js" "$FEATURE_ID"
