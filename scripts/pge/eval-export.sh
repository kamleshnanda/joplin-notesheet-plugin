#!/usr/bin/env bash
# Evaluator's authoritative export capture — thin wrapper around
# eval-export.js.
#
# Strategy: like eval-screenshot.sh, Joplin is launched by
# `launch-joplin.sh` with `--remote-debugging-port=$PGE_CDP_PORT`
# (default 8315) which exposes the renderer's Chrome DevTools Protocol.
# eval-export.js attaches over CDP, drives the real "Export .xlsx"
# button in the live webview, captures the emitted blob, unzips it,
# and writes a sidecar `<out>.manifest.txt` listing the `xl/media/*`,
# `xl/charts/*`, and `xl/drawings/*` parts with a PASS/EMPTY summary.
#
# This closes the gap where an image/chart RENDERED fine (green
# screenshot) but the EXPORTED .xlsx was silently empty (Buffer
# undefined in the webview) — a render screenshot never opens the
# exported bytes.
#
# Usage:
#   bash eval-export.sh <feature-id>
#
# Output: prints the exported .xlsx path on stdout.
#
# Env (all optional, forwarded to the JS):
#   PGE_CDP_PORT       CDP port (default 8315; must match launch-joplin.sh).
#   PGE_NOTE_ID        Skip note lookup; open this id directly.
#   PGE_OUT            Override output path (verification mode).
#   PGE_DEBUG_CONTEXTS Print every CDP page's title/URL to stderr.
set -euo pipefail

FEATURE_ID="${1:-}"
if [ -z "$FEATURE_ID" ]; then
    echo "usage: eval-export.sh <feature-id>" >&2
    exit 2
fi

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"

# 1. Ensure Joplin is up (Web Clipper API + CDP both responding).
"$REPO_ROOT/scripts/pge/launch-joplin.sh" >&2

# 2. Hand off to the Node driver. It prints the exported .xlsx path on
# stdout; we pass that through unchanged so the evaluator agent can
# capture it from `$(bash eval-export.sh ...)`.
exec node "$REPO_ROOT/scripts/pge/eval-export.js" "$FEATURE_ID"
