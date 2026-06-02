#!/usr/bin/env bash
# Idempotent: ensure Joplin desktop (DEV profile) is running and the
# Web Clipper Data API is responding on localhost:41184. Generator
# and evaluator both call this; safe to run repeatedly.
#
# Uses the `--env dev` profile (~/.config/joplindev-desktop/) instead
# of the operator's main Joplin profile, so PGE cycles never touch
# real notes. This matches the operator's existing workflow.
set -euo pipefail

API_BASE="${JOPLIN_API:-http://localhost:41184}"
TIMEOUT_SECONDS="${JOPLIN_BOOT_TIMEOUT:-30}"
JOPLIN_BIN="${JOPLIN_BIN:-/Applications/Joplin.app/Contents/MacOS/Joplin}"
JOPLIN_ENV_FLAG="${JOPLIN_ENV:---env dev}"

# 1. Already up? (Web Clipper /ping returns 'JoplinClipperServer')
if curl -sf -m 2 "$API_BASE/ping" >/dev/null 2>&1; then
    echo "joplin: already up at $API_BASE"
    exit 0
fi

# 2. Joplin process running but Web Clipper not? Surface the gap.
if pgrep -f 'Joplin.app/Contents/MacOS/Joplin' >/dev/null 2>&1; then
    cat >&2 <<'EOF'
joplin: app is running but Web Clipper Data API is not responding.

Open Joplin → Tools → Options → Web Clipper, click "Enable Web
Clipper service", and try again. The harness creates and inspects
test notes via the Data API; without it, every PGE cycle will fail
at evaluation time.

If you launched Joplin without `--env dev`, quit it and re-launch
this script — it will start Joplin with the dev profile so PGE
cycles never touch real notes.
EOF
    exit 1
fi

# 3. Launch in dev-profile mode.
echo "joplin: launching '$JOPLIN_BIN $JOPLIN_ENV_FLAG'..."
nohup "$JOPLIN_BIN" $JOPLIN_ENV_FLAG >/dev/null 2>&1 &
disown

# 4. Poll until the Data API responds, or until timeout.
deadline=$(( $(date +%s) + TIMEOUT_SECONDS ))
while true; do
    if curl -sf -m 2 "$API_BASE/ping" >/dev/null 2>&1; then
        echo "joplin: up (Web Clipper API responding)"
        exit 0
    fi
    if [ "$(date +%s)" -ge "$deadline" ]; then
        cat >&2 <<EOF
joplin: timed out after ${TIMEOUT_SECONDS}s waiting for Web Clipper
Data API at $API_BASE.

Possible causes:
  - First launch of dev profile? Web Clipper is OFF by default — open
    Joplin → Tools → Options → Web Clipper and enable it once. The
    setting persists across restarts so this is a one-time step.
  - Joplin first launch is slow; bump JOPLIN_BOOT_TIMEOUT.
  - Port $API_BASE not actually 41184. Set JOPLIN_API env var.
EOF
        exit 1
    fi
    sleep 1
done
