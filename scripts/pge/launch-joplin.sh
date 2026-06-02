#!/usr/bin/env bash
# Idempotent: ensure Joplin desktop (DEV profile) is running and the
# Web Clipper Data API is responding on localhost:41184. Generator
# and evaluator both call this; safe to run repeatedly.
#
# Uses the `--env dev` profile (~/.config/joplindev-desktop/) instead
# of the operator's main Joplin profile, so PGE cycles never touch
# real notes. This matches the operator's existing workflow.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TIMEOUT_SECONDS="${JOPLIN_BOOT_TIMEOUT:-30}"
JOPLIN_BIN="${JOPLIN_BIN:-/Applications/Joplin.app/Contents/MacOS/Joplin}"
JOPLIN_ENV_FLAG="${JOPLIN_ENV:---env dev}"

# 1. Already up? Discover the dev profile's Web Clipper port and probe.
if HOSTPORT=$("$SCRIPT_DIR/discover-api-port.sh" 2>/dev/null); then
    echo "joplin: already up at http://$HOSTPORT"
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
    if HOSTPORT=$("$SCRIPT_DIR/discover-api-port.sh" 2>/dev/null); then
        echo "joplin: up at http://$HOSTPORT"
        exit 0
    fi
    if [ "$(date +%s)" -ge "$deadline" ]; then
        cat >&2 <<EOF
joplin: timed out after ${TIMEOUT_SECONDS}s waiting for Web Clipper
Data API.

Possible causes:
  - First launch of dev profile? Web Clipper is OFF by default. Open
    the dev Joplin → menu bar → Joplin → Settings → Web Clipper, and
    click "Enable Web Clipper Service". The setting persists.
  - Joplin first launch is slow; bump JOPLIN_BOOT_TIMEOUT.
  - The dev profile's port is recorded in
    ~/.config/joplindev-desktop/log-clipper.txt. discover-api-port.sh
    reads it from there.
EOF
        exit 1
    fi
    sleep 1
done
