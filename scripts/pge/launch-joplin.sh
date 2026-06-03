#!/usr/bin/env bash
# Idempotent: ensure Joplin desktop (DEV profile) is running, the Web
# Clipper Data API is responding, AND the Chrome DevTools Protocol
# endpoint is exposed so the evaluator can attach via Playwright.
# Generator and evaluator both call this; safe to run repeatedly.
#
# Uses the `--env dev` profile (~/.config/joplindev-desktop/) instead
# of the operator's main Joplin profile, so PGE cycles never touch
# real notes. We additionally pass `--remote-debugging-port=$PGE_CDP_PORT`
# (default 8315) which Joplin's Electron main process forwards to the
# renderer; Playwright's `chromium.connectOverCDP(...)` attaches to it
# from `eval-screenshot.js`. (`_electron.launch` does NOT work — it
# injects `--inspect=0` which Joplin rejects with a fatal modal.)
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TIMEOUT_SECONDS="${JOPLIN_BOOT_TIMEOUT:-30}"
JOPLIN_BIN="${JOPLIN_BIN:-/Applications/Joplin.app/Contents/MacOS/Joplin}"
JOPLIN_ENV_FLAG="${JOPLIN_ENV:---env dev}"
CDP_PORT="${PGE_CDP_PORT:-8315}"

api_up() {
    "$SCRIPT_DIR/discover-api-port.sh" 2>/dev/null
}

cdp_up() {
    curl -sf -m 2 "http://localhost:$CDP_PORT/json/version" >/dev/null 2>&1
}

# 1. Already up? Both Web Clipper API AND CDP must respond.
if HOSTPORT=$(api_up); then
    if cdp_up; then
        echo "joplin: already up at http://$HOSTPORT (CDP on :$CDP_PORT)"
        exit 0
    else
        cat >&2 <<EOF
joplin: app is running and Web Clipper API is up, but the CDP
endpoint on localhost:$CDP_PORT is not responding.

This usually means Joplin was launched without
\`--remote-debugging-port=$CDP_PORT\`. Quit Joplin (Cmd+Q) and re-run
this script — it will relaunch with the flag set.

If you need a different CDP port, set PGE_CDP_PORT before invoking.
EOF
        exit 1
    fi
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

# 3. Launch in dev-profile mode with CDP port exposed.
echo "joplin: launching '$JOPLIN_BIN $JOPLIN_ENV_FLAG --remote-debugging-port=$CDP_PORT'..."
nohup "$JOPLIN_BIN" $JOPLIN_ENV_FLAG --remote-debugging-port="$CDP_PORT" >/dev/null 2>&1 &
disown

# 4. Poll until BOTH the Data API and the CDP endpoint respond, or
# until timeout. Either alone is not enough — the evaluator needs CDP
# to attach for screenshots, and the harness needs the Data API to
# create/look up test notes.
deadline=$(( $(date +%s) + TIMEOUT_SECONDS ))
while true; do
    HOSTPORT=$(api_up || true)
    if [ -n "$HOSTPORT" ] && cdp_up; then
        echo "joplin: up at http://$HOSTPORT (CDP on :$CDP_PORT)"
        exit 0
    fi
    if [ "$(date +%s)" -ge "$deadline" ]; then
        cat >&2 <<EOF
joplin: timed out after ${TIMEOUT_SECONDS}s waiting for Joplin to
become ready.

  Web Clipper API: ${HOSTPORT:-NOT RESPONDING}
  CDP on :$CDP_PORT: $(cdp_up && echo OK || echo "NOT RESPONDING")

Possible causes:
  - First launch of dev profile? Web Clipper is OFF by default. Open
    the dev Joplin → menu bar → Joplin → Settings → Web Clipper, and
    click "Enable Web Clipper Service". The setting persists.
  - Joplin first launch is slow; bump JOPLIN_BOOT_TIMEOUT.
  - The dev profile's port is recorded in
    ~/.config/joplindev-desktop/log-clipper.txt. discover-api-port.sh
    reads it from there.
  - CDP port collision? Another process may hold :$CDP_PORT. Set
    PGE_CDP_PORT to a free port and try again.
EOF
        exit 1
    fi
    sleep 1
done
