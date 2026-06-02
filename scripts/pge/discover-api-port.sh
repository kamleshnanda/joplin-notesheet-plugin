#!/usr/bin/env bash
# Discover the Joplin Web Clipper Data API port for the dev profile
# and print it. Other scripts source this to avoid hardcoding 41184
# (which is the MAIN profile's default).
#
# Joplin allocates the clipper port at first start (and tries 41184,
# then climbs by one if it's taken — which it is on this machine
# because the main profile holds 41184). The chosen port is recorded
# in `~/.config/joplindev-desktop/log-clipper.txt`.
#
# Usage:
#   PORT=$(./scripts/pge/discover-api-port.sh)
#   curl http://localhost:$PORT/ping
#
# If $JOPLIN_API is already set in the environment, that wins; this
# script just echoes its host:port and exits.

set -euo pipefail

# Operator override.
if [ -n "${JOPLIN_API:-}" ]; then
    # Strip http:// prefix if present.
    echo "${JOPLIN_API#http://}" | sed 's,^//,,'
    exit 0
fi

PROFILE_DIR="${JOPLIN_PROFILE_DIR:-$HOME/.config/joplindev-desktop}"
LOG="$PROFILE_DIR/log-clipper.txt"

if [ -f "$LOG" ]; then
    # Pull the most recent port. Lines look like:
    #   2026-06-02 14:23:34: Starting Clipper server on port 27583
    PORT=$(grep -E 'Starting Clipper server on port [0-9]+' "$LOG" | tail -1 | awk '{print $NF}')
    if [ -n "$PORT" ]; then
        # Verify the port is actually responding.
        if curl -sf -m 2 "http://localhost:$PORT/ping" >/dev/null 2>&1; then
            echo "localhost:$PORT"
            exit 0
        fi
    fi
fi

# Fallback: probe a small range. Joplin starts at 41184 and increments.
for port in 41184 41185 41186 27583 27584 27585; do
    if curl -sf -m 1 "http://localhost:$port/ping" >/dev/null 2>&1; then
        echo "localhost:$port"
        exit 0
    fi
done

echo "discover-api-port: no Joplin Web Clipper Data API responding" >&2
exit 1
