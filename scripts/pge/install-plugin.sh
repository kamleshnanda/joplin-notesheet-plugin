#!/usr/bin/env bash
# Build (if requested), then install the Notesheet .jpl into the
# Joplin DEV profile (~/.config/joplindev-desktop/plugins/), clearing
# the plugin's compiled cache so the new build is picked up. Joplin
# must be quit before this runs — plugins don't hot-reload, and a
# running Joplin holds locks that can corrupt the cache wipe.
#
# Workflow this matches (operator's existing process):
#   1. Quit Joplin (Cmd+Q).
#   2. ./scripts/pge/install-plugin.sh
#   3. ./scripts/pge/launch-joplin.sh    (re-launches with --env dev)
#
# The wrapper run-cycle.sh handles steps 1+2+3 in order.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
JPL="$REPO_ROOT/publish/com.kamleshnanda.joplin-notesheet.jpl"
PLUGIN_ID="com.kamleshnanda.joplin-notesheet"
PROFILE_DIR="${JOPLIN_PROFILE_DIR:-$HOME/.config/joplindev-desktop}"
PLUGIN_DIR="$PROFILE_DIR/plugins"
CACHE_DIR="$PROFILE_DIR/cache/$PLUGIN_ID"
SKIP_BUILD="${SKIP_BUILD:-0}"

# 1. Refuse if Joplin is currently running. The cache wipe only
# works cleanly with the app quit.
if pgrep -f 'Joplin.app/Contents/MacOS/Joplin' >/dev/null 2>&1; then
    cat >&2 <<EOF
install-plugin: Joplin is currently running. Quit it (Cmd+Q) first,
then re-run this script. The plugin cache must be wiped without the
app holding locks on it.
EOF
    exit 1
fi

# 2. Build.
if [ "$SKIP_BUILD" != "1" ]; then
    echo "install-plugin: building .jpl (set SKIP_BUILD=1 to skip)..."
    (cd "$REPO_ROOT" && npm run dist --silent)
fi
if [ ! -f "$JPL" ]; then
    echo "install-plugin: build did not produce $JPL" >&2
    exit 1
fi

# 3. Wipe the compiled-plugin cache so the new build is freshly
# unpacked instead of reusing stale artifacts.
if [ -d "$CACHE_DIR" ]; then
    echo "install-plugin: clearing plugin cache at $CACHE_DIR"
    rm -rf "$CACHE_DIR"
fi

# 4. Install.
mkdir -p "$PLUGIN_DIR"
cp "$JPL" "$PLUGIN_DIR/$PLUGIN_ID.jpl"
echo "install-plugin: copied to $PLUGIN_DIR/$PLUGIN_ID.jpl"

cat <<EOF
install-plugin: done. Now run:
    ./scripts/pge/launch-joplin.sh
EOF
