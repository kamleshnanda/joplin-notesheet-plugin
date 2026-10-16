#!/usr/bin/env bash
# Neutral deploy step for the PGE harness: build a given commit and install it
# into the DEV-profile Joplin, without touching the working tree.
#
# WHY this exists: the evaluator must test the shipped behaviour of a commit
# on its own, without waiting on (or trusting) the generator's live run. It
# is not allowed to build in the working tree (the generator may be mid-edit
# there), so this script builds the requested ref in a throwaway git worktree,
# installs that exact .jpl, relaunches the dev Joplin and PROVES what is
# installed (byte-identical .jpl + the commit SHA stamped into the bundle).
#
# Usage:
#   scripts/pge/deploy-build.sh [<git-ref>]      (default: HEAD)
# Writes a JSON summary to stdout and to /tmp/pge-deploy-last.json:
#   { ref, sha, jpl, sha256, version, installedMatches, bundleHasSha }
# Exit codes: 0 deployed and verified, 1 failed, 3 refused (regular-profile
# Joplin running — never touched).
#
# Safety: only ever stops a Joplin started with `--env dev` (the harness
# profile). A regular-profile Joplin is the operator's real notes: if one is
# running, this refuses and exits 3.
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
REF="${1:-HEAD}"
PLUGIN_ID="com.kamleshnanda.joplin-notesheet"
PROFILE_DIR="${JOPLIN_PROFILE_DIR:-$HOME/.config/joplindev-desktop}"
JOPLIN_BIN_RE='Joplin.app/Contents/MacOS/Joplin'

say() { echo "deploy-build: $*" >&2; }

SHA="$(git -C "$REPO_ROOT" rev-parse --verify "$REF^{commit}")"
SHORT="$(git -C "$REPO_ROOT" rev-parse --short "$SHA")"
say "ref $REF → $SHORT"

# 1. Refuse if a regular-profile Joplin is running (any Joplin main process
#    whose args lack --env dev).
if ps -axo args | grep "$JOPLIN_BIN_RE" | grep -v grep | grep -v -- '--type=' | grep -qv -- '--env dev'; then
    say "a regular-profile Joplin is running — refusing to touch it. Quit it and retry."
    exit 3
fi

# 2. Build the ref in a throwaway worktree.
WT="$(mktemp -d "/tmp/pge-deploy-$SHORT-XXXX")"
cleanup() { git -C "$REPO_ROOT" worktree remove --force "$WT" >/dev/null 2>&1 || rm -rf "$WT"; }
trap cleanup EXIT
git -C "$REPO_ROOT" worktree add --detach "$WT" "$SHA" >/dev/null 2>&1
# Reuse the main checkout's node_modules when the lockfile is identical;
# otherwise install exactly what that commit's lockfile pins.
if cmp -s "$REPO_ROOT/package-lock.json" "$WT/package-lock.json" && [ -d "$REPO_ROOT/node_modules" ]; then
    # .gitignore's `node_modules/` doesn't match a SYMLINK, which would make
    # the build stamp `-dirty`. A bare pattern in the repo-local exclude file
    # (shared by all worktrees, never committed) covers it.
    EXCLUDE="$(git -C "$REPO_ROOT" rev-parse --path-format=absolute --git-common-dir)/info/exclude"
    mkdir -p "$(dirname "$EXCLUDE")"
    grep -qx 'node_modules' "$EXCLUDE" 2>/dev/null || echo 'node_modules' >>"$EXCLUDE"
    ln -s "$REPO_ROOT/node_modules" "$WT/node_modules"
    say "node_modules: reused (lockfile identical)"
else
    say "node_modules: lockfile differs at $SHORT — running npm ci in the worktree"
    npm --prefix "$WT" ci --silent >&2
fi
say "building $SHORT ..."
npm --prefix "$WT" run dist --silent >&2
BUILT="$WT/publish/$PLUGIN_ID.jpl"
[ -f "$BUILT" ] || { say "build produced no .jpl"; exit 1; }
OUT_JPL="/tmp/pge-deploy-$SHORT.jpl"
cp "$BUILT" "$OUT_JPL"

# 3. Stop ONLY the dev-profile Joplin, then install + relaunch.
DEV_PIDS="$(ps -axo pid,args | grep "$JOPLIN_BIN_RE --env dev" | grep -v grep | awk '{print $1}' || true)"
if [ -n "$DEV_PIDS" ]; then
    say "stopping dev-profile Joplin ($DEV_PIDS)"
    # shellcheck disable=SC2086
    kill -TERM $DEV_PIDS
    for _ in $(seq 1 30); do
        pgrep -f "$JOPLIN_BIN_RE" >/dev/null 2>&1 || break
        sleep 1
    done
    if pgrep -f "$JOPLIN_BIN_RE" >/dev/null 2>&1; then
        say "Joplin did not exit within 30s"
        exit 1
    fi
fi
PGE_JPL="$OUT_JPL" SKIP_BUILD=1 "$REPO_ROOT/scripts/pge/install-plugin.sh" >&2
"$REPO_ROOT/scripts/pge/launch-joplin.sh" >&2

# 4. Prove what is installed.
INSTALLED="$PROFILE_DIR/plugins/$PLUGIN_ID.jpl"
MATCH=false
cmp -s "$INSTALLED" "$OUT_JPL" && MATCH=true
HAS_SHA=false
tar -xzOf "$INSTALLED" index.js 2>/dev/null | grep -q "$SHORT" && HAS_SHA=true
VERSION="$(tar -xzOf "$INSTALLED" manifest.json | python3 -c 'import json,sys; print(json.load(sys.stdin)["version"])')"
SUM="$(shasum -a 256 "$OUT_JPL" | cut -d' ' -f1)"
python3 - "$REF" "$SHA" "$OUT_JPL" "$SUM" "$VERSION" "$MATCH" "$HAS_SHA" <<'PY' | tee /tmp/pge-deploy-last.json
import json, sys
ref, sha, jpl, s256, ver, match, has = sys.argv[1:]
print(json.dumps({"ref": ref, "sha": sha, "jpl": jpl, "sha256": s256, "version": ver,
                  "installedMatches": match == "true", "bundleHasSha": has == "true"}, indent=2))
PY
if [ "$MATCH" != true ] || [ "$HAS_SHA" != true ]; then
    say "verification FAILED (installedMatches=$MATCH bundleHasSha=$HAS_SHA)"
    exit 1
fi
say "deployed and verified $SHORT"
