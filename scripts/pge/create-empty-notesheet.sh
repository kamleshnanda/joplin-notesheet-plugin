#!/usr/bin/env bash
# Create a fresh Notesheet note via Joplin's Data API. Does NOT
# trigger the plugin's "New Spreadsheet" command path — that's a UI
# flow we'd need Playwright to drive. Instead, this writes a note
# whose body is the smallest valid notesheet fence; the plugin's
# Custom Editor activates when the note is opened in Joplin.
#
# Used for the harness smoke (feature-1-smoke-red-cell) where the
# acceptance criterion is "the seed shows up in cell A1 of a freshly
# opened Notesheet note." The generator's job is to make the
# initial-snapshot helper produce that seed; this script is the
# evaluator's tool to OBSERVE the seed.
#
# Output: prints the new note's id on stdout.
#
# IMPORTANT: this script writes the BODY that the editor sees on
# load. The minimal fence the plugin recognizes is:
#
#     ```notesheet v=1
#     {}
#     ```
#
# The plugin opens that note in its Custom Editor and Univer
# bootstraps with no cells. (Pre-PR-#22 the production
# `emptySnapshot()` carried an A1 smoke seed — that seed has since
# been removed; the harness smoke fixture is now built independently
# in `create-seeded-notesheet.js`, not by relying on production
# code shipping the seed.)
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
TITLE="${1:-PGE smoke note $(date +%H-%M-%S)}"

bodyFile=$(mktemp -t pge-smoke-body.XXXX.md)
trap 'rm -f "$bodyFile"' EXIT

cat > "$bodyFile" <<'EOF'
```notesheet v=1
{}
```
EOF

node "$REPO_ROOT/scripts/pge/joplin-api.js" create-note \
    --title "$TITLE" \
    --body-file "$bodyFile"
