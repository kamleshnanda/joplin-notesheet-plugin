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
# load. If the generator's seed lives in code (the New Spreadsheet
# command handler), this script's body should be the minimal fence
# the plugin recognizes as "no snapshot yet, please scaffold one."
# Currently that's exactly what an empty fenced block is:
#
#     ```notesheet v=1
#     {}
#     ```
#
# When the plugin opens that note for the first time, its Custom
# Editor activates, sees an empty snapshot, and (per the smoke
# spec's expectation after the generator's fix) seeds A1 with
# `harness-smoke-OK` in red.
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
