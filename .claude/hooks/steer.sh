#!/usr/bin/env bash
# Copyright 2026 Anthropic PBC
# SPDX-License-Identifier: Apache-2.0
# Run from the project root so relative paths (AGENT_STOP, STEER.md,
# .claude/.evidence-reads, test-results.json) resolve correctly even after
# the session has cd'd into a subdirectory. CLAUDE_PROJECT_DIR is set by
# Claude Code for hooks; fall back to this script's own location.
cd "${CLAUDE_PROJECT_DIR:-$(dirname "$0")/../..}" 2>/dev/null || exit 0
# If STEER.md has content, surface it to the agent once and clear the file.
# Write to STEER.md (or pipe from a UI) to redirect the agent mid-run.
# Note: this is a convenience channel, not a trust boundary; if the agent has
# Write access to the project it can write STEER.md itself.
f="${AGENT_STEER_FILE:-./STEER.md}"
if [ -s "$f" ]; then
  note=$(cat "$f")
  reason=$(python3 -c 'import json,sys; print(json.dumps("OPERATOR STEERING: " + sys.argv[1] + "\n\nPause what you were about to do, incorporate this guidance, then continue toward the feature goal."))' "$note" 2>/dev/null) || exit 0
  printf '{"decision":"block","reason":%s}\n' "$reason"
  : > "$f"
fi
