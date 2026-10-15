#!/usr/bin/env bash
# Copyright 2026 Anthropic PBC
# SPDX-License-Identifier: Apache-2.0
# Run from the project root so relative paths (AGENT_STOP, STEER.md,
# .claude/.evidence-reads, test-results.json) resolve correctly even after
# the session has cd'd into a subdirectory. CLAUDE_PROJECT_DIR is set by
# Claude Code for hooks; fall back to this script's own location.
cd "${CLAUDE_PROJECT_DIR:-$(dirname "$0")/../..}" 2>/dev/null || exit 0
# Halt every tool call while ./AGENT_STOP exists. `touch AGENT_STOP` to engage; `rm AGENT_STOP` to resume.
if [ -e "${AGENT_STOP_FILE:-./AGENT_STOP}" ]; then
  cat <<'JSON'
{"decision":"block","reason":"Kill switch engaged: AGENT_STOP file exists. Agent is halted. Remove the file to resume."}
JSON
fi
