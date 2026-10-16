#!/usr/bin/env bash
# Copyright 2026 Anthropic PBC
# SPDX-License-Identifier: Apache-2.0
# Run from the project root so relative paths (AGENT_STOP, STEER.md,
# .claude/.evidence-reads, test-results.json) resolve correctly even after
# the session has cd'd into a subdirectory. CLAUDE_PROJECT_DIR is set by
# Claude Code for hooks; fall back to this script's own location.
cd "${CLAUDE_PROJECT_DIR:-$(dirname "$0")/../..}" 2>/dev/null || exit 0
#
# Commit tracked changes at the end of every session so work is durable across
# restarts. Uses `commit -am` (tracked files only) on purpose: ephemeral
# artifacts (screenshots, logs, scratch files) shouldn't land in history. The
# agent is expected to `git add` new source files itself per CLAUDE.md.
#
# Fails silently if commit can't be made (no git user.name, hook rejects, etc);
# check `git log` periodically when relying on this as a backstop.
if git rev-parse --git-dir >/dev/null 2>&1; then
  if ! git diff --quiet || ! git diff --cached --quiet; then
    git commit -am "session checkpoint: $(date '+%Y-%m-%d %H:%M')" >/dev/null 2>&1
  fi
fi
exit 0
