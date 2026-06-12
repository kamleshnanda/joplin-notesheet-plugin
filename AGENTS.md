# AGENTS.md — operating guide for AI coding agents

This file is for AI coding agents working in this repo (Claude Code,
Cursor, and any other agent that reads `AGENTS.md`).

## First: read CONTRIBUTING.md — all of it applies to you

[`CONTRIBUTING.md`](./CONTRIBUTING.md) is the single source of truth for
**development, testing, code quality, dependency hygiene, high-standards,
and milestone/PR process.** Every rule there binds you exactly as it binds
a human contributor. This file does **not** repeat any of it.

What follows is **only** the operational layer that is specific to how an
agent should behave — things a human contributor wouldn't need.

## Milestone assignment is the maintainer's call — always ask

Per CONTRIBUTING's milestone rule, the _which-milestone_ decision is **not
yours to make**. Before recording any change in the README "Milestones"
table, ask the operator to choose, with exactly these options:

1. **Current milestone** — attach `[#N](url)` to the in-flight row.
2. **New milestone** — a new row with a purpose-stating name (this is how
   codebase-health work earns its own line; cf. M11, M14, M20).
3. **Future documented milestone** — fold into a planned (⏳) row's scope.

Use a structured question (e.g. `AskUserQuestion`). Do **not** self-assign
by merge date, scope heuristic, or "is this major" judgment — those were
explicitly rejected in favor of just asking.

## Stale build is the FIRST suspect for any reported render bug

When the operator reports that something **renders** wrong in Joplin (not a
snapshot-data or export bug), suspect a stale `.jpl` before suspecting the
code. This has cost real time twice (M13 rotated text, M17 percent axis —
both were builds that predated the fix). Before debugging:

1. `ls -la publish/*.jpl` and the installed copy under the Joplin dev
   profile — check the mtime.
2. `git log -1 --format=%ci -S "<symbol>" -- <file>` — when did the fix land?
3. If the `.jpl` predates the fix → rebuild (`npm run dist`), reinstall, and
   have the operator re-test. That is often the whole fix.
4. Verify the symbol is actually in the bundle (`tar -xzf` the `.jpl`, grep
   a non-minified string token — function names get mangled).

**PGE install/launch order is quit → install → launch.** `install-plugin.sh`
uninstalls-then-reinstalls; running it against a _live_ Joplin removes the
plugin from that session (symptom: "UserWebviewIndex frame did not appear").

## Screenshot / image handling — shell first, read sparingly

Reading images through the tool API is rate-limited and brittle when
chained. When the operator pastes a screenshot path (e.g. on `~/Desktop`):

- **Extract structured data from the shell first** — `sips -g pixelWidth
-g pixelHeight`, `unzip -p`, `grep` — before ever opening an image.
- **Read at most ~2 images per session**, one at a time, ideally with
  separate turns between them. Do **not** loop `sips -Z N` resizing an
  image hoping a smaller version slips past the limit — the limit is
  per-request and chaining compounds it (this wasted ~5–6 turns once).

## PGE harness (planner → generator → evaluator)

The long-running autonomous harness has its own runtime contract in
[`.claude/CLAUDE.md`](./.claude/CLAUDE.md): proof-before-passing
(build → install → screenshot → Read-as-evidence → `test-results.json`),
the `verify-gate` evidence hook, `PROGRESS.md` handoff discipline, and the
`scripts/pge/` utilities. If you are running under that harness, follow it.
If you are a general coding agent (not the PGE loop), you still follow
CONTRIBUTING.md + this file, but the PGE-specific ceremony doesn't apply.

## Persistent memory (Claude Code)

Project-specific lessons are kept in Claude's file-based memory. They are
**background context, not a substitute** for CONTRIBUTING.md or this file —
when a memory and a committed guideline disagree, the committed file wins
(it's reviewable and versioned; memory is not). If you learn a durable,
repo-wide rule, propose adding it to CONTRIBUTING.md (or AGENTS.md if it's
agent-only) so humans and other agents get it too — don't let it live only
in one agent's memory.
