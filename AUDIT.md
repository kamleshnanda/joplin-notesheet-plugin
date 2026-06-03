# PGE harness audit log

Append-only log of decisions and actions during PGE-driven cycles.
Each entry: ISO-8601 UTC timestamp, actor (operator | claude | subagent
name | generator | evaluator), short heading, and one paragraph or a
bullet list. Add new entries at the bottom; don't rewrite history.

When in doubt, log it. The reason this file exists is that PGE
sessions span multiple agents and processes; without a trace, we
can't reconstruct what happened or why.

---

## 2026-06-03T...Z — claude — resume PGE harness, fan out

Resuming from `m13-redo/pge-harness` at commit `950826a`. Goal:
end-to-end PASS on the smoke (`feature-1-smoke-red-cell`).

Plan:

1. **Subagent A — harness rewrite.** Replace `eval-screenshot.sh`
   (Bash + macOS `screencapture`) with Playwright + CDP-attach.
   Operator confirmed Joplin accepts
   `--remote-debugging-port=8315`; that opens the renderer's CDP so
   Playwright can attach (instead of `_electron.launch` which
   passes `--inspect=0` and trips Joplin's "Unknown flag" modal).
   Update `launch-joplin.sh` to pass the flag. Verify by capturing
   a screenshot of an existing dev-profile note.

2. **Subagent B — Jest tests for the smoke seed.** The smoke feature
   asks: when a new Notesheet is opened from an empty fence (`{}`
   body), the editor must seed cell A1 with text `harness-smoke-OK`
   in red (#FF0000). Subagent writes Jest tests against the
   snapshot scaffold helper, NOT against the rendered UI. The
   visual gate is the evaluator's runtime screenshot; Jest tests
   are a fast pre-check.

3. **Integrate** results from A + B. Commit each cleanly.

4. **Run the PGE cycle.** `scripts/pge/run-cycle.sh` invokes
   generator (implements the seed in `src/index.ts` /
   `src/snapshot.ts` to make Jest tests pass and visual evaluator
   PASS) then evaluator (uses the new driver to capture its own
   screenshot, grades).

5. **Commit + push + PR.**

Hard rules carried over from yesterday:

- Joplin dev profile only (`--env dev`). NEVER touch main profile.
- Never invoke Joplin with `--help` or other unknown flags.
- Joplin must be quit before `install-plugin.sh` (cache wipe).
- Token at `.claude/joplin-token.local`, gitignored.

Audit entries below this one are appended chronologically.

---

## 2026-06-03T01:40Z — subagent A (claude) — eval-screenshot CDP-attach driver

Replaced the macOS `screencapture` workaround with Playwright's
`chromium.connectOverCDP('http://localhost:8315')`. `launch-joplin.sh`
now passes `--remote-debugging-port=$PGE_CDP_PORT` (default 8315) and
gates "ready" on BOTH the Web Clipper API responding AND the CDP
endpoint serving `/json/version` (otherwise the evaluator would
attach-fail later with no clear diagnosis). `eval-screenshot.sh` is
now a thin Bash wrapper that calls launch-joplin then `node
eval-screenshot.js`. The JS rewrites use CDP attach instead of
`_electron.launch` (Joplin rejects `_electron.launch`'s implicit
`--inspect=0` with a fatal modal). On verification, Joplin's CDP
exposed 4 pages: a "DevTools" page (Joplin opens devtools in dev mode),
two plugin sandboxes (`plugin_index.html?pluginId=...js-draw` and
`...backup`), and the editor renderer at `file://...index.html` with
title "Joplin". Picked the editor via a small scoring heuristic
(index.html in URL +10, title contains "Joplin" +5) which wins
unambiguously (score 16 vs 11/0/0). For the "Univer rendered"
detection I tried `.univer-render-canvas`, `canvas.univer-render-canvas`,
`#joplin-plugin-content`, and `.univer-container` (5s each) and on
miss fall back to a 5s `waitForTimeout` — this is a documented gap;
the dev profile currently has no Notesheet plugin installed so I
couldn't probe the live selector, future work is to inspect the
rendered DOM once the plugin is installed and pick a stable hook.
Verification: opened the existing "PGE smoke note test1" note via
`open joplin://x-callback-url/openNote?id=...`, captured an 89 KB PNG
of the real Joplin renderer (notebook tree + note list + editor pane
showing the empty `{}` body), saved at
`/tmp/pge-verify-eval-screenshot.png` for operator inspection. The
canonical `screenshots/<feature-id>/` dir was not polluted (the
verification PNG was deleted after copying to /tmp). joplin:// is
dispatched via subprocess `open` (not `page.evaluate`) because the
URL handler lives in Joplin's main process, not the renderer.

---
