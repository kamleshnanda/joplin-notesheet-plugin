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

## 2026-06-02T15:00Z — claude — integrate subagents A + B

Both subagents returned cleanly:

- A (`512d495`): CDP-attach driver. Verified end-to-end with the
  existing dev-profile note "PGE smoke note test1": CDP exposes 4
  pages (devtools + 2 plugin sandboxes + editor); scoring picks the
  editor unambiguously (16 vs 11/0/0). Captured /tmp/pge-verify2.png
  (89 KB), shows real Joplin pixels with the empty `{}` body. The
  Notesheet plugin is NOT yet installed in the dev profile, so
  `.univer-render-canvas` doesn't exist yet — that's the generator's
  first job (build .jpl, install-plugin.sh, then create the seed).
- B (`b576440`): Jest pin-down at `emptySnapshot()` in
  `src/snapshot.ts:65-85`. 5 tests; 3 currently RED (A1 value, A1
  style→#FF0000, wrap→extract round-trip) — exactly the shape the
  generator needs to flip green. Caveat from B: jest cannot prove
  visual rendering; the visual gate (evaluator screenshot) remains
  load-bearing. Documented this in test header.

Operator hint received this turn: a `connectOverCDP` sample using
`page.waitForLoadState('networkidle')`. Folded that into
`eval-screenshot.js` (line ~213) bounded at 10s, with a stderr
warning on timeout — protects against attaching mid-navigation
during Joplin's startup-async-chunk loading.

Open issue surfaced by verification: Notesheet plugin not in the dev
profile yet. Generator's PGE/CLAUDE.md proof-before-passing sequence
explicitly requires `npm run dist && bash install-plugin.sh` before
opening the fixture, so this resolves itself once the generator
runs. Logging here so the operator (or a future debugger) doesn't
mistake "no Univer canvas" for a driver bug.

Next: invoke `scripts/pge/run-cycle.sh` to drive the smoke generator
+ evaluator end-to-end.

---
