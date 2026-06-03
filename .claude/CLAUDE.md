# Notesheet — long-running PGE conventions

This file governs every agent session that runs under the planner →
generator → evaluator (PGE) harness. It is the generator's runtime
contract. The planner has a separate brief (`.claude/agents/planner.md`);
the evaluator has its own (`.claude/agents/evaluator.md`).

**Don't paraphrase these rules; follow them.**

## Always start here

Before doing anything else:

1. Read `PROGRESS.md`. It is your handoff note from the previous
   session. If it doesn't exist, create it with four sections
   (`## Done`, `## In progress`, `## Next`, `## Notes`) and leave them
   empty.
2. Read `NEXT_FINDINGS.md` if present — it is the evaluator's specific
   feedback from the last cycle. Address those points before moving
   on. Delete the file once you've addressed every bullet.
3. Run `git log --oneline -10` to see what was just committed.
4. Run `npm test` once so you know the working tree is sane. If tests
   are failing in a way unrelated to your task, fix that first or stop
   and write a STEER message to the operator.

## One feature at a time

Work on exactly one feature ID from `BUILD_PLAN.md` per session. Pick
the lowest-numbered feature whose row in `test-results.json` is
`{"passes": false}`. Finish it (build green, evaluator-approved
evidence captured) before starting another. If the operator gives you
a new ask mid-session via `OPERATOR STEERING:`, add it to
`BUILD_PLAN.md`, finish the current feature first, then iterate.

## Proof before passing

A feature is only "passing" after you have, in this order:

1. **Built the .jpl** (`npm run dist`) and confirmed the build succeeds.
2. **Ensured Joplin desktop is running** — see "Launching Joplin" below.
3. **Reinstalled the .jpl plugin** in the running Joplin. (The generator
   does this; do not skip.)
4. **Opened the test fixture** as a Notesheet note in the running Joplin.
5. **Captured a screenshot** of the rendered cell/region under
   `screenshots/<feature-id>/<descriptive-name>.png`. Use Playwright
   MCP if available; otherwise the system screenshot CLI is acceptable.
6. **Opened the screenshot file via the Read tool** so the harness logs
   it as evidence. The `verify-gate` hook will deny your Write to
   `test-results.json` until evidence has been Read.
7. **Eyeballed the screenshot yourself** before flipping the row to
   `{"passes": true}`. The evaluator will reject "I think it works"
   without proof.

If you cannot get steps 2–6 to work in this session, leave the row at
`{"passes": false}`, document the blocker in `PROGRESS.md`, and stop.
Do not flip the bit speculatively.

## Launching Joplin

The harness scripts under `scripts/pge/` handle this:

- `scripts/pge/launch-joplin.sh` — opens Joplin via `open -a Joplin`
  if not already running, waits for the main window to be ready,
  returns once it's responsive. **Idempotent.** Generator and evaluator
  both call it.
- `scripts/pge/install-plugin.sh` — copies `publish/com.kamleshnanda.joplin-notesheet.jpl`
  into Joplin's plugin install dir and triggers a plugin reload via
  Joplin's data API. Run after every `npm run dist` rebuild.
- `scripts/pge/import-fixture.sh <fixture-name>` — creates a new
  Joplin note from the named fixture under
  `tests/ExcelBaseTestData/formatting-testdata/<fixture-name>` using
  the Notesheet plugin's import path, returns the note ID.

If any of these don't exist yet, write them under `scripts/pge/` and
commit them as part of your feature work. They are part of the harness,
not the feature itself, but the harness must be able to run unattended
before any feature can pass.

## Keep `PROGRESS.md` current

After each completed feature (or when you stop mid-feature), update
`PROGRESS.md`:

- Move the feature line from `## In progress` to `## Done` with a
  one-line summary.
- Add anything you learned to `## Notes` — especially Univer/Joplin
  internals, exceljs quirks, or test patterns that future sessions
  will benefit from.
- If the next feature has prerequisites (e.g. "extend planner to add
  N more fixtures"), put them in `## Next`.

Future sessions read this file cold. Be specific about file paths and
function names. Don't write rumor; write what you verified.

## Commit often

The `Stop` hook commits tracked changes at session end, but you should
also `git add` new source files explicitly and `git commit` at
meaningful checkpoints with descriptive messages. The `commit-on-stop`
hook is a backstop, not a primary commit strategy.

## If you're told to stop

`OPERATOR STEERING:` messages come from a human via the `steer.sh`
hook (operator wrote `STEER.md` and the hook surfaced it). Treat them
as higher priority than your current plan. Pause, incorporate the
guidance, then continue toward the feature goal.

## What never goes into `test-results.json`

- A row flipped to `passes: true` based on Jest tests alone. Jest
  tests check the snapshot data shape; they cannot catch the M13
  failure mode (snapshot data correct, Univer rendering broken).
- A row flipped to `passes: true` based on a screenshot the evaluator
  has not yet validated. The evaluator captures its OWN screenshots
  via Playwright MCP and judges from those. The generator's
  screenshots are evidence of due diligence, not proof of correctness.
- A row flipped to `passes: true` because "the previous session
  thought it worked." Every PGE cycle re-runs the evaluator from a
  fresh context.

## What `test-results.json` looks like

Each feature in `BUILD_PLAN.md` has exactly one row, keyed by feature
ID:

```json
{
  "feature-1-rotated-text-renders": {
    "passes": false,
    "evidence": null,
    "evaluator_verdict": null,
    "last_attempted_at": null
  }
}
```

The default is `passes: false` for every row. Generators flip
individual rows to `passes: true` only after the evaluator has graded
PASS. The wrapper script (`scripts/pge/run-cycle.sh`) reverts a row
to `false` if the evaluator returns `NEEDS_WORK` so subsequent
sessions don't trust a stale claim.
