# Operator ask — cycle 1 (PGE harness smoke)

Smoke-test the planner → generator → evaluator harness end-to-end on a
**knowingly-broken trivial feature**. The intent is to verify every
moving part of the loop:

1. Planner reads this ask, writes BUILD_PLAN.md + test-results.json
   (one row, `passes: false`).
2. Generator picks up the row, must actually implement code +
   build + install + screenshot to flip the row.
3. Evaluator (separate fresh-context process) opens its OWN
   screenshot via Playwright MCP attached to running Joplin and
   judges against the spec.
4. Default-FAIL evidence gate (verify-gate hook) blocks
   test-results.json writes until the generator has Read evidence.
5. Loop wraps cleanly when evaluator returns PASS.

## The smoke feature

**Brand-new Notesheet note** — created via the existing "New
Spreadsheet" command — should open showing cell A1 filled with the
literal text `harness-smoke-OK` rendered in **red** (#FF0000).

This is intentionally not currently implemented. The generator has
to:
- Edit `src/index.ts` (the New Spreadsheet command handler) to seed
  cell A1 with `harness-smoke-OK` and a red font color in the initial
  snapshot.
- Build `.jpl`, install in Joplin, create a fresh Notesheet, screenshot.
- Confirm A1 actually renders red text in the screenshot.

If at any step the harness wiring breaks (Playwright can't attach,
screenshot doesn't save, evaluator can't read it, default-FAIL gate
mis-fires), that's the smoke catching a real harness bug — fix the
harness, re-run.

## Out of scope

- Permanence: the seed only needs to land in newly-created notes from
  this session forward. Don't migrate existing notes.
- Localization or i18n.
- Any of the M13 work tracked in tasks #28-#32. Those come AFTER the
  smoke passes.

## Success criteria

- `test-results.json` shows the smoke row at `{"passes": true,
  "evaluator_verdict": "PASS"}`.
- A screenshot under `screenshots/feature-1-smoke-red-cell/` shows
  cell A1 with red `harness-smoke-OK` text in a freshly-created note.
- The screenshot the evaluator graded against was captured by the
  EVALUATOR via Playwright MCP, not the generator.
- `git log` shows commits with descriptive messages.
- No bypass of the default-FAIL gate.
