# Notesheet PGE — build plan (M13/C: rotated-text round-trip)

> **Cycle context.** This is the first real-feature cycle through the
> PGE harness, post-smoke. The smoke (`feature-1-smoke-red-cell`) is
> done and lives in `PROGRESS.md` `## Done`. The harness has been
> hardened (Univer canvas selector + pixel sidecar). This file
> overwrites the prior smoke plan; the smoke row is removed from
> `test-results.json`.

## Operator ask

(See `OPERATOR_ASK.md`.) Restore Workstream C from the reverted M13
PR #16 — rotated-text round-trip — and prove it via the harness.
PR #16 added rotation import/export plus six Jest tests, all of
which passed. The PR was reverted because Univer **did not actually
render the text rotated**: snapshot data correct, visible output
broken. That is exactly the failure mode the PGE harness was built
for. This cycle ships the feature and proves the loop catches that
class of bug going forward.

## Features

### feature-1-m13-rotated-text-renders

**Spec**

When the user imports
`tests/ExcelBaseTestData/formatting-testdata/MergedCellsAndAlignment.xlsx`
into a Notesheet note (via the Tools menu **Import .xlsx as
Notesheet** command, or the editor's Import button), the resulting
note opens in the Univer editor with row 6 of the "Merged and
Alignment" sheet displaying three cells whose text is **visibly
rotated** in the rendered canvas:

- **A6** "Rotated 45 degrees" — text leans up-and-to-the-right
  (CCW 45°)
- **B6** "Rotated 90 degrees" — text runs vertically with the
  baseline pointing up (CCW 90°)
- **C6** "Rotated -45 degrees" — text leans down-and-to-the-right
  (CW 45°)

Exporting the same note back to `.xlsx` via the editor's Export
button (or `snapshotToXlsxBuffer`) must produce a workbook where
exceljs reads `cell.alignment.textRotation` as `45`, `90`, and `-45`
for A6/B6/C6 respectively. The snapshot shape is Univer's
`ITextRotation`: `style.tr = { a: <angle> }` for plain rotation and
`style.tr = { a: 0, v: 1 }` for stacked-vertical (the latter is not
exercised by the fixture; covered in Jest only).

**Acceptance criteria**

The evaluator must verify ALL of the following:

1. **Visual — rotation is visible in pixels.** The evaluator's
   Playwright-captured screenshot of the imported note (taken via
   `scripts/pge/eval-screenshot.js` against the running Joplin dev
   profile) shows three rotated text strings in the row-6 region of
   the Univer canvas. Specifically:
   - A6's text glyphs run along an upward-right diagonal (NOT
     horizontal). The leftmost character sits lower than the
     rightmost.
   - B6's text glyphs are stacked along a vertical axis with the
     baseline rotated 90° CCW (each character's bottom faces the
     right edge of the cell).
   - C6's text glyphs run along a downward-right diagonal (the
     leftmost character sits higher than the rightmost).
   - The evaluator must explicitly state in its verdict that all
     three angles are visually distinguishable from horizontal and
     from each other. "Three rotated strings" is not enough —
     the verdict must call out the per-cell direction.
2. **Pixel sidecar — text ink is not on a single horizontal row.**
   The `<screenshot>.pixels.json` sidecar emitted alongside the
   evaluator's screenshot shows the row-6 vertical slab carries
   text-coloured pixels distributed across multiple Y rows of the
   slab (rotated text breaks the dominant-horizontal-ink pattern that
   plain text produces). The evaluator captures the row-6 region and
   confirms text-ink coverage on more than one canvas row within the
   slab. (The harness's existing colour-histogram emit is fine; the
   evaluator picks the row-6 vertical band when grading.)
3. **Jest — reverted PR #16 tests are restored and green.** All of
   the following pass under `npm test`:
   - `tests/m13RotatedText.test.ts` is restored byte-equivalent to
     commit `415b4a4` (5 tests: vertical-stacked import,
     no-rotation absence-of-tr, explicit-zero absence-of-tr, the
     12-angle round-trip sweep at ±15°/30°/45°/60°/75°/90°, and the
     vertical-stacked round-trip).
   - `tests/m12FixtureRoundTrip.test.ts`'s `KNOWN SHORTCOMING —
     rotated-text cells lose their rotation on import (→ M13)`
     test (currently at line 420) is flipped to a positive
     pin-down: A6/B6/C6 of `MergedCellsAndAlignment.xlsx` arrive
     with `style.tr` set to `{ a: 45 }`, `{ a: 90 }`, `{ a: -45 }`
     respectively.
   - The total Jest passing count moves from the current baseline
     (176 at last green) to at least 182, matching PR #16's
     post-feature count, with no `KNOWN SHORTCOMING — rotated-text`
     test left referencing M13.
4. **Round-trip — angles survive export → re-import.** A new (or
   restored) Jest test imports `MergedCellsAndAlignment.xlsx`, calls
   `snapshotToXlsxBuffer` on the resulting snapshot, loads the buffer
   with a fresh exceljs `Workbook`, and asserts that the "Merged and
   Alignment" sheet's A6/B6/C6 carry `alignment.textRotation` of 45,
   90, and -45 respectively. (PR #16 added this assertion to
   `m12FixtureRoundTrip.test.ts` alongside the flipped pin-down;
   restoring it is sufficient.)
5. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** The generator captures its own screenshot via
   `scripts/pge/install-plugin.sh` + `eval-screenshot.js` as evidence
   of due diligence under `screenshots/feature-1-m13-rotated-text-renders/`,
   then opens that screenshot via the Read tool so the
   `verify-gate` hook unlocks the `test-results.json` write. The
   evaluator captures its OWN screenshot independently — the
   generator's drop is corroborating, not authoritative.

**Out of scope**

- **Rotation inside merged cells.** Row 1 of the fixture has merges
  (A1:B2 and C1:D1); row 6 does not. If rotation incidentally renders
  inside a merge, fine; if it doesn't, that's a separate follow-up,
  not a blocker for this feature.
- **Per-row-block rotation in OOXML xfs** (rotation set on a row
  default style rather than per-cell). The fixture doesn't exercise
  it; don't add code paths for it here.
- **Rich text combined with rotation.** OOXML rotation is cell-level
  only, not per-run, so there is nothing to combine. Rich-text-only
  is a separate workstream (M13/D, future cycle).
- **Stacked-vertical mode in the visual gate.** OOXML's `'vertical'`
  / textRotation=255 is covered in `m13RotatedText.test.ts` (Jest
  only). The fixture has no stacked cell, so the evaluator does not
  grade it visually.
- **README's "Known shortcomings — rotated text" entry.** Leave it.
  The operator explicitly punted the README edit to a follow-up
  cycle — this BUILD_PLAN does not include a docs-edit acceptance
  criterion.
- **Stale `tr: { a: 0 }` cleanup pass.** Do not retroactively scrub
  prior snapshots. The export side simply must not write rotation
  attributes when `tr` is absent or zero — covered by the
  no-rotation Jest cases.

**Suggested fixture**

`tests/ExcelBaseTestData/formatting-testdata/MergedCellsAndAlignment.xlsx`
sheet "Merged and Alignment", row 6 — A6 +45°, B6 +90°, C6 -45°.

The harness invokes Import .xlsx as Notesheet on this fixture (via
`scripts/pge/import-fixture.sh MergedCellsAndAlignment.xlsx` or the
equivalent Joplin Data API call, depending on what the harness
already wires up). The eval-screenshot script captures the editor's
Univer main canvas; the evaluator visually inspects row 6.

**Related risks**

- **The reverted PR #16's `src/xlsx.ts` changes are the right
  starting point.** The 20-line addition (`style.tr = { a, v? }` on
  import, reversed on export) was correct under exceljs 4.x and
  matches Univer 0.23's `ITextRotation` shape. The Jest tests
  passing tells you the import/export math is sound — what failed
  was the runtime renderer, not the data plumbing. Don't rewrite
  the angle math.
- **Investigate before patching.** If the screenshot shows
  horizontal text after restoring PR #16 verbatim, the fix is
  downstream of `xlsx.ts`. Likely suspects, in order:
  1. **Style placement.** Univer reads styles from the snapshot's
     top-level `styles[id]` map keyed by the cell's `s` reference.
     If `tr` ends up inline on the cell (not on `styles[id]`),
     Univer's resolver ignores it. Confirm against the actual
     emitted snapshot via DevTools, not against what the test
     assertion claims.
  2. **`tr` being stripped by a later pass.** `src/snapshot.ts`,
     `src/index.ts`'s save path, or any of the wrapper helpers
     might be doing a shallow style merge that drops unknown keys.
     `git grep` for the existing recognised style keys (`bg`, `cl`,
     `bd`, `bl`, `it`, `un`, `n`, `ht`, `vt`) and check whether `tr`
     is in any allowlist.
  3. **Univer plugin gating.** Some Univer presets only honour `tr`
     when a specific plugin is registered. Compare the plugin set
     in `src/editor` against Univer 0.23's docs for the
     `IRenderConfig` / cell-renderer module that consumes `tr`.
     If a missing plugin is the issue, register it; do not fall
     back to drawing rotation manually.
- **Univer style lookup is by `s` reference (M13 lesson, smoke
  confirmed).** The smoke proved that `cl.rgb` only renders when on
  `styles['pge-smoke-red']` and the cell carries `s: 'pge-smoke-red'`.
  `tr` follows the same rule. Any "fix" that puts `tr` directly on a
  cellData entry will Jest-pass and visually-fail — exactly the
  M13 failure mode.
- **Don't symptom-patch the test on a Jest failure post-rebuild.**
  If a Jest assertion regresses, run `git diff package-lock.json`
  first. exceljs's `alignment.textRotation` surface drifted between
  3.x and 4.x (string `'vertical'` vs number 255 in older versions).
  We are intentionally on 4.4.0; do not edit the test to make a
  silent downgrade pass — fix the dependency drift instead.
  Reference: `feedback_dependency_hygiene.md` in operator memory.
- **Pre-existing typecheck bug fixed during smoke** —
  `tests/exportTableRoundTrip.test.ts:334` was changed from
  `'dashed'` to `'mediumDashed'` to satisfy exceljs 4's BorderStyle
  enum (smoke session). Do not revert that; it is unrelated to
  rotation but blocks `npm run dist` if it regresses.
- **The feature touches the same file (`src/xlsx.ts`) as M9/M10/M12.**
  Run the full Jest suite, not just the rotation tests, before flipping
  the row. Borders, hyperlinks, table styles, chart export, and
  alignment all share the import/export pipeline.

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file as the template. Each feature gets:
- `### feature-N-<kebab-id>` heading (stable; never renamed once
  written)
- **Spec**: one paragraph naming the user-observable change
- **Acceptance criteria**: numbered list of observable evidence
  (visual, pixel-sidecar, Jest, runtime). No data-shape-only
  assertions; no "code does X" — only outcomes the evaluator can
  inspect.
- **Out of scope**: explicit non-goals so the evaluator does not
  penalise for them
- **Suggested fixture(s)**: which file(s) under
  `tests/ExcelBaseTestData/formatting-testdata/` exercise this
- **Related risks**: regression hot-spots and prior-bug pointers

The lowest-numbered feature with `passes: false` in
`test-results.json` is what the generator works on next.
