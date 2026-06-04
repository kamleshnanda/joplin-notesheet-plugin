# Operator ask — M14: SheetJS Community migration spike (Phase 1 only)

This is the **spike** PR. We do NOT ship the migration in this PR. The
end of the spike is a documented decision: GO / NO-GO / CONDITIONAL.

## Why this matters

`exceljs` has gone quiet (last release Dec 2024) and ships stale
transitives (uuid@8 with a moderate CVE we tolerate as not-reachable;
glob@7 deprecated). It's the entire foundation of `src/xlsx.ts` —
1959 lines of import/export logic for `.xlsx` files. The plan is to
migrate to SheetJS Community (or `xlsx-js-style`, the community fork
with basic cell styling on top of SheetJS Community). The spike
proves whether that migration is feasible without losing the
formatting fidelity we shipped in M9 / M10 / M12 / M13/A through E.

## The constraint

**ZERO feature regressions on shipped behaviour.** Phase 2 (the actual
migration, separate PR) must produce a snapshot that's byte-identical
or structurally equivalent to what `exceljs` produces today, for every
fixture in `tests/ExcelBaseTestData/`. The spike's job is to prove
either:

(a) That parity is achievable with reasonable effort, with a clear
    plan for any divergences, OR
(b) That SheetJS lacks a capability we depend on — at which point we
    DON'T migrate, document the gap, and revisit if SheetJS adds it
    upstream.

The operator's directive: *"I DO NOT WANT FEATURE REGRESSION ON
FEATURES THAT ARE ALREADY WORKING."* That is the bar.

## What ships in this PR

1. **Add `xlsx-js-style` (NOT `xlsx`) as a devDep.** Plain SheetJS
   Community (`xlsx`) does NOT support cell styles. The fork
   `xlsx-js-style@1.2.0` adds basic cell styling on top. Confirm at
   spike time whether it's enough for our needs (theme palette
   round-trip, named-style banding, hyperlinks Pattern A + Pattern B,
   rotated text, rich text within a cell). Document gaps explicitly.

2. **Parallel module `src/xlsxSheetJS.ts`** implementing the same
   public surface as `src/xlsx.ts`:
   - `xlsxBufferToSnapshot(buffer): Promise<UniverSnapshot>`
   - `snapshotToXlsxBuffer(snapshot): Promise<ArrayBuffer>`
   - `class NotesheetImportError extends Error` (same `code` /
     `cause` shape)
   - The two resource-name string constants

   Coverage goal for the spike: try to handle every code path
   `src/xlsx.ts` handles. If something can't be done with SheetJS
   (e.g. theme palette extraction via `xl/theme/theme1.xml` raw read),
   document it and either implement a workaround or note "this path
   is the migration's blocker".

3. **Comparison test bed `tests/xlsxParserParity.test.ts`.** For
   every fixture in `tests/ExcelBaseTestData/formatting-testdata/`
   AND `tests/ExcelBaseTestData/chart-testdata/`:
   - Run both parsers on the same input
   - Deep-equal their resulting snapshots' `cellData`, `styles`,
     `defaultStyle`, `sheetOrder`, `sheets[*].rowCount` and
     `columnCount`, `resources` (filter to known-stable fields)
   - For round-trip: import → export → re-import; assert idempotency
     within each parser AND cross-parser
   - When divergence is expected (e.g. `xlsx-js-style` doesn't
     surface theme-tinted borders), `expect.toEqual(divergence)`
     against a documented-known-divergence list rather than fail.

   This test must produce a clear matrix of "what works / what
   diverges / what blocks the migration" — the matrix becomes input
   to the decision document.

4. **Decision document `docs/m14-sheetjs-spike.md`.** Markdown only,
   committed. Sections:
   - **Capability matrix.** One row per .xlsx feature dimension
     Notesheet uses today (workbook value read/write, cell styles —
     bg/fg/font, borders, merged cells, named tables + table styles,
     hyperlinks Pattern A, hyperlinks Pattern B, rotated text, rich
     text per-run, theme palette `<a:clrScheme>`, conditional
     formatting rules, custom `<dxf>` records, drawings/charts,
     formulas + structured refs, defined names). Columns: exceljs
     (current), SheetJS (xlsx-js-style), gap notes.
   - **Migration cost estimate.** For each gap, a 1-3-day estimate of
     either "SheetJS API supports it, just port the code" or "needs a
     workaround (raw XML read like our `readThemeClrScheme`)" or
     "blocked, no path".
   - **Recommendation.** GO / NO-GO / CONDITIONAL with the
     conditions named explicitly.

5. **Regression test scaffolding for Phase 2.** Even though Phase 1
   doesn't ship the migration, leave behind the test bed Phase 2 will
   need to gate its merge:
   - **Golden-snapshot tests.** For each fixture, `xlsxBufferToSnapshot`
     produces a known snapshot. Capture those snapshots as JSON files
     under `tests/golden-snapshots/` keyed by fixture name. Add a test
     that asserts current `xlsx.ts` produces the golden. Phase 2 will
     run the same test against `xlsxSheetJS.ts` and Phase 2 ships only
     when the assertion passes (or every divergence is documented +
     accepted).
   - **Feature-smoke test list.** A bullet list in
     `docs/m14-sheetjs-spike.md` of every Notesheet feature the
     operator listed as MUST-NOT-REGRESS:
     - Snapshot creation via `New Spreadsheet` command
     - Snapshot editing in Univer (cell value changes round-trip)
     - Univer toolbar formula bar
     - Univer formula evaluation (basic + structured refs)
     - Named tables: insert / right-click row+col operations
     - Chart insertion + live updates
     - Anchored chart drag/resize
     - .xlsx import via Tools menu
     - .xlsx export via editor button
     - Note navigation (open note, switch sheets, sidebar nav)
     - Univer rendering pixel correctness (M13/E reference fidelity)

     For each, state how Phase 2 will validate (existing Jest, new
     Jest, PGE harness, or manual).

## What does NOT ship in this PR

- **Production swap.** `src/xlsx.ts` stays as the production import.
  Don't change `src/index.ts` or `src/editorView.tsx`. The new
  `src/xlsxSheetJS.ts` is dead code from Notesheet's runtime
  perspective; only the test bed exercises it.
- **`xlsx` direct dependency.** `package.json` `dependencies` stays
  unchanged. Only `devDependencies` may add `xlsx-js-style`.
  Phase 2 will swap `dependencies`.
- **README "Known shortcomings" edits.** Punted to Phase 2.
- **Univer renderer follow-up** for M13/E's `bd.b` colour mismatch.
  Separate cycle.
- **Inter-banded-row strips for `TableStyleMedium4`.** Separate
  cycle.

## Acceptance criteria for the spike PR

The spike PR can merge when ALL of these are satisfied:

1. **Existing 209 Jest tests stay green.** Spike adds no production
   code change; the existing suite must be unaffected.
2. **`tests/xlsxParserParity.test.ts` runs and produces a matrix.**
   Either as a JSON output committed to the repo, or printed in the
   test output and surfaced in the decision document.
3. **`docs/m14-sheetjs-spike.md` is committed** with capability
   matrix, migration cost estimate, recommendation, and the
   feature-smoke test list for Phase 2.
4. **Golden snapshots committed under `tests/golden-snapshots/`** for
   at least the M12+M13 fixtures (the operator-owned ones), so
   Phase 2 has a deterministic target.
5. **`npm run dist` builds clean.** No new transitive deprecation
   noise from `xlsx-js-style` worse than what `exceljs` already
   carries.

## Out of scope for the spike

- **Performance benchmarking.** Both libraries parse a 1MB workbook
  in well under a second; the migration's value isn't speed, it's
  longevity / dependency hygiene / better CF/style coverage. Don't
  spend time on perf comparisons unless the spike surfaces a 10x+
  delta.
- **Univer rendering pixel parity.** That's a Phase-2 concern. The
  spike compares snapshot data shape only.
- **xlsx-js-style maintenance state.** Read its npm page; if it
  hasn't shipped in 18+ months, note that as a risk in the
  recommendation but don't block on it.

## Sequencing

1. Generator runs the spike: parallel module + test bed + decision
   doc + golden snapshots.
2. Spike PR opens; CI confirms existing tests still green and the
   parser parity matrix produces.
3. Operator reviews the recommendation. If GO, Phase 2 (migration)
   starts in a new branch / new PR.
4. If NO-GO, the spike PR can still merge (the parity matrix and
   decision doc are valuable artefacts even without a migration) OR
   close it without merging — operator's call at review time.

## Suggested fixtures

Every existing `tests/ExcelBaseTestData/` file. The spike's value
comes from running both parsers across the WHOLE matrix, not
cherry-picking. If a fixture trips SheetJS (or `xlsx-js-style`),
that's a finding for the matrix.

## Related risks

- **`xlsx-js-style` is a fork.** Its `xlsx@0.18.5` base is two years
  old. The fork itself last shipped in 2022 per npm. If it's
  abandoned, Phase 2's migration replaces one-quiet-library
  (`exceljs`) with another (`xlsx-js-style`). The decision doc must
  consider this honestly.
- **Cell styling is the biggest open question.** `xlsx` (plain
  SheetJS Community) explicitly doesn't support styling. `xlsx-js-style`
  claims to add it. Verify how much of M12/M13 it can actually carry:
  named-style banding (TableStyleMedium4 with theme accents), rich
  text per-run formatting, theme-tinted borders, conditional formatting
  dxfs. If the fork can't carry these, the migration is NO-GO.
- **Theme palette extraction.** Our `readThemeClrScheme` reads
  `xl/theme/theme1.xml` directly via JSZip and parses with regex.
  That's parser-agnostic — the same code works with SheetJS. Confirm
  the migration retains it.
- **Pattern B hyperlink detection.** Same — our `readNamedHyperlinkCells`
  is raw-XML via JSZip. Parser-agnostic. Confirm.
- **Chart export post-processing.** `src/charts/xlsxChart.ts` patches
  the zip exceljs writes. Confirm the SheetJS-written zip has the
  same structure for that patcher to work. If not, Phase 2 needs to
  port the patcher too.
- **Don't symptom-patch.** If a parity test fails because SheetJS
  emits a different cellXf id, do NOT change the test to accept it.
  Either find the equivalent SheetJS API or document the divergence
  in the matrix. Test parity is the whole point of the spike.

## Output

Reply with a paragraph describing what shipped (parallel module, test
bed, decision doc, golden snapshots), the recommendation (GO / NO-GO
/ CONDITIONAL), and any blockers found. Then stop. The operator will
review.
