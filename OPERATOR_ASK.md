# Operator ask — M15: Conditional formatting (full round-trip on exceljs)

## Why this matters

Conditional formatting is one of the most-used Excel formatting features
that Notesheet currently silently drops on import. The current behaviour
is documented as a "Known shortcoming" — when an Excel workbook with CF
rules is imported, the rules are stripped and the cell values render
plain. Round-trip back to .xlsx loses the rules entirely.

The M14 SheetJS spike removed the earlier "wait for SheetJS to make CF
cheaper" rationale (xlsx-js-style's `!cf` is also undefined on
indexed-cellXf import). M15 ships on the existing `exceljs` parser. No
parser blocker.

## The feature

When a user imports
`tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
into a Notesheet note, the CF rules MUST:

1. **Import** as Univer-shape rules under the snapshot's
   `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource (the Univer CF
   plugin's stable resource key).
2. **Render** in the Univer canvas with the right visual decoration:
   - Color scale (column A) — red→yellow→green gradient over `A2:A11`
     based on cell-value percentile.
   - Data bar (column C) — proportional blue horizontal bars.
   - cellIs > 50 (column E) — pink fill on cells whose value is `> 50`.
   - top10 / top-3 by rank (column G) — green fill on the top 3 cells.
   - iconSet 3-arrows (column I) — directional-arrow glyphs based on
     percentile bands (0–33% / 33–67% / 67–100%).
3. **Edit** via Univer's CF UI panel — users can add / modify / delete
   rules from the Notesheet editor.
4. **Export** back to .xlsx via the editor's Export button. The
   resulting workbook reads cleanly in real Excel and re-imports into
   Notesheet with the same rules.

The five rule types in the existing fixture cover the common cases.
Out of scope: text-based rules (`beginsWith`/`endsWith`), time-period
rules (`thisMonth`), unique-values / duplicate-values, formula-based
rules. Document these as M15 follow-up scope, not M15 blockers.

## Acceptance criteria

The evaluator must verify ALL of:

1. **Visual — all 5 rule types render correctly.** PGE evaluator's
   Playwright-captured screenshot of the imported note shows:
   - **Column A (colorScale)**: A2 through A11 carry distinct
     background colours along a red→yellow→green gradient. The cell
     with the lowest value is red-ish (`#F8696B` family); the cell at
     the median is yellow-ish (`#FFEB84` family); the cell with the
     highest value is green-ish (`#63BE7B` family). Evaluator must
     explicitly call out the gradient direction and the three named
     hues.
   - **Column C (dataBar)**: each cell C2..C11 has a proportional
     horizontal bar in blue (`#638EC6` family) whose width scales with
     value. Cells with smaller values have shorter bars; the largest
     value fills the full cell width.
   - **Column E (cellIs > 50)**: cells whose value is `> 50` carry a
     pink/light-red fill (`#FFC7CE` per the source dxf); cells `<= 50`
     are unfilled. Evaluator must verify at least one cell in each
     bucket.
   - **Column G (top-3 rank)**: exactly 3 cells in G2..G11 carry a
     light-green fill (`#C6EFCE` per the source dxf). The other 7 are
     unfilled.
   - **Column I (iconSet 3-Arrows)**: each cell I2..I11 has a small
     arrow glyph (red-down, yellow-flat, or green-up) corresponding to
     the value's bucket within the 0–33% / 33–67% / 67–100% percentile
     range.

2. **Pixel sidecar — colour signal lands in the histogram.** The
   `<screenshot>.pixels.json` sidecar emitted alongside the
   evaluator's screenshot, sampled over each CF column band, contains
   ink hits that match the spec. Specifically:
   - Column A band: red ink ≥ 30 hits AND green ink ≥ 30 hits over
     A2..A11 (gradient endpoints).
   - Column C band: blue ink ≥ 30 hits over C2..C11 (data bars).
   - Column E band: pink ink (R≥220, G∈[180,220], B∈[180,220]) ≥ 20
     hits over the > 50 cells.
   - Column G band: light-green ink ≥ 20 hits over the top-3 cells.
   - Column I band: at least one of red/green/yellow ink ≥ 5 hits
     (icon glyphs are small; the threshold is intentionally loose).

   Add new region helpers (`cfColumnRegion(col)`) and band aggregates
   (`pinkInk`, `lightGreenInk`) to `eval-screenshot.js`.

3. **Snapshot fidelity — Excel reference → snapshot data shape.**
   `tests/excelReferenceFidelity.test.ts` asserts that
   `xlsxBufferToSnapshot('ConditionalFormatting-Variants.xlsx')`
   produces a snapshot whose
   `resources['SHEET_CONDITIONAL_FORMATTING_PLUGIN']` carries 5 rules
   matching the source XML structurally:
   - Rule 0: type=`colorScale`, ref=`A2:A11`, 3 cfvo
     (min/percentile=50/max), 3 colours
     (#F8696B / #FFEB84 / #63BE7B).
   - Rule 1: type=`dataBar`, ref=`C2:C11`, 2 cfvo (min/max), colour
     #638EC6.
   - Rule 2: type=`highlightCell` (subType=`number`), ref=`E2:E11`,
     operator=`greaterThan`, formula=`50`, fill bg=#FFC7CE.
   - Rule 3: type=`highlightCell` (subType=`rank`), ref=`G2:G11`,
     rank=3, isBottom=false, fill bg=#C6EFCE.
   - Rule 4: type=`iconSet`, ref=`I2:I11`, iconSet=`3Arrows`, 3 cfvo
     (percent 0/33/67).

   This is the import-side gate. Authored failing first, then made to
   pass — that ordering is the structural integrity check
   (`feedback_pge_fidelity_test_gap.md`).

4. **Canvas fidelity — Joplin canvas vs Excel reference.**
   `tests/excelCanvasFidelity.test.ts` adds 5 tests (one per CF
   column) that sample the dominant colour of each column band in
   BOTH the operator-captured Excel reference screenshot AND the
   Joplin canvas screenshot, asserting parity within Δ ≤ 16 RGB per
   channel. Reference screenshot from real Excel:
   `screenshots/excel-reference/ConditionalFormatting-Variants.png`
   (operator captures it as part of the cycle).

   This is the render-side gate. The M13/E lesson stands: the
   snapshot-data fidelity test alone is not sufficient — the
   render-vs-render test catches Univer-renderer bugs the
   snapshot-data test can't see.

5. **Round-trip — CF rules survive export → re-import.** New Jest
   test (or extend `tests/m12FixtureRoundTrip.test.ts`):
   - Import `ConditionalFormatting-Variants.xlsx` → snapshot
   - Call `snapshotToXlsxBuffer(snapshot)` → buffer
   - Re-load buffer with a fresh exceljs `Workbook`
   - Assert all 5 rules survived: same `type`, same `ref`, same
     `cfvo`, same `color` / `fill.bgColor`, same operator/rank where
     applicable.
   - The KNOWN SHORTCOMING test in `tests/m12FixtureRoundTrip.test.ts`
     about CF being dropped is flipped to a positive pin-down.

6. **Univer CF preset is registered in the editor.** New imports in
   `src/editorView.tsx`:
   - `import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';`
   - `import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';`
   - `import sheetsCfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';`
   - Added to the presets array passed to `createUniver()`.
   - Added to the locales merge.

7. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** Generator captures own evidence under
   `screenshots/feature-1-m15-conditional-formatting/`, reads via the
   Read tool to satisfy the verify-gate hook.

8. **`npm test` — all 209+ existing tests stay green plus new ones.**
   Target: 209 → ≥ 220 (new pin-downs across snapshot fidelity,
   canvas fidelity, and round-trip).

## Out of scope

- **Text-based CF rules** (`beginsWith`, `endsWith`, `containsText`,
  etc). Not in the fixture; document as M15-followup.
- **Time-period CF rules** (`thisMonth`, `last7Days`, etc). Same.
- **Unique-values / duplicate-values rules.** Same.
- **Formula-based CF rules** (`type="expression"`). Same.
- **Stop-If-True semantics.** OOXML CF rules can carry `stopIfTrue` —
  if a higher-priority rule matches, lower-priority ones don't apply.
  Document the current behaviour (always apply, even after a higher
  match) and revisit if it surfaces as a real-world issue.
- **CF inside merged cells.** The fixture doesn't exercise it.
- **Custom dxf records beyond fill-bg.** The fixture uses fill-bg only.
  If exceljs surfaces a font-colour or border on a cellIs/top10 dxf,
  preserve them; otherwise out of scope for M15.
- **Univer's "manage rules" panel UI customisations.** Use the panel
  exactly as Univer ships it — no theming, no relabelling.
- **Editing CF via Joplin's note import path.** The CF UI panel is
  the editor surface; users won't edit CF via a markdown table or a
  fenced-code modification. Editor-only editing.
- **README "Known shortcomings — Conditional formatting" cleanup.**
  Punted to a follow-up docs PR (precedent: M13/C, M13/D, M13/E all
  punted their entries to PRs #21 and the like).

## Suggested fixtures

- `tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
  — already in the repo. 5 rules, one per column (A/C/E/G/I).
  This is the spec's central fixture.
- `screenshots/excel-reference/ConditionalFormatting-Variants.png`
  — operator-captured reference. Used by the canvas-fidelity test
  as ground truth. Generator creates the empty placeholder; operator
  captures from real Excel before the evaluator runs.

The harness invokes `import-fixture.sh ConditionalFormatting-Variants.xlsx`
and the canvas-targeted screenshot via `eval-screenshot.js`. Region
helpers covering each of the 5 CF columns must be added to
`REGION_BY_FEATURE` and `TITLE_PREFIX_BY_FEATURE`.

## Related risks

- **Univer's CF preset must be wired into the editor.** Step 6 above.
  Without this, the snapshot's CF data lands but the canvas paints
  nothing — Jest tests pass on snapshot fidelity, runtime visual gate
  fails. M13's exact failure mode. Don't ship without confirming the
  preset is loaded.
- **`SHEET_CONDITIONAL_FORMATTING_PLUGIN` is the resource key** Univer
  expects on the snapshot. It's a string constant exported from
  `@univerjs/sheets-conditional-formatting`. Reuse the export, don't
  hard-code the string.
- **exceljs's CF data shape doesn't match Univer's 1:1.** Translation
  table:
  - exceljs `colorScale` → Univer `colorScale` (cfvo + colour arrays
    map directly).
  - exceljs `dataBar` → Univer `dataBar` (single colour, two cfvo).
  - exceljs `cellIs` → Univer `highlightCell` (subType `number`,
    operator + formula carry over).
  - exceljs `top10` → Univer `highlightCell` (subType `rank`, rank
    value, isBottom flag — exceljs uses `bottom: true` on `top10`
    rules to mean bottom-N; Univer uses `isBottom: true`).
  - exceljs `iconSet` → Univer `iconSet` (iconSet name maps directly:
    `3Arrows`, `3TrafficLights1`, etc; cfvo bands carry over).
  Each mapping is well-defined; build a per-type translator.
- **Round-trip ARGB drift.** exceljs surfaces colours as `argb`
  strings (`FF638EC6` — leading FF is alpha). Univer expects `#RRGGBB`
  (no alpha). The export side has to flip this back. Don't drop the
  alpha on import then accidentally reintroduce it on export — both
  sides go through `argbToHex` / `hexToArgb` helpers.
- **The synthesizer interaction.** M12's
  `synthesizeTableStyleAssignments` bakes per-cell `bg`/`cl` from the
  table style. CF rules are applied by Univer's renderer ON TOP of
  the cell style. Confirm Univer paints CF over the synthesized
  table-style fill, not under it. If it paints under, the
  table-style's fill will mask the CF colour and the visual gate
  fails. The fix shape would be to mark CF-bound cells in the
  synth-styles sidecar so the M12 synthesizer skips painting them.
- **iconSet rendering.** Univer's CF engine ships icon glyphs for
  the standard iconSet names (`3Arrows`, `3TrafficLights1`,
  `5Arrows`, etc.). Confirm `3Arrows` works without additional asset
  bundling. If glyphs render as missing characters, that's a
  Univer-side icon-font issue — investigate whether the preset's CSS
  includes the icon font or whether we need an extra import.
- **The PGE harness's `:variant` suffix scheme** (M13/E set the
  precedent) supports per-fixture screenshots. M15 has one fixture
  but FIVE columns to grade. Either:
    (a) Single screenshot, multiple region helpers
        (`cfColumnRegion('A')` … `cfColumnRegion('I')`), tests sample
        each region from the same image.
    (b) Multiple screenshots with `:variant=column-A` etc.
  Option (a) is cleaner for this case — all 5 rules visible in one
  shot.
- **Don't symptom-patch the test on a Jest failure post-rebuild.** If
  a test regresses after wiring the CF preset, run
  `git diff package-lock.json` first. The CF preset may have pulled
  new transitives. Reference: `feedback_dependency_hygiene.md`.
- **The feature touches `src/xlsx.ts` (the parser's import + export
  paths) and `src/editorView.tsx` (preset wiring).** Run the full
  Jest suite before flipping the row. Borders, hyperlinks, table
  styles, chart export, alignment, rotation, rich text, and theme
  banding all share the import/export pipeline. Nothing here should
  regress.
- **Test gap warning.** Per `feedback_pge_fidelity_test_gap.md`,
  pin-downs anchored to our own emit are no test at all. The
  reference-fidelity test's expected RGBs come from the source XML,
  not from the parser output. The canvas-fidelity test's expected
  RGBs come from the Excel reference screenshot, not from our
  render. Both layers' assertions are anchored upstream of our code.
