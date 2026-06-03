# Operator ask — M13/E: theme-aware banding

Third real-feature cycle through the PGE harness. M13/C (rotated text,
PR #19) and M13/D (rich text, PR #20) shipped via this same loop.
M13/E is the last M13 workstream — theme-aware banding accuracy —
and finishes M13.

## Why this matters

Notesheet's import path bakes per-cell `bg` / `cl` / `bd` into
`cellData[row][col].s` for every cell of an imported `.xlsx` table,
synthesizing the colours from `EXCEL_TABLE_STYLE_BY_NAME` in
`src/charts/excelTableStyles.ts`. That catalog is hardcoded against
the **Office 2016+ Aptos theme** (accent1 `#156082`, accent3
`#196B24` green, etc.).

When a user imports a workbook that ships a non-Aptos `<a:clrScheme>`
— for example the project-owned fixture
`FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`
whose accent3 is `#A5A5A5` (grey) and whose
`xl/tables/table1.xml` declares `TableStyleMedium4` (an accent3-keyed
style) — the in-Joplin render paints **green** header + bands instead
of the grey Excel would render. The exported `.xlsx` round-trips the
source clrScheme correctly (verified by an existing pin-down in
`tests/m12FixturePinDowns.test.ts`), so re-opening the file in Excel
shows the right colours; only the Joplin-side render is wrong.

The README's "Theme-aware banding" entry under "Known shortcomings"
flags this as `→ M13/E`. PR #21 is the most recent edit to that
entry.

## The feature

When a user imports any workbook that ships its own `<a:clrScheme>`
(any non-Aptos theme), the in-Joplin render of every named-style
table (`TableStyleLight*`, `TableStyleMedium*`, `TableStyleDark*`)
matches what Excel renders for that workbook. Concretely, the
synthesizer must derive per-cell `bg` / `cl` / `borderColor` from the
**source workbook's clrScheme** when present, falling back to the
hardcoded Aptos catalog only when no clrScheme is detected.

The two project-owned fixtures pin down the two halves of this:

- **Aptos** (`FormattingSmorgasboard.xlsx`, `TableStyleMedium4`,
  accent3 `#196B24`) — header bg renders **green** `#196B24`, band
  rows light green `#C2F1C8` (or whatever the catalog already
  emits). This is the existing behaviour; M13/E must not regress it.
- **Classic** (`FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`,
  `TableStyleMedium4`, accent3 `#A5A5A5`) — header bg renders
  **grey** `#A5A5A5`, band rows light grey (whatever
  `tint(+0.6, #A5A5A5)` produces under Excel's tint formula). This
  is the failure mode being fixed.

## Acceptance criteria

The evaluator must verify ALL of:

1. **Visual — Aptos and Classic fixtures render distinct table
   colours.** Two evaluator screenshots under
   `screenshots/feature-1-m13-theme-aware-banding/`, one per fixture:
   - `eval-aptos-*.png`: header row of the imported `ProjectTracker`
     table renders **green** (Aptos accent3 `#196B24`). The
     evaluator must explicitly call out the green hue and that
     banding rows are light green.
   - `eval-classic-*.png`: header row of the imported
     `ProductCatalog` table renders **grey** (Classic accent3
     `#A5A5A5`), NOT green. The evaluator must explicitly call out
     that the header is grey, not green, and that banding rows
     match (light grey, not light green).
2. **Pixel sidecar — header colour signal differentiates Aptos vs
   Classic.** Two `<screenshot>.pixels.json` sidecars over a header-
   row region helper (analogous to `richTextA1A2Region`):
   - Aptos sidecar: green ink `(R≤80, G≥80, B≤80)` ≥ 30 hits over
     the header band; grey ink (`R≈G≈B`, mid-range) below 10 hits.
   - Classic sidecar: grey ink (`R∈[140,180]`, `G∈[140,180]`,
     `B∈[140,180]`) ≥ 30 hits; green ink below 10 hits.
   The harness needs new aggregates `greenInk` (already present
   from M13/D) and `greyInk` (new) plus a `tableHeaderRowRegion`
   helper that picks the row 1 band. Per-fixture
   `REGION_BY_FEATURE` and `TITLE_PREFIX_BY_FEATURE` entries with
   suffixes (e.g. `feature-1-m13-theme-aware-banding:aptos` and
   `:classic`) — or a single entry with the harness importing
   both fixtures into separate notes.
3. **Jest — theme-aware synthesis is pinned down.** New positive
   pin-down tests in `tests/m12FixturePinDowns.test.ts` (or a new
   `tests/m13ThemeAwareBanding.test.ts`):
   - Importing the Classic fixture: the header cell `s` style's `bg`
     resolves to `#A5A5A5` (grey), NOT `#196B24` (green).
   - Importing the Aptos fixture: the header cell `s` style's `bg`
     resolves to `#196B24` (green) — regression check.
   - The total Jest passing count moves from 195 (M13/D baseline)
     to at least 197 (M13/E adds ≥2 new pin-downs). No `KNOWN
     SHORTCOMING — theme-aware banding` test is left referencing
     M13/E.
4. **Round-trip — exported tables still carry `tableStyleInfo`
   pointing at the same style name.** A regression check that the
   round-tripped table.xml's `name="TableStyleMedium4"` survives —
   M13/E only changes the in-memory snapshot's per-cell baked
   colours, not what gets exported. The existing pin-down in
   `m12FixturePinDowns.test.ts` ("Aptos fixture: round-trip
   preserves table name + ProjectTracker columns") must stay green.
5. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** Generator captures its own screenshots under
   `screenshots/feature-1-m13-theme-aware-banding/` for both
   fixtures, then opens them via the Read tool so the
   `verify-gate` hook unlocks the `test-results.json` write.
   Evaluator captures its own independently.

## Out of scope

- **Theme-tinted borders** (the other "Known shortcomings" entry).
  Border `{theme: N, tint: T}` references are resolved against
  whichever clrScheme is loaded at import time. That's already
  correct in `resolveExceljsColor` (M12). M13/E only addresses
  table-style synthesis, not direct border references.
- **Conditional formatting** preserving theme colours. M16.
- **Custom `<tableStyle>` entries** authored in the workbook (not
  built-in). Those have real `<dxf>` records the importer can read
  directly; M13/E only fixes the built-in lookup-catalog path.
- **First-column / last-column emphasis** (`showFirstColumn`,
  `showLastColumn`). Not modelled by `EXCEL_TABLE_STYLE_BY_NAME`
  today; out of scope here.
- **Column stripes** (`showColumnStripes`). Same — rare, not
  modelled.
- **Recomputing tints from the source clrScheme using ECMA-376
  HSL-L tint maths from scratch.** M13/E may take the simpler
  path: when a non-Aptos clrScheme is present, swap the catalog
  entry's accent reference for the source's accent and apply
  the catalog's pre-computed tint. The catalog's tint maths are
  already correct (PR #14 verified them); only the accent input
  changes. Implementation may use whichever approach is cleaner —
  the spec is the user-observable colour, not the algorithm.
- **README "Known shortcomings — Theme-aware banding" edit.** Punt
  to a follow-up like the rotated-text and rich-text entries did.
  PR #21 set the precedent.

## Suggested fixtures

- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard.xlsx`
  (Aptos, accent3 `#196B24`, `TableStyleMedium4`, `ProjectTracker`).
  Pre-existing fixture; the regression sentinel.
- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`
  (Classic, accent3 `#A5A5A5`, `TableStyleMedium4`, `ProductCatalog`).
  Pre-existing fixture; the failure-mode sentinel.

The harness imports both via
`scripts/pge/import-fixture.sh <fixture-name>` (already wired) into
two separate notes. Region helper `tableHeaderRowRegion` covers row 1
of the visible canvas. The eval-screenshot script needs to be
extended so it can capture both notes in one run — either two
invocations with different `:aptos` / `:classic` suffixes on the
feature ID, or a single invocation that opens both notes in
sequence.

## Related risks

- **The synthesizer is deeply wired into M12.** `synthesizeTableStyleAssignments`
  in `src/xlsx.ts:943` consumes `EXCEL_TABLE_STYLE_BY_NAME[table.styleName]`
  and bakes the result into per-cell records. Changing the input
  source from "static catalog" to "static catalog rebased against
  source clrScheme" must not change the field shape, the
  `addedFields` tagging, or the sidecar resource format. The M12
  pin-downs (synth-styles sidecar, applyFont=1 invariant) are the
  guards.

- **The hardcoded catalog encodes accent INDEXES indirectly through
  pre-computed RGBs.** `TableStyleMedium2` is accent1, Medium3 is
  accent2, Medium4 is accent3, … (every 7 entries shifts to the
  next accent). To rebase against a non-Aptos scheme, you need a
  table mapping each style name → which accent (1..6) it uses, plus
  a way to apply the catalog's tint relative to that accent. One
  approach: precompute the tint amount per style per band slot
  (header, evenRow, oddRow, totals, border) and store that in a
  parallel catalog, then rebase by recomputing
  `applyTint(sourceClrScheme[accent], tintAmount)` at import time.
  Another: keep the existing catalog, and on import detect non-
  Aptos themes and remap each accent in the result by HSL channel
  swap. Either works; the simpler one wins.

- **The Jest pin-down for "synth-styles sidecar" expects header
  tags to include `bg` and `bl` and 3 borders.** That contract is
  shape, not colour — but verify it stays green after the rebase.

- **The Aptos fixture has hand-edited cell-level colours on top of
  table-style synthesis.** Cell A2 carries a hand-applied `#F4B183`
  border. Synthesis must not overwrite hand-edits — the existing
  `existingCellStyles` map and `addedFields` are how M12 already
  handles this. A theme-aware rebase must preserve that
  precedence.

- **The Classic fixture's `ProductCatalog` table also uses
  `TableStyleMedium4`** — same style name, different accent.
  That's the whole point of the test: ONE catalog entry must
  produce two distinct rendered colours depending on the source
  workbook's clrScheme. If you "fix" by adding a second hardcoded
  catalog entry, you've missed the design — Excel ships ~thousands
  of accent permutations, the catalog can't enumerate them.

- **The PGE harness has only screenshotted one note per cycle so
  far.** M13/E is the first feature requiring TWO independent
  screenshots in one session. The eval-screenshot script needs an
  extension: either accept a `--fixture` arg, or look up multiple
  fixtures in `REGION_BY_FEATURE` keyed by `feature-id:variant`,
  or take both screenshots in one run. The cleanest approach is
  probably the variant suffix — keeps backward compatibility with
  prior cycles' single-screenshot path.

- **Don't symptom-patch by editing the catalog itself.** Hardcoding
  Classic-accent3 alongside Aptos-accent3 is the wrong shape (see
  preceding bullet). The fix must be data-driven from the source
  clrScheme.

- **The feature touches `src/xlsx.ts` (the same file as M9–M13/D)
  and `src/charts/excelTableStyles.ts`.** Run the full Jest suite,
  not just the new pin-downs, before flipping the row. Borders,
  hyperlinks, table styles, chart export, alignment, rotation, and
  rich text all share the import/export pipeline. The m12, m13/C,
  and m13/D round-trip tests must stay green.

- **Region-by-feature pixel sampling.** The eval-screenshot script
  needs `tableHeaderRowRegion` (or whatever name fits) plus a
  `greyInk` band aggregate. M13/D's `richTextA1A2Region` and
  `redInk`/`blueInk` aggregates are the template.
