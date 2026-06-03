# Notesheet PGE — build plan (M13/E: theme-aware banding accuracy)

> **Cycle context.** This is the THIRD real-feature cycle through the
> PGE harness. M13/C (rotated text round-trip, PR #19) and M13/D
> (rich text within one cell, PR #20) have shipped via this same
> loop, and the README cleanup PR #21 just merged. Their rows live
> in `PROGRESS.md` `## Done` and are intentionally absent from this
> `test-results.json` (the harness keeps only the in-flight cycle's
> rows). Harness scripts under `scripts/pge/` (`launch-joplin.sh`,
> `install-plugin.sh`, `import-fixture.{ts,sh}`,
> `eval-screenshot.{js,sh}`, `prep-joplin-window.sh`,
> `create-seeded-notesheet.js`) are present and proven on M13/C +
> M13/D. The pixel-sidecar (`<screenshot>.pixels.json`) and
> region-by-feature lookup (`REGION_BY_FEATURE` /
> `TITLE_PREFIX_BY_FEATURE` in `eval-screenshot.js`) are the
> machine-checkable signal alongside the visual screenshot.
>
> **First multi-screenshot cycle.** Until now every cycle has captured
> exactly one note per session. M13/E is the first feature whose
> evidence requires TWO independent screenshots (Aptos fixture +
> Classic fixture) — one rendered note per source theme. The harness
> needs a small extension to support that; this plan picks the
> approach.

## Operator ask

(See `OPERATOR_ASK.md`.) Notesheet's import path bakes per-cell
`bg` / `cl` / `bd` into `cellData[row][col].s` for every cell of an
imported `.xlsx` table, synthesizing the colours from
`EXCEL_TABLE_STYLE_BY_NAME` in `src/charts/excelTableStyles.ts`. That
catalog is hardcoded against the Office 2016+ Aptos theme (accent3
`#196B24` green). When the source workbook ships its own non-Aptos
`<a:clrScheme>` — for example
`FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`
whose accent3 is `#A5A5A5` (grey) — the in-Joplin render paints
**green** header + bands instead of the **grey** Excel renders. The
exported `.xlsx` already round-trips the source clrScheme correctly
(M12 work — verified by the `Classic fixture: exported theme1.xml
carries source accents` pin-down at
`tests/m12FixturePinDowns.test.ts:193`); only the in-Joplin-side
synthesis is theme-blind. M13/E is the last M13 workstream and
finishes M13.

## Harness extension chosen for this cycle

The operator ask raises two options for the multi-screenshot
problem; this plan picks **the variant-suffix approach** for these
reasons:

1. **Backward compatibility.** Every prior feature is keyed by a
   plain feature ID in `REGION_BY_FEATURE` and
   `TITLE_PREFIX_BY_FEATURE`. A multi-fixture single invocation
   would either rewrite those table values (breaking the M13/D
   screenshot path) or special-case M13/E.
2. **Smaller blast radius.** Adding entries
   `feature-1-m13-theme-aware-banding:aptos` and
   `feature-1-m13-theme-aware-banding:classic` keys is purely
   additive — no existing entry changes. The lookup logic in
   `eval-screenshot.js` becomes "if the FEATURE_ID env var contains
   a `:`, look up the suffixed key; otherwise fall back as today."
3. **Two clean screenshot files.** Each fixture lands its own PNG
   + `.pixels.json` sidecar at deterministic paths
   (`screenshots/feature-1-m13-theme-aware-banding/eval-aptos-*.png`
   and `…/eval-classic-*.png`), which the evaluator can grade
   independently.
4. **Reusable.** Future features (e.g. multi-sheet workbooks,
   per-locale comparisons) can adopt the same `:variant` suffix
   without further harness changes.

The generator implements this by:
- Adding two suffixed keys to both lookup tables in
  `scripts/pge/eval-screenshot.js`.
- Adding a `tableHeaderRowRegion` helper covering the
  ProjectTracker / ProductCatalog header row (row 1, A:G slab) on
  the Univer canvas.
- Adding a `greyInk` aggregate to `samplePixelsAt` (template:
  `R∈[140,180]` AND `G∈[140,180]` AND `B∈[140,180]`, abs(R−G) ≤ 10,
  abs(G−B) ≤ 10, to gate "grey ink, not coloured ink").
- Calling `eval-screenshot.sh feature-1-m13-theme-aware-banding:aptos`
  and `eval-screenshot.sh feature-1-m13-theme-aware-banding:classic`
  one after the other, importing the matching fixture before each
  call via `import-fixture.sh`.
- The `verify-gate` hook unlocks the `test-results.json` write
  after Read of BOTH PNGs (one per variant).

The generator is free to extend the harness in additional ways
(e.g. a `--fixture` flag) but MUST NOT regress the prior-cycle
single-key path; the M13/C and M13/D entries must keep working
without rename.

## Features

### feature-1-m13-theme-aware-banding

**Spec**

When a user imports any `.xlsx` workbook whose `<a:clrScheme>` is
non-Aptos (i.e. anything other than the Office 2016+ Aptos accents
`#156082 / #E97132 / #196B24 / #0F9ED5 / #A02B93 / #4EA72E`), the
in-Joplin render of every named-style table (`TableStyleLight*`,
`TableStyleMedium*`, `TableStyleDark*`) matches the colours Excel
would render against the workbook's own theme — derived per cell
from the source clrScheme rather than from the hardcoded Aptos
catalog.

The two project-owned fixtures pin down the two halves:

- **Aptos** (`FormattingSmorgasboard.xlsx`, `TableStyleMedium4`,
  accent3 `#196B24`, table `ProjectTracker` over `A1:G10`). This is
  the regression sentinel — the existing behaviour must not regress.
  Header row paints **green** `#196B24`; banded rows paint pastel
  green `#84E291`.
- **Classic**
  (`FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`,
  same `TableStyleMedium4`, accent3 `#A5A5A5`, table `ProductCatalog`
  over `A1:F10`). This is the failure-mode sentinel — header row
  paints **grey** `#A5A5A5`, banded rows paint a light grey derived
  by applying the Medium4 catalog tint to grey (NOT to Aptos green).

The same `EXCEL_TABLE_STYLE_BY_NAME[Medium4]` lookup must produce
TWO distinct rendered outputs depending on the source workbook's
clrScheme. Hardcoding a Classic-Medium4 catalog entry alongside the
Aptos one is the wrong shape (Excel ships ~thousands of accent
permutations). The fix is data-driven from the source clrScheme.

The synthesizer's existing field-tagging contract
(`SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN` resource, `addedFields` per
cell, the `existingCellStyles` precedence rule that hand-edited
colours win over synthesis) is preserved unchanged. M13/E only
changes WHICH RGB the synthesizer chooses; the shape of the output
and the export-time strip behaviour stay identical.

**Acceptance criteria**

The evaluator must verify ALL of the following from a fresh context:

1. **Visual — Aptos and Classic fixtures render distinct table
   colours.** Two evaluator-captured screenshots under
   `screenshots/feature-1-m13-theme-aware-banding/`, one per
   fixture, taken via
   `bash scripts/pge/eval-screenshot.sh feature-1-m13-theme-aware-banding:aptos`
   and
   `bash scripts/pge/eval-screenshot.sh feature-1-m13-theme-aware-banding:classic`
   after the matching fixture has been imported into Joplin via
   `scripts/pge/import-fixture.sh`. Each screenshot is taken
   against the Univer main canvas
   (`canvas[id^="univer-sheet-main-canvas"]`) inside the editor's
   `UserWebviewIndex.html` frame, after `prep-joplin-window.sh`
   has filled the window and hidden the side panes:
   - **`eval-aptos-*.png`**: header row of the imported
     `ProjectTracker` table at row 1 of the visible canvas renders
     in **a recognisable green hue** (Aptos accent3 `#196B24`,
     dark green; the catalog's pastel-green band `#84E291`
     alternates on data rows). The evaluator verdict text MUST
     explicitly call out the green header and the green-tinted
     banding.
   - **`eval-classic-*.png`**: header row of the imported
     `ProductCatalog` table at row 1 of the visible canvas renders
     in **a recognisable grey hue** (Classic accent3 `#A5A5A5`,
     mid-grey; light-grey banded rows on alternating data rows),
     NOT green. The evaluator verdict text MUST explicitly call
     out that the header is grey, not green, and that banding rows
     are light grey, not light green.
   - The two screenshots side-by-side must be visibly different in
     hue at the header band, not "both green" and not "both grey."
2. **Pixel sidecar — header colour signal differentiates Aptos vs
   Classic.** Two `<screenshot>.pixels.json` sidecars, one per
   fixture, sampled over a new region helper covering the header
   row (`tableHeaderRowRegion`), with these aggregates:
   - **Aptos sidecar** (`eval-aptos-*.pixels.json`): `greenInk`
     (R≤80 AND G≥150 AND B≤80) **≥ 30 hits** over the header
     band — proves the Aptos green header is rendering.
     `greyInk` (the new aggregate; `R∈[140,180]`, `G∈[140,180]`,
     `B∈[140,180]`, with `abs(R−G) ≤ 10` and `abs(G−B) ≤ 10`)
     **< 10 hits** — proves grey is NOT bleeding into the Aptos
     case.
   - **Classic sidecar** (`eval-classic-*.pixels.json`): `greyInk`
     **≥ 30 hits** over the header band — proves the Classic
     grey header is rendering. `greenInk` **< 10 hits** — proves
     green is NOT being painted on the Classic fixture (the
     symptom we're fixing).
   - The harness extension adds the `tableHeaderRowRegion` helper
     plus a `greyInk` count to `samplePixelsAt` in
     `scripts/pge/eval-screenshot.js`. M13/D's `richTextA1A2Region`
     and `redInk`/`blueInk` aggregates are the template.
3. **Jest — theme-aware synthesis is pinned down.** New positive
   pin-down tests in `tests/m12FixturePinDowns.test.ts` (or a new
   sibling file `tests/m13ThemeAwareBanding.test.ts` — the
   generator chooses):
   - **Classic fixture pin-down**: importing
     `FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`
     produces a snapshot whose `ProductCatalog` header cell (the
     synthesized one — typically row 0, col 0 of the table range)
     resolves its `s` style record to `bg.rgb === '#A5A5A5'`
     (grey, the Classic accent3), NOT `'#196B24'` (Aptos green).
   - **Aptos fixture regression pin-down**: importing
     `FormattingSmorgasboard.xlsx` produces a snapshot whose
     `ProjectTracker` header cell resolves its `s` style record
     to `bg.rgb === '#196B24'` (Aptos green).
   - **Banded-row colour pin-down (Classic)**: a banded data row
     of the Classic table resolves to a light-grey
     `bg.rgb` — value matches `tint(+0.6, '#A5A5A5')` per the
     ECMA-376 HSL-L tint formula the catalog uses for Medium4's
     even rows; if the implementation reuses the catalog's
     pre-computed `bandedRowEvenBg` against accent3, it will be a
     greyscale value derived from `#A5A5A5`. **Pin down the
     greyscale shape** (e.g. `expect(bg).toMatch(/^#[A-F0-9]{2}\1\1$/i)`
     where the bytes are equal — grey is `RR === GG === BG`),
     NOT a specific RGB. The exact byte values may differ
     between the two valid implementation paths the operator
     calls out (re-tint from clrScheme vs catalog accent-swap),
     but EITHER path produces a greyscale.
   - The total Jest passing count moves from the M13/D baseline
     **195** (verify via `npm test` at start of session) to
     **at least 197** — at minimum the two header-colour
     pin-downs (Classic grey + Aptos regression). A third
     pin-down for the banded-row shape pushes this to ≥198,
     preferred but not blocking.
   - **No `KNOWN SHORTCOMING — theme-aware banding` test is left
     in the suite referencing `→ M13/E`**. If
     `m12FixtureRoundTrip.test.ts` carries one (currently does
     not by name; the closest is the `KNOWN SHORTCOMING —
     theme-tinted borders … (→ M13)` test at line 158, which is
     a different scope — **borders**, not table-style banding,
     and remains an acknowledged shortcoming under M13/E's
     out-of-scope; do not flip that one).
4. **Round-trip — the existing M12 invariants stay green.** The
   feature must NOT regress any of these pin-downs in
   `tests/m12FixturePinDowns.test.ts`:
   - `Classic fixture: exported theme1.xml carries source accents
     (#4472C4, not #4F81BD)` (line 193) — round-trip preserves
     the Classic clrScheme.
   - `Aptos fixture: exported theme1.xml carries source accents
     (#156082, #196B24, hlink #467886)` (line 179) — round-trip
     preserves the Aptos clrScheme.
   - `import emits SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN with
     header tags for table cells` (line 226) — the synth-styles
     sidecar still tags `bg`, `bl`, and 3 borders on the header
     cell. Shape is invariant; only RGB values may differ.
   - `import emits SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN with
     the source clrScheme` (line 239) — the source clrScheme
     resource is still emitted unchanged.
   - `synthesized header (no source cell font) exports without
     per-cell font` (line 270) — `applyFont=1` invariant on
     synth-only cells.
   - `Aptos fixture: round-trip preserves table name, style,
     stripes, totals row, columns` (line 359) — `name=
     "TableStyleMedium4"` survives in the exported `table1.xml`.
   - `Classic fixture: round-trip preserves table name +
     ProductCatalog columns` (line 378) — same for the Classic
     fixture.
   The full Jest suite (`npm test`) is run before flipping the
   row, NOT just the two new pin-downs. M13/E touches
   `src/xlsx.ts` (the same file as every prior milestone's
   import/export) and possibly `src/charts/excelTableStyles.ts`;
   borders, hyperlinks, table styles, chart export, alignment,
   rotation (M13/C), and rich text (M13/D) all share the
   pipeline.
5. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** The generator captures its OWN screenshots for both
   variants under `screenshots/feature-1-m13-theme-aware-banding/`
   (`generator-evidence-aptos.png` and
   `generator-evidence-classic.png` plus their pixel sidecars),
   then opens BOTH files via the Read tool so the `verify-gate`
   hook unlocks the `test-results.json` write. Generator-evidence
   screenshots are corroborating, not authoritative — the
   evaluator captures its own independently and grades from those.

**Out of scope**

- **Theme-tinted borders.** The other "Known shortcomings" entry
  about `{theme: N, tint: T}` border references is unrelated.
  M12's `resolveExceljsColor` already resolves those against
  whichever clrScheme is loaded at import time; that's correct.
  M13/E only addresses **table-style banding synthesis**, not
  direct border references.
- **Conditional formatting preserving theme colours.** M16. The
  Classic fixture happens to ship two `<conditionalFormatting>`
  blocks (color scales on D2:D9 and F2:F9); those are dropped on
  import today and stay dropped post-M13/E.
- **Custom `<tableStyle>` entries authored in the workbook.** The
  built-in lookup-catalog path is the only one M13/E fixes.
  Custom table styles have real `<dxf>` records the importer can
  read directly.
- **First-column / last-column emphasis** (`showFirstColumn`,
  `showLastColumn`). Not modelled by `EXCEL_TABLE_STYLE_BY_NAME`
  today; remains out of scope.
- **Column stripes** (`showColumnStripes`). Same — rare, not
  modelled.
- **README "Known shortcomings — Theme-aware banding" edit.**
  Punt to a follow-up like the M13/C and M13/D edits did. PR #21
  set the precedent. The evaluator does NOT penalise if the docs
  edit is deferred.
- **Recomputing tints from the source clrScheme using ECMA-376
  HSL-L tint maths from scratch.** M13/E may take the simpler
  path — when a non-Aptos clrScheme is present, swap the catalog
  entry's accent reference for the source's accent and apply the
  catalog's pre-computed tint. The catalog's tint maths are
  already correct (PR #14 verified). Either implementation
  approach (re-tint vs accent-swap) is acceptable; the spec is
  the user-observable colour, not the algorithm.

**Suggested fixtures**

- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard.xlsx`
  (Aptos, accent3 `#196B24`, `TableStyleMedium4`, table
  `ProjectTracker` over `A1:G10`). Pre-existing fixture; the
  regression sentinel.
- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`
  (Classic, accent3 `#A5A5A5`, same `TableStyleMedium4`, table
  `ProductCatalog` over `A1:F10`). Pre-existing fixture; the
  failure-mode sentinel.

The harness imports both via `scripts/pge/import-fixture.sh
<fixture-name>` (already wired) into TWO separate notes with
distinct titles — the generator picks title prefixes consistent
with the variant-suffix scheme (e.g. `PGE M13E aptos eval ` and
`PGE M13E classic eval `) and registers them in
`TITLE_PREFIX_BY_FEATURE` keyed off
`feature-1-m13-theme-aware-banding:aptos` and `…:classic`.

**Related risks**

- **The synthesizer is deeply wired into M12.**
  `synthesizeTableStyleAssignments` in `src/xlsx.ts:943` consumes
  `EXCEL_TABLE_STYLE_BY_NAME[table.styleName]` and bakes the
  result into per-cell records. Changing the input source from
  "static catalog" to "static catalog rebased against source
  clrScheme" must NOT change:
  - The field shape (`bg`, `cl`, `bl`, `bd.{t,r,b,l}`).
  - The `addedFields` tagging (header still gets `['bg', 'bl',
    'bd.t', 'bd.b', 'bd.l']` for the top-left corner cell — the
    pin-down at line 236 is the sentinel).
  - The `existingCellStyles` precedence rule (hand-edited cell
    colours like the `#F4B183` border on Aptos A2 still win).
  - The synth-styles sidecar resource format.
  Implementation is free to call into an internal helper that
  takes `(table, sourceClrScheme)` and returns a
  `Record<string,string>` mapping (header bg, header fg,
  bandedEvenBg, etc.) — but the consumer of those values must
  stay the existing overlay function.

- **The hardcoded catalog encodes accent INDEXES indirectly via
  pre-computed RGBs.** `TableStyleMedium2` is accent1, Medium3 is
  accent2, **Medium4 is accent3**, … (every 7 entries shifts to
  the next accent). To rebase against a non-Aptos scheme you need
  a table mapping each style name → which accent (1..6) it uses
  PLUS a way to apply the catalog's pre-computed tint relative
  to that accent. Two valid approaches:
  1. **Per-style accent index + tint amount.** Precompute the
     tint amount per style per band slot (header, evenRow,
     oddRow, totals, border) and store that in a parallel
     catalog. At import time, look up the source clrScheme's
     accent and `applyTint(sourceAccent, tintAmount)` to get
     the final RGB.
  2. **Accent channel swap.** Keep the existing catalog. On
     import detect a non-Aptos clrScheme and remap each accent
     in the catalog result by HSL channel swap from the Aptos
     accent → source accent.
  Either works; the simpler one wins. **Don't symptom-patch by
  adding a Classic-Medium4 catalog entry alongside the Aptos
  one** — Excel ships ~thousands of accent permutations and the
  catalog can't enumerate them. The fix MUST be data-driven from
  the source clrScheme.

- **The synth-styles sidecar pin-down expects header tags to
  include `bg` AND `bl` AND 3 borders.** That contract is shape,
  not colour — but verify it stays green after the rebase. A
  rebase that accidentally drops `bl` (header-bold) or one of
  the borders breaks line 236.

- **The Aptos fixture has hand-edited cell-level colours on top
  of table-style synthesis.** Cell A2 carries a hand-applied
  `#F4B183` border (the canonical "user-added cell border" case
  per the fixture README). Synthesis must NOT overwrite
  hand-edits — the existing `existingCellStyles` map and
  `addedFields` are how M12 handles this. A theme-aware rebase
  must preserve that precedence. Add a Jest assertion if a
  bug-prone code path was touched.

- **The Classic fixture's `ProductCatalog` table also uses
  `TableStyleMedium4`** — same style name, different accent.
  That's the whole point of the test: ONE catalog entry must
  produce TWO distinct rendered colours depending on the source
  workbook's clrScheme.

- **The PGE harness has only screenshotted ONE note per cycle so
  far.** M13/E is the first feature requiring TWO independent
  screenshots in one session. The plan uses the variant-suffix
  approach (see "Harness extension chosen for this cycle"
  above). The generator MUST add suffixed entries to BOTH
  `REGION_BY_FEATURE` and `TITLE_PREFIX_BY_FEATURE` and MUST NOT
  rename or remove the existing M13/C and M13/D entries — those
  are evidence-bearing for prior cycles' rows in `PROGRESS.md`
  `## Done`.

- **`tableHeaderRowRegion` y-band.** Univer's column header
  consumes ~y=0–18; row 1 of the data area starts immediately
  after that and is ~19px tall at default zoom. Header band is
  approximately y=19–37 (mirror of the M13/D y-band note for A1).
  Use a generous slab (e.g. y=18–40, h≈22) to absorb default
  row-height variation — but **carve out the active-cell
  selection border** (`rgb(44,83,241)`, 1–2 px around the
  selected cell). On a freshly imported note the active cell is
  A1 by default, which lands inside the header band; the
  selection border could contribute blue pixels that fail the
  `greenInk` Aptos gate or the `greyInk` Classic gate. Either
  move the selection off A1 before sampling (programmatic
  click on a far-down cell via Univer's API), OR sample a
  header column away from A1 (e.g. col B / col C / col D —
  ProjectTracker's "Spent" column is roughly mid-table). The
  generator picks one.

- **Don't symptom-patch the Jest test on a regression
  post-rebase.** If a Jest assertion regresses after restoring
  theme-aware synthesis, run `git diff package-lock.json` first.
  exceljs's `Workbook.getTheme()` / clrScheme accessors changed
  between 3.x and 4.x. We are intentionally on 4.4.0; do NOT
  edit a test or downgrade exceljs to make a symptom go away —
  fix the dependency drift instead. Reference:
  `feedback_dependency_hygiene.md` in operator memory.

- **The feature touches `src/xlsx.ts`, the same file as every
  prior milestone's import/export.** Run the FULL Jest suite,
  not just the new pin-downs, before flipping the row. M12,
  M13/C, and M13/D round-trip tests must stay green.

- **Window prep is mandatory before evaluator screenshots.**
  `scripts/pge/prep-joplin-window.sh` (added in M13/C) fills the
  Joplin window, hides the sidebar and note list, and closes
  DevTools. A small / panes-up window crops the Univer canvas
  horizontally, which can make the right edge of a wide table
  partially offscreen and produce false-negative pixel-sidecar
  readings. The harness wires this in ahead of
  `eval-screenshot.js` automatically; do NOT bypass it.

- **Pre-existing typecheck fix on
  `tests/exportTableRoundTrip.test.ts` must stay.** Smoke session
  changed `'dashed'` → `'mediumDashed'` on line 334 (with
  matching assertion on line 349) to satisfy exceljs 4's
  `BorderStyle` enum; reverting it would break `npm run dist`.
  Unrelated to theme-aware banding but blocks the .jpl build.

- **Resource-name spelling.** The synth-styles sidecar resource
  name is `SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN` and the clrScheme
  resource name is `SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN`.
  Don't mistype them — the pin-downs at lines 226 and 239 do an
  exact `r.name ===` lookup.

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file (and the prior M13/C and M13/D BUILD_PLAN.md
preserved in git history at `dc80505` and `420d583`) as the
template. Each feature gets:

- `### feature-N-<kebab-id>` heading (stable; never renamed once
  written).
- **Spec**: one paragraph naming the user-observable change.
- **Acceptance criteria**: numbered list of observable evidence
  (visual, pixel-sidecar, Jest, runtime). NO data-shape-only
  assertions; NO "code does X" — only outcomes the evaluator can
  inspect from a fresh context.
- **Out of scope**: explicit non-goals so the evaluator does not
  penalise for them.
- **Suggested fixture(s)**: which file(s) under
  `tests/ExcelBaseTestData/formatting-testdata/` exercise this.
- **Related risks**: regression hot-spots and prior-bug pointers,
  including pointers into prior PR / commit hashes when prior
  work is the starting point.

The lowest-numbered feature with `passes: false` in
`test-results.json` is what the generator works on next.
