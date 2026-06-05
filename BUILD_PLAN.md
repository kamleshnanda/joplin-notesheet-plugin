# Notesheet PGE — build plan (M15: Conditional formatting — full round-trip on exceljs)

> **Cycle context.** This is the FOURTH real-feature cycle through the
> PGE harness, and the FIRST post-M14 milestone. M13/A–E shipped via
> this loop (PRs #19, #20, #21, #22, #23) and M14 was an explicit
> NO-GO documented at `docs/m14-sheetjs-spike.md` and merged via
> PR #25 (commit `f423d5a`). Harness scripts under `scripts/pge/`
> (`launch-joplin.sh`, `install-plugin.sh`, `import-fixture.{ts,sh}`,
> `eval-screenshot.{js,sh}`, `prep-joplin-window.sh`,
> `create-seeded-notesheet.js`) are present and proven. The pixel
> sidecar (`<screenshot>.pixels.json`) plus `REGION_BY_FEATURE` /
> `TITLE_PREFIX_BY_FEATURE` lookup in `eval-screenshot.js` are the
> machine-checkable signals alongside the visual screenshot.
>
> **The two-layer fidelity test pattern is now the project default.**
> M13/E established it the hard way (rework #2 / #3, PR #22): the
> snapshot-data fidelity test (`tests/excelReferenceFidelity.test.ts`)
> catches import-side bugs by anchoring expected RGBs to the
> operator-captured Excel reference PNG; the canvas-vs-Excel fidelity
> test (`tests/excelCanvasFidelity.test.ts`) catches render-side bugs
> by sampling the Joplin canvas screenshot vs the same reference PNG.
> M15 must extend BOTH layers — five tests per layer, one per CF
> column. See "Test gap warning" in Risks.
>
> **Single-feature cycle, big feature.** The five CF rule types
> (colorScale, dataBar, cellIs, top10, iconSet) ship together in one
> PR per the operator's ask — they share the import/export
> translation pipeline and the preset-wiring step. Decomposing into
> five sequential features would force five round-trips through the
> harness for what is structurally one piece of work.
>
> **Single Excel reference screenshot, multiple regions.** Per the
> operator ask: option (a) — one
> `screenshots/excel-reference/ConditionalFormatting-Variants.png`
> capture, with five `cfColumnRegion('A'|'C'|'E'|'G'|'I')` helpers
> sampling the five CF column bands from the same image. Cleaner than
> a `:variant` per column.

## Operator ask

(See `OPERATOR_ASK.md`.) Notesheet's import path currently drops all
conditional-formatting rules silently. The KNOWN SHORTCOMING test at
`tests/m12FixtureRoundTrip.test.ts:206` pins the current behaviour:
imported snapshots have no CF resource, exported workbooks have zero
`<conditionalFormatting>` blocks. The M14 SheetJS spike removed the
"wait for SheetJS to make CF cheaper" rationale —
`xlsx-js-style`'s `!cf` is also undefined on indexed-cellXf import,
so M15 ships on the existing `exceljs` parser. No parser blocker.

The fixture
`tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
is already in the repo and contains all five rule types (verified):
column A `colorScale`, column C `dataBar`, column E `cellIs > 50`,
column G `top10` (top-3 rank), column I `iconSet 3Arrows`.

## Harness extensions chosen for this cycle

The operator ask raises one harness extension; this plan picks
**option (a) — single screenshot, multiple region helpers** for these
reasons:

1. **All 5 rules visible in one shot.** The fixture authors A2..I11
   in a single sheet area; one Univer canvas screenshot at default
   zoom captures every CF column. No need for per-column scrolling
   or :variant suffixes.
2. **Cheaper to grade.** The evaluator captures ONE eval screenshot
   and ONE pixel sidecar; the sidecar carries five aggregates
   (`redInk`, `greenInk`, `blueInk`, `pinkInk`, `lightGreenInk`)
   each gated against a different `cfColumnRegion`.
3. **Reusable.** Future single-fixture multi-region features (e.g.
   M16 if it covers something with multiple visible bands) can adopt
   the same `cfColumnRegion(col)` shape.
4. **Backward compat.** No existing `REGION_BY_FEATURE` /
   `TITLE_PREFIX_BY_FEATURE` entries change. Adding
   `feature-1-m15-conditional-formatting` keys is purely additive.

The generator implements this by:

- Adding ONE entry to both lookup tables in
  `scripts/pge/eval-screenshot.js`:
  - `TITLE_PREFIX_BY_FEATURE['feature-1-m15-conditional-formatting']
    = 'PGE M15 CF eval '`.
  - `REGION_BY_FEATURE['feature-1-m15-conditional-formatting']
    = 'cfAllColumns'` — a composite region kind that triggers the
    helper to emit five sub-region samples in the sidecar.
- Adding `cfColumnRegion(col, dprScale)` to `eval-screenshot.js`,
  similar in shape to `tableHeaderRowRegion`. The `col` argument is
  the Excel column letter ('A', 'C', 'E', 'G', 'I'); the helper
  returns `{x, y, w, h}` covering the CF data band (rows 2–11) for
  that column. DPR-scale-aware per the M13/E lesson.
- Extending `samplePixelsAt` to write a per-column nested object
  under a top-level `cfColumns` key in the sidecar:
  ```
  cfColumns: {
    A: { redInk, greenInk, ...top },
    C: { blueInk, ...top },
    E: { pinkInk, ...top },
    G: { lightGreenInk, ...top },
    I: { redInk, greenInk, yellowInk, ...top },
  }
  ```
- Adding two new aggregate counters to `samplePixelsAt`:
  - `pinkInk`: `R≥220 AND G∈[180,220] AND B∈[180,220]` (matches
    `#FFC7CE` family).
  - `lightGreenInk`: `R∈[180,220] AND G≥220 AND B∈[180,220]`
    (matches `#C6EFCE` family).
  Reuse the existing `redInk`, `greenInk`, `blueInk` aggregates
  (defined for M13/D and M13/E). Re-tightening them is out of scope.
- A `yellowInk` aggregate: `R≥200 AND G≥200 AND B≤120` (icon-set
  middle band; loose because glyphs are small).

The generator MUST NOT regress prior cycles' single-key paths
(`feature-1-m13-rotated-text-renders`,
`feature-1-m13-rich-text-renders`,
`feature-1-m13-theme-aware-banding:aptos`/`:classic`) — those entries
are evidence-bearing for `## Done` rows.

## Excel reference screenshot — capture timing

`screenshots/excel-reference/ConditionalFormatting-Variants.png` does
NOT exist at planner time. The operator captures it from real Excel
as part of THIS cycle, same pattern as M13/E with the Aptos and
Classic Smorgasboard references. Generator should NOT block on it:

- Phase 1 work (preset wire, exceljs↔Univer translation, round-trip
  test, snapshot-fidelity test scaffolded with TODO assertions) can
  proceed without the reference screenshot.
- Phase 2 work (`excelCanvasFidelity.test.ts` extension and the
  evaluator's pixel-sidecar gating) NEEDS the reference. Generator
  captures own evidence first; if the reference screenshot lands
  before Phase 2 begins, fold it in. If it lands after, the
  evaluator runs the canvas-fidelity tests against the reference
  once it's present.
- Generator-evidence screenshots under
  `screenshots/feature-1-m15-conditional-formatting/` are the
  pre-evaluator due-diligence; they do NOT substitute for the Excel
  reference. The reference is ground truth for snapshot-fidelity
  RGBs and canvas-fidelity colour parity.

## Features

### feature-1-m15-conditional-formatting

**Spec**

When a user imports
`tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
into a Notesheet note, the workbook's five conditional-formatting
rules — colour scale on column A, data bar on column C, cellIs > 50
fill on column E, top-3 rank fill on column G, 3-arrow iconSet on
column I — render correctly on the Univer canvas in Joplin's editor.
Users can add, edit, and delete CF rules from the Univer "Manage
Conditional Format Rules" panel (shipped by
`@univerjs/preset-sheets-conditional-formatting`). Exporting the note
back to .xlsx via the editor's Export button produces a workbook that
re-opens cleanly in real Excel and re-imports into Notesheet with all
five rules preserved structurally (type, ref, cfvo, colour, operator,
rank where applicable).

The KNOWN SHORTCOMING test at
`tests/m12FixtureRoundTrip.test.ts:206` flips from a negative
("zero `<conditionalFormatting>` blocks", "no CF resource") to a
positive pin-down ("five rules survive round-trip with the source
ref/type/colours intact").

The five-rule scope is explicit. Out-of-scope rule types
(text-based, time-period, unique/duplicate, formula-based,
stopIfTrue, CF inside merged cells, custom dxf records beyond
fill-bg) remain unsupported but documented as M15 follow-up.

**Acceptance criteria**

The evaluator must verify ALL of the following from a fresh context:

1. **Visual — all 5 CF rule types render correctly in the Joplin
   canvas.** A single evaluator-captured screenshot under
   `screenshots/feature-1-m15-conditional-formatting/`, taken via
   `bash scripts/pge/eval-screenshot.sh feature-1-m15-conditional-formatting`
   after `import-fixture.sh
   ConditionalFormatting-Variants.xlsx` and `prep-joplin-window.sh`
   has filled the window and hidden the side panes. The screenshot
   shows the Univer main canvas
   (`canvas[id^="univer-sheet-main-canvas"]`) inside the editor's
   `UserWebviewIndex.html` frame. The evaluator's verdict text
   MUST explicitly call out:
   - **Column A (colorScale)**: cells A2..A11 carry distinct
     background fills along a red→yellow→green gradient. The lowest-
     value cell renders red-ish (`#F8696B` family — i.e. R≥200, G≤140,
     B≤140); the median cell renders yellow-ish (`#FFEB84` family —
     R≥200, G≥200, B≤180); the highest-value cell renders green-ish
     (`#63BE7B` family — R≤180, G≥180, B≤180). The verdict explicitly
     names the gradient direction (low→high) and the three named hue
     buckets.
   - **Column C (dataBar)**: each of C2..C11 has a horizontal bar in
     blue (`#638EC6` family — R≤140, G≤180, B≥180) whose width scales
     with the cell value. The smallest-value bar is visibly shorter
     than the largest; at least one cell shows a bar covering more
     than half the cell width and at least one shows a bar shorter
     than a quarter.
   - **Column E (cellIs > 50)**: cells whose numeric value is `> 50`
     carry a pink/light-red fill (`#FFC7CE` family — R≥220,
     G∈[180,220], B∈[180,220]). Cells whose value is `<= 50` are
     unfilled. The verdict explicitly names at least one filled cell
     and at least one unfilled cell.
   - **Column G (top-3 rank)**: exactly THREE cells in G2..G11 carry
     a light-green fill (`#C6EFCE` family — R∈[180,220], G≥220,
     B∈[180,220]). The other seven are unfilled. The verdict
     explicitly states "exactly three".
   - **Column I (iconSet 3-Arrows)**: each cell I2..I11 displays a
     small arrow glyph — red-down for low-band (0–33%), yellow-flat
     for mid-band (33–67%), green-up for high-band (67–100%). The
     verdict explicitly names at least one of each colour glyph
     observed in the screenshot.

2. **Pixel sidecar — colour signal lands in the histograms per CF
   column.** The
   `screenshots/feature-1-m15-conditional-formatting/eval-*.pixels.json`
   sidecar contains a `cfColumns` object with FIVE per-column
   sub-objects, each with the right ink-aggregate signal sampled from
   `cfColumnRegion(col)`:
   - `cfColumns.A`: `redInk ≥ 30` AND `greenInk ≥ 30` (gradient
     endpoints both visible in the band).
   - `cfColumns.C`: `blueInk ≥ 30` (data bars).
   - `cfColumns.E`: `pinkInk ≥ 20` (some > 50 cells render pink).
   - `cfColumns.G`: `lightGreenInk ≥ 20` (top-3 cells render
     light-green).
   - `cfColumns.I`: at least one of `redInk`, `greenInk`,
     `yellowInk` is `≥ 5` (icon glyphs are small; the threshold is
     intentionally loose).

   The harness extension adds `cfColumnRegion(col, dprScale)` plus
   `pinkInk`, `lightGreenInk`, `yellowInk` aggregates to
   `samplePixelsAt` in `scripts/pge/eval-screenshot.js`. M13/E's
   `tableHeaderRowRegion` plus `greyInk` aggregate is the template.
   DPR scaling is mandatory per the M13/E Phase 4 fix.

3. **Snapshot fidelity — exceljs CF data → Univer CF resource shape.**
   New test block in `tests/excelReferenceFidelity.test.ts` adds 5
   tests that import
   `ConditionalFormatting-Variants.xlsx` via
   `xlsxBufferToSnapshot()` and assert the resulting snapshot's
   `resources` array contains an entry whose `name` matches the
   Univer CF plugin resource key (`SHEET_CONDITIONAL_FORMATTING_PLUGIN`,
   re-exported from `@univerjs/sheets-conditional-formatting` — DON'T
   hard-code the string; import the constant). The data inside MUST
   carry 5 rules whose translated shape matches the source XML
   structurally:
   - Rule 0 — `colorScale`, ref `A2:A11`, 3 cfvo (min /
     percentile=50 / max), 3 colours `#F8696B` / `#FFEB84` / `#63BE7B`.
   - Rule 1 — `dataBar`, ref `C2:C11`, 2 cfvo (min / max), colour
     `#638EC6`.
   - Rule 2 — `highlightCell` (subType `number`), ref `E2:E11`,
     operator `greaterThan`, formula `50`, fill bg `#FFC7CE`.
   - Rule 3 — `highlightCell` (subType `rank`), ref `G2:G11`,
     `rank: 3`, `isBottom: false`, fill bg `#C6EFCE`.
   - Rule 4 — `iconSet`, ref `I2:I11`, iconSet `3Arrows`, 3 cfvo
     (percent 0 / 33 / 67).

   Each test reads the source `xl/worksheets/sheet1.xml` directly
   (via the in-test fixture-XML reader pattern already in
   `excelReferenceFidelity.test.ts`) so expected RGBs and refs are
   sourced from the file, NOT from our parser's emit. This is the
   import-side gate. Per `feedback_pge_fidelity_test_gap.md`,
   pin-downs anchored to our own emit are no test at all — the
   reference values come from upstream.

4. **Canvas fidelity — Joplin canvas vs Excel reference.** New
   tests in `tests/excelCanvasFidelity.test.ts` (one per CF column,
   five total) sample the dominant fill colour of each column band in
   BOTH the operator-captured Excel reference screenshot
   (`screenshots/excel-reference/ConditionalFormatting-Variants.png`)
   AND the latest Joplin canvas eval screenshot
   (`screenshots/feature-1-m15-conditional-formatting/eval-*.png` by
   mtime). The five tests assert per-channel parity within Δ ≤ 16
   for the column's dominant CF colour:
   - Column A — pick the median row's dominant fill (yellow-ish
     `#FFEB84` family — middle of the gradient; most stable
     measurement).
   - Column C — sample the right edge of a long-bar cell and the
     centre of a short-bar cell separately; assert blue-bar regions
     in both PNGs land within Δ ≤ 16 of each other.
   - Column E — sample a known-`>50` cell (say E2, value 80 per the
     fixture); assert pink fill parity.
   - Column G — sample one of the top-3 cells (the highest-value
     cell in G2..G11); assert light-green fill parity.
   - Column I — sample the icon-glyph centroid for one row in each
     band (low / mid / high); assert at least one of (red / yellow /
     green) ink hits in the corresponding band parity-checks against
     the same band in the reference within Δ ≤ 24 (icon glyphs
     anti-alias more aggressively than fills; looser tolerance per
     the M13/E "1-2px strip" lesson).

   This is the render-side gate. M13/E's lesson stands: the
   snapshot-data fidelity test alone is not sufficient — the
   render-vs-render test catches Univer-renderer bugs the
   snapshot-data test can't see. **The canvas-fidelity test is
   skipped (Jest `test.skip`) when the reference PNG is absent and
   re-enabled once the operator captures it.** The skip path keeps
   the suite green while the reference is being produced; the
   evaluator MUST verify the test is unskipped and passing before
   PASS.

5. **Round-trip — five CF rules survive export → re-import.** New
   test (or extend `tests/m12FixtureRoundTrip.test.ts`):
   - Import `ConditionalFormatting-Variants.xlsx` →
     snapshot.
   - Call `snapshotToXlsxBuffer(snapshot)` → buffer.
   - Re-load buffer with a fresh `new ExcelJS.Workbook()`.
   - Assert all FIVE source rules survive: each rule's `type`,
     `ref` (the cell-range string the worksheet returns), `cfvo`
     array (length, types, values), and colour (`#RRGGBB` /
     `dxf.fill.bgColor.argb`) match the source within
     argbToHex / hexToArgb symmetry. For cellIs (rule 2): `operator`
     and `formulae[0]` survive. For top10 (rule 3): `rank` and
     `bottom` survive.
   - The KNOWN SHORTCOMING test at
     `tests/m12FixtureRoundTrip.test.ts:206` is FLIPPED to a
     positive pin-down whose body now asserts the round-trip
     preserves the rules. Old assertion bodies removed; new ones
     reference the same source XML extraction.

6. **Univer CF preset is registered in `src/editorView.tsx`.**
   The editor mounts the CF preset alongside the existing six
   presets (Core, Sort, Filter, Table, Drawing, HyperLink). New
   imports in `src/editorView.tsx`:
   - `import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';`
   - `import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';`
   - `import sheetsCfEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';`
   The preset is added to the `presets` array passed to
   `createUniver()` (after `UniverSheetsHyperLinkPreset()`), and the
   locale is merged into the `locales[LocaleType.EN_US]` object.
   The CF preset must also be added to `package.json` as a direct
   dependency (it's currently transitively pinned at 0.23.0 in the
   lockfile via `@univerjs/presets`'s peer graph; making it a direct
   dep ensures the editor wiring is reproducible).

   Verification: the editor's "Manage Rules" panel button is
   reachable (whether by toolbar icon or menu) — the evaluator
   verdict mentions the panel rendering in the canvas screenshot or
   indicates it's accessible via a known click path. The visual gate
   in criterion (1) is the load-bearing one; this criterion is the
   MEANS. Without preset wiring, criterion (1) would fail with
   "snapshot data correct, canvas blank" — exactly M13's failure
   mode.

7. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** Generator captures own evidence under
   `screenshots/feature-1-m15-conditional-formatting/`
   (`generator-evidence.png` + `.pixels.json` sidecar), reads via
   the Read tool to satisfy the `verify-gate` hook.

8. **`npm test` — all existing tests stay green plus new ones.**
   Baseline at session start (verify with `npm test`): **209**. M15
   adds:
   - 5 new tests in `excelReferenceFidelity.test.ts` (snapshot
     fidelity per CF column).
   - 5 new tests in `excelCanvasFidelity.test.ts` (canvas fidelity
     per CF column).
   - 1 new round-trip test (or the flipped `m12FixtureRoundTrip`
     KNOWN SHORTCOMING).
   - Existing KNOWN SHORTCOMING at
     `tests/m12FixtureRoundTrip.test.ts:206` flipped (no net change
     in test count for that one — same `test()` block, body
     rewritten).
   Target: **209 → ≥ 220**. The full Jest suite is run before
   flipping the row, NOT just the new tests. M15 touches `src/xlsx.ts`
   (the import + export pipeline shared with every prior milestone)
   and `src/editorView.tsx` (preset wiring); borders, hyperlinks,
   table styles, chart export, alignment, rotation (M13/C), rich
   text (M13/D), theme banding (M13/E) all share the import/export
   pipeline. Nothing here should regress.

**Out of scope**

- **Text-based CF rules** (`beginsWith`, `endsWith`, `containsText`,
  `notContainsText`). Not in the fixture. Document as M15 follow-up.
- **Time-period CF rules** (`thisMonth`, `last7Days`, `today`,
  `yesterday`, `tomorrow`, `lastWeek`, `thisWeek`, `nextWeek`,
  `lastMonth`, `nextMonth`). Same.
- **Unique-values / duplicate-values rules.** Same.
- **Formula-based CF rules** (`type="expression"`). Same.
- **Stop-If-True semantics.** OOXML CF rules can carry `stopIfTrue` —
  if a higher-priority rule matches, lower-priority ones don't apply.
  M15 always applies all rules; document the gap.
- **CF inside merged cells.** The fixture doesn't exercise it.
- **Custom dxf records beyond fill-bg.** The fixture uses fill-bg
  only. If exceljs surfaces a font-colour or border on a
  cellIs/top10 dxf, preserve them; otherwise out of scope for M15.
- **Univer "manage rules" panel UI customisations.** Use the panel
  exactly as Univer ships it — no theming, no relabelling. The
  preset's CSS import (`/lib/index.css`) is the ONLY styling
  surface M15 touches.
- **Editing CF via Joplin's note-import path.** The CF UI panel is
  the editor surface; users do NOT edit CF via a markdown table or
  fenced-code modification. Editor-only editing.
- **README "Known shortcomings — Conditional formatting" cleanup.**
  Punted to a follow-up docs PR (precedent: M13/C, M13/D, M13/E
  punted theirs to PRs #21 and #23). The evaluator does NOT
  penalise if the docs edit is deferred.
- **iconSet variants beyond `3Arrows`.** The fixture uses only
  `3Arrows`. Other iconSets (`3TrafficLights1`, `4RedToBlack`,
  `5Arrows`, etc.) probably work via the same translation but are
  not under test in M15.
- **CF preserving theme-tinted colours via the source clrScheme.**
  M15 treats CF colours as literal `argb` strings. Theme refs
  (`type="theme"` inside `<dxf>`) are out of scope; document if
  encountered in the wild.

**Suggested fixtures**

- `tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
  — already in the repo. 5 rules, one per column (A/C/E/G/I). This
  is the spec's central fixture. Verified contents:
  - colorScale on `A2:A11`, three cfvo (min / percentile=50 / max),
    three colours #F8696B / #FFEB84 / #63BE7B.
  - dataBar on `C2:C11`, two cfvo (min/max), colour #638EC6.
  - cellIs `>50` on `E2:E11`, dxf fill bg #FFC7CE.
  - top10 (rank=3, bottom=false) on `G2:G11`, dxf fill bg #C6EFCE.
  - iconSet 3Arrows on `I2:I11`, three cfvo (percent 0 / 33 / 67).
- `screenshots/excel-reference/ConditionalFormatting-Variants.png`
  — operator-captured reference (does NOT exist at planner time).
  Used as ground truth by the canvas-fidelity tests once captured.
  Generator does NOT block on this; canvas-fidelity tests skip when
  the file is absent and unskip when present.

The harness invokes
`scripts/pge/import-fixture.sh ConditionalFormatting-Variants.xlsx`
to land the fixture into a Joplin note headlessly, then
`eval-screenshot.sh feature-1-m15-conditional-formatting` for the
canvas-targeted PNG and pixel sidecar.

**Related risks**

- **CF preset wiring is the single point of failure.** Without
  registering `UniverSheetsConditionalFormattingPreset` in
  `createUniver()`, the snapshot's CF data lands but Univer's
  renderer paints nothing — the M13 failure mode (snapshot data
  correct, canvas blank, Jest passes on snapshot fidelity, runtime
  visual gate fails). The first thing the generator does after
  exceljs↔Univer translation is wire the preset; do NOT
  defer that to "I'll add it after I see the data lands."

- **`SHEET_CONDITIONAL_FORMATTING_PLUGIN` is the resource key**
  Univer expects on the snapshot. It's a string constant exported
  from `@univerjs/sheets-conditional-formatting`. **Reuse the
  export, do NOT hard-code the string** — if Univer ever renames it,
  a hard-coded string in `src/xlsx.ts` becomes a silent bug. The
  pin-down test in `excelReferenceFidelity.test.ts` likewise reads
  the constant from the package, not a string literal.

- **exceljs's CF data shape doesn't match Univer's 1:1.**
  Translation table (build a per-type translator):
  - exceljs `type: 'colorScale'` → Univer `type: 'colorScale'` (cfvo
    + colour arrays map directly). exceljs surfaces colours as
    `argb: 'FFRRGGBB'`; Univer expects `#RRGGBB`. Strip leading
    `FF` via `argbToHex`. cfvo `type` values (`min`, `percentile`,
    `max`) carry over verbatim.
  - exceljs `type: 'dataBar'` → Univer `type: 'dataBar'`. Single
    colour, two cfvo. Direction (`gradient`, `solidFill`,
    `negativeColor`) — exceljs may not surface these; ship defaults
    matching Univer's preset defaults if absent.
  - exceljs `type: 'cellIs'` → Univer `type: 'highlightCell'`,
    `subType: 'number'`. `operator` (`greaterThan`,
    `lessThan`, etc.) carries over. `formulae[0]` becomes Univer's
    `formula` field. dxf fill bg becomes Univer's `style.bg.rgb`.
  - exceljs `type: 'top10'` → Univer `type: 'highlightCell'`,
    `subType: 'rank'`. `rank` value carries over. exceljs's
    `bottom: true` (bottom-N) maps to Univer's `isBottom: true`.
  - exceljs `type: 'iconSet'` → Univer `type: 'iconSet'`. iconSet
    name carries over verbatim (`3Arrows`, `3TrafficLights1`,
    `5Arrows`). cfvo bands carry over.

- **Round-trip ARGB drift.** exceljs surfaces colours as `argb`
  strings (`FF638EC6` — leading `FF` is alpha). Univer expects
  `#RRGGBB` (no alpha). Both sides go through `argbToHex` /
  `hexToArgb` helpers (already exist in `src/xlsx.ts` from M12). Do
  NOT drop the alpha on import then accidentally reintroduce it on
  export — the round-trip test in criterion (5) is the regression
  sentinel.

- **CF-on-top-of-table-style precedence.** M12's
  `synthesizeTableStyleAssignments` (`src/xlsx.ts:943`) bakes
  per-cell `bg`/`cl` from the table style. CF rules are applied by
  Univer's renderer ON TOP of the cell style. **Confirm Univer
  paints CF over the synthesized table-style fill, not under it.**
  If it paints under, the table-style's fill will mask the CF
  colour and criterion (1) fails. Mitigation if needed: mark
  CF-bound cells in the synth-styles sidecar
  (`SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN`) so the M12 synthesizer
  skips painting per-cell `bg` on those cells. The current fixture
  doesn't ship a `tableStyle` so this risk doesn't bite the gate
  shot — but the M12-synthesized fixtures (Aptos, Classic) and the
  M15 fixture share `src/xlsx.ts` parsing, so a bad change to one
  could regress the others.

- **iconSet glyph rendering.** Univer's CF engine ships icon glyphs
  for the standard iconSet names (`3Arrows`, `3TrafficLights1`,
  `5Arrows`, etc.). Confirm `3Arrows` renders without additional
  asset bundling. If glyphs render as missing characters
  (□ / ?), that's a Univer-side icon-font issue — investigate
  whether the preset's CSS includes the icon font or whether we
  need an extra import. The CSS import
  (`@univerjs/preset-sheets-conditional-formatting/lib/index.css`)
  should pull the font; if it doesn't, the preset's
  `package.json` `style` / `exports` field will tell.

- **Test gap warning** (`feedback_pge_fidelity_test_gap.md`).
  Pin-downs anchored to our own emit are no test at all. The
  reference-fidelity test's expected RGBs come from the source
  `xl/worksheets/sheet1.xml`, NOT from the parser output. The
  canvas-fidelity test's expected RGBs come from the operator-
  captured Excel reference screenshot, NOT from our render. Both
  layers' assertions are anchored UPSTREAM of our code. M13/E
  rework #2 / #3 is the precedent — those tests landed BEFORE the
  fix (failing first, then passing) to prove the gap was real. The
  M15 generator should follow the same ordering: write the
  reference-fidelity test first against the source XML, watch it
  fail (no CF resource yet emitted), then make it pass.

- **Don't symptom-patch the test on a Jest failure post-rebuild.**
  If a test regresses after wiring the CF preset, run
  `git diff package-lock.json` first. The CF preset may have pulled
  new transitives. Reference: `feedback_dependency_hygiene.md`.
  Do NOT downgrade exceljs or any other dep to make a symptom go
  away — fix the dependency drift instead.

- **The feature touches `src/xlsx.ts` (parser's import + export
  paths) and `src/editorView.tsx` (preset wiring) — the same files
  every prior milestone touches.** Run the FULL Jest suite, not
  just the new tests. M9 borders, M12 hyperlinks + table styles,
  M10 chart export, M12 alignment, M13/C rotation, M13/D rich text,
  M13/E theme banding all share the import/export pipeline. Nothing
  here should regress. Specifically, verify the M13/E pin-downs
  in `m12FixturePinDowns.test.ts` (Aptos + Classic header colours,
  totals top + bottom borders) stay green — those are sensitive to
  any synthesizer change, and a CF-precedence fix could touch the
  synthesizer.

- **The Excel reference screenshot capture is operator-side.** The
  generator must NOT block on it. Implement the canvas-fidelity
  tests with `test.skip` semantics that detect file absence and
  unskip when present (the M13/E pattern: read latest by mtime,
  fail clearly if missing). Generator captures own evidence
  (screenshot + sidecar) regardless; the verify-gate hook unlocks
  the `test-results.json` write after the generator-evidence PNG is
  Read.

- **Window prep is mandatory before evaluator screenshots.**
  `scripts/pge/prep-joplin-window.sh` (added in M13/C) fills the
  Joplin window, hides the sidebar and note list, and closes
  DevTools. A small / panes-up window crops the Univer canvas
  horizontally; column I (the rightmost CF column) is most
  vulnerable to being clipped offscreen. The harness wires this in
  ahead of `eval-screenshot.js` automatically; do NOT bypass it.

- **DPR scaling on `cfColumnRegion`.** M13/E's
  `tableHeaderRowRegion` was Retina-broken until Phase 4 (hardcoded
  y-offsets assumed DPR=1, broke on DPR=2 macOS). New region
  helpers MUST scale x/y/w/h by `canvas.width / canvas.clientWidth`
  at sample time. Use the M13/E pattern: write offsets in CSS px,
  multiply by the DPR ratio at sample time.

- **Univer column-letter strip occupies y≈0–18 at default zoom.** The
  CF data band starts at row 2 (= second visible row), so y starts
  ~y=18 + rowHeight = ~y=37. With ten data rows at 19px each, the
  full band ends ~y=37 + 10*19 = ~y=227. `cfColumnRegion` should
  cover `y=37..227, h=190` for a single column at default zoom (CSS
  px). Generator may need to tune empirically against the first
  generator-evidence screenshot.

- **Active-cell selection on A1 colours the column-A region with
  blue (`rgb(44,83,241)`).** On a freshly imported note Univer
  defaults the active selection to A1, which lands at the top-left
  of the colorScale band. The blue selection border could
  contribute false-positive `blueInk` hits to the column-A region —
  but column A is gated on `redInk` AND `greenInk`, not blueInk, so
  this doesn't bite the spec. Still, if `redInk < 30` on a first
  capture, programmatically click off A1 (Univer command bus via the
  webview frame) before re-capturing.

- **The flipped KNOWN SHORTCOMING test must keep its
  `test.skip`-or-`test()` marker semantics aligned with the suite.**
  `tests/m12FixtureRoundTrip.test.ts:206` is currently `test(...)`
  asserting NO CF resource. Flipping its body (asserting the
  preserved rules) keeps it as `test()`. Don't accidentally promote
  to `test.skip` or `test.todo` — the suite count needs the test.

- **Generator MUST add the CF preset to `package.json`
  dependencies** as `@univerjs/preset-sheets-conditional-formatting`
  at the version pinned in the lockfile (currently 0.23.0,
  consistent with the other Univer presets). After the install,
  `git diff package.json package-lock.json` and surface every
  version change — npm's transitive resolution decisions are not
  implicitly approved (`feedback_dependency_hygiene.md`). Major
  version changes in transitive deps must be flagged.

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file (and the prior M13/C, M13/D, and M13/E BUILD_PLAN.md
preserved in git history at `dc80505`, `420d583`, and `fca1cbc`) as
the template. Each feature gets:

- `### feature-N-<kebab-id>` heading (stable; never renamed once
  written).
- **Spec**: one paragraph naming the user-observable change.
- **Acceptance criteria**: numbered list of observable evidence
  (visual, pixel-sidecar, Jest, runtime). Two-layer fidelity tests
  (snapshot + canvas) anchored UPSTREAM of our code (source XML for
  snapshot, operator-captured Excel reference PNG for canvas) are
  the project default for "match Excel" features. NO data-shape-only
  assertions; NO "code does X" — only outcomes the evaluator can
  inspect from a fresh context.
- **Out of scope**: explicit non-goals so the evaluator does not
  penalise for them.
- **Suggested fixture(s)**: which file(s) under
  `tests/ExcelBaseTestData/formatting-testdata/` exercise this.
- **Related risks**: regression hot-spots and prior-bug pointers,
  including pointers into prior PR / commit hashes when prior work
  is the starting point.

The lowest-numbered feature with `passes: false` in
`test-results.json` is what the generator works on next.
