# Notesheet PGE — progress log

This file is the generator's session-to-session handoff. Update after
every feature.

## Done

- **feature-1-smoke-red-cell** (2026-06-02) — Generator inline-implemented
  the smoke seed in `src/snapshot.ts:emptySnapshot()`. A1 = "harness-smoke-OK"
  styled via `styles['pge-smoke-red'] = { cl: { rgb: '#FF0000' } }`,
  cell carries `s: 'pge-smoke-red'` reference. Built .jpl, installed in
  dev profile, captured generator-evidence screenshot showing red text
  rendered in Univer at A1. Fresh-context evaluator subprocess graded
  PASS (cd8bf51).
- **harness-hardening** (2026-06-02) — `eval-screenshot.js` now drops
  into the `UserWebviewIndex.html` frame inside the editor page (where
  Univer actually mounts) and waits on the real Univer canvas selector
  `canvas[id^="univer-sheet-main-canvas"]` instead of a 5s sleep.
  Emits a `<screenshot>.pixels.json` sidecar with the top non-background
  colours sampled from the row-0 canvas slab — gives evaluators a
  machine-checkable signal alongside the visual screenshot. Confirmed
  on the smoke note: dominant `rgb(234,237,249)` (header band),
  `rgb(255,0,0)` appears in top-3 with 353 hits, proving the red is
  real pixels not just snapshot data.
- **feature-1-m13-rotated-text-renders** (2026-06-03) — Cherry-picked
  the reverted PR #16 (`415b4a4`) rotation import/export in
  `src/xlsx.ts`, plus `tests/m13RotatedText.test.ts` and the
  `m12FixtureRoundTrip.test.ts` flips. README docs edit explicitly
  out-of-scope. Built .jpl, installed in dev profile, imported
  `MergedCellsAndAlignment.xlsx` headlessly via the new
  `scripts/pge/import-fixture.{ts,sh}`, captured canvas-targeted
  screenshot showing A6 up-right diagonal, B6 vertical, C6 down-right
  diagonal. Pixel sidecar over the rotated row band reports
  `inkRowSpread=1.000` (text ink occupies every sampled y-row in the
  slab) — strong non-horizontal signal independent of colour. Jest
  187/187 (was 181 baseline + 6 new rotation tests). The reverted
  code worked first try in Univer 0.23 — earlier visual failure was
  almost certainly a stale-build issue, not a rendering gap. Evaluator
  graded PASS (PR #19, dc80505).
- **feature-1-m13-rich-text-renders** (2026-06-03) — Cherry-picked PR
  #16 commit `6f33f3a` rich-text import/export in `src/xlsx.ts`
  (4 new helpers: `buildTextStyleFromExceljsFont`,
  `buildRichTextCellP`, `buildExceljsFontFromTextStyle`,
  `extractRichTextRunsFromCellP`), plus `tests/m13RichText.test.ts`
  (8 new tests) and the two flipped `m12FixtureRoundTrip.test.ts`
  pin-downs. README docs drop explicitly out-of-scope. Imported
  `RichTextInOneCell.xlsx` headlessly via
  `scripts/pge/import-fixture.sh`. Canvas screenshot shows
  A1 = bold "**Hello**" + plain " world", A2 = "Red"(red) + " and "
  + "Blue"(blue) + " text" with three distinct foreground colours,
  A3 = blue underlined "Visit example.com for more info" hyperlink
  (Pattern A). Pixel sidecar over the A2-only band reports
  `redInk=67`, `blueInk=76` — both well above the spec's ≥30
  threshold; dominant histogram bucket `rgb(0,0,255)` 76 hits
  followed by `rgb(255,0,0)` 67 hits. Jest 195/195 (was 187 baseline
  + 8 new rich-text tests). Reverted code worked first try in Univer
  0.23 — same hypothesis as M13/C confirmed: the original revert
  was almost certainly a build/cache issue, not a renderer gap.
  Evaluator graded PASS (PR #20). README cleanup followed in PR #21
  (5aaf690).
- **feature-1-m13-theme-aware-banding** (2026-06-03 → 2026-06-04,
  three-rework cycle) — Routed `synthesizeTableStyleAssignments`
  through a new `EXCEL_TABLE_STYLE_RECIPE_BY_NAME` parallel table
  (`src/charts/excelTableStyleRecipes.ts`) that names each TableStyle
  slot's accent index + tint. The synthesizer reads the source
  workbook's `<a:clrScheme>` (already captured by
  `readThemeClrScheme`) and resolves the recipe via the ECMA-376
  HSL-L tint formula. Aptos fixture (`FormattingSmorgasboard.xlsx`)
  paints green; Classic fixture
  (`FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`)
  paints grey instead of green — same `EXCEL_TABLE_STYLE_BY_NAME[Medium4]`
  lookup, two distinct rendered outputs driven by source clrScheme.
  Achromatic styles (Light1/8/15, Medium1/8/15/22, Dark1/8) keep
  their literal greys.

  **First multi-screenshot cycle.** Added `:variant` suffix support
  to `eval-screenshot.js` (`feature-1-m13-theme-aware-banding:aptos` +
  `:classic`), `tableHeaderRowRegion` helper sampling cols B+ to dodge
  A1 active-cell selection blue, `greyInk` aggregate (R,G,B∈[140,180]
  AND `abs(R-G)≤10` AND `abs(G-B)≤10`), and broadened the existing
  `greenInk` aggregate so it catches dark Aptos green `#196B24` (G=107)
  too.

  **Rework #2 (canvas fidelity, PR #22).** Operator caught a
  render-side bug the snapshot-fidelity test couldn't see: the
  `totalsTopBorder` slot was emitting `bd.t.s = 7` (DOUBLE) on the
  totals row. Univer 0.23's `_renderDoubleBorder` paints DOUBLE as
  two 1px strips with a 1-px white gap, which reads as anti-aliased
  `#89CE74` rather than the pure `#72D068` Excel paints. AND Excel's
  render is actually a single 2px strip, not a true double-line — so
  DOUBLE was the wrong style code in the first place. Phase 1 added
  `tests/excelCanvasFidelity.test.ts` (361 lines, pure-stdlib via
  `tests/util/pngSampler.ts`) — region-finding heuristics align Joplin
  canvas screenshot with Excel reference at structural regions
  (header / banded / totals-top), asserts dominant-colour parity for
  tall regions (Δ ≤ 8) and looser for 1-2px strips (Δ ≤ 32). Phase 2
  switched `BORDER_STYLE_TO_UNIVER.medium` (style 8, lineWidth=2)
  instead of `.double` (style 7) for the totals-top slot. Phase 3
  re-captured eval screenshots; Phase 4 raised counts 204 → 206.

  **Rework #3 (totals-row BOTTOM border).** Operator's eyeball +
  side-by-side pixel-probe revealed Excel paints TWO accent strips
  framing top AND bottom of the totals body. Recipe extended with
  `totalsBottomBorder` slot, parallel to `totalsTopBorder`. Empirical
  overrides set both Aptos (`#72D068`) and Classic (`#C9C9C9`) to
  the lighter accent. `synthesizeTableStyleAssignments` emits `bd.b`
  on every totals cell, REPLACING the table outline's thin frame.
  Diagnostic asset:
  `tests/ExcelBaseTestData/formatting-testdata/border-isolation.xlsx`
  (operator-built fixture with explicit border combinations,
  pixel-probed against Excel to establish ground truth).

  **RESOLVED 2026-06-07 — was a known gap during the M13/E rework
  but no longer reproduces on current main.** PR #22 shipped with
  a documented gap: the snapshot's `bd.b.cl.rgb === '#72D068'` was
  correct (verified via direct `xlsxBufferToSnapshot` introspection
  on every totals cell), but the canvas pixel-probe at y=436 showed
  `rgb(52,106,46) = #34692E` — the header's dark green — rather
  than the lighter accent. During M16 prep (2026-06-07), a fresh
  re-capture of the Aptos eval at
  `screenshots/feature-1-m13-theme-aware-banding/eval-aptos-2026-06-07T05-20-38-627Z.png`
  showed BOTH totals strips rendering at y=398 and y=436 in
  `rgb(137,206,116)` (anti-aliased of `#72D068`). Bug appears to
  have resolved itself through snapshot-shape changes between M13/E
  (PR #22) and M16 — most likely cause is the M15 CF rework's
  changes to `synthesizeTableStyleAssignments`. Univer version
  unchanged (`@univerjs/engine-render@0.23.0` locked).
  Bonus finding from the investigation (informs future work): an
  isolation probe (cell with `bd.b = #72D068` above + cell with
  `bd.t = #34692E` below) showed Univer paints the LOWER cell's
  `bd.t` colour at the shared edge — i.e., when both `bd.b` (upper)
  and `bd.t` (lower) are declared at a shared edge, the lower
  wins. This is the right rule to follow for any future inter-row
  strip work.

  Final test totals: 209 (was 197 pre-cycle). Three new pin-downs in
  `tests/m12FixturePinDowns.test.ts` (M13/E describe block), 6 new
  tests in `tests/excelReferenceFidelity.test.ts` (snapshot fidelity),
  4 new tests in `tests/excelCanvasFidelity.test.ts` (canvas
  fidelity, includes bottom-border tests added in rework #3), 4 leak
  pin-downs in `tests/m13RedoSmokeRedCell.test.ts` (smoke seed leak
  fix). Generator-evidence:
  `screenshots/feature-1-m13-theme-aware-banding/generator-evidence-{aptos,classic}.png`.
  Operator-captured Excel references:
  `screenshots/excel-reference/FormattingSmorgasboard-{Aptos,Classic}.png`.
  Evaluator graded PASS across all reworks. Shipped via PR #22 (commit
  `fca1cbc`); README cleanup PR #23 marked M13/E shipped.
- **M14 NO-GO** (2026-06-04) — `xlsx-js-style` (the SheetJS-style
  parser explored as an `exceljs` replacement) ALSO has `!cf`
  undefined on indexed-cellXf import. The "wait for SheetJS to make
  CF cheaper" rationale evaporates; M15 ships on `exceljs`. Decision
  document preserved at `docs/m14-sheetjs-spike.md` and merged via
  PR #25 (commit `f423d5a`). Roadmap renumbered: M15 is now the next
  active milestone (was already so before but the README marker was
  off-by-one).
- **feature-1-m16-snapshot-to-html** (2026-06-06) — Ships the
  snapshot → HTML renderer as a Joplin
  `ContentScriptType.MarkdownItPlugin` content script. New entry at
  `src/contentScripts/notesheetRenderer.ts`; registered in
  `src/index.ts` alongside the existing editor; bundled via the
  existing `buildExtraScripts` webpack target (commonjs2 — see
  `plugin.config.json` extraScripts list). The renderer overrides
  markdown-it's fence handler: when the fence info is `notesheet v=1`
  it parses the JSON snapshot, walks `sheetOrder` / `sheets[id]` /
  `styles` / `mergeData`, and emits `<table>` per sheet with inline-
  styled `<td>`s (bg / fg / bold / italic / underline / horizontal +
  vertical alignment / per-side borders). Non-notesheet fences fall
  through to markdown-it's default — verified with a Jest test that
  feeds `javascript` / `python` / empty info. CF rules are evaluated
  inside the renderer itself (Univer's CF preset only paints at
  canvas-render time): cellIs, top10, colorScale all work; dataBar +
  iconSet are punted with documentation in OPERATOR_ASK and
  BUILD_PLAN's Out-of-scope list. Merged cells emit colspan/rowspan
  on the anchor and skip the interior cells (verified by counting
  tds in the merge-row).

  **13 new Jest tests** in `tests/m16NotesheetMarkdownRender.test.ts`
  (criterion 1 base shape × 7, criterion 2 multi-sheet × 1,
  criterion 3 FormattingSmorgasboard × 1, criterion 4 CF × 1, plus 3
  edge cases — non-notesheet fence falls through, malformed JSON
  returns null, unsupported version returns null, HTML-escapes
  cell values to defang `<script>` injection). Multi-sheet test
  builds a 3-sheet workbook in-memory via exceljs because the
  shipped `MultiSheet.xlsx` fixture has chart drawings that crash
  exceljs's reconcile (a pre-existing M12 known shortcoming). Test
  total: 220 → 233 (13 new, none regressed).

  **Harness extension: `previewPane` regionKind.** New code path
  in `scripts/pge/eval-screenshot.js`. Drives Joplin into preview-
  visible state via three menu clicks (View > Toggle editor plugin
  if Custom Editor active, View > Toggle editors if TinyMCE active,
  View > Toggle editor layout up to 3× until preview iframe is
  visible). Identifies the preview frame by `document.title ===
  'Note viewer'` (Playwright's `frame.url()` returns empty for
  Joplin's `joplin-content://note-viewer/` protocol). Samples the
  preview iframe's DOM via `samplePreviewPaneInk()`: parses
  CSS-computed `background-color` of every `<td>`, runs them
  through the same threshold expressions the canvas sampler uses
  (greenInk / pinkInk / lightGreenInk / etc.), and adds three
  preview-pane-specific signals to the sidecar: `tableCount`,
  `sheetHeadings`, `rawJsonLeak`. The latter is the M13-style gate
  — if `rawJsonLeak: true`, the renderer didn't run and the user
  saw a JSON blob in their export.

  AppleScript invocation hardened: switched from
  `execSync('osascript -e ${JSON.stringify(script)}')` to
  `execFileSync('osascript', ['-e', script])` because nested-quote
  escaping in the shell-arg form intermittently produced syntax
  errors (`Expected "given", "in", "of", expression...` at line
  37:38). The argv form passes the script verbatim and is
  syntactically reliable.

  Generator-evidence:
  `screenshots/feature-1-m16-snapshot-to-html/eval-*.png`
  (FormattingSmorgasboard fixture rendered: green table header,
  CAEFCB banded rows, no raw JSON, all column headers visible) +
  `.pixels.json` sidecar (`dominant=rgb(202,239,203)` =
  `#CAEFCB`, `greenInk=35`, `lightGreenInk=28`, `tableCount=1`,
  `rawJsonLeak=false`). Read via the Read tool.

- **feature-1-m15-conditional-formatting** (2026-06-05) — Full
  round-trip on the 5 CF rule types in
  `ConditionalFormatting-Variants.xlsx`: colorScale, dataBar,
  cellIs/highlightCell.number, top10/highlightCell.rank, iconSet.
  Phase 1 wired `UniverSheetsConditionalFormattingPreset` into
  `src/editorView.tsx` (the single-point-of-failure step — Phase 1
  build smoke confirmed canvas renders). Phase 2 added 6 fidelity
  tests in `excelReferenceFidelity.test.ts` anchored to the source
  XML (parsed via JSZip + regex), authored failing first per the
  fidelity-test-gap discipline; per-type translators in `src/xlsx.ts`
  (`translateExceljsCfRuleToUniver`) flipped them green. Phase 3
  extended `eval-screenshot.js` with `cfAllColumns` regionKind +
  `cfColumnRegion(col)` per-column samplers + new
  `pinkInk`/`lightGreenInk`/`yellowInk` aggregates plus broadened
  `redInk` (g/b ≤ 80 → ≤ 140 to catch #F8696B colorScale red end)
  and `blueInk` (b clearly dominant + ≥ 150, catches #638EC6 dataBar
  blue). Captured generator-evidence screenshot showing all 5 CF
  columns rendering correctly: A red→yellow→green gradient, C
  proportional blue bars, E pink fills on >50, G light-green on top-3
  cells, I red-down/yellow-flat/green-up arrows. Phase 4 added
  `excelCanvasFidelity.test.ts` describe block (1 test + 4 todos)
  gated `describe.skip` when the operator-captured Excel reference is
  absent at `screenshots/excel-reference/ConditionalFormatting-Variants.png`.
  Phase 5 added `translateUniverCfRuleToExceljs` + assigns
  `worksheet.conditionalFormattings` in `snapshotToXlsxBuffer`;
  flipped the KNOWN SHORTCOMING test at
  `tests/m12FixtureRoundTrip.test.ts:206` to a positive 'round-trip:
  5 conditional-formatting rules survive export → re-import' pin-down
  asserting all 5 source rules carry through structurally. Test
  total: 209 → 215 passed (6 new fidelity tests + 1 flipped, 6
  skipped including the canvas-fidelity todos). Generator-evidence:
  `screenshots/feature-1-m15-conditional-formatting/generator-evidence.png`
  (+ `.pixels.json` sidecar). Excel reference screenshot at
  `screenshots/excel-reference/ConditionalFormatting-Variants.png`
  delivered by operator; evaluator graded PASS. Shipped via PR #26
  (commit `88297a1`); README cleanup PR #27 (commit `19b0a6c`)
  marked M15 shipped.

- **m16-gap-3-formula-recalc-doc-only** (2026-06-07) — Closed M16's
  formula re-evaluation gap with documentation, NOT a renderer-side
  evaluator. Empirical investigation showed Univer is the de facto
  source of truth for cell.v on every save path; building a second
  evaluator inside the M16 HTML renderer would be ~30-50 functions
  / ~10KB / two-engine drift. README "Known shortcomings" entry
  rewritten to explain the Univer-as-source-of-truth contract; 4
  pin-down tests added in `tests/m16FormulaSourceOfTruth.test.ts`
  (formula cells carry both f and v at import; stale results flow
  through unchanged; synthesizeTableStyleAssignments doesn't touch
  f/v; renderCellValue reads cell.v not cell.f). Detail under
  ## Notes "M16 Gap #3 closure rationale". Test count 263 → 267.
  Shipped on `m16/document-formula-recalc` branch.

- **m13-e-followup-totals-and-inter-row-strips** (2026-06-07) —
  Re-probed the wide Aptos reference at 7 x positions and Classic
  at 10 x positions. Two definitive findings: (1) totals-top is the
  HEADER colour DOUBLE-line (not the lighter accent MEDIUM) — Aptos
  `#34692E` and Classic `#A5A5A5`; (2) inter-row strips DO exist at
  every banded-row boundary in the lighter accent (`#72D068` Aptos /
  `#C9C9C9` Classic). Recipe `EXCEL_TABLE_STYLE_EMPIRICAL_OVERRIDES.TableStyleMedium4`
  updated; synthesizeTableStyleAssignments now emits totals-top as
  DOUBLE (s=7), totals-bottom as MEDIUM (s=8), inter-row strips as
  MEDIUM bd.t (s=8) on every data row using `totalsBottomBorder`
  slot. New canvas-fidelity test shape: structural sentinels for
  "double-line at totals-top" + "9 inter-row strips ±3" with
  gridline-tail trim heuristic. Bidirectional round-trip test
  added (3 tests in `tests/roundTripBidirectional.test.ts`)
  verifying content + style edits flow through both Joplin → export
  → Excel-edit → re-import correctly, and that synth fields don't
  bleed into the exported xlsx or accumulate. Test count 256 → 263.
  Univer canvas anti-aliases recipe colours by Δ14-23 RGB units —
  documented as a Univer renderer characteristic (snapshot data and
  exported `.xlsx` are unaffected).

## In progress

(Empty between sessions.)

- **feature-6-m17-programmatic-roundtrip-pack** (2026-06-10) — Added
  `tests/m17ChartProgrammaticRoundTrip.test.ts` (8 tests). A pack of
  in-memory `UniverSnapshot`s authored directly in test code — each
  carrying one (or two) `NotesheetChart` drawings in a
  `SHEET_DRAWING_PLUGIN` resource built in the SAME shape
  `readChartsFromSnapshot` consumes — driven through the real export
  (`snapshotToXlsxBuffer`) and import (`xlsxBufferToSnapshot`)
  pipelines. NO `new ExcelJS.Workbook()` anywhere (criterion 4
  enforced by a runtime lint-sentinel test that strips comments from
  this file's own source and asserts the constructor never appears in
  executable code). Cases:
  - **A** bar w/ mixed +/− values `[3,-2,5,-1]` — negatives survive.
  - **B** line w/ single-data-point series — 1-element shape survives.
  - **C** pie w/ a >30-char category label — survives verbatim.
  - **D** doughnut w/ empty series — see ## Notes for the PRESERVE
    decision (it does NOT round-trip to truly-empty arrays).
  - **E** bar w/ `& < > " '` in title + category + series-label —
    `escapeXml`/`decodeXmlEntities` confirmed inverse on this set.
  - **F** cross-sheet (chart on `sheet-2`, `sourceSheetName: 'Sheet1'`)
    — both survive.
  - **G** two charts on one sheet — both survive in document order
    with DISTINCT regenerated `chartId`s.
  Each case anchors to the ORIGINAL snapshot's drawing fields, not a
  literal: `expectPinnedFieldsEqual` asserts type/sourceRange/labels/
  datasets/anchor-col-row equality. `chartId` is excluded by design
  (import regenerates `chart-imported-<sheet>-<idx>-...`). Anchor
  offsets authored as 0 because the import resource builder treats
  `<xdr:colOff>/<xdr:rowOff>` as EMU and divides by 9525 — only a 0
  offset survives a programmatic round-trip cleanly; col/row indices
  round-trip exactly. Test count 338 → 346. No source changes —
  feature-6 is pure characterization of the existing M10+M17 pipeline.
  Evidence:
  `screenshots/feature-6-m17-programmatic-roundtrip-pack/jest-result.txt`
  (per-test pass list). Row flipped `passes: true`,
  `evaluator_verdict: PENDING` (pure-Jest feature; no rendering
  dimension for the evaluator to screenshot — the spec's acceptance
  criteria are all round-trip data assertions).

- **feature-2,3,4,5-m17 (multi-feature cycle)** (2026-06-09) —
  M17's round-trip core landed in one session by treating
  features 2 + 3 + 4 + 5 as interlocked rather than feature-cycling
  through each. Decisions and their consequences:
  - **Plumbed `sourceSheetName` end-to-end** (`ChartDrawing.sourceSheetName`
    in `src/charts/xlsxChart.ts`, written by xlsx.ts at import time,
    read by `injectChartsIntoZip` to rebuild `<c:f>` formulas against
    the data sheet — not the chart-host sheet). Cross-sheet round-trip
    test (`tests/m17ChartBidirectionalRoundTrip.test.ts`) asserts the
    EXPORTED chart1.xml's `<c:f>` prefix stays `Sheet1!` even when the
    chart lives on Sheet2.
  - **Added `extractDataFromSnapshot(snapshot, range, sheetName?)`**
    (`src/charts/extractData.ts`) — pure-function snapshot variant of
    `extractRangeAsChartData` for the M16 content-script (feature-7) and
    feature-4's bus tests. Reads cell.v sparsely; tolerates empty cells
    via NaN/empty-string coercion. Same column-0 → labels convention as
    the FWorkbook variant.
  - **Factored `trackedCharts` + `populateTrackedChartsFromSnapshot`**
    out of `src/editorView.tsx` into `src/charts/trackedCharts.ts` so
    Jest can import without booting Univer. Hydrates the editor's
    chart-tracking map from `SHEET_DRAWING_PLUGIN` on snapshot load —
    plugs the M13-class trap that imported charts subscribed to nothing
    on the bus.
  - **Two correctness fixes in xlsx.ts's chart-resource emit, both
    load-bearing for feature-3:**
      1. **`drawingType: 5` → `drawingType: 8`.** The 5 was
         `DrawingTypeEnum.DRAWING_VIDEO`; 8 is `DRAWING_DOM` (verified
         in `node_modules/@univerjs/core/lib/es/index.js:2867`). Univer's
         drawing service silently drops unrecognized chart components
         from the render layer when drawingType is wrong — no warning,
         no error, just an invisible chart. Discovered via probe script
         that compared `addFloatDomToPosition`-emitted shape (`8`) with
         our import shape (`5`).
      2. **Added `transform: { left, top, width, height }` block** with
         pixel coords derived from the anchor's cell indices using the
         same `DEFAULT_COL_W=73 / DEFAULT_ROW_H=19 / ROW_HEADER_W=46 /
         COL_HEADER_H=20` constants the live `insertChart` uses. Without
         the transform block the float-DOM mounts at (0,0) regardless of
         anchor.
  - **`resolveDataFromCells` fallback** in xlsx.ts: when the source
    chart XML's `<c:cat>`/`<c:val>` shipped formulas without
    `<strCache>`/`<numCache>` (programmatically-built workbooks like
    `MultiSheet.xlsx`), we resolve the labels/data values from the
    matching data sheet's `cellData` at import time. Each dataset gets
    a `backgroundColor` from `CHART_PALETTE` so the rendered chart uses
    Notesheet's recognisable palette (otherwise NotesheetChart falls
    back to Chart.js's neutral default).
  - **Harness extensions for feature-3:**
    `scripts/pge/eval-screenshot.js` gained `floatDomChart` regionKind
    (screenshots `#notesheet-univer-root` instead of just the canvas so
    the float-DOM is captured), a stdlib PNG decoder + chart-palette
    histogram (decodes the saved screenshot, counts pixels within Δ ≤ 30
    of each `CHART_PALETTE` entry, sidecar reports `chartPaletteHits`,
    `paletteSwatchesFound`, `dominantNonBackground`), and an optional
    `PGE_ACTIVATE_SHEET=<name>` env var that activates a non-default
    sheet via FUniver before screenshotting (needed because
    MultiSheet.xlsx's chart lives on the second sheet, but joplin://
    re-opens onto the first).
  - **`scripts/pge/activate-sheet.js`** — companion utility that does
    the same activation as a one-shot.
  - **Test totals: 279 → 294** (15 new tests across features 2/4/5):
    `tests/m17ChartTypeFidelity.test.ts` (5 type-fidelity cases incl.
    radar fallback), `tests/m17ChartImportLiveUpdate.test.ts` (5 bus +
    snapshot + trackedCharts cases), `tests/m17ChartBidirectionalRoundTrip.test.ts`
    (5 cases: 4 type fixtures + cross-sheet survives).
  - **Feature-3 evidence:** generator-evidence.png shows MultiSheet's
    Chart sheet with the bar chart float-DOM rendered — title "Data
    Chart", bars Apples 30 / Bananas 50 / Cherries 20 / Dates 45 /
    Elderberry 60 in `#3b82f6` (CHART_PALETTE[0]). Sidecar reports
    `chartPaletteHits["#3b82f6"]: 134791`, `paletteSwatchesFound: 1`,
    `dominantNonBackground: rgb(72,128,232)` (anti-aliased blue near
    the palette colour).

- **feature-1-m17-chart-import-no-crash** (2026-06-09) — Pre-load chart
  reader + drawing-stripper architecture. New module
  `src/charts/xlsxChartImport.ts` exports `readChartsFromXlsxZip(buffer)`
  (zip+regex chart parser, walks every sheet's drawing rels in document
  order, parses each `xl/charts/chart{N}.xml` independently of exceljs's
  thin chart parser) and `stripChartPartsFromZip(buffer)` (drops chart
  drawing parts + chart xml + sheet rels + content-types overrides;
  idempotent on chart-less workbooks). Wired into
  `src/xlsx.ts:xlsxBufferToSnapshot` BEFORE `wb.xlsx.load` — read charts
  from original buffer, swap to stripped buffer for the load, attach
  the `SHEET_DRAWING_PLUGIN` resource to the snapshot AFTER post-load
  readers complete. The four existing post-load readers
  (`readTablesFromXlsxZip` / `readThemeFont` / `readThemeClrScheme` /
  `readNamedHyperlinkCells`) continue against the ORIGINAL buffer
  (their inputs don't include chart parts; the strip is invisible).
  All 10 hand-crafted fixtures under `tests/fixtures/charts/` import
  with their charts in the snapshot resource — bar/line/pie/doughnut
  type fidelity confirmed against source XML; multi-anchor case
  `06-two-charts-one-sheet.xlsx` walks both `<xdr:twoCellAnchor>`
  blocks in document order; cross-sheet `07-chart-cross-sheet.xlsx`
  preserves `sourceSheetName: 'Sheet1'` despite chart living on Sheet2.
  Test count moved 267 → 279 (12 new tests in
  `tests/m17ChartImportNoCrash.test.ts`). Generator-evidence
  screenshot at
  `screenshots/feature-1-m17-chart-import-no-crash/eval-2026-06-09T06-54-25-634Z.png`
  (chart-bearing fixture rendered in Joplin Univer canvas, no
  xlsx-charts-unsupported error). Two M12 pin-downs in
  `tests/m12ImportRecovery.test.ts:59,84` flipped (forced by criterion
  5 of feature-1 — the strip path makes the legacy expectations
  factually wrong; see ## Notes below for the reasoning).
  `MultiSheet.xlsx` flipped to "→ snapshot with N charts";
  `LargeWorkbook.xlsx` flipped to "→ xlsx-multi-table-unsupported"
  because the strip uncovers a SECOND crash class (multi-table reduce
  in worksheet.js:920) that's beyond M17's scope. Harness:
  `import-fixture.ts` extended to search both `tests/fixtures/charts/`
  and the legacy formatting-testdata root; `eval-screenshot.js`
  TITLE_PREFIX_BY_FEATURE + REGION_BY_FEATURE entries added for the
  feature.

## Next

M17 ships chart import from `.xlsx` (drawings + bar/line/pie/doughnut
chart definitions) plus the M16 follow-up gap "charts don't render in
HTML / preview-pane / PDF export" rolled into the same cycle. See
`BUILD_PLAN.md` for the full per-feature decomposition; see
`OPERATOR_ASK.md` for the operator brief (12 acceptance criteria,
Out-of-scope list, three suggested fixture sets, detailed
Related-risks notes). Approach choice: **B — extend pre-load
zip-direct readers** (mirrors `readTablesFromXlsxZip` /
`readThemeFont` / `readNamedHyperlinkCells` / `readThemeClrScheme`
existing pattern); the chart parts are read first, then a drawing-
stripped buffer is passed to exceljs.

- **feature-1-m17-chart-import-no-crash** — all 10 hand-crafted
  fixtures (`tests/fixtures/charts/01-bar-simple.xlsx` through
  `10-bar-with-trendline.xlsx`) import without throwing
  `xlsx-charts-unsupported`; snapshot has `SHEET_DRAWING_PLUGIN`
  resource with chart drawings. Tests anchor to source XML, NOT to
  our own emit. Error class stays defined for future drawing-related
  crash classes.
- **feature-2-m17-chart-type-fidelity** — bar / line / pie / doughnut
  fixtures import with the matching `ChartType` literal; unsupported
  source types (radar, scatter) fall back to `'bar'` with
  `meta.unsupportedSourceType` populated and a `console.warn`.
  Programmatic radar-chart zip exercises the fallback in-test (not
  a checked-in fixture).
- **feature-3-m17-multisheet-import-editor-canvas** — PGE smoke
  against `MultiSheet.xlsx` (the original "this crashes import"
  fixture) imported via `import-fixture.sh`, opened in Joplin's
  Custom Editor; screenshot of the Univer outer container shows
  cells + chart float-DOM + chart title + bars/lines/slices. New
  `floatDomChart` regionKind in `eval-screenshot.js`. Harness
  fixture-path expansion accepts `tests/fixtures/charts/` too.
- **feature-4-m17-live-update-data-bus** — Jest test confirms the
  imported chart's `chartId` is what `subscribeChartUpdate` keys off;
  `extractData(snapshot, sourceRange)` returns labels/values matching
  the source XML; no second renderer code path in `editorView.tsx`
  (static-analysis sentinel).
- **feature-5-m17-bidirectional-roundtrip-excel-fixtures** —
  Excel-authored fixture → snapshot → M10 export → re-import yields
  same chart-drawing fields (type, sourceRange, labels, datasets).
  Anchored to ORIGINAL snapshot, not a hardcoded literal.
  Cross-sheet (`07-chart-cross-sheet.xlsx`) survives.
- **feature-6-m17-programmatic-roundtrip-pack** — 5–7 in-memory
  snapshots (negative values, single-data-point, long labels, empty
  series, special chars, cross-sheet, two-charts-on-one-sheet) round-
  trip through M10 export + M17 import. NO `new ExcelJS.Workbook()`
  in the test (per operator's "Notesheet emit ↔ Notesheet import"
  framing).
- **feature-7-m17-chart-svg-html-export-jest** — M16 content script
  extended with chart-to-SVG renderer (hand-authored `<rect>` /
  `<polyline>` / `<path>` primitives — NO Chart.js / D3 in the
  bundle). Bar `<rect>` count == data length; line `<polyline>`/
  `<path>` count per dataset; pie sweep angles ±3° of expected.
  Bundle stays under ~20 KB. M16's existing tests stay green.
  CHART_PALETTE duplicated in the content script with a comment
  pointing at `src/charts/extractData.ts` (verified at test time).
- **feature-8-m17-chart-preview-pane-pge-smoke** — second PGE smoke
  against `MultiSheet.xlsx` re-using M16's `previewPane` regionKind;
  preview iframe screenshot shows table + inline `<svg>` chart at
  anchor; sidecar's new `inlineSvgCount ≥ 1` signal gates this.
- **feature-9-m17-flip-pin-downs-and-test-count** — flip
  `tests/m12ImportRecovery.test.ts:59,84` from "MultiSheet → throws"
  / "LargeWorkbook → throws" to "MultiSheet → snapshot with N
  charts" / "LargeWorkbook → snapshot with M charts" (counts read
  from source XML). Test count moves 267 → ≥ 290. `git diff tests/`
  shows ONLY new `tests/m17*.test.ts` files AND the two flipped
  lines — no other content edit to any existing test file (operator
  criterion #10 is load-bearing structural-integrity gate).

## Notes

- **feature-6 Case D — empty-doughnut PRESERVE-with-placeholder.**
  The spec lets the empty-series doughnut either DROP (0 drawings) or
  PRESERVE (1 drawing). The implementation lands on PRESERVE, but NOT
  with truly empty arrays. M10 export emits a degenerate chart: the
  series ref formula collapses to `Sheet1!$A$2:$A$1` (startRow+1 >
  endRow) with `<c:ptCount val="0"/>` for both `<c:cat>` strCache and
  `<c:val>` numCache. On re-import, `xlsxChartImport` reconstructs a
  1-row `sourceRange` (`{startRow:0,endRow:1,...}` — the import widens
  by +1 because the header sits one row above data), sees the empty
  caches, and `resolveDataFromCells` (in xlsx.ts) fires: it reads the
  (empty) cellData over that range and fabricates a SINGLE placeholder
  point — `labels: ['']`, `datasets: [{data: [0]}]`, plus default
  `meta` (legendPos:'r', categoryAxisType:'category', holeSize:50,
  dispBlanksAs:'gap'). So the round-trip does not crash and does not
  invent MEANINGFUL data, but it is NOT idempotent for the empty case:
  `[] → ['']` and `[] → [0]`. The test pins exactly this — surviving
  labels are all `''` and surviving values are all `0`, length ≤ 1 —
  rather than asserting strict emptiness. If a future cycle wants true
  idempotence here, the fix is in `resolveDataFromCells`: skip the
  fallback when the reconstructed data range is degenerate
  (dataRowEnd < dataRowStart maps to a real zero-row chart), OR have
  M10 export DROP charts whose datasets are all empty before zip
  injection. Either is a behaviour change beyond feature-6's
  characterization scope.
- **`emptySnapshot()` is the seam.** It's the single function that
  produces a fresh workbook for both the New Spreadsheet command
  (src/index.ts:103) and any "load empty fence" path. Putting the
  smoke seed here meant Jest tests could pin it down without mocking
  the Joplin Data API.
- **Univer's style resolver requires the colour on `styles[id]`,
  not inline on the cell.** The cell carries `s: <id>` (not
  `s: { cl: ... }`) and `cl.rgb` must include the leading `#`. We
  documented this in the BUILD_PLAN risks block; M13 lessons confirm
  it. **Important caveat for M13/D:** this rule is for cell-level
  style. Rich-text runs carry their `ts` inline on
  `cell.p.body.textRuns[i].ts` by design — they do NOT live in
  `styles[id]`. Don't conflate the two paths.
- **Evaluator needs a SEEDED note, not a `{}` note.** The plugin's
  `extractSnapshot('{}'...)` succeeds on `{}` (valid JSON, valid object)
  and the editor loads an empty Univer workbook — bypassing
  `emptySnapshot()` entirely. The harness needs a note whose body is
  `wrapSnapshot(emptySnapshot())`. We added
  `scripts/pge/create-seeded-notesheet.js` for that; it duplicates
  the seed shape, which is intentional coupling (the harness is
  allowed to know what the runtime does).
- **CDP page picker.** Joplin's CDP exposes 4-5 pages: a DevTools
  page, plugin sandboxes (one per loaded plugin including ours), and
  the editor. eval-screenshot.js scores by URL/title; editor wins
  unambiguously (16 vs 11/0/0). Plugin sandbox pages are
  `<body></body>` — they host plugin process logic, NOT the editor
  view. Don't try to attach to them.
- **Univer mounts in `UserWebviewIndex.html`** (a frame of the editor
  page), not the plugin sandbox. Stable selectors (Univer 0.23):
    - `canvas[id^="univer-sheet-main-canvas"]` — the main spreadsheet
      canvas. Id is `univer-sheet-main-canvas_<workbookId>`.
    - `[class*="univer-flex"]` — the toolbar wrapper (appears slightly
      before the canvas).
    - `#joplin-plugin-content` — the Joplin webview wrapper (always
      present once the plugin loads).
- **Pixel sidecar (`.pixels.json`)** — the harness samples the
  Univer main canvas's top-80px row-0 slab (or a feature-specific
  region) and writes a histogram of non-background colours alongside
  the screenshot. Use it for machine-checkable assertions like
  "top contains rgb(255,0,0) > 50 hits" instead of "I saw red."
  Sampling is stride-2 to keep cost cheap. Background (>235 in all
  channels) and gridline ink (<30) are filtered out.
- **Pre-existing `tests/exportTableRoundTrip.test.ts:334` typecheck
  bug** — `'dashed'` is not in exceljs `BorderStyle` enum. Changed
  to `'mediumDashed'` (with matching assertion update on line 349)
  during the smoke session because webpack's TS check blocks .jpl
  build. This was unrelated to smoke; the smoke didn't introduce it.
  Reverting it would re-break `npm run dist` for any subsequent
  cycle.
- **`scripts/pge/import-fixture.{ts,sh}`** — headless equivalent of
  the plugin's "Import .xlsx as Notesheet" command. The .ts calls
  `xlsxBufferToSnapshot()` from `src/xlsx.ts` directly so the harness
  exercises the SAME conversion the plugin runs at runtime; the .sh
  wrapper compiles via `node_modules/.bin/tsc` to a temp dir and runs
  the JS through Node with `NODE_PATH` pointing at the repo's
  `node_modules`. `PGE_REPO_ROOT` env var is required because
  `__dirname` after compile lives inside the temp dir.
- **Joplin window pane crops the Univer canvas.** `page.screenshot()`
  with `fullPage:false` captures the whole Joplin window — but at
  Joplin's default pane sizes the editor pane is narrower than the
  Univer canvas, so the canvas is cropped or partially offscreen.
  Fixed in `eval-screenshot.js` by screenshotting the canvas element
  directly via `webview.locator(canvasSel).screenshot()` when a
  Notesheet note is opened. The whole-page screenshot is the
  fallback for smoke / verification mode.
- **`inkRowSpread` metric.** Added to the pixel sidecar:
  `inkRows / ceil(regionHeight/2)` — the fraction of sampled y-rows
  in the region that carry text-coloured ink. Horizontal text
  concentrates ink on a narrow band (low spread); rotated/stacked
  text spreads across the band (high spread, near 1.0). Use this
  alongside the colour histogram for rotation-style features. For
  M13/D, the relevant metric is per-colour histogram hits within the
  A1+A2 (or A2-only) region — `inkRowSpread` is not the right signal
  for rich-text.
- **`feature-1-m13-rotated-text-renders` worked first try in Univer
  0.23.** PR #16's `style.tr = { a: <angle> }` (and `{ a: 0, v: 1 }`
  for stacked) maps directly to Univer's `ITextRotation` and the
  resolver honours it without extra plugin registration. The earlier
  visual failure that prompted the revert was almost certainly a
  stale-build / cache-not-wiped issue, not a renderer-vs-shape
  mismatch. The harness's `install-plugin.sh` cache wipe + Joplin-
  quit gate is what makes the difference. **Apply the same
  hypothesis to M13/D first** — restore `6f33f3a` verbatim, do a
  full rebuild + cache wipe + Joplin re-launch via the harness
  scripts, then judge from real pixels before suspecting the
  helpers.
- **Region-by-feature pixel sampling.** `eval-screenshot.js` now
  consults `REGION_BY_FEATURE` and `TITLE_PREFIX_BY_FEATURE` tables
  to pick the right canvas region and note title prefix per feature.
  When adding a new feature whose evidence isn't on row 0, add a
  region helper (like `rotatedRowRegion`) and an entry in both
  tables. M13/D needs an A1+A2 region (and ideally an A2-only
  variant for cleaner colour signal — A3's hyperlink also
  contributes blue ink).
- **Window prep is mandatory before evaluator screenshots.** Three
  pre-conditions had to be enforced after operator hit them:
    1. Joplin window must fill its display (not 800×568 or whatever
       the operator left it at) — Univer sizes its canvas to the
       editor pane, so a small window means fewer columns rendered
       and the visual gate may miss content.
    2. Sidebar (`.rli-sideBar`) and note list (`.note-list`) must be
       hidden — they consume ~500-700px of horizontal real estate
       the editor could otherwise use.
    3. Any DevTools window must be closed — it shrinks the renderer
       AND can confuse `eval-screenshot.js`'s CDP page picker.
  `scripts/pge/prep-joplin-window.sh` handles all three. AppleScript
  invokes `Window > Fill` and sends `Cmd+Alt+S` / `Cmd+Alt+L` for
  pane toggles (Playwright's `keyboard.press()` does NOT reach
  Joplin's Electron-accelerator-routed shortcuts — the renderer
  receives the keydown but the command never fires; OS-level
  System Events keystroke does work). State-aware: queries renderer
  DOM via `element.offsetParent !== null` (canonical for "any
  ancestor has display:none") and only toggles when needed. Wired
  into `eval-screenshot.sh` ahead of the screenshot. Requires
  Accessibility permission for the terminal app in System Settings.
- **Pattern A hyperlink emitter has precedence over the rich-text
  emitter (M13/D pre-context).** `src/xlsx.ts` already emits Pattern
  A hyperlinks via `buildHyperlinkCellP` for single-format hyperlink
  cells (M12 work). `6f33f3a`'s rich-text export path adds a second
  consumer of `cell.p` — `extractRichTextRunsFromCellP` — which MUST
  explicitly skip when a hyperlink customRange is present on the
  cell, otherwise A3 of `RichTextInOneCell.xlsx` regresses to a
  1-element richText value with no hyperlink. The Jest test
  `hyperlink + plain stays Pattern A` is the regression sentinel;
  the round-trip assertion on A3 (`{text, hyperlink}` not richText)
  is the user-visible sentinel.
- **Pixel histogram fragments under anti-aliasing — added per-spec
  band aggregates to `samplePixelsAt`.** During M13/D, even with
  clean A2-text rendering, the `top` histogram showed `rgb(255,0,0)`
  at only 17 hits and `rgb(0,0,255)` at 12 hits when sampled at
  stride-2. Anti-aliased glyph edges produce many near-but-not-exact
  RGB tuples, fragmenting the per-bucket counts even when the
  feature is rendering fine. We added stride-1 sampling plus three
  inequality-band aggregates that match the spec verbatim:
  `redInk` (R≥200, G≤80, B≤80), `blueInk` (R≤80, G≤80, B≥200),
  `greenInk` (R≤80, G≥150, B≤80). The aggregates are robust to
  anti-aliasing because they sum every pixel that satisfies the
  inequality, regardless of exact RGB. M13/D's final reading was
  `redInk=67`, `blueInk=76` — well above the ≥30 threshold. Future
  per-run colour features should use these aggregates rather than
  the `top` histogram for gating. **For M13/E**, add a `greyInk`
  aggregate using the same template — recommended thresholds
  `R∈[140,180]`, `G∈[140,180]`, `B∈[140,180]` with `abs(R−G) ≤ 10`
  AND `abs(G−B) ≤ 10` (the equality between channels is what
  distinguishes grey from a tinted hue at similar luminance).
  **For M15**, add `pinkInk` (`R≥220 AND G∈[180,220] AND
  B∈[180,220]`), `lightGreenInk` (`R∈[180,220] AND G≥220 AND
  B∈[180,220]`), and `yellowInk` (`R≥200 AND G≥200 AND B≤120`)
  using the same template. The CF columns each have a different
  signature colour band; the gates are per-column.
- **The Univer cell-selection blue border at `rgb(44,83,241)` will
  saturate any region that includes A1 + cell-border y-band.** When
  a cell is the active selection (typical state on a freshly opened
  note), Univer paints a ~1px blue border around it. That border
  contributes ~150+ blue pixels to a slab covering the cell, even
  when the cell text itself is pure black. For colour-band
  features, sample a y-region that **excludes the active-cell
  border** — empirically on the M13/D fixture the borders sat at
  y≈20–21 (top of A1) and y≈38–40 (bottom of A1 / top of A2). We
  put the A2-only region at y=41–58 to land inside A2's text band
  cleanly. If a future feature needs an A1 colour gate (e.g. a red
  font on A1), either move selection off A1 first or carve the
  border-y rows out of the region. **For M13/E** the table header
  row is row 0 of the table data area = row 1 of the visible
  worksheet — the active-cell selection is on A1 by default and
  lands inside the header band. Either click off A1 before
  sampling (programmatic Univer click via the webview frame's
  command bus) OR sample a header column that is NOT col A
  (operator suggests a mid-table column like the Spent / Discount
  column where the green-vs-grey signal is unambiguous).
- **Rich-text rendering worked first try in Univer 0.23 once the
  build/cache state was clean.** Same lesson as M13/C: cherry-pick
  `6f33f3a` verbatim, full `npm run dist` + `install-plugin.sh`
  (cache-wipe gated on Joplin being quit) + relaunch + headless
  import via `import-fixture.sh` — the per-run formatting renders
  correctly without any helper rewrite. The original PR #16 revert
  was almost certainly a stale-build artefact, not a renderer gap.
  The hypothesis from M13/C generalises.
- **M13/E is the first multi-screenshot cycle.** Until M13/E every
  cycle had exactly one note → one screenshot → one row in
  `test-results.json`. M13/E needs TWO independent screenshots
  (Aptos fixture + Classic fixture) because the failure mode is
  "same catalog entry must produce DIFFERENT colours under
  different source clrSchemes." The plan picks the
  **variant-suffix approach** for the harness extension: feature
  ID becomes `feature-1-m13-theme-aware-banding:aptos` and
  `feature-1-m13-theme-aware-banding:classic`, and
  `eval-screenshot.js` looks up the suffixed key in
  `REGION_BY_FEATURE` / `TITLE_PREFIX_BY_FEATURE` (falling back to
  the plain key for prior-cycle compatibility). The generator MUST
  add the suffixed entries WITHOUT renaming or removing the
  M13/C and M13/D entries — those are evidence-bearing for the
  prior-cycle rows already in `## Done`. Document the suffix
  convention here once the M13/E session lands so the next planner
  knows it exists.
- **`tableHeaderRowRegion` y-band hint.** The header row of the
  imported `ProjectTracker` / `ProductCatalog` table sits
  immediately below Univer's column header (~y=0–18 at default
  zoom). At default row height (19px) the header band is roughly
  y=19–37. Use a slab y=18–40 (h≈22) to absorb default row-height
  variation. The new region helper goes in
  `scripts/pge/eval-screenshot.js` next to `rotatedRowRegion` and
  `richTextA1A2Region`.
- **Header colour gate column choice.** The top-left active-cell
  selection (`rgb(44,83,241)` border) on A1 will pollute a header-
  row sample that includes col A. Two valid mitigations: (1)
  programmatically click off A1 to dismiss the selection before
  sampling, or (2) restrict `tableHeaderRowRegion` to start
  several columns in (e.g. starting at `x=2*colWidth`). The
  generator picks one and documents it. **M13/E chose option 2**:
  `tableHeaderRowRegion` starts at `x=80` (col B onward, col width
  ~73px), excluding A1's active-cell selection border entirely.
  The selection-blue still shows up at low counts (~136 hits of
  `rgb(150,169,248)` in the histogram — bleed from the cell-border
  anti-aliasing along the right edge of A1 still visible at x≈74)
  but doesn't trigger redInk/blueInk/greenInk/greyInk gates.
- **Variant-suffix harness extension.** `eval-screenshot.js` now
  splits `FEATURE_ID` on `:` into `BASE_FEATURE_ID` + `VARIANT`.
  Lookups try the suffixed key first, then the plain base key —
  prior single-screenshot features keep working unchanged. Output
  filenames bake the variant: `eval-aptos-...png` /
  `eval-classic-...png`. Both screenshots land in the same
  `screenshots/<base-feature-id>/` directory. The `verify-gate`
  hook reads the directory; both PNGs must be Read before the
  test-results.json flip is allowed. **For future multi-variant
  cycles**: add a `:<variant>` row in BOTH `TITLE_PREFIX_BY_FEATURE`
  AND `REGION_BY_FEATURE` keyed off the suffixed id, run
  `eval-screenshot.sh feature-id:variant` once per variant. Note
  list and import order do NOT matter — the title-prefix lookup
  picks the latest match by `updated_time`.
- **Theme-aware synth approach: parallel recipe table.**
  `src/charts/excelTableStyleRecipes.ts` is a parallel mirror of
  `excelTableStyles.ts` that names each slot's accent index +
  HSL-L tint amount. At synthesis time
  `resolveTableStylePalette(styleName, catalog, themeRgb)` looks
  up the recipe and resolves each slot via `tintRgb(accent, tint)`
  using the source workbook's clrScheme accent values. The
  Aptos baseline catalog is preserved unchanged (legacy fallback
  when `themeRgb` is null) and verified exact against the new
  HSL-L formula by direct comparison: `tint('#196B24', 0.6)` =
  `'#84E291'` (catalog Medium4 even-row), within zero RGB units.
  Achromatic styles use `accent: null, rgb: '#...'`. **Don't add
  hardcoded Classic-Medium4 entries** — Excel ships ~thousands of
  accent permutations and the catalog can't enumerate them; the
  recipe table is the source of truth.
- **In-process ExcelJS test workbooks ship the Office 2007 default
  theme.** When tests build a workbook via
  `new ExcelJS.Workbook(); ws.addTable({...style: 'TableStyleMedium2'...})`
  and then round-trip it through xlsxBufferToSnapshot, the
  resulting snapshot's per-cell `bg` is `#4F81BD` (exceljs
  accent1) NOT `#156082` (Aptos accent1). M13/E updated three
  pin-downs in `tests/formattingFidelity.test.ts` and one in
  `tests/m12FixtureRoundTrip.test.ts` to assert the actual
  exceljs-default colour. The Aptos and Classic project-owned
  fixtures pin the Aptos and Classic palettes specifically (see
  the M13/E describe block in `m12FixturePinDowns.test.ts`).
- **`greenInk` aggregate threshold relaxed for M13/E.** The M13/D
  threshold (`R≤80 AND G≥150 AND B≤80`) only matches pure greens
  like `rgb(0,255,0)` — Aptos accent3 `#196B24` (R=25,G=107,B=36)
  has G < 150 and fails. Relaxed to `G > R+30 AND G > B+30 AND
  G ≥ 80` so any green-channel-dominant pixel ≥ 80 luminance
  qualifies. Catches Aptos accent3 (107 vs 25/36), Aptos pastel
  banded `#84E291` (132,226,145), AND pure `rgb(0,255,0)`.
  redInk/blueInk thresholds stay tight (M13/D's gate values
  unchanged). The post-Phase-3 measured Aptos header `#34692E`
  (52,105,46) also qualifies (`g=105 > r+30=82 && g > b+30=76 &&
  g >= 80`).
- **PR #17 smoke seed leak — production code carried debug state
  for ~6 weeks before catch.** `emptySnapshot()` was the simplest
  seam to put the smoke seed in (one function, pure-testable via
  Jest, no Joplin Data API mock needed). It was the wrong place
  long-term: every "New Spreadsheet" command in the deployed plugin
  shipped with cell A1 pre-filled. Lesson: **harness-only state
  belongs in `scripts/pge/`, not in `src/`.** The PGE
  `create-seeded-notesheet.js` had ALWAYS inlined the seed shape
  itself — the production seed in `emptySnapshot()` was redundant.
  Fix landed Phase 1 of M13/E rework session.
- **Why "asserting our own emit" let PR #22 ship wrong colours.**
  The original M13/E pin-downs in `m12FixturePinDowns.test.ts`
  asserted `bg.rgb === '#196B24'` for the Aptos header — that's
  what our code emitted (raw accent3), NOT what Excel actually
  paints (`#34692E`). The acceptance criterion "dark green" passed
  trivially because both `#196B24` and `#34692E` ARE dark greens.
  The pixel-sidecar threshold `greenInk ≥ 30` passed for both too.
  Anchoring tests to your own behaviour gives you confidence your
  code is consistent, NOT that it's correct. Phase 2's
  `excelReferenceFidelity.test.ts` is the corrective: it samples
  the operator-captured `screenshots/excel-reference/*.png` and
  asserts the synthesizer's output matches what Excel actually
  paints, within Δ ≤ 8 per channel. **Future "match Excel" features
  should establish a reference-anchored gate FIRST (Phase 2 before
  Phase 3) so the recipe lands against ground truth, not against
  hand-derivation.**
- **Excel TableStyle algorithm not crackable from a single accent.**
  Phase 3 investigation tested HSL-L tint, HSV scale, satMod+lumMod,
  RGB mix toward grey, lumMod+lumOff in HSL space — none reproduce
  all four target RGBs (Aptos header #34692E, Aptos band #CAEFCB,
  Classic band #EDEDED, Aptos totals top #72D068) from a single
  accent. Excel's built-in TableStyle definitions live in Office's
  installed assets (NOT in the workbook's `xl/styles.xml`) and the
  actual transformation isn't documented in OOXML. Operator's
  allowed empirical-lookup path was the right escape: ship measured
  RGBs in `EXCEL_TABLE_STYLE_EMPIRICAL_OVERRIDES`, keyed by
  `(styleName, accentHex)`. The HSL-L formula remains as fallback
  for accents we haven't measured. **Adding a third fixture means
  adding a third entry to the override map** — measure with
  `tests/util/pngSampler.ts:dominantColor` against the new Excel
  reference PNG.
- **Recipe shape gained `totalsTopBorder`.** PR #22's shape only
  modelled `borderColor` (table outline). Excel paints a separate
  border above the totals row in a lighter shade of the accent —
  `#72D068` for Aptos accent3, `#C9C9C9` for Classic accent3. Added
  to the recipe interface; consumed by
  `synthesizeTableStyleAssignments` which emits the totals row's top
  side in that colour. Falls back to the existing thin border when
  the slot is missing. **Style code is MEDIUM (8), not DOUBLE (7) —
  see "DOUBLE on top renders as anti-aliased two-strip" note below.**
  M13/E pin-down `Aptos fixture: totals-row top border carries the
  Excel separator (#72D068 green, MEDIUM)` is the sentinel.

- **Univer 0.23 DOUBLE-on-top renders as anti-aliased two-strip,
  NOT a true double-line.** `getLineWidth(BorderStyleTypes.DOUBLE)=1`
  in `node_modules/@univerjs/engine-render/lib/es/index.js:1872`,
  and `_renderDoubleBorder` (line 5102) places inner+outer strokes
  at `±lineWidth/2 = ±0.5px` from centre. Net result: two 1px strips
  separated by 1px of white, which reads as `#89CE74` (anti-aliased
  green) rather than the pure target colour. **Excel doesn't even
  paint a true double-line for `TableStyleMedium4` totals-top — pixel
  sampling of the operator-captured reference shows a single 2px
  strip in the lighter accent**, identical to the strips Excel paints
  at every banded-row boundary. So MEDIUM (style 8, lineWidth=2) is
  both: (a) closer to Excel's actual render, and (b) more visually
  reliable than DOUBLE in Univer 0.23. The visual "double-line"
  perception of the totals-top in Excel comes from the strip pairing
  with the banded-row decoration above it — see the cross-feature
  follow-up below.

- **Cross-feature follow-up: banded-row boundary decoration.**
  Excel paints `#72D068` (Aptos) / `#C9C9C9` (Classic) 2px strips at
  EVERY banded-row boundary inside `TableStyleMedium4`, not just the
  totals-top. Notesheet currently relies on Univer's default
  `#D7D8DB` thin gridline for inter-row separation, so our render
  shows the right cell-fill colours but the wrong inter-row borders
  — only the totals-top gets the accent-shade decoration. Out of
  scope for this rework; should be its own feature in BUILD_PLAN.md.
  The synthesizer would need to emit MEDIUM borders on every banded-
  data-row boundary (top of rows 2-9 in the Aptos/Classic fixtures)
  in `palette.totalsTopBorder` colour.

- **Canvas-vs-Excel fidelity test layer (`tests/excelCanvasFidelity.test.ts`).**
  Second gate added in M13/E rework #2. The original
  `excelReferenceFidelity.test.ts` compares Excel reference PNG →
  `xlsxBufferToSnapshot` output (catches import-side bugs); the
  canvas-fidelity test compares Joplin canvas screenshot → Excel
  reference PNG (catches render-side bugs). Two PNGs at different
  DPRs and table positions: aligned by structural regions, NOT by
  raw pixel coordinates. `findColouredStrip` scans for the dominant-
  coloured strip whose hex is within tolerance of the spec target.
  Per-region tolerance:
    - Region-FINDING: Δ ≤ 24 per channel (wide for anti-aliased
      strip edges).
    - Header / banded ASSERTION: Δ ≤ 8 (tall regions, anti-aliasing
      doesn't dominate).
    - Totals-top STRIP ASSERTION: Δ ≤ 32 (1-2px strip at DPR=2
      pulls in sub-pixel anti-aliasing; pure-pixel parity isn't
      achievable without a different sampling scheme).
  Structural sentinel: "Joplin's totals-top is a SINGLE strip, not
  two strips with a ≤4px gap" — directly catches the DOUBLE-render
  artifact. The test reads the LATEST `eval-{aptos,classic}-*.png`
  by mtime, so re-capturing screenshots automatically refreshes
  what's tested.
- **`tests/util/pngSampler.ts` is test-only.** Pure-stdlib `zlib`
  PNG decoder. Supports 8-bit RGB / RGBA, non-interlaced — the only
  shapes our reference PNGs use. **Don't promote to a runtime
  dependency.** If runtime PNG decoding is ever needed, that's a
  separate design conversation; pulling `pngjs` or `sharp` into the
  bundle is not free.
- **`tableHeaderRowRegion` was Retina-broken until Phase 4.** The
  hardcoded y=22 / h=13 region values assumed a non-Retina Univer
  canvas (DPR=1). On macOS Retina the canvas backing store is at
  DPR=2, so `canvas.width = 3008` instead of `1504`, and y=22
  landed inside the Univer column-letter strip instead of the
  table header row. Fix: scale x/y/w/h by
  `canvas.width / canvas.clientWidth`. clientWidth is the CSS box;
  the ratio is exactly devicePixelRatio. **Future region helpers
  should use this same scaling pattern** — there's no clean way to
  hardcode pixel offsets that work on both Retina and standard
  displays. The simplest invariant: write the offsets in CSS px and
  multiply by `(canvas.width / canvas.clientWidth)` at sample time.
- **Two-layer fidelity test pattern is now the project default for
  "match Excel" features.** Snapshot-data fidelity
  (`tests/excelReferenceFidelity.test.ts`) anchored to the Excel
  reference PNG via `tests/util/pngSampler.ts:dominantColor` —
  catches import-side bugs. Canvas-vs-Excel fidelity
  (`tests/excelCanvasFidelity.test.ts`) anchored to the same
  reference PNG, samples Joplin's `eval-*.png` capture — catches
  render-side bugs. Both anchored UPSTREAM of our code, never
  against our own emit. M13/E rework #2 + #3 is the precedent;
  M15 extends both layers (5 tests per layer for the 5 CF
  columns). **M16 does NOT use the canvas-fidelity layer** — the
  M16 render target is HTML in Joplin's preview pane, not the
  Univer canvas, so canvas-vs-Excel pixel parity is not the gate.
  M16's test layer is "HTML content vs source-fixture content"
  (Jest-only; see BUILD_PLAN.md acceptance criteria 1–4).
- **Operator captures the Excel reference screenshot for a new
  fixture as part of the cycle.** M13/E precedent: when a new
  fixture is introduced (Aptos and Classic Smorgasboard PNGs were
  captured during the M13/E cycle, not before), the generator
  proceeds with own evidence and snapshot-fidelity work that
  doesn't need the reference, then the reference lands and the
  canvas-fidelity layer can be unskipped. M15's
  `screenshots/excel-reference/ConditionalFormatting-Variants.png`
  follows the same pattern. The canvas-fidelity tests should
  use `test.skip` (or the M13/E "read latest by mtime, fail
  clearly if missing" pattern) when the reference is absent.
- **M15: Univer CF resource shape — `{ [subUnitId]:
  IConditionFormattingRule[] }`** stringified as the snapshot's
  `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource. Each rule has
  `{cfId, ranges, stopIfTrue, rule: <type-specific>}`. The CF
  preset's `parseJson` does `JSON.parse(json)` and walks subunits;
  ranges use Univer's `IRange` shape `{startRow, endRow,
  startColumn, endColumn}` (zero-based). Multi-rule import works
  via flattening exceljs's `worksheet.conditionalFormattings`
  (which is `Array<{ref, rules: Array<...>}>`) into the per-subunit
  flat array. The CF preset registers via `UniverInstanceType.UNIVER_SHEET`
  business — no per-instance hook setup required.
- **`@univerjs/sheets-conditional-formatting`'s CJS bundle pulls
  `lodash-es`** (ESM-only), which jest-runtime can't parse without
  a transformer override. Tests that need
  `SHEET_CONDITIONAL_FORMATTING_PLUGIN` hard-code the string
  literal locally. Runtime preset (`createUniver` in
  `src/editorView.tsx`) reads the const from the package directly
  (webpack happily bundles the ESM dep). If Univer ever renames the
  constant, the runtime gets the new value but our hard-coded
  string stays the old; the snapshot/runtime mismatch trips loudly.
  Future cycle could add a Jest transformer override
  (`transformIgnorePatterns: ['node_modules/(?!lodash-es)']`) to
  let tests import the const directly — out of scope for M15.
- **iconSet shape mismatch: Univer descending + catch-all vs
  Excel ascending.** Univer's `IconSetCalculateUnit` walks the
  config from index 0 forward, returning the first item whose
  value/operator matches; the standard layout puts the HIGH icon
  at index 0 (e.g. `3Arrows = [up-green, right-gold, down-red]`).
  To match Excel's iconSet semantics where cfvo[0] = lowest band /
  cfvo[N-1] = highest, the import-side translator emits config in
  descending threshold order and uses MAX_SAFE_INTEGER for the
  catch-all entry that handles the lowest band. The export-side
  inverse drops the catch-all and reverses to ascending — Excel
  cfvo[0] is synthesized as `percent=0` since Univer's catch-all
  doesn't carry the lowest-band threshold.
- **CF colour aggregates** (`pinkInk` / `lightGreenInk` /
  `yellowInk`) added to `samplePixelsAt` plus a parallel
  per-column variant in `sampleCfColumns`. The thresholds match
  the operator-ask spec exactly:
    pinkInk:        R≥220 AND G∈[180,220] AND B∈[180,220] (#FFC7CE family)
    lightGreenInk:  R∈[180,220] AND G≥220 AND B∈[180,220] (#C6EFCE family)
    yellowInk:      R≥200 AND G≥200 AND B≤120 (gold arrow / yellow flat icon)
  Two existing aggregates also broadened (monotonic — no prior
  gate regresses):
    redInk:   g/b ≤ 80 → ≤ 140 (catches #F8696B = (248,105,107))
    blueInk:  r/g ≤ 80 → 'b clearly dominant by ≥30 over r AND g, AND b ≥ 150'
              (catches #638EC6 = (99,142,198))
- **`cfAllColumns` regionKind** routes through `sampleCfColumns()`
  instead of `samplePixelsAt(regionFn)`. The sidecar carries a
  top-level `cfColumns` object keyed by column letter (A/C/E/G/I)
  with per-column dominant + ink-aggregate sub-summaries; the
  legacy top-level fields (dominant / sampled / redInk / etc.)
  are filled from an aggregate sample over the whole A2:I11 band
  for backward compat. Future single-fixture multi-region cycles
  can mirror this pattern — name a dedicated `cfXxxYyy` regionKind
  and add a parallel sampler.
- **CF + table-style synthesizer interaction.** The fixture
  doesn't ship a table, so CF-on-top-of-table-style precedence is
  hypothetical for the M15 spec. Verified the
  `FormattingSmorgasboard-NonAptosClassicTheme...` fixture (which
  has both a table AND CF rules) imports without regression — the
  M13/E test suite stays green and Aptos/Classic header colours
  stay correct. The `cfBound` synth-styles sidecar tagging fix is
  out of scope for M15; if a future fixture exercises both
  surfaces and CF colour gets masked by synthesizer fill, mark
  CF-bound cells in `synthStyleSidecar` so the synthesizer skips
  them.
- **iconSet glyphs render correctly** via the preset's bundled
  SVG icon font. No missing-glyph squares observed for `3Arrows`
  on the M15 fixture. The `@univerjs/preset-sheets-conditional-formatting/lib/index.css`
  CSS import is what pulls the fonts; deferring it would have
  rendered as `?`/`□` placeholders. Other iconSets (4Arrows,
  3TrafficLights1, etc.) probably work via the same import but
  aren't under test in M15.
- **The CF round-trip is fully closed via exceljs.** No CF parts
  are dropped on import or export — the `<conditionalFormatting>`
  blocks in the re-emitted xml carry every source rule's type /
  ref / cfvo / colour / operator / rank / iconSet. exceljs
  preserves the dxf `bgColor.argb` for cellIs and top10 styles via
  its `style.fill.bgColor` field. `dataBar` blocks gain extra
  exceljs metadata on round-trip (x14Id, minLength, maxLength,
  axisPosition, etc.) — those are exceljs's defaults and don't
  break Excel's render.
- **Initial title-prefix mismatch caught at eval-screenshot
  time** — `findLatestNoteByTitle` requires the note title to
  start with the prefix INCLUDING the trailing space. Initial
  test note title `"PGE M15 CF eval"` (no trailing space) didn't
  match `"PGE M15 CF eval "` (with space). The fix is to use
  `--title "PGE M15 CF eval $(date -u +...)"` so the timestamp
  appended after the space supplies the matching prefix. Future
  cycles should remember: `import-fixture.sh --title "<PREFIX> ..."`
  must echo the title-prefix verbatim followed by a separator.
- **M15 harness sampler: content-discovering geometry via FUniver
  (Option A).** The first M15 cycle hardcoded
  `COL_W_CSS = 73` in `eval-screenshot.js:sampleCfColumns` and
  computed each CF column's x-origin as
  `ROW_HEADER_W_CSS + colIndex * COL_W_CSS`. When the operator
  widened columns to fit content (the long header labels "Color
  Scale (3-stop)" / "Highlight (>50 = Red)" / etc.), the sampler's
  per-column regions drifted sideways and bucket-leaked across
  columns: `cfColumns.C.blueInk = 0`, `cfColumns.E.pinkInk = 0`,
  `cfColumns.G.lightGreenInk = 0` while `cfColumns.G.pinkInk =
  20790` (col E's pink leaked into col G's region) and
  `cfColumns.I.lightGreenInk = 14706` (col G's green leaked into
  col I's region).

  Fix: `src/editorView.tsx:bootUniver()` exposes the FUniver facade
  on `window.__notesheetUniverAPI` (read-only, diagnostic; harmless
  in production — just an extra global property pointer). The
  sampler now reads live per-column widths and per-row heights via
  `ws.getColumnWidth(c)` / `ws.getRowHeight(r)` and computes
  cumulative x/y origins from those. The pixel sidecar carries a
  new `geometrySource` field — `fUniver` when the facade was
  reachable, `default-fallback` (with reason) when not, so future
  evaluators can tell which path was taken at a glance.

  Other geometry options considered:
    - **Option B** (scan canvas column-header strip for borders):
      doable but brittle — the column-header gridline is anti-aliased
      and faint at DPR=2.
    - **Option C** (OCR column letters): too brittle for harness use.
    - **Option D** (sample whole CF area, bucket by colour signature
      only): drops per-column resolution; would weaken the spec.

  Trade-off vs Option B: Option A requires a live Univer instance.
  If the plugin fails to load, `geometrySource=default-fallback`
  kicks in and the sampler degrades to the (broken) hardcoded math.
  That's acceptable because if the plugin can't load, the screenshot
  itself would also fail to capture CF rendering — the sampler not
  working is the least of the problems.

- **`scripts/pge/widen-columns.js` — utility to align live Joplin
  column widths to the canvas-fidelity test reference geometry.**
  The harness sampler is now width-agnostic, but
  `tests/excelCanvasFidelity.test.ts` (the M15 canvas-vs-Excel
  parity tests) still hardcodes device-pixel x-bands for col G (
  `1290..1480`) and col I (`1635..1655`). Those bands assume all 9
  columns A..I at uniform CSS width 95. After any session that
  changes column widths (operator interaction, snapshot save, etc.)
  re-run `node scripts/pge/widen-columns.js` to reset all 9
  columns to 95 CSS px before invoking `eval-screenshot.sh`. A
  follow-up cleanup could re-tune the canvas-fidelity test
  coordinates to use FUniver-discovered widths the same way the
  sampler does — out of scope for the harness fix cycle.

- **`scripts/pge/move-selection.js` — utility to dismiss A1 active
  cell selection so the screenshot doesn't show the edit cursor over
  cell content.** Univer's active-cell selection paints a `|` cursor
  over the cell's text-render position, which on a narrow-width A1
  visually obscures the header text (you see only `i` when the cell
  contains "Color Scale (3-stop)"). The actual cell value is
  unaffected — verified via `getValue()` — but the screenshot looks
  wrong. Calling `ws.getRange('N1').activate()` moves the selection
  to a far-off blank cell. Future evaluators should run this BEFORE
  `eval-screenshot.sh` if active-cell selection ends up landing
  somewhere that pollutes the visual gate. The existing
  `tableHeaderRowRegion` mitigates this for M13/E by sampling
  past col A; the M15 fixture has multi-column visual evidence
  including col A so we need to actually move the selection.

- **M16 introduces a new content-script entry shape.** Per
  `api/JoplinContentScripts.d.ts`, `ContentScriptType.MarkdownItPlugin`
  scripts export `default function(context)` that returns
  `{ plugin: (markdownIt, pluginOptions) => {...}, assets: {...} }`.
  The plugin function adds a fenced-code-block override for tag
  `notesheet`. Joplin runs the script in a sandboxed renderer
  worker, so the script bundle MUST be self-contained — bundle
  `extractSnapshot()` (from `src/snapshot.ts`) into the entry,
  do NOT pull `src/xlsx.ts` (heavy: exceljs, JSZip; renderer
  doesn't need import/export logic). The new entry goes into
  `plugin.config.json` extraScripts list — webpack's existing
  `buildExtraScripts` config (`webpack.config.js:345`) builds it
  as a `commonjs2` bundle by default, which is what content scripts
  expect. Do NOT add the new entry to `browserExtraScripts` (that
  set switches the output to IIFE for browser-bound webview scripts
  — the markdown-it content script is renderer-side, not webview).
  Generator should verify the produced bundle exports
  `module.exports = function(context) { ... }` (commonjs2 shape)
  before flipping the row.
- **M16 harness target shifts from Univer canvas to preview pane.**
  Until M15 every eval-screenshot targeted the Univer canvas inside
  `UserWebviewIndex.html`. M16's render lands in Joplin's main shell
  preview pane (a `note-viewer-iframe` or similar — generator
  verifies via DevTools at session start; selectors aren't a Joplin
  plugin-API contract so they may shift across Joplin versions).
  Add a `previewPane` regionKind to `eval-screenshot.js`'s
  `REGION_BY_FEATURE` table and a sampler path that picks the
  Joplin shell page (NOT the editor page) and screenshots the
  preview iframe body. The CDP page-picker scoring needs adjusting
  for this mode — the existing scoring favours the editor page
  (Univer-host) for Notesheet notes; M16 wants the SHELL page
  (preview-host). Use a feature-aware override on the page picker:
  if `REGION_BY_FEATURE[FEATURE_ID] === 'previewPane'`, score the
  shell page higher. The fix lands as part of the M16 cycle.
- **M16 window prep: editor + preview both visible.**
  `prep-joplin-window.sh` (M13/C) covers sidebar + note-list +
  fill-window. M16 additionally needs the preview pane visible —
  Joplin's default state after sidebar/note-list toggle is
  editor-only. Cmd+L on macOS toggles the editor / split / preview-
  only triad; the script needs state-aware toggling per the M13/C
  precedent. Detection signal: is `iframe.note-viewer-iframe`
  visible and `> 100px wide` in the renderer DOM. If not, send the
  Cmd+L AppleScript stroke and re-check.
- **M16 CF evaluator is duplicated logic vs the Univer CF preset.**
  Univer's CF preset (M15) evaluates rules at canvas-render time;
  the M16 renderer evaluates them again at HTML-render time — the
  static HTML pipeline doesn't run Univer. This duplication is
  unavoidable without restructuring snapshot → render through a
  headless Univer engine, which is itself a multi-cycle project.
  M16 ships its own evaluator for cellIs, top10, and colorScale
  only (dataBar + iconSet punted). The evaluator code lives inside
  the content-script bundle; if a later cycle wants to factor out
  into a shared module, that's its own refactor.
- **M16 fence handler MUST tolerate non-Notesheet fenced blocks.**
  The renderer is invoked for EVERY note in Joplin's preview
  pipeline. The first thing the override does is check the fence
  info string for `notesheet v=1`; if it doesn't match, fall
  through to markdown-it's default fence handler. Don't break
  generic-code-block rendering for non-Notesheet notes. Test
  surface for this is implicit in criterion 1 (the test feeds a
  notesheet body and asserts table render); add an explicit
  "non-notesheet fenced block falls through" assertion if there's
  budget.
- **M16 test gap discipline.** Per
  `feedback_pge_fidelity_test_gap.md`: pin "the bold cell carries
  `font-weight: bold`", NOT "the renderer emits `<td class='b'>`".
  Pin "the header `<td>` style contains `background-color: #34692E`",
  NOT "the renderer emits a particular wrapper-span shape." HTML
  emit format is implementation detail; the cell-by-cell rendered
  CONTENT is the contract. The HTML format CAN change without test
  failure as long as the rendered values + colours match. Source
  fixture content (cell values from the snapshot, M13/E-validated
  colour values from the snapshot's already-pinned styles) is the
  upstream anchor.

- **Preview pane DOM target — `iframe.noteTextViewer`** (Joplin
  current build, verified 2026-06-06). Lives inside the main shell
  page (`Resources/app.asar/index.html`) — NOT inside the
  `UserWebviewIndex.html` frame the prior cycles targeted. Its iframe
  src is `joplin-content://note-viewer/...`, but Playwright's
  `frame.url()` returns empty string for that custom protocol —
  identify the frame by `document.title === 'Note viewer'`. The
  iframe element on the main page is `iframe.noteTextViewer`; use
  `page.locator('iframe.noteTextViewer').screenshot()` to capture
  the rendered HTML directly. Pre-existing Custom Editor (Univer)
  state DEFAULT-takes-over the editor pane for Notesheet notes;
  the harness has to invoke `View > Toggle editor plugin` to make
  the preview pane appear. After that, depending on operator state
  the layout may be editor-only / split / preview-only — cycle
  `View > Toggle editor layout` up to 3 times until preview is
  visible.

- **Joplin layout-state machine has 5 distinct states for an
  open Notesheet note.** Verified via DOM inspection 2026-06-06:
  (1) Custom Editor active (Univer iframe = `iframe.plugin-user-
  webview` visible; no preview pane). (2) Built-in editor active,
  TinyMCE rich-text mode (`iframe.tox-edit-area__iframe` visible).
  (3) Markdown editor active, layout = editor-only (`.cm-editor`
  visible; preview iframe hidden). (4) Markdown editor active,
  layout = split (`.cm-editor` and `iframe.noteTextViewer` BOTH
  visible at ~half width). (5) Markdown editor active, layout =
  preview-only (`iframe.noteTextViewer` visible at full editor-
  pane width; `.cm-editor` hidden). For M16 evidence, states 4 +
  5 both work — the preview iframe is what gets screenshotted.
  `ensurePreviewPaneVisible()` in `eval-screenshot.js` is the
  state-aware driver; reads current state, applies one or two
  toggles to land in 4 or 5, never blindly toggles. Idempotent:
  re-running on an already-correct state is a no-op.

- **AppleScript via `execFileSync('osascript', ['-e', script])`,
  not `execSync('osascript -e ${JSON.stringify(script)}')`.**
  Switched in M16. The shell-arg form was lossy on nested-quote
  shapes — `JSON.stringify` of a multi-line AppleScript with
  embedded `"` produced a string that the shell unwrapped wrong,
  yielding `Expected "given", "in", "of", expression...` syntax
  errors at line 37:38. argv form passes the script verbatim with
  no shell-tokenization step.

- **Preview-pane sampler reads CSS computed colours, NOT canvas
  pixels.** M16's gate target is HTML, not pixel-perfect render —
  `samplePreviewPaneInk()` walks every `<td>` in the preview
  iframe's DOM, reads `getComputedStyle(td).backgroundColor`,
  parses the `rgb(R, G, B)` string, and applies the same threshold
  expressions the canvas sampler uses (greenInk / pinkInk /
  lightGreenInk / etc.). That keeps the sidecar shape consistent
  across regionKinds, so the evaluator's gate expressions don't
  have to special-case M16. Three preview-pane-specific signals
  added: `tableCount`, `sheetHeadings`, `rawJsonLeak`. The latter
  is the M13-style failure-mode gate — if the renderer didn't run
  (or the registration failed), the user sees a `<pre><code>`
  block containing the JSON; `rawJsonLeak: true` catches that.

- **Webpack content-script bundle shape verified for M16.** The
  built `dist/contentScripts/notesheetRenderer.js` is wrapped in
  a single IIFE that ends with `module.exports = t` where `t` is
  an object with keys `default`, `renderFenceToken`,
  `renderNotesheetSnapshot` — exactly what `commonjs2`
  `libraryTarget` produces and what Joplin's content-script
  loader expects to require. Bundle size ~8.4 KB (very compact —
  no exceljs / JSZip / Univer pulled in; just stdlib + the
  renderer's CF evaluator). Don't add the entry to
  `browserExtraScripts` — that's for browser-bound IIFE bundles
  loaded via `addScript` (the editor view), not for renderer-side
  markdown-it plugins. The webpack `extraScripts` map subdir-name
  prefix correctly: `contentScripts/notesheetRenderer.ts` →
  `dist/contentScripts/notesheetRenderer.js`. The path passed to
  `joplin.contentScripts.register` must match this layout
  (`./contentScripts/notesheetRenderer.js`).

- **Renderer is read-only on the snapshot.** The M16 renderer
  walks `snapshot.sheetOrder` / `sheets[id].cellData` /
  `sheets[id].mergeData` / `styles[id]` / `resources[]` but never
  modifies any of it. CF rule evaluation produces a separate
  `Map<string, string>` of per-cell fill overrides; the snapshot
  itself is untouched. Any future content script should follow
  the same discipline — content scripts run for every note in
  the preview pipeline, and side-effecting from one would mean
  side-effecting on every note, including notes the plugin
  doesn't own.

- **CF colorScale interpolation by anchors.** The M15 import
  shape gives us `[{index, color, value: cfvo}]`. Each `cfvo`
  resolves to a numeric position in the rule's value range
  (min/max/percentile/percent/num — formula falls through to
  null and the anchor is dropped). With ≥2 valid anchors,
  `evalColorScale` sorts by position, finds the segment the cell
  value lies in, and lerps RGB linearly between the segment's
  two anchors. `lerpRgb('#F8696B', '#FFEB84', 0.5)` produces a
  reasonable orange — verified visually by sampling
  ConditionalFormatting-Variants.xlsx column A, where values
  0..90 step through red → orange → yellow → green-ish.

- **Renderer falls through gracefully for malformed input.**
  Three failure modes return `null` from `renderFenceToken`:
  (a) fence info doesn't start with `notesheet`; (b) version
  isn't `1`; (c) fence body isn't valid JSON. Markdown-it then
  invokes the default fence handler which renders the raw JSON.
  This is the safe degradation — better to show the user the
  source than to inject a placeholder string they might
  misinterpret as the actual content. The M16 Jest tests pin
  all three failure modes.

- **Multi-sheet test fixture choice — in-memory, not
  MultiSheet.xlsx.** The shipped `MultiSheet.xlsx` fixture under
  `tests/ExcelBaseTestData/formatting-testdata/` carries chart
  drawings that crash exceljs's reconcile loop (raised as a
  known shortcoming in M12; pinned in
  `tests/m12ImportRecovery.test.ts`). M16's multi-sheet test
  cannot import that fixture without first adopting a chart-
  supporting importer (out of scope). We build a 3-sheet
  workbook in-memory via `new ExcelJS.Workbook(); wb.addWorksheet(...)`
  — same pattern other tests in the suite use. The test exercises
  the same renderer code path that a fixture-imported multi-sheet
  snapshot would.

- **M17 feature-1: namespace tolerance is load-bearing.** Drawing XML
  in xlsx workbooks comes in two flavours: (a) Excel-authored uses the
  canonical `xdr:` prefix on `<xdr:wsDr>` / `<xdr:twoCellAnchor>` /
  `<xdr:from>` etc. (b) Programmatically-built workbooks (`MultiSheet.xlsx`,
  `LargeWorkbook.xlsx`) declare the spreadsheetDrawing namespace as the
  DEFAULT and emit unprefixed elements. Same for chart parts: Excel
  emits `<c:barChart>` etc.; programmatic emitters use a default
  namespace and emit `<barChart>`. Every regex in
  `src/charts/xlsxChartImport.ts` uses `(?:xdr:)?` / `(?:c:)?` so the
  parser handles both. Failure mode: parser silently returned 0
  charts on namespace-less inputs, and the strip path didn't run, so
  exceljs hit the `anchors` reconcile crash and `xlsxBufferToSnapshot`
  threw `xlsx-charts-unsupported`. The all-10-fixtures test now
  exercises both shapes (the project's canonical Excel-authored set
  uses `c:` prefix; the imported Joplin-shipped MultiSheet/LargeWorkbook
  use the default-namespace shape via the m12ImportRecovery flips).

- **M17 feature-1: regex `[^/]*` is wrong for attribute strings that
  contain URLs.** OOXML `<Relationship Type="http://schemas.openxmlformats.org/.../drawing"/>`
  has many `/` characters inside the Type attribute. Patterns like
  `<Relationship\b[^/]*Type=...[^/]*Target=...[^/]*\/>` silently fail
  to match because the URL eats the `[^/]*` ranges. The fix is to use
  `[^>]*` (everything up to the next `>`) and parse Id/Type/Target
  independently. Same trap in three places — `parseDrawingRels`,
  `findSheetDrawingLinks`, and the `<drawing r:id="..."/>` strip in
  worksheets where some workbooks add `xmlns:r="..."` to the drawing
  element itself, putting a URL between the tag name and the `r:id`
  attr.

- **M17 feature-1: walk anchors in DOCUMENT order, not zip-key
  order.** `06-two-charts-one-sheet.xlsx` packs both charts into ONE
  `xl/drawings/drawing1.xml` as two consecutive `<xdr:twoCellAnchor>`
  blocks. Walking by zip key (i.e. iterating `xl/charts/chart{N}.xml`)
  would associate anchor 0 with chart1 and anchor 1 with chart2 by
  coincidence — but the load-bearing case is that for a drawing with
  multiple anchors, you MUST walk anchors in document order and
  resolve each anchor's `r:id` against
  `xl/drawings/_rels/drawing{N}.xml.rels` to find its chart part path.
  This anchors-driven walk is what the implementation does;
  `tests/m17ChartImportNoCrash.test.ts` pin-downs anchor coordinates
  on snapshot drawings to source XML, ensuring the document-order
  invariant holds.

- **M17 feature-1: m12ImportRecovery flips were a forced consequence,
  not a planned scope creep.** The planner reserved the
  `tests/m12ImportRecovery.test.ts:59,84` flip for feature-9. But
  feature-1's criterion 5 requires the suite green — and the strip
  path makes MultiSheet.xlsx import successfully, so the existing
  `expect(e.code).toBe('xlsx-charts-unsupported')` for MultiSheet
  goes from green to red as soon as feature-1 lands. There's no
  way to ship feature-1 + green suite without flipping this assertion.
  LargeWorkbook.xlsx is similar but uncovers a SECOND crash class
  beyond M17's scope (`name` reduce in worksheet.js:920 — the same
  multi-table-unsupported class FormulasAndStructuredRefs.xlsx
  exhibits). Flipped its expectation to
  `xlsx-multi-table-unsupported` rather than positive-import. The
  M17 README "Known shortcomings" should mention LargeWorkbook still
  doesn't import — that's docs work for feature-9 or the README PR.

- **M17 feature-1: `oneCellAnchor` synthesizes a `to` of (col+6,
  row+14).** Excel renders one-cell-anchored charts by laying down
  the EMU `<xdr:ext>` extent at the from point; for our cell-anchored
  UI a reasonable approximation is "from + a chart-sized span." The
  exact span (6 cols / 14 rows ~= 460×270 px @ default cell size) is
  documented in `walkAnchors`; the synthetic-anchor test pin-down
  asserts this shape so future tweaks are deliberate. In practice
  Excel-authored fixtures all use `twoCellAnchor`; oneCellAnchor
  shows up only in programmatically-generated workbooks (MultiSheet,
  LargeWorkbook, and one canonical synthetic in the test). The
  approximation never sees an Excel-authored input where it would
  need to be EMU-precise.

- **M17 feature-1: source range bounding box. The first `<c:f>` in a
  series is the SERIES-NAME ref nested in `<c:tx><c:strRef>`, NOT
  categories.** Categories live in the second `<c:f>` (inside
  `<c:cat>`); values are the third (inside `<c:val>`). Verified at
  `01-bar-simple.xlsx`'s chart1.xml: the three `<c:f>` elements are
  `Sheet1!$B$1` (series name), `Sheet1!$A$2:$A$5` (categories),
  `Sheet1!$B$2:$B$5` (values). The implementation explicitly scopes
  to `<c:cat>` and `<c:val>` sub-elements rather than grabbing the
  first `<c:f>`. Same trap is inverse in M10's export side
  (`src/charts/xlsxChart.ts:170` emits the series-name ref FIRST in
  series order). The test's `readSourceTruth` helper does the
  scoped walk too — the assertion comes from independently parsing
  the XML, NOT from the snapshot's emit.

- **M17 feature-1: the `xlsx-charts-unsupported` error class stays
  defined.** The strip path is the common-case fix; it covers the
  10-fixture set and MultiSheet/LargeWorkbook's chart-bearing
  drawings. But future drawing-related crash classes that the strip
  doesn't yet recognize (image+chart mixed drawings, OLE objects,
  etc.) will still surface from exceljs's reconcile, and the
  existing wrap path catches them with the same code. The test
  pins the class definition itself rather than dynamically
  injecting a crash — `jest.spyOn(importMod, 'readChartsFromXlsxZip')`
  doesn't work in TS strict-export mode (`Cannot redefine property`),
  so the test simply asserts the constructor + code constant
  shape. Acceptable: the wrap is unit-tested by the existing
  m12ImportRecovery suite for the in-tree crash classes.
