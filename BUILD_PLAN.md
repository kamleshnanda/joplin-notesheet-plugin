# Notesheet PGE — build plan (M17: Chart import from `.xlsx` + chart rendering in HTML / preview-pane export)

> **Cycle context.** Sixth real-feature cycle. M13/A–E shipped via this
> harness (PRs #19–#23), M14 was an explicit NO-GO (`docs/m14-sheetjs-spike.md`),
> M15 conditional formatting shipped (PRs #26 + #27), and M16 snapshot →
> HTML rendering shipped (PR #28). M16 deferred two follow-ups that
> M17 picks up:
>
> 1. **The original M16 brief listed "charts don't render in HTML
>    export" as an explicit gap.** Static charts in the M16 content
>    script were punted because the import-side path
>    (`xlsxBufferToSnapshot`) couldn't read chart drawings yet — a
>    snapshot built from an Excel-authored chart workbook would have
>    no `SHEET_DRAWING_PLUGIN` resource to walk in the first place.
>    M17 closes the import gap AND adds inline SVG emit in the same
>    cycle so both halves arrive together.
> 2. **`xlsxBufferToSnapshot` currently throws
>    `NotesheetImportError('xlsx-charts-unsupported', ...)`** for any
>    workbook whose drawings exceljs's reconcile loop trips on (`anchors`
>    crash class, classified at `src/xlsx.ts:1745`). The README's
>    "Known shortcomings" line says `M17 addresses this`; this is the
>    standing promise.
>
> **Two-layer fidelity test pattern remains the project default for
> "match Excel" features.** M13/E established it (snapshot fidelity vs
> source XML, canvas fidelity vs Excel reference PNG); M15 extended both
> layers; M16 ran HTML-content fidelity vs source-fixture content
> (Jest-only). M17 inherits both anchor disciplines:
>
> - **Excel-authored fixture tests** (criteria 1, 2, 5) anchor to the
>   source Excel chart XML inside the `.xlsx`, parsed independently
>   via stdlib `JSZip` + regex / DOMParser. Expected values come from
>   the source XML, NOT from "what `xlsxBufferToSnapshot` happens to
>   emit." (Per `feedback_pge_fidelity_test_gap.md` — the M13/E mistake
>   was asserting our own emit and missing the actual rendered
>   colour.)
> - **Programmatic round-trip tests** (criterion 6) anchor to the
>   ORIGINAL in-memory snapshot we author. Build snapshot →
>   `snapshotToXlsxBuffer` → `xlsxBufferToSnapshot` → assert RE-IMPORTED
>   chart matches ORIGINAL chart. Tests Notesheet-emit ↔ Notesheet-
>   import inverseness, NOT Excel parity.
> - **HTML / SVG export tests** (criterion 7) anchor to structural SVG
>   element counts derived from the input snapshot — `<rect>` count
>   == data length, `<path>` count == slice count, ±3° on pie sweep
>   angles. Not pinned to byte-string output; renderer can change
>   emit style without breaking the test as long as counts match.
>
> **Test-suite no-regression discipline (criterion #10).** This cycle's
> evaluator runs `git diff tests/` and confirms the ONLY changes in
> `tests/` are: NEW files (`tests/m17*.test.ts` family) AND the two
> flipped pin-down lines in `tests/m12ImportRecovery.test.ts:59,84`.
> No content edit to any other existing test file. The eighteen
> existing test files listed in operator criterion #10 stay untouched
> and pass. This prevents the "262 → 287 by deleting 5 existing tests
> and adding 30 new ones" failure mode that almost slipped through
> M13/E.
>
> **Approach choice for the import-side architecture.** Two viable
> approaches were called out in the operator brief: (A) catch the
> exceljs throw, parse charts from the zip, retry with a drawing-
> stripped buffer; (B) add a NEW pre-load zip-direct reader
> `readChartsFromXlsxZip`, build a drawing-stripped buffer, then
> pass the stripped buffer to exceljs.
>
> **This plan picks Approach B.** Rationale: pre-stripping is the
> cleanest way to keep exceljs's reconcile loop functional for
> everything that's NOT a chart drawing (cells, tables, CF, theme,
> hyperlinks); avoids a retry path that would mask future crash
> classes we haven't classified yet; keeps `NotesheetImportError`
> defined for genuine non-chart crashes.
>
> **Important — call-ordering distinction the operator brief
> conflated.** The four existing zip-direct readers
> (`readTablesFromXlsxZip`, `readThemeFont`, `readThemeClrScheme`,
> `readNamedHyperlinkCells`) are NOT pre-load — they all run AFTER
> `await wb.xlsx.load(buffer)` (verified at `src/xlsx.ts:1764-1779`).
> They work because exceljs's load already succeeded for non-chart
> workbooks; they're parallel parsers reading the SAME successfully-
> loaded buffer, not pre-strippers. M17's `readChartsFromXlsxZip`
> is fundamentally different: it MUST run BEFORE `wb.xlsx.load`
> because the load currently throws on chart drawings. The flow is:
>
> 1. Read chart parts from the ORIGINAL buffer.
> 2. Build a STRIPPED buffer in memory by removing every chart
>    drawing reference: `xl/drawings/drawing*.xml`,
>    `xl/drawings/_rels/drawing*.xml.rels`, the `<drawing r:id="...">`
>    element from each `xl/worksheets/sheet{N}.xml`, the matching
>    drawing rel from `xl/worksheets/_rels/sheet{N}.xml.rels`, and
>    the chart Override entries from `[Content_Types].xml`.
> 3. Call `wb.xlsx.load(STRIPPED_BUFFER)` — succeeds cleanly because
>    exceljs's reconcile loop no longer encounters chart drawings.
> 4. Continue calling the EXISTING post-load readers
>    (`readTablesFromXlsxZip` etc.) against the ORIGINAL buffer
>    (their inputs do not include chart parts; they're unaffected
>    by the strip).
>
> Document the strip recipe as a single helper
> `stripChartPartsFromZip(buffer): Promise<Buffer>` with its own
> unit test that round-trips a non-chart-bearing fixture buffer
> through it and asserts byte-equality (idempotency) plus a
> chart-bearing fixture and asserts post-strip
> `[Content_Types].xml` no longer mentions chart Overrides AND
> post-strip `xl/worksheets/sheet1.xml` no longer carries
> `<drawing r:id=...>`.
>
> **Forward-compat note.** A future zip-direct reader that walks
> `<drawing>` elements (e.g. for image import in M18+) MUST decide
> whether to consume the ORIGINAL or STRIPPED buffer — pick at
> implementation time. Document the choice in `xlsx.ts` comments
> at the strip call-site. Add a unit test asserting that all 4
> existing post-load readers return identical output before and
> after the strip step on each of the 10 chart fixtures.
>
> **`SHEET_DRAWING_PLUGIN` snapshot resource shape (from
> `src/charts/xlsxChart.ts:310` — `readChartsFromSnapshot`).** The
> resource's `data` field is a JSON-stringified map:
> ```
> { [subUnitId]: { data: { [drawingId]: ISheetDrawing }, order: string[] } }
> ```
> Each `ISheetDrawing` filtered by `componentKey === 'NotesheetChart'`
> carries:
> ```
> {
>   componentKey: 'NotesheetChart',
>   data: {
>     chartId: string,
>     type: 'bar' | 'line' | 'pie' | 'doughnut',
>     title: string,
>     sourceRange: { startRow, endRow, startColumn, endColumn },
>     labels: string[],
>     datasets: Array<{ label?: string, data: number[] }>,
>   },
>   axisAlignSheetTransform: {
>     from: { column, columnOffset, row, rowOffset },
>     to:   { column, columnOffset, row, rowOffset },
>   },
> }
> ```
> M10's export pipeline (`readChartsFromSnapshot`) already documents
> this shape. M17's import path writes the SAME shape directly into
> the snapshot's `resources` array — we are outside Univer's command
> bus during snapshot construction, so we cannot route through
> `ICommandService.executeCommand('sheet.command.add-floating-dom', ...)`.
> Direct resource write is the right path.
>
> **SVG palette parity.** The static SVG renderer in the content script
> MUST use the same `CHART_PALETTE` array used by the live Chart.js
> renderer (`src/charts/extractData.ts:11`) so a single chart is the
> same colour in the editor and in the exported PDF. The planner
> picks: **duplicate the palette in the content script with a comment
> pointing at the source.** Rationale: the content-script bundle is a
> separate `commonjs2` output target (M16 confirmed it ships at ~8.4 KB
> with NO Chart.js / JSZip / exceljs in it); pulling
> `src/charts/extractData.ts` directly drags `chart.js` types and the
> data-extraction logic into the bundle. An **8-entry hex array**
> (`['#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6', '#ec4899',
> '#06b6d4', '#84cc16']` — verified at `src/charts/extractData.ts:11-20`)
> with a comment (`// MUST match src/charts/extractData.ts:CHART_PALETTE`)
> is simpler and stays under M16's bundle target.
>
> **Bundle size discipline.** M16 shipped at ~8.4 KB. Adding chart-to-
> SVG rendering grows it. Don't pull Chart.js, D3, or any visualization
> library into the content script — hand-author SVG primitives (`<rect>`
> for bars, `<polyline>` / `<path>` for lines, `<path>` with arc
> commands for pie / doughnut slices). **Target stays under ~20 KB after
> M17 lands.** The planner sets this as a soft gate; if the bundle
> creeps higher, the generator flags it before flipping the row.
>
> **Multi-feature cycle.** M17 decomposes into 8 user-observable
> features. Earlier features establish the import path; mid-features
> verify type fidelity, live-update wiring, and round-trip robustness;
> later features cover the HTML/SVG export side and the live-fixture
> PGE smokes. Each feature is independently evaluator-gradable from a
> fresh context.

## Operator ask

(See `OPERATOR_ASK.md` for the full brief — 12 numbered acceptance
criteria, extensive Out-of-scope list, three suggested fixture sets,
detailed Related-risks notes.) The user-facing change: a workbook
authored in Excel that contains a bar / line / pie / doughnut chart
opens in Notesheet without throwing. The chart shows up live in the
Univer editor as a Chart.js float-DOM at its anchor position; the
chart re-renders when the user edits the source range; the chart
round-trips through Notesheet's M10 export back to Excel; the chart
also renders as inline SVG in Joplin's preview pane and in the M16
PDF / HTML export. The README's "Known shortcomings — Workbooks with
chart drawings… M17 addresses this" entry can be removed.

## Harness extension chosen for this cycle

The PGE harness needs three additions for M17:

- **`floatDomChart` regionKind** in `eval-screenshot.js`'s
  `REGION_BY_FEATURE` table. Until M17 every screenshotted note had
  Univer rendering content as a canvas; charts are float-DOM (HTML
  overlays on top of the canvas). The existing canvas-screenshot path
  captures the canvas underneath the float-DOM but NOT the float-DOM
  itself. The new regionKind screenshots the outer Univer container
  (`.univer-render-wrapper` or equivalent — selector is verified at
  generator time) instead of just the canvas. Reference: M16's
  `previewPane` regionKind precedent for a non-canvas region.
- **`scripts/pge/import-fixture.sh` fixture-path expansion.** The
  script currently hardcodes
  `tests/ExcelBaseTestData/formatting-testdata/`. Expand the fixture-
  path search to ALSO accept `tests/fixtures/charts/` so the chart
  fixtures can be imported via the same harness path.
- **`samplePreviewPaneInk()` extension** — a new `inlineSvgCount`
  signal counting `<svg>` elements inside the preview iframe DOM. The
  M17 preview-pane gate is `inlineSvgCount ≥ 1` for chart-bearing
  fixtures. The existing `tableCount` / `sheetHeadings` / `rawJsonLeak`
  signals stay unchanged.

Generator owns these as part of feature-3 (editor-canvas smoke) and
feature-8 (preview-pane smoke) respectively. They are part of the
harness, not separate features themselves.

## Features

### feature-1-m17-chart-import-no-crash

**Spec**

When a user imports a chart-bearing `.xlsx` (any of the 10 hand-
crafted fixtures under `tests/fixtures/charts/01-*` through
`10-*.xlsx`) via Notesheet's existing "Import .xlsx" path, the
import does NOT throw `NotesheetImportError('xlsx-charts-unsupported',
...)`. The returned snapshot carries a `SHEET_DRAWING_PLUGIN`
resource with at least one chart drawing per fixture, in the shape
`readChartsFromSnapshot()` already consumes (M10 export side). Each
chart's `chartId`, `type`, `sourceRange`, `anchor`, `labels`, and
`datasets` are populated from the source `xl/charts/chart{N}.xml` /
`xl/drawings/drawing{N}.xml` parts independently of exceljs's chart
parser (which is thin / uneven and is what crashes today).

The implementation extends the existing zip-direct reader pattern with
a new `readChartsFromXlsxZip(buffer)` that runs BEFORE
`wb.xlsx.load(buffer)`. The chart parts are read first, then a
drawing-stripped buffer is constructed in-memory and passed to
exceljs (which loads cleanly). The `xlsx-charts-unsupported` error
class stays defined for any future drawing-related crash class we
haven't classified.

**Acceptance criteria**

The evaluator must verify ALL of the following from a fresh context:

1. **Each of the 10 chart fixtures imports without throwing.** A new
   Jest test `tests/m17ChartImportNoCrash.test.ts` loads each file
   under `tests/fixtures/charts/01-bar-simple.xlsx` through
   `10-bar-with-trendline.xlsx` via `xlsxBufferToSnapshot()` and
   asserts:
   - No exception thrown.
   - Returned snapshot is a valid object with non-empty `sheetOrder`.
   - Snapshot has a `resources` array containing an entry whose
     `name === 'SHEET_DRAWING_PLUGIN'`.
   - Parsing that resource's `data` JSON via the M10 reader pattern
     (`readChartsFromSnapshot(snapshot)`) returns at least 1 chart
     drawing for every fixture except `06-two-charts-one-sheet.xlsx`,
     which returns exactly 2.
   - **Multi-anchor walk discipline.**
     `06-two-charts-one-sheet.xlsx` packs BOTH charts into a SINGLE
     `xl/drawings/drawing1.xml` with two `<xdr:twoCellAnchor>`
     blocks (verified via `unzip -p
     tests/fixtures/charts/06-two-charts-one-sheet.xlsx
     xl/drawings/drawing1.xml`). For each
     `xl/drawings/drawingN.xml`, the implementation MUST walk
     its `<xdr:twoCellAnchor>` / `<xdr:oneCellAnchor>` /
     `<xdr:absoluteAnchor>` elements **in document order**. For
     each anchor, look up its `r:id` against
     `xl/drawings/_rels/drawingN.xml.rels` to resolve the chart
     part path, then parse THAT chart's xml. Walking
     `xl/charts/chart{N}.xml` in zip-key order is **wrong** — it
     mis-aligns charts with anchors when one drawing has multiple
     anchors. Pin-down: assert
     `06-two-charts-one-sheet.xlsx`'s drawing has chart1.xml
     mapped to anchor 0 (the FIRST `<xdr:twoCellAnchor>` block in
     `drawing1.xml`) and chart2.xml mapped to anchor 1 (the
     SECOND block). The exact `<xdr:from>`/`<xdr:to>` coordinates
     for each anchor are read from the source XML for assertion —
     **NOT hardcoded in the test** (the M13/E our-own-emit trap;
     parse them at test time so a future re-saved fixture flows
     through).
2. **Chart fields match source XML — anchored upstream.** The same
   test parses `xl/charts/chart{N}.xml` and `xl/drawings/drawing{N}.xml`
   directly from the fixture's zip (using `JSZip` + regex or
   stdlib XML reading — same pattern as `tests/util/pngSampler.ts`,
   pure-stdlib, no new deps) and asserts the snapshot's chart drawing
   matches the source XML on:
   - `type` (parsed from `<c:barChart>` / `<c:lineChart>` / `<c:pieChart>`
     / `<c:doughnutChart>` element name).
   - `sourceRange` — parsed from the `<c:f>` CHILD of
     `<c:cat><c:strRef>` (categories ref) AND the `<c:f>` CHILD of
     `<c:val><c:numRef>` (values ref). The bounding box of cat-ref
     unioned with val-ref across all `<c:ser>` is the source range.
     **The FIRST `<c:f>` in the chart XML is the series-name ref
     nested inside `<c:tx><c:strRef>`, NOT the categories ref —
     verified by `unzip -p tests/fixtures/charts/01-bar-simple.xlsx
     xl/charts/chart1.xml | grep '<c:f>'` (returns `Sheet1!$B$1`,
     `Sheet1!$A$2:$A$5`, `Sheet1!$B$2:$B$5` in that order — the
     categories ref is the SECOND, not the first). Same shape in
     03-pie-single, 09-bar-percent-axis, 07-chart-cross-sheet.**
     Row/col indices decoded from the `Sheet!$A$1:$B$5` ref string.
   - `anchor` (parsed from the drawing XML's
     `<xdr:from>`/`<xdr:to>` `<xdr:col>`/`<xdr:row>` numbers).
   - `chartId` is non-empty (we don't pin a specific id format —
     either a synthetic uuid or a normalized form from the drawing
     part is acceptable).
   The expected values come from the source XML, NEVER from what
   `xlsxBufferToSnapshot` produces. Per `feedback_pge_fidelity_test_gap.md`.
3. **Synthetic robustness cases — drawing-shape coverage beyond
   the 10 hand-crafted fixtures.** The 10 in-tree fixtures all use
   simple `<xdr:twoCellAnchor>` + non-empty `<c:cat>` + the standard
   drawing-rels shape. The Jest test additionally covers (built
   in-memory in the test, NOT checked-in fixtures):
   - **`<xdr:oneCellAnchor>` drawing.** A synthetic .xlsx with a
     chart anchored via `oneCellAnchor` (one corner, plus an
     EMU `<xdr:ext>` extent). Asserts: import does NOT throw;
     produces 1 chart drawing with `anchor.fromCol/fromRow`
     populated and either an inferred `to` (e.g. `from + ext`
     converted to cell coords) or null `to` — pin one shape and
     document the choice in `## Notes`.
   - **Missing chart-target rels.** A synthetic .xlsx with a
     drawing whose `_rels/drawing1.xml.rels` points at a
     `chartZ.xml` that doesn't exist in `xl/charts/`. This is
     the original 'anchors' crash class exceljs's reconcile loop
     trips on. Asserts: import does NOT throw; produces 0 chart
     drawings (the dangling reference is silently dropped); a
     `console.warn` is logged with the missing target path.
   - **`<c:cat>` omitted (numeric-axis chart shape).** Synthetic
     chart XML with only `<c:val>` and no `<c:cat>` — Excel emits
     this for numeric scatter / X-Y plots. Asserts: import does
     NOT throw; the chart is dropped with a warn (M17's
     ChartType union doesn't include 'scatter'; document the
     drop in `## Notes`).
4. **The `xlsx-charts-unsupported` error class stays defined.** The
   `NotesheetImportError` class is still exported, the
   `xlsx-charts-unsupported` code is still a recognized literal, and
   the error path triggers on a synthetic input that produces an
   `'anchors'` exceljs crash NOT covered by the new
   `readChartsFromXlsxZip` path. The Jest test uses `jest.spyOn` on
   `readChartsFromXlsxZip` to make it return `[]`, then loads
   `MultiSheet.xlsx`, and asserts the error still throws — verifying
   we haven't deleted the safety net for future crash classes.
5. **`npm test` runs the full suite green** with this feature's new
   test file in place. No existing test file is edited as part of
   this feature's implementation. (Feature-9 below specifically flips
   the two pin-downs in `m12ImportRecovery.test.ts`.)

**Out of scope**

- Live editor rendering of the chart (covered by feature-3).
- Round-trip back to xlsx (covered by feature-5).
- Cross-sheet source range resolution (covered by feature-2 and
  feature-5).
- Chart fidelity-of-render (M13/E-style canvas-vs-Excel pixel parity)
  — Notesheet's chart palette is `CHART_PALETTE`, not Excel's per-chart
  style XML. Out of M17 entirely.

**Suggested fixture(s)**

- `tests/fixtures/charts/01-bar-simple.xlsx`
- `tests/fixtures/charts/02-line-multi-series.xlsx`
- `tests/fixtures/charts/03-pie-single.xlsx`
- `tests/fixtures/charts/04-doughnut.xlsx`
- `tests/fixtures/charts/05-bar-special-chars.xlsx`
- `tests/fixtures/charts/06-two-charts-one-sheet.xlsx`
- `tests/fixtures/charts/07-chart-cross-sheet.xlsx`
- `tests/fixtures/charts/08-drag-resized.xlsx`
- `tests/fixtures/charts/09-bar-percent-axis.xlsx`
- `tests/fixtures/charts/10-bar-with-trendline.xlsx`

(All 10 already shipped — they are the M10 export ground-truth set.)

**Related risks**

- `readChartsFromXlsxZip` runs BEFORE exceljs's load. The drawing-
  stripping step must not corrupt the zip — `[Content_Types].xml`
  references and sheet rels referencing drawing parts must also be
  cleaned up, otherwise exceljs's clean-load path may itself trip on
  dangling references. Document the strip recipe in the implementation
  and unit-test it.
- Source range cell references that target an entire column (`Sheet1!$A:$A`)
  vs a finite range (`Sheet1!$A$1:$A$5`) — the former is rejected by
  M10's export; the import should reject (drop with `console.warn`)
  the same way to keep the surface symmetric.
- `chartId` uniqueness across two charts on one sheet
  (`06-two-charts-one-sheet.xlsx`). Both must be present in the
  snapshot's drawing resource as separate entries.
- Don't symptom-patch test failures. If a typecheck or test regresses
  after wiring the import path, run `git diff package-lock.json` first
  per `feedback_dependency_hygiene.md`. The new path uses only stdlib
  + JSZip + the existing chart utilities — no new dependency should
  appear.
- **Untouchable test files**: this feature MUST NOT edit any of:
  `m13RotatedText.test.ts`, `m13RichText.test.ts`,
  `m12FixturePinDowns.test.ts`, `m12FixtureRoundTrip.test.ts`,
  `excelReferenceFidelity.test.ts`, `excelCanvasFidelity.test.ts`,
  `m15ConditionalFormatting.test.ts` (note: this file does not exist
  in the current tree — operator's criterion #10 lists it but the
  M15 cycle put its tests inside `excelReferenceFidelity.test.ts`
  and `excelCanvasFidelity.test.ts`; this rule still binds for any
  file that DOES exist),
  `m16NotesheetMarkdownRender.test.ts`, `m16FormulaSourceOfTruth.test.ts`,
  `roundTripBidirectional.test.ts`, `xlsxChart.test.ts`,
  `formattingFidelity.test.ts`, `exportTableRoundTrip.test.ts`,
  `borders.test.ts`, `hyperlinks.test.ts`, `numberFormats.test.ts`,
  `mergedCells.test.ts`, `themeFonts.test.ts`. (Several of those file
  names — `borders.test.ts`, `hyperlinks.test.ts`, etc. — also do
  not exist as separate files in the current tree; the operator's
  list is a precaution against accidental coverage deletion. The
  rule binds for whatever subset DOES exist when the generator runs.)

---

### feature-2-m17-chart-type-fidelity

**Spec**

For each supported chart type (bar / line / pie / doughnut), a fixture
of that type imports with the matching Notesheet `ChartType` literal
in the snapshot's chart drawing. Fixtures whose source type is NOT in
the supported four (radar, scatter, area, etc.) fall back to `'bar'`
in the imported drawing AND add a `meta.unsupportedSourceType` field
carrying the source type string for evaluator visibility, plus emit
a single `console.warn` line during import. The `ChartType` union
remains the existing four-type union from `src/charts/xlsxChart.ts:24`.

**Acceptance criteria**

1. **Type round-trip — Jest test `tests/m17ChartTypeFidelity.test.ts`.**
   - `01-bar-simple.xlsx` imports as a chart with `type === 'bar'`.
   - `02-line-multi-series.xlsx` imports as `type === 'line'`.
   - `03-pie-single.xlsx` imports as `type === 'pie'`.
   - `04-doughnut.xlsx` imports as `type === 'doughnut'`.
   The expected `type` value comes from independently parsing the
   source `xl/charts/chart{N}.xml`'s top-level chart-type element name
   (`<c:barChart>` / `<c:lineChart>` / `<c:pieChart>` /
   `<c:doughnutChart>`) — anchored UPSTREAM, not pinned to our own
   emit literal.
2. **Multi-series count for line.** `02-line-multi-series.xlsx`'s
   imported drawing has `datasets.length === 3` (the source XML has
   three `<c:ser>` elements; the test parses the source to confirm
   the expected count, then asserts equality).
3. **Unsupported-type fallback.** A small in-memory fixture (built
   programmatically via `JSZip` from a hand-authored `chart1.xml`
   with `<c:radarChart>` as the chart-type element, dropped into a
   minimal `xlsx` zip skeleton) imports as `type === 'bar'`. The
   chart drawing carries `meta.unsupportedSourceType === 'radar'`.
   `console.warn` is called at least once during the import (test
   uses `jest.spyOn(console, 'warn')`).
4. **`ChartType` union is unchanged.** TypeScript compile succeeds
   without widening the existing
   `export type ChartType = 'bar' | 'line' | 'pie' | 'doughnut'`
   (`src/charts/xlsxChart.ts:24`). The fallback path produces
   `'bar'` — it does not introduce a fifth string literal into the
   union.

**Out of scope**

- Implementing native rendering for radar / scatter / area charts.
- Trendline support for `10-bar-with-trendline.xlsx` — the bar chart
  imports correctly; the trendline itself is dropped silently or with
  a documented gap. (Operator's brief lists trendlines as out of
  scope.)
- Chart titles' rich-text formatting — M17 ships plain-text titles
  only. (Operator's brief lists this as out of scope.)

**Suggested fixture(s)**

- `tests/fixtures/charts/01-bar-simple.xlsx`
- `tests/fixtures/charts/02-line-multi-series.xlsx`
- `tests/fixtures/charts/03-pie-single.xlsx`
- `tests/fixtures/charts/04-doughnut.xlsx`
- Programmatically built radar-chart zip (in-memory in the test, NOT
  a checked-in fixture — the point is to exercise the fallback, not
  to support radar long-term).

**Related risks**

- The fallback path must NOT break the bar / line / pie / doughnut
  cases — each type's parser keys off the chart-type element name
  exactly. A regex like `/<c:(bar|line|pie|doughnut)Chart\b/` returns
  the matched type; everything else falls to `'bar'`.
- **Untouchable test files**: same list as feature-1.

---

### feature-3-m17-multisheet-import-editor-canvas

**Spec**

When the user imports
`tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx` (the
original "this crashes import" fixture, currently throwing
`xlsx-charts-unsupported`) via the Notesheet plugin's "Import .xlsx"
flow in Joplin, the resulting note opens in the Univer editor with:

- The cells from each sheet rendered (already works once import
  doesn't throw — the M5 + M9 paths handle this).
- The chart float-DOM visible at its anchor position.
- The chart's title text visible.
- The chart bars / lines / pie slices visible (live Chart.js render).

The PGE harness imports the fixture via `import-fixture.sh
MultiSheet.xlsx`, opens the resulting note in Joplin's Custom Editor,
and screenshots the editor canvas region INCLUDING the float-DOM
overlay (not just the canvas underneath). The new `floatDomChart`
regionKind makes this possible.

**Acceptance criteria**

1. **Import succeeds without error.** The PGE harness's
   `import-fixture.sh MultiSheet.xlsx` returns a valid Joplin note ID
   without raising the `xlsx-charts-unsupported` error.
2. **Editor canvas + float-DOM screenshot at
   `screenshots/feature-3-m17-multisheet-import-editor-canvas/eval-*.png`.**
   Captured via `bash scripts/pge/eval-screenshot.sh feature-3-m17-multisheet-import-editor-canvas`.
   The screenshot's `floatDomChart` regionKind targets the outer
   Univer container (NOT just the main canvas — that would crop
   off the chart). The screenshot shows:
   - At least one rendered table region (cells from the imported
     sheet — confirmed by visible row/col headers in the canvas).
   - At least one Chart.js float-DOM visible at its anchor (the user
     would recognise it as a chart — bars, lines, or pie slices
     visible).
   - The chart's title text visible.
   - No raw "Cannot read properties of undefined (reading 'anchors')"
     error shown anywhere in the UI.
3. **Pixel sidecar `eval-*.pixels.json` includes** at least one
   non-background colour from the `CHART_PALETTE` array (which is
   an **8-entry hex array**, verified at
   `src/charts/extractData.ts:11-20`; the test reads the palette
   directly and checks the histogram for membership). The dominant
   colour is NOT a near-white background like `rgb(255,255,255)`
   (which would mean the chart float-DOM rendered as empty /
   errored).
4. **The Jest suite stays green.** Running `npm test` after
   harness-only changes (`scripts/pge/eval-screenshot.js`,
   `import-fixture.sh`) leaves all existing tests passing. No new
   Jest test is required for this feature — the gate is the
   screenshot.
5. **Generator-evidence** at
   `screenshots/feature-3-m17-multisheet-import-editor-canvas/generator-evidence.png`
   plus its `.pixels.json` sidecar; both Read via the Read tool to
   satisfy the `verify-gate` hook.

**Out of scope**

- Pixel-perfect chart rendering parity with Excel. Notesheet's
  Chart.js renderer uses `CHART_PALETTE`, not the source workbook's
  per-chart style XML. The chart looks like a Notesheet chart,
  recognisably the same data, NOT pixel-identical.
- Pixel-perfect anchor-position parity. Excel's anchors carry EMU
  sub-cell offsets that Notesheet rounds to the nearest cell. The
  chart appears within ±1 cell of where Excel places it.
- Live update through the data bus (covered by feature-4).
- Round-trip preservation through M10 export (covered by feature-5).

**Suggested fixture(s)**

- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx` — the
  original "this crashes import" fixture, already shipped.

**Related risks**

- **`floatDomChart` regionKind is new harness work.** The Univer
  outer-container selector needs verification at generator time
  (DevTools session). Candidate selectors include
  `.univer-render-wrapper`, `.univer-app-layout`, or the Univer-mounted
  div inside `UserWebviewIndex.html`. Use the M16 `previewPane`
  precedent for graceful fallback (try multiple selectors in order).
- **`import-fixture.sh` fixture-path expansion.** The script currently
  hardcodes `tests/ExcelBaseTestData/formatting-testdata/`. Extend it
  to ALSO accept `tests/fixtures/charts/` so future cycles' chart
  fixtures import via the same path.
- **Window prep.** `prep-joplin-window.sh` must run before the
  screenshot — the existing fill-window + sidebar/note-list-hide
  routine. No additional state needed for editor-canvas screenshots
  (the editor pane is the default after the existing prep).
- **Active-cell selection bleed.** Univer paints a blue selection
  border on A1 by default. If A1 is inside the visible region, the
  selection border may show in the screenshot. `move-selection.js`
  is available to move selection off A1 if needed; the float-DOM
  chart is unlikely to overlap A1 (charts in
  `MultiSheet.xlsx` typically anchor to row > 5).
- **Untouchable test files**: same list as feature-1.

---

### feature-4-m17-live-update-data-bus

**Spec**

An imported chart drawing wires up to Notesheet's existing chart
data bus (`subscribeChartUpdate` from `src/charts/dataBus.ts`) on the
SAME code path as a Notesheet-authored chart. When the source range's
cells change in the editor, the chart re-renders. This feature pins
the import-to-bus contract via Jest, NOT via a runtime drag-test
(the runtime gate is implicit in feature-3's screenshot — the chart
must already be subscribed to render at all).

**Acceptance criteria**

1. **Jest test
   `tests/m17ChartImportLiveUpdate.test.ts`** asserts:
   - Import `01-bar-simple.xlsx` to a snapshot via
     `xlsxBufferToSnapshot()`.
   - The snapshot has exactly one chart drawing in the
     `SHEET_DRAWING_PLUGIN` resource.
   - The chart drawing's `chartId` is a non-empty string.
   - Calling **`extractDataFromSnapshot(snapshot, sourceRange)`**
     (a NEW production helper M17 adds to
     `src/charts/extractData.ts` — the existing exported function
     is `extractRangeAsChartData(workbook, range)` which takes a
     Univer FWorkbook, NOT a snapshot — verified at
     `src/charts/extractData.ts:60`; M17 needs the snapshot variant
     for the M16 content-script renderer (feature-7) to extract
     chart data without booting Univer) against the imported chart's
     `sourceRange` returns labels and dataset values matching
     what the source `xl/charts/chart{N}.xml` declared. Expected
     values come from independently parsing the source XML's
     `<c:cat><c:strRef><c:strCache><c:pt><c:v>` and
     `<c:val><c:numRef><c:numCache><c:pt><c:v>` elements — NOT from
     `xlsxBufferToSnapshot`'s output.
2. **Bus-id parity.** The test confirms that the imported chart's
   `chartId` is what a `subscribeChartUpdate(chartId, listener)` call
   would key off. (Concretely: the test reads the chart-drawing
   `chartId`, calls `subscribeChartUpdate(chartId, listener)`, then
   `pushChartUpdate(chartId, fakeChartData)`, and asserts the
   listener fires with `fakeChartData`. The bus listener fires
   even with no other code path active — that's just the bus
   contract; the test below proves the editor actually wires the
   import to the bus.)
3. **`trackedCharts` Map population on snapshot load.** The most
   dangerous M13-class trap: today `editorView.tsx:177` declares
   `trackedCharts = new Map()`, and it's populated ONLY by
   `insertChart()` at `editorView.tsx:290`. Imported snapshots have
   NO code path to populate it. So a chart imported from .xlsx
   subscribes to nothing, and editing a source-range cell does
   not trigger `pushChartUpdate` on the imported chart's `chartId`
   — the live editor renders the chart with stale data forever.
   M17 MUST add a snapshot-load → `trackedCharts.set(chartId, ...)`
   code path. The Jest test asserts: import a snapshot, simulate
   the editorView's snapshot-consumer code path (whatever name
   the generator picks — likely a new `populateTrackedChartsFromSnapshot`
   helper exported from `editorView.tsx` for testability), then
   assert `trackedCharts.has(importedChartId) === true` AND
   `trackedCharts.get(importedChartId).sourceRange` equals the
   imported chart drawing's `sourceRange`. **WITHOUT this, feature-3's
   PGE screenshot is the only signal — and a stale-data chart looks
   identical to a fresh one in a single screenshot. This is exactly
   the M13 failure mode (snapshot data correct, runtime broken)
   the harness was built to prevent.**
4. **End-to-end live update through trackedCharts.** Builds on
   criterion 3. The Jest test:
   - Imports a snapshot with one chart.
   - Calls the snapshot-load population helper.
   - Subscribes a listener to the imported chart's `chartId`.
   - Simulates an edit to a cell inside the chart's source range
     (modify `cell.v`, then trigger whatever function `editorView.tsx`
     uses to fan out source-range edits to chart subscribers — at
     `editorView.tsx:200` today it iterates `trackedCharts.values()`
     and calls `pushChartUpdate(chart.id, fresh)` per chart whose
     range contains the edited cell).
   - Asserts the listener fires with new data reflecting the edit.
5. **No second renderer code path.** Static-analysis sentinel: a
   `grep` test (or simple file-content check) asserts that
   `src/editorView.tsx`'s chart-mount logic doesn't grow a second
   chart-render branch keyed off "imported vs authored" — both flow
   through the same `NotesheetChart` component +
   `subscribeChartUpdate` subscription path.

**Out of scope**

- A runtime-driven drag test of the source range. The bus contract is
  the testable surface; runtime drag-and-drop is brittle in CI and
  not the gate.
- Validating that `extractData` works correctly for cells the snapshot
  hasn't computed yet — `extractData` reads `cell.v` (cached value)
  and tolerates missing cells today. M17 relies on that existing
  behaviour.

**Suggested fixture(s)**

- `tests/fixtures/charts/01-bar-simple.xlsx` — simplest single-series
  case. The test exercises the bus contract on this fixture only;
  feature-2 covers per-type fidelity for the others.

**Related risks**

- The chart's source range may reference cells that don't exist in
  the snapshot. `extractData` already tolerates this; the test exercises
  a fixture where every source-range cell DOES exist
  (`01-bar-simple.xlsx`'s data is `Sheet1!$A$1:$B$5`).
- Whole-column references (`Sheet1!$A:$A`) — out of scope, see
  feature-1's risks.
- **Untouchable test files**: same list as feature-1.

---

### feature-5-m17-bidirectional-roundtrip-excel-fixtures

**Spec**

A chart imported from one of the hand-crafted fixtures and then
exported back through Notesheet's existing M10 pipeline produces an
xlsx that re-imports with the same chart drawing. This pins import
+ export are inverse operations on the chart subset, parallel to
how M9 (tables) and M15 (CF) already pin inverse round-trip on their
respective subsets.

**Acceptance criteria**

1. **Jest test `tests/m17ChartBidirectionalRoundTrip.test.ts`**
   asserts for each of the four type-anchor fixtures
   (`01-bar-simple.xlsx`, `02-line-multi-series.xlsx`,
   `03-pie-single.xlsx`, `04-doughnut.xlsx`):
   - Import the fixture to a snapshot via `xlsxBufferToSnapshot()`.
   - Export the snapshot via `snapshotToXlsxBuffer()` (M10 path).
   - Re-import the resulting buffer via `xlsxBufferToSnapshot()`.
   - The second snapshot's chart drawing matches the first snapshot's
     chart drawing on:
     - `type` (string equality).
     - `sourceRange` (deep equality on `{startRow, endRow, startColumn,
       endColumn}`).
     - `labels` (array equality).
     - `datasets[*].label` and `datasets[*].data` (deep equality).
   The expected values come from the FIRST snapshot, NOT from a
   hardcoded literal — this is a snapshot-vs-snapshot inverseness
   test (fundamentally the M9 pattern, not the M13/E pattern). It
   does NOT pin to source XML because M10's emit may legitimately
   produce slightly different XML (palette colours, spPr details)
   that nonetheless re-imports to the same logical chart.
2. **Cross-sheet survives.** A fifth round-trip case using
   `07-chart-cross-sheet.xlsx` (chart anchored on Sheet2,
   referencing Sheet1 data). After round-trip, the chart drawing
   is still on the same sheet-id (Sheet2) AND its `sourceRange`
   still resolves to Sheet1's cells (NOT to Sheet2). **This
   requires extending `ChartDrawing.sourceRange` (currently
   `{ startRow, endRow, startColumn, endColumn }` at
   `src/charts/xlsxChart.ts:34`) with a new optional
   `sourceSheetId?: string` field.** Today's M10 export at
   `xlsxChart.ts:551` rebuilds cell-ref formulas as
   `<chartContainingSheetName>!$A$2:$A$5` because there's no
   sheet field; on cross-sheet imports that's wrong. M17 plumbs
   `sourceSheetId` through both M10 export and M17 import, then
   feature-5's cross-sheet test:
   - Imports `07-chart-cross-sheet.xlsx` and asserts the
     chart's `sourceSheetId` resolves to Sheet1 (data sheet),
     not Sheet2 (chart sheet). The expected sheet name comes
     from the source XML (`<c:f>Sheet1!$A$2:$A$5</c:f>`),
     parsed independently.
   - After M10 export + re-import, asserts the round-tripped
     chart's `sourceSheetId` STILL resolves to Sheet1. Catches
     "M10 silently rebuilds the formula prefix as the chart's
     containing sheet" — which would let the cached labels/values
     pass equality but break Excel's re-evaluation from cells.
   - Asserts the EXPORTED xlsx's `xl/charts/chart1.xml`'s
     `<c:f>` formula prefix is still `Sheet1!`, not `Sheet2!`
     (verify via JSZip + regex on the exported buffer).

**Out of scope**

- Pixel-byte-equal XML output. M10's emit may differ from Excel's
  in `spPr`, `extLst`, and other "Microsoft extension" blocks —
  those are dropped on re-import without affecting the chart's
  data round-trip.
- Anchor-position byte parity. The EMU sub-cell offsets are rounded
  to nearest cell on import; the inverse rounds back, but precision
  isn't guaranteed.
- Programmatic edge-case round-trips (covered by feature-6).
- **Bar orientation (`<c:barDir val="bar"/>` horizontal vs `val="col"`
  vertical) round-trip.** 5 of 10 fixtures are horizontal bars
  (`01-bar-simple`, `05-bar-special-chars`, `06-two-charts-one-sheet`
  chart 1, `09-bar-percent-axis`, `10-bar-with-trendline`); the
  remaining 2 'bar' fixtures (`07`, `08`) are vertical. M10's
  exporter at `src/charts/xlsxChart.ts:108` hardcodes
  `<c:barDir val="col"/>` and Notesheet's `ChartType` union has
  only `'bar'` with no orientation flag. M17 IMPORTS preserves
  the source bar direction in the chart drawing (so feature-1's
  XML-anchored test passes), but M17 EXPORT collapses everything
  to vertical. Document the gap in M17's README "Known
  shortcomings" entry. Round-trip 01-bar-simple and observe that
  the re-imported chart drawing's orientation field reverts to
  vertical — pin this as a documented behaviour, NOT a bug. A
  future cycle can extend `ChartType` to `'bar' | 'barH' | ...`
  or add `meta.barDir` to ChartDrawing.

**Suggested fixture(s)**

- `tests/fixtures/charts/01-bar-simple.xlsx`
- `tests/fixtures/charts/02-line-multi-series.xlsx`
- `tests/fixtures/charts/03-pie-single.xlsx`
- `tests/fixtures/charts/04-doughnut.xlsx`
- `tests/fixtures/charts/07-chart-cross-sheet.xlsx`

**Related risks**

- M10's chart export reads `SHEET_DRAWING_PLUGIN` via
  `readChartsFromSnapshot`. M17's import writes to the same resource
  shape. If the shapes diverge (e.g. import uses
  `transform.from.col` but export reads `axisAlignSheetTransform.from.column`),
  the round-trip silently drops the chart. The test is the contract;
  the shape is documented at the top of this BUILD_PLAN.
- M10's chart export expects a minimum `sourceRange` (header row +
  ≥1 data row + label col + ≥1 data col). If an Excel-authored chart
  has fewer rows/cols (e.g. a chart over `A1:B1` only — header but
  no data), M10's export drops it silently. M17's import should
  honour the same minimum or document the gap.
- **`07-chart-cross-sheet.xlsx` is misnamed.** The fixture has the
  chart anchored on Sheet2 referencing Sheet1 data (verified via
  `unzip -p tests/fixtures/charts/07-chart-cross-sheet.xlsx
  xl/charts/chart1.xml | grep '<c:f>'` → all three formula refs
  resolve to `Sheet1!`, `Sheet1!$A$2:$A$5`, `Sheet1!$B$2:$B$5`).
  But BOTH the categories ref AND the values ref are on Sheet1 —
  they're not split across DIFFERENT sheets. So this fixture
  exercises "chart-on-different-sheet-than-data," NOT
  "categories-and-values-on-different-sheets" (which Excel
  permits). The latter case is OUT OF SCOPE for M17 — when M17
  resolves both refs, M10's export rebuilds both refs as the
  chart's containing sheet. Document at `xlsxChart.ts`.
- **Untouchable test files**: same list as feature-1.

---

### feature-6-m17-programmatic-roundtrip-pack

**Spec**

A pack of 5–7 in-memory snapshots authored DIRECTLY in test code, each
carrying one chart in `SHEET_DRAWING_PLUGIN`, exercises Notesheet's
own emit-and-import round-trip on edge cases that the Excel-authored
fixture set doesn't cover or covers thinly. Each case is built by
constructing a `UniverSnapshot` programmatically (NOT via
`new ExcelJS.Workbook()` — exceljs's chart write API is thin / uneven;
M10 explicitly bypasses it), calling `snapshotToXlsxBuffer()`, then
`xlsxBufferToSnapshot()`, and asserting equality on the chart fields
the operator cares about (`type`, `sourceRange`, `anchor`, `labels`,
`datasets`).

**Acceptance criteria**

1. **Jest test `tests/m17ChartProgrammaticRoundTrip.test.ts` covers
   the following cases at minimum:**
   - **Case A — bar chart with negative values.** Datasets contain
     mixed positive and negative numbers (e.g. `[3, -2, 5, -1]`).
     Round-trip preserves all values including negatives.
   - **Case B — line chart with single-data-point series.** Labels
     length 1, datasets[0].data length 1. Round-trip preserves the
     1-element shape.
   - **Case C — pie chart with very long category labels.** At
     least one label > 30 chars. Round-trip preserves the label
     string verbatim.
   - **Case D — doughnut chart with empty series (zero rows of
     data).** Labels length 0, datasets[0].data length 0. Round-trip
     does NOT crash; the resulting drawing is either preserved with
     empty arrays, or silently dropped — the test accepts either
     outcome (DROP results in zero chart drawings; PRESERVE results
     in one chart drawing with empty arrays). Document which path
     the implementation takes in `## Notes`.
   - **Case E — bar chart with special chars in title and category
     names.** Title contains `&`, `<`, `>`, `"`, `'`. Category names
     contain `&`, `<`, `>`. Round-trip preserves the strings
     verbatim. (This pins the M10 `escapeXml` + the M17 reverse-
     unescape are inverse on the supported character set.)
   - **Case F — cross-sheet chart.** A snapshot with two sheets;
     chart on Sheet2 references data on Sheet1. Round-trip preserves
     the cross-sheet reference. (Parallels `07-chart-cross-sheet.xlsx`
     but generated, so the test is self-contained.)
   - **Case G — two charts on one sheet.** Two chart drawings in
     `SHEET_DRAWING_PLUGIN` for the same `subUnitId`. Round-trip
     preserves both with distinct `chartId`s. (Parallels
     `06-two-charts-one-sheet.xlsx` but generated.)
2. **Each case anchors to its ORIGINAL snapshot, NOT to the
   re-imported snapshot or to a hardcoded literal.** The test reads
   the original snapshot's chart-drawing fields, runs the round-trip,
   reads the re-imported snapshot's chart-drawing fields, and asserts
   `originalChart.X === reimportedChart.X` for each pinned field.
3. **Bundle assertion.** The full case set produces ≥ 5 distinct
   round-trip cases (counted as Jest `test()` calls inside the
   `describe('feature-6: programmatic round-trip pack', ...)` block).
4. **No `new ExcelJS.Workbook()` is invoked anywhere in the test.**
   (Lint sentinel: a top-of-file comment forbids it; the test exercises
   ONLY the M10 export pipeline + the M17 import pipeline.)

**Out of scope**

- Exhaustive chart-type coverage (the pack stays at 5–7 cases per
  operator brief; the four-type fidelity is feature-2's job).
- Asserting the round-tripped buffer opens cleanly in Excel itself
  (the M10 export pipeline already covers that via
  `tests/xlsxChart.test.ts`).
- Programmatic round-trips through unsupported types (radar, scatter)
  — feature-2 covers fallback.

**Suggested fixture(s)**

- None on disk — every case is built in-memory in the test.

**Related risks**

- **`SHEET_DRAWING_PLUGIN` resource shape mismatch.** Building a
  snapshot programmatically requires authoring the resource JSON in
  the SAME shape `readChartsFromSnapshot` expects (componentKey ===
  'NotesheetChart', `axisAlignSheetTransform` populated, etc.). The
  shape is documented at the top of this BUILD_PLAN. The test should
  use a shared `buildChartSnapshot(opts)` helper inside the test
  file to reduce per-case repetition.
- **Empty-data degenerate case (D).** M10's export currently emits
  `<c:numCache><c:ptCount val="0"/></c:numCache>` for empty datasets;
  the import path needs to tolerate this. The test pins whatever
  behaviour the implementation lands on (preserve OR drop), and the
  generator documents the choice in `## Notes`.
- **Untouchable test files**: same list as feature-1.

---

### feature-7-m17-chart-svg-html-export-jest

**Spec**

The M16 content script (`src/contentScripts/notesheetRenderer.ts`)
is extended to walk the snapshot's `SHEET_DRAWING_PLUGIN` resource
and emit static inline SVG for each chart at its anchor position,
inside or alongside the rendered HTML table. The static SVG uses the
SAME `CHART_PALETTE` colours that the live Chart.js renderer uses
(duplicated in the content script with a comment pointing at
`src/charts/extractData.ts`). Static SVG works in Joplin's PDF export
(no JS runtime at view time); the editor preview pane shows the SAME
SVG.

The renderer falls through cleanly when the `SHEET_DRAWING_PLUGIN`
resource is absent — the M16 base table-render path is unaffected,
and non-notesheet fenced blocks still fall through to markdown-it's
default renderer.

**Acceptance criteria**

1. **Jest test `tests/m17ChartInHtmlExport.test.ts` covers FIVE
   distinct render cases:**
   - **Bar chart fixture.** Builds a programmatic bar-chart snapshot
     in-memory (using the same `buildChartSnapshot` helper feature-6
     introduces, or an inline equivalent). Runs the M16 renderer
     on the snapshot. Asserts the resulting HTML contains exactly
     one `<svg>` element AND the count of `<rect>` elements inside
     that SVG equals the dataset's `data.length`. **Structural
     floors (prevent the "0 rects, valid input, still passes"
     degenerate-emit hole):** every `<rect>` MUST have
     `width > 0` AND `height > 0`. The dataset's data length MUST
     be ≥ 1 (the test's input has 4 points; assert `rectCount >= 1`
     in addition to `rectCount === data.length`).
   - **Line chart fixture.** Programmatic line-chart snapshot.
     Renders an `<svg>` containing at least one `<polyline>` OR one
     `<path>` element per dataset. The count of polylines+paths
     dedicated to data lines (excluding axis lines / titles) equals
     `datasets.length`. **Structural floor:** each line element's
     `points` attr (or `d` attr for path) parses to ≥ 2
     comma-separated coordinate pairs (a "line" with 1 point is
     degenerate).
   - **Pie chart fixture.** Programmatic pie-chart snapshot. Renders
     an `<svg>` containing one `<path>` per slice (one per data
     point in `datasets[0].data`). The pie sweep angles approximate
     the data proportions within ±3° (tolerance accounts for
     `pathArcCommand` rounding). The test independently computes the
     expected sweep angles from the input data
     (`angle_i = data[i] / sum(data) * 360`) and parses each
     `<path>`'s `d` attribute to extract the sweep — same M13/E
     "expected from input, NOT from emit" discipline. **Structural
     floors:** sum of all sweep angles is within ±1° of 360, AND
     no individual sweep angle is exactly 0.
   - **No-charts fall-through.** A snapshot with NO
     `SHEET_DRAWING_PLUGIN` resource (M16's existing test-case
     shape) renders as a table-only HTML with NO `<svg>` elements.
     The M16 base test cases still pass — feature-7 implementation
     does NOT regress them.
   - **CF + chart on the same sheet.** A snapshot with both a
     `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource (one cellIs
     rule painting B2:B5 pink) AND a `SHEET_DRAWING_PLUGIN`
     resource (one bar chart anchored at B2:F20). Renders to
     HTML containing BOTH the CF-coloured `<td>` (the cellIs
     paint flows through M16's existing CF evaluator) AND the
     chart `<svg>`, in document order: the `<table>` block
     first, then the `<svg>`. No `position:absolute` /
     `z-index` styling — the chart sits AFTER the table, never
     overlaid. (Hand-authored SVG inline in HTML can't easily
     layer over a colour-painted `<td>`; the explicit non-overlay
     behaviour avoids inventing a layout system.)
2. **`CHART_PALETTE` colour parity.** At least one case asserts the
   first dataset's primary colour in the SVG output matches
   `CHART_PALETTE[0]` — currently `#3b82f6` (the actual first entry
   of the 8-entry array at `src/charts/extractData.ts:11-20`).
   The test reads the palette from `src/charts/extractData.ts`
   directly (e.g. via a `tests/util/m17ChartPaletteSync.ts` helper
   that imports the symbol — NOT a fs+regex parse, which can fail
   silently if the source format shifts) so a future palette change
   updates both the runtime AND the test atomically.
3. **No new visualization-library dependency added.** `git diff
   package.json package-lock.json` after the implementation lands
   shows no new entry for `chart.js`, `d3`, `victory`, or any
   visualization lib in the content-script's transitive set.
   (Verification: `npm ls --json | jq` over the content-script
   bundle's deps — or simply read the bundle's text and assert no
   `chart.js` symbol leaked in.)
4. **Bundle size soft-gate.** After M17 lands, the M16 content
   script bundle stays under ~20 KB (was ~8.4 KB pre-M17; the
   chart-to-SVG renderer is the only major new logic). The generator
   measures the bundle size during the feature-7 build, surfaces
   the number to the operator if it exceeds 20 KB, and flags it
   here in `## Notes`. This is a soft gate — the feature does NOT
   fail review on bundle size alone, but a sudden 50 KB jump merits
   investigation before flipping the row.

**Out of scope**

- Pixel-perfect SVG-vs-Chart.js parity. The static SVG is recognisably
  the same chart with the same palette and the same data, NOT
  pixel-identical.
- ARIA / `<title>` / `<desc>` accessibility extras. M17 ships
  `<title>` with the chart name as the minimum (operator brief).
- DataBar / iconSet / colorScale CF rendering interactions with the
  chart layer — CF rules paint cells, charts render at anchor; the
  two layers don't interact in the static SVG output.
- iconSet glyph SVG inside CF columns (still M16-followup; M17 only
  adds chart-as-SVG, not CF-icon-as-SVG).

**Suggested fixture(s)**

- None on disk — every case is built in-memory in the test using a
  shared snapshot-builder helper (shared with feature-6 via
  `tests/util/chartSnapshotBuilder.ts` or an in-test factory).

**Related risks**

- **Bundle size.** Hand-author the SVG primitives (rect / polyline /
  path with arc commands). Don't pull D3 or Chart.js. Operator's
  related-risks block is explicit on this; the soft gate is ~20 KB.
- **Zero / negative-value edge cases.** Bar charts with negative
  values: bars below the zero line. Pie charts with all-zero data:
  degenerate "no data to plot" emit (NOT a divide-by-zero crash).
  feature-6's case A and case D exercise both; the renderer handles
  them without throwing.
- **CHART_PALETTE duplication.** The content-script's hardcoded
  palette MUST match `src/charts/extractData.ts:CHART_PALETTE`
  byte-for-byte. The test imports the source palette as a TS
  symbol via `tests/util/m17ChartPaletteSync.ts` (which re-exports
  `CHART_PALETTE` from `src/charts/extractData.ts`) and asserts
  array equality with the renderer's hardcoded copy. Do NOT use
  `fs.readFileSync` + regex — a malformed regex passes silently
  and the drift goes undetected. Document this contract in the
  renderer file's header comment.
- **CF + chart on the same sheet — z-order is "no overlay."**
  Snapshots may carry both `SHEET_CONDITIONAL_FORMATTING_PLUGIN`
  AND `SHEET_DRAWING_PLUGIN`. Feature-7 criterion 1's CF+chart
  case enforces document-order rendering (table first, then SVG)
  — never `position:absolute` / `z-index` overlay. Hand-authored
  SVG can't easily layer over a CF-painted `<td>`; the explicit
  non-overlay behaviour avoids inventing a layout system. The
  user gets both the CF-coloured cells AND the chart, just not
  visually overlapping.
- **Untouchable test files**: this feature MUST NOT edit
  `tests/m16NotesheetMarkdownRender.test.ts`, but it SHOULD verify
  (by running the suite) that those tests stay green after the
  content-script extension lands.

---

### feature-8-m17-chart-preview-pane-pge-smoke

**Spec**

When the user opens the imported `MultiSheet.xlsx` note in Joplin
and views Joplin's preview pane (markdown-rendered HTML), the chart
appears as inline SVG at its anchor position. The PGE harness
captures a screenshot of the preview iframe and the
`samplePreviewPaneInk()` sidecar reports `inlineSvgCount ≥ 1`.

This is the second PGE smoke that exercises a single fixture import
(parallels feature-3 against the editor canvas). One fixture import,
two screenshot-grades.

**Acceptance criteria**

1. **Preview-pane screenshot at
   `screenshots/feature-8-m17-chart-preview-pane-pge-smoke/eval-*.png`.**
   Captured via `bash scripts/pge/eval-screenshot.sh feature-8-m17-chart-preview-pane-pge-smoke`
   after `import-fixture.sh MultiSheet.xlsx` lands the note and
   `prep-joplin-window.sh` fills the window with preview-pane visible.
   Re-uses M16's `previewPane` regionKind from `eval-screenshot.js`.
   The screenshot shows:
   - The rendered HTML table with cells from each sheet (M16's
     existing render path — confirmed by `<table>` + cell-bg
     visible).
   - At least one inline `<svg>` chart element at the chart's anchor
     position.
   - The chart's title text (the SVG `<title>` content or a `<text>`
     element with the title string) visible.
   - No raw JSON `{"id":...}` leak text shown in the preview.
2. **Pixel sidecar `eval-*.pixels.json`** carries
   `inlineSvgCount ≥ 1` (new sampler signal added to
   `samplePreviewPaneInk()`). The existing M16 signals are still
   present:
   - `tableCount ≥ 1` (the M16 base render still works).
   - `rawJsonLeak === false` (the renderer still ran).
   - `sheetHeadings.length ≥ 1` (multi-sheet headings still
     emit).
3. **The `ensurePreviewPaneVisible()` harness routine still works.**
   M16's window-prep + layout-toggling routine drives Joplin from
   any starting state into "preview pane visible" via the AppleScript-
   based menu clicks. If it has rotted since M16 shipped, the
   feature-8 implementation fixes it as part of the cycle (per
   operator's related-risks note).
4. **Generator-evidence** at
   `screenshots/feature-8-m17-chart-preview-pane-pge-smoke/generator-evidence.png`
   plus its `.pixels.json` sidecar; both Read via the Read tool.

**Out of scope**

- Pixel-byte-equal SVG between editor canvas (Chart.js) and preview
  pane (static SVG). They use the same palette and data; they look
  recognisably the same. Pixel parity is not the gate.
- Joplin Export → PDF round-trip pixel test. The PDF export pipeline
  is downstream of HTML; if HTML preview-pane works, PDF works
  (gated by Joplin's Electron / Chromium PDF renderer, NOT M17's
  emit). The PDF gate is "no crash" — covered by feature-7's Jest
  output as a structural check, not a runtime PDF capture.
- Multi-pane layout state-machine work beyond what M16 already does.

**Suggested fixture(s)**

- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx` —
  same fixture as feature-3, exercised through the preview-pane
  region instead of the editor canvas.

**Related risks**

- **`samplePreviewPaneInk()` extension** is new harness work for
  this cycle — adds the `inlineSvgCount` signal. The function counts
  `<svg>` elements inside the preview iframe DOM via
  `document.querySelectorAll('svg').length`. Test the new signal
  works on a single-table no-chart preview (returns 0) and on a
  chart-bearing preview (returns ≥ 1).
- **Preview iframe DOM stability.** M16 verified
  `iframe.noteTextViewer` as the preview-pane element on the current
  Joplin build (verified 2026-06-06). M17 inherits this; if the
  selector changes between Joplin versions, the screenshot fails
  loudly with "no preview-pane element found" — same fallback
  pattern as M13/E's `tableHeaderRowRegion`.
- **Layout-state toggling.** `ensurePreviewPaneVisible()` cycles
  through 5 distinct Joplin layout states (Custom Editor active /
  TinyMCE active / markdown editor-only / split / preview-only).
  States 4 and 5 both work for evidence capture.
- **Untouchable test files**: same list as feature-1. (No Jest test
  file is added for this feature — the gate is the screenshot.)

---

### feature-9-m17-flip-pin-downs-and-test-count

**Spec**

The two pin-down lines in `tests/m12ImportRecovery.test.ts:59,84`
flip from "MultiSheet.xlsx → xlsx-charts-unsupported with a friendly
message" / "LargeWorkbook.xlsx → xlsx-charts-unsupported (same crash
class as MultiSheet)" to positive pin-downs that assert
"MultiSheet.xlsx → snapshot with N charts" / "LargeWorkbook.xlsx →
snapshot with M charts". The third pin-down on line 96
(`FormulasAndStructuredRefs.xlsx → xlsx-multi-table-unsupported`)
stays UNCHANGED — that's a different crash class, not addressed by
M17. The error class itself
(`NotesheetImportError(code='xlsx-charts-unsupported', ...)`) stays
DEFINED for any future drawing-related crash class we haven't
classified.

After the flips and all M17 features land, the full test suite count
moves from 267 to ≥ 290 (≥ 23 new Jest tests across features 1, 2,
4, 5, 6, 7 plus the 2 flipped pin-downs).

**Acceptance criteria**

1. **`tests/m12ImportRecovery.test.ts:59`** is the ONLY content edit
   to that file. The flipped test asserts:
   - `MultiSheet.xlsx` imports without throwing.
   - The returned snapshot has a non-empty `SHEET_DRAWING_PLUGIN`
     resource.
   - The number of chart drawings in that resource is ≥ 1 (the
     test reads the expected count from the source XML — counting
     `xl/charts/chart{N}.xml` files inside the zip — and asserts
     equality).
2. **`tests/m12ImportRecovery.test.ts:84`** is the ONLY content edit
   to the second flipped test. Same shape as criterion 1 but for
   `LargeWorkbook.xlsx`.
3. **No other content edit to `m12ImportRecovery.test.ts`.** The
   `FormulasAndStructuredRefs.xlsx` test stays untouched. The
   importable-fixtures `test.each` block stays untouched. (Note:
   if `MultiSheet.xlsx` and `LargeWorkbook.xlsx` are still listed
   in the importable-fixtures array, they CAN be added; this is a
   structural addition, not a content edit. But typical placement
   is to leave them in the crashing-fixtures `describe` and just
   flip the test bodies.)
4. **Test count gate (with cardinality floor — count alone is not
   enough).** Before the M17 cycle starts, the generator records
   the baseline `npm test` count (operator says 267 is the
   M16-shipping baseline; verify with a clean checkout). After all
   M17 features land, `npm test` reports ≥ 290 tests, all green.
   No skipped tests beyond M15's existing `excelCanvasFidelity`
   describe-skip-when-reference-PNG-absent block. **The count
   alone is reachable trivially via `test.each([fixture1...10])`
   expansion (10 fixtures × 2 assertions = 20 "tests" before any
   substantive feature work).** So the count floor is paired with
   a **cardinality floor on distinct surfaces**:
   - feature-1's test file has at least one `test()` call covering
     each of: bar / line / pie / doughnut chart types + the
     multi-anchor case (06) + the cross-sheet case (07) + the
     missing-chart-target synthetic + the oneCellAnchor synthetic
     + the no-`<c:cat>` synthetic = **at least 9 distinct cases**.
   - feature-2 has at least 1 positive case per chart type
     (4 cases) + at least 1 fallback case = **5 distinct cases**.
   - feature-4 has the 5 distinct asserts in criterion 1, 2, 3, 4,
     5 = **5 distinct test cases minimum**.
   - feature-5 has at least 1 round-trip case per chart type +
     the cross-sheet-survives case = **5 distinct cases minimum**.
   - feature-6's programmatic round-trip pack = **at least 5
     distinct cases** (operator brief says 5–7).
   - feature-7's SVG export = **5 distinct cases** (bar / line /
     pie / no-charts-fall-through / CF-plus-chart).
   The evaluator counts distinct `test()` calls per file and
   confirms each minimum.
5. **`git diff tests/` shows ONLY a strict subset of the
   following.** No other addition or deletion is permitted in
   `tests/`:
   - **NEW files matching `tests/m17*.test.ts`** (the seven new
     test files from features 1, 2, 4, 5, 6, 7).
   - **NEW files at `tests/util/m17*.ts`** (helpers shared
     between the m17 tests, e.g. `m17ChartPaletteSync.ts`,
     `m17BuildChartSnapshot.ts`).
   - **The two flipped lines in
     `tests/m12ImportRecovery.test.ts:59,84`** (specifically: the
     `expect(e.code).toBe('xlsx-charts-unsupported')` assertions
     replaced with the positive `MultiSheet.xlsx → snapshot with
     N charts` / `LargeWorkbook.xlsx → ...` assertions; line
     numbers may shift slightly as the test bodies grow).
   - **No content edit to any other existing test file in
     `tests/`** — including `tests/util/`, `tests/__mocks__/`,
     `tests/ExcelBaseTestData/`, and `tests/fixtures/`. The
     mock/util constraint is load-bearing: silently changing what
     a mock returns or what a test util computes shifts existing
     tests' actual values to match new expected values, which is
     the M13/E silent-modify failure mode wearing a disguise.
   - **No new file at any path that LOOKS like an existing
     untouchable file's name** — specifically forbidden: any new
     file named `tests/m15*.test.ts`, `tests/borders.test.ts`,
     `tests/hyperlinks.test.ts`, `tests/numberFormats.test.ts`,
     `tests/mergedCells.test.ts`, `tests/themeFonts.test.ts`,
     `tests/m13*.test.ts` other than what already exists. (These
     names appear in the operator's untouchable list but several
     don't exist on disk yet — adding them now under the M17
     namespace would muddy the gate. Use `tests/m17*.test.ts`
     for everything M17.)
   The evaluator runs `git diff --name-status tests/` at gate
   time and confirms each line in the output matches one of the
   four allowed shapes above. Any other line fails the gate.

**Out of scope**

- Removing the `xlsx-charts-unsupported` error class itself. It
  stays defined for future drawing-related crash classes. (Operator
  criterion #9 is explicit: the error class itself stays defined;
  only the two pin-down ASSERTIONS flip.)
- README "Known shortcomings" entries' update — punted to a follow-
  up docs PR per the M13/C, M13/D, M13/E, M15, M16 precedent. The
  evaluator does NOT penalise if the docs edit is deferred.
- Touching the `m12FixtureRoundTrip.test.ts` file — it has its own
  separate set of pin-downs that are NOT chart-related; M17 does
  NOT modify them.

**Suggested fixture(s)**

- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx`
- `tests/ExcelBaseTestData/formatting-testdata/LargeWorkbook.xlsx`
  (currently throws same error class)

**Related risks**

- This feature is the LAST one to land in the cycle, after features
  1–8 are all green. It depends on every M17 import-side feature
  working correctly — a partial M17 ship would fail this feature's
  `MultiSheet.xlsx → snapshot with N charts` assertion.
- **Untouchable test files**: see operator criterion #10. The
  evaluator runs `git diff tests/` post-implementation. The ONLY
  changes allowed in `tests/` are: NEW files in the
  `tests/m17*.test.ts` family AND the two flipped lines in
  `tests/m12ImportRecovery.test.ts`. NO content edit to any other
  existing test file. This is load-bearing — it prevents the "262 →
  287 by deleting 5 existing tests and adding 30 new ones" failure
  mode.

---

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file (and the prior M13/C, M13/D, M13/E, M15, M16
BUILD_PLAN.md files preserved in git history) as templates. Each
feature gets:

- `### feature-N-<kebab-id>` heading (stable; never renamed once
  written).
- **Spec**: one paragraph naming the user-observable change.
- **Acceptance criteria**: numbered list of observable evidence
  (HTML content, pixel-sidecar, Jest, runtime). Tests anchored
  UPSTREAM of our code (source XML for snapshot fidelity; original
  in-memory snapshot for round-trip; structural SVG element counts
  for HTML export; CHART_PALETTE source-of-truth for colour parity).
  NO assertions pinned to our own emit; NO "code does X" — only
  outcomes the evaluator can inspect from a fresh context.
- **Out of scope**: explicit non-goals so the evaluator does not
  penalise for them.
- **Suggested fixture(s)**: which file(s) under
  `tests/fixtures/charts/` or
  `tests/ExcelBaseTestData/formatting-testdata/` exercise this.
- **Related risks**: regression hot-spots and prior-bug pointers,
  including pointers into prior PR / commit hashes when prior
  work is the starting point. ALWAYS include the **Untouchable
  test files** sentinel for cycles where structural-integrity is
  load-bearing (anything past M13).

The lowest-numbered feature with `passes: false` in
`test-results.json` is what the generator works on next.
