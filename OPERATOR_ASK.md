# Operator ask — M17: Chart import from `.xlsx`

## Why this matters

`xlsxBufferToSnapshot()` (`src/xlsx.ts:1727`) currently throws
`NotesheetImportError('xlsx-charts-unsupported', ...)` whenever
exceljs's reconcile pipeline trips on chart drawings. README's
"Known shortcomings" line `M17 addresses this` is the standing
promise. The user-visible failure today: a workbook authored in
Excel that contains a bar chart (the most common case) cannot be
opened in Notesheet at all — the import is rejected with a friendly
message, but the workbook is still inaccessible.

Notesheet already has a fully-shipped chart pipeline on the EXPORT
side (M7/M8 anchored Chart.js + M10 native OOXML write-back via
zip surgery). M17 builds the IMPORT-side counterpart. The user's
mental model: "I author a chart in Excel, paste the workbook into
Notesheet, the chart shows up live and re-renders when I edit the
source range." Same shape as a Notesheet-authored chart.

## The mechanism

Excel charts live in three OOXML parts inside the workbook zip:
1. `xl/drawings/drawing*.xml` — the anchor (cell range the float
   sits over) plus a relationship to the chart XML
2. `xl/charts/chart*.xml` — the chart definition (type, title,
   series, source range references)
3. `xl/charts/style*.xml` + `xl/charts/colors*.xml` — chart style
   metadata (M10 already parses these; can be ignored on import)

The exceljs reconcile loop crashes because it looks up the chart's
drawing target and `anchors` is undefined — the crash class is
already classified at `src/xlsx.ts:1745`. M17 needs to:
1. Read the chart parts directly from the zip (parallel to the
   existing `readTablesFromXlsxZip` / `readThemeFont` /
   `readNamedHyperlinkCells` zip-direct readers in `src/xlsx.ts`)
   BEFORE exceljs's reconcile crashes — or after, recovering from
   the throw with a chart-aware fallback path.
2. Map each parsed chart definition into Notesheet's existing
   `ChartDrawing` shape (`src/charts/xlsxChart.ts:25`) and emit a
   `SHEET_DRAWING_PLUGIN` snapshot resource that mirrors what the
   in-app chart authoring command produces.
3. Resolve the source data ranges (chart XML uses `Sheet1!$A$1:$B$5`
   formula-style references) into the snapshot's existing cell
   coordinates so Univer's float-DOM rendering picks up the chart
   verbatim — same code path as a Notesheet-authored chart.

## The feature

When a user imports an `.xlsx` with chart drawings:

1. The import does NOT throw. A note is created with cells, tables,
   and formatting (the existing M5 + M9 paths) AND chart float-DOMs
   anchored at the original positions.
2. Each imported chart renders live in the Univer editor: drag a
   source-range cell, the chart re-renders. Same `subscribeChartUpdate`
   pubsub the in-app charts use — there is no second renderer in
   the editor.
3. Round-trip: import an Excel chart, export the same notebook
   back to `.xlsx`, the resulting file opens in Excel with the
   chart visible. M10's existing export pipeline picks the chart
   up from the snapshot and writes it back — M17 only adds the
   import side; xlsx export is already shipped.
4. **Charts also render in the M16 HTML / preview-pane / PDF
   export pipeline.** The M16 content script
   (`src/contentScripts/notesheetRenderer.ts`) is extended to walk
   the snapshot's `SHEET_DRAWING_PLUGIN` resource and emit static
   inline SVG for each chart at its anchor position. Same data
   source as the editor float-DOM (no second data extraction
   path). Static SVG works in Joplin's PDF export (no JS runtime
   at view time). M16's existing "charts don't render in HTML
   export" gap closes as part of M17.
5. Existing tests stay green. `m12ImportRecovery.test.ts` flips
   from "MultiSheet.xlsx → xlsx-charts-unsupported" to
   "MultiSheet.xlsx → snapshot with N charts". The error class
   stays defined for non-chart drawing crashes.

## Acceptance criteria

The evaluator must verify ALL of:

1. **Import doesn't crash on chart-bearing fixtures.** A new Jest
   test loads each of `tests/fixtures/charts/01-bar-simple.xlsx`
   through `10-bar-with-trendline.xlsx` (already shipped, used as
   M10 export ground truth) via `xlsxBufferToSnapshot()` and
   asserts:
   - No `NotesheetImportError` thrown.
   - The returned snapshot has a `SHEET_DRAWING_PLUGIN` resource
     containing at least one chart drawing per fixture.
   - The chart's `chartId`, `type`, `sourceRange` and `anchor`
     fields match the source XML (use the same XML reader pattern
     as `tests/util/pngSampler.ts` — pure-stdlib, no new deps).
2. **Chart type round-trip.** A second Jest test asserts:
   - `01-bar-simple.xlsx` imports as `type: 'bar'` (not 'line', etc.)
   - `02-line-multi-series.xlsx` imports as `type: 'line'`
   - `03-pie-single.xlsx` imports as `type: 'pie'`
   - `04-doughnut.xlsx` imports as `type: 'doughnut'`
   The four types match Notesheet's existing `ChartType` union
   (`src/charts/xlsxChart.ts:23`). Charts of types we don't support
   (e.g. radar, scatter) fall back to bar with a console.warn —
   document the fallback in the chart drawing's `meta.unsupportedSourceType`
   field for evaluator visibility.
3. **MultiSheet.xlsx import works end-to-end.** The PGE harness
   imports `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx`
   (currently throws `xlsx-charts-unsupported`) via
   `scripts/pge/import-fixture.sh`, opens the resulting Notesheet
   note, and screenshots the Univer canvas. The screenshot shows:
   - The cells from each sheet rendered (this part already works
     once import doesn't throw)
   - The chart float-DOM visible at its anchor position
   - The chart's title text visible
   - The chart bars/lines/pie slices visible (live Chart.js render)
4. **Live update through the data bus.** A Jest test asserts:
   - Import `01-bar-simple.xlsx` to a snapshot
   - The snapshot has a `subscribeChartUpdate` ID matching
     `chartId` from the snapshot's chart drawing
   - When a NEW `extractDataFromSnapshot(snapshot, sourceRange)`
     production helper (M17 adds it; the existing
     `extractRangeAsChartData(workbook, range)` in
     `src/charts/extractData.ts:60` takes a Univer FWorkbook
     not a snapshot) is called against the imported source range,
     it returns the same labels and dataset values the source XML
     declared.
   - **Snapshot-load populates `trackedCharts`.** The test asserts
     that a snapshot-load → `trackedCharts.set(chartId, ...)`
     code path runs at editor boot time. Today
     (`src/editorView.tsx:177`) `trackedCharts` is populated ONLY
     by `insertChart()` at line 290; imported snapshots have no
     code path to populate it, so a chart imported from .xlsx
     subscribes to nothing and edits to source-range cells do
     not re-render the chart in the live editor. M17 must add
     this code path or the chart silently shows stale data.
     This is exactly the M13 failure mode (snapshot data
     correct, runtime broken) — see BUILD_PLAN feature-4
     criteria 3-4.
   This proves the import-to-bus wiring is the SAME as the
   authoring-to-bus wiring; the chart isn't a separate code path.
5. **Bidirectional round-trip — Excel-authored fixtures.** A Jest
   test:
   - Imports `01-bar-simple.xlsx` to snapshot
   - Calls `snapshotToXlsxBuffer()` (the existing M10 export)
   - Reloads the resulting buffer via `xlsxBufferToSnapshot()`
   - Asserts the second snapshot has the same chart drawing as
     the first (same type, source range, labels, dataset values)
   This pins that import + export are inverse operations on the
   chart subset, just like they are on cells + tables today.
6. **Programmatic round-trip pack — generated edge cases.** A Jest
   test authors 5–7 chart-bearing buffers in-memory via Notesheet's
   own M10 export path (NOT via `new ExcelJS.Workbook()` — exceljs's
   chart write API is thin and uneven; the M10 pipeline is the
   real chart writer), then re-imports each via M17 and asserts
   the round-tripped chart matches the source snapshot. Cover at
   minimum:
   - Bar chart with negative values
   - Line chart with single-data-point series
   - Pie chart with very long category labels (>30 chars)
   - Doughnut chart with empty series (zero rows of data)
   - Bar chart with special chars (`&`, `<`, `>`) in title and
     category names — verifies the M10 escapeXml + the M17 reverse
   - Cross-sheet chart (chart on Sheet2 referencing Sheet1 data) —
     parallels `07-chart-cross-sheet.xlsx` but generated, so the
     test is self-contained
   - Two charts on one sheet — parallels
     `06-two-charts-one-sheet.xlsx`
   The generated cases stay tight to the in-tree-fixture set (don't
   sprawl); the point is round-trip robustness, not exhaustive
   chart-type coverage. Each generated case is built by directly
   constructing a `UniverSnapshot` with `SHEET_DRAWING_PLUGIN`
   resource, calling `snapshotToXlsxBuffer()`, then
   `xlsxBufferToSnapshot()` and asserting equality on the chart
   fields the operator cares about (type, sourceRange, anchor,
   labels, datasets).
7. **Chart in HTML / preview-pane export — Jest tests.** Four
   new Jest tests in `tests/m17ChartInHtmlExport.test.ts` exercise
   the M16 content-script extension:
   - Bar chart fixture renders an `<svg>` with one `<rect>` per
     data point. The total count of `<rect>` elements equals the
     dataset's data length. Test fixture: a generated bar snapshot
     (in-memory, not a hand-crafted file).
   - Line chart fixture renders an `<svg>` with at least one
     `<polyline>` or `<path>` element per dataset.
   - Pie chart fixture renders an `<svg>` with one `<path>` per
     slice (data point). Pie sweep angles approximate the data
     proportions within ±3°.
   - The renderer falls through to markdown-it default when the
     `SHEET_DRAWING_PLUGIN` resource is absent (no chart
     fixtures). M16's existing tests stay green — the new SVG
     emit must not break the table render path or the
     non-notesheet fence fall-through.
8. **PGE harness — chart in editor canvas (criterion 3 above)
   AND chart in preview pane.** A SECOND PGE smoke targets the
   M16 preview-pane region (re-using M16's `previewPane`
   regionKind from `eval-screenshot.js`). Imports
   `MultiSheet.xlsx`, switches Joplin to a layout that shows the
   preview pane (M16 `ensurePreviewPaneVisible()` precedent),
   captures a screenshot of the preview iframe. The screenshot
   shows the rendered HTML table + at least one inline `<svg>`
   chart at the chart's anchor position. The preview-pane sampler
   (`samplePreviewPaneInk()`) is extended with an
   `inlineSvgCount` signal; the gate is `inlineSvgCount ≥ 1` for
   chart-bearing fixtures.
9. **`xlsx-charts-unsupported` error class is no longer thrown
   for the existing fixtures.** `tests/m12ImportRecovery.test.ts:59`
   "MultiSheet.xlsx → xlsx-charts-unsupported with a friendly
   message" flips to "MultiSheet.xlsx → snapshot with N charts"
   (positive pin-down). Same for `tests/m12ImportRecovery.test.ts:84`
   ("LargeWorkbook.xlsx → xlsx-charts-unsupported"). The error
   class itself stays defined for any FUTURE crash class we
   haven't classified yet — don't remove it.
10. **No regression in any existing test file.** `npm test` runs
    the full suite green. The evaluator runs `git diff tests/`
    after the implementation lands and confirms:
    - The ONLY tests/ changes are: NEW files (`tests/m17*.test.ts`
      family) AND the two flipped pin-down lines in
      `tests/m12ImportRecovery.test.ts:59,84`. No content edit to
      any other existing test file in `tests/`.
    - All of the following test files stay UNTOUCHED and pass:
      `m13RotatedText.test.ts`, `m13RichText.test.ts`,
      `m12FixturePinDowns.test.ts`, `m12FixtureRoundTrip.test.ts`,
      `excelReferenceFidelity.test.ts`, `excelCanvasFidelity.test.ts`,
      `m15ConditionalFormatting.test.ts`,
      `m16NotesheetMarkdownRender.test.ts`,
      `m16FormulaSourceOfTruth.test.ts`,
      `roundTripBidirectional.test.ts`, `xlsxChart.test.ts`,
      `formattingFidelity.test.ts`, `exportTableRoundTrip.test.ts`,
      `borders.test.ts`, `hyperlinks.test.ts`,
      `numberFormats.test.ts`, `mergedCells.test.ts`,
      `themeFonts.test.ts`. (Evaluator: `git diff tests/` lists
      no change to these files; `npm test` reports them green.)
    - The `m12ImportRecovery.test.ts` file's other tests (the
      non-flipped ones) stay green and untouched.
    This criterion is load-bearing: it prevents the "262 → 287
    by deleting 5 existing tests and adding 30 new ones" failure
    mode. The test count target below sets the lower bound; this
    criterion sets the structural-integrity bound.
11. **`npm run dist` succeeds and `npm test` count moves from 267
    to ≥ 290** (≥ 23 new Jest tests across criteria 1, 2, 4, 5,
    6, 7, plus the two flipped pin-downs).
12. **README's "Known shortcomings" entries are updated.** Two
    entries change:
    - `README.md:61` ("Workbooks with chart drawings... M17
      addresses this") is removed or replaced with whatever
      residual gap survives (radar, scatter, stock charts —
      document those if not supported).
    - The M16 "charts don't render in HTML export" entry (added
      during the M16 cycle) is removed; M17 closes that gap.
    The milestone table row for M17 flips to ✅ with PR link.

## Out of scope

- **Excel chart types beyond bar/line/pie/doughnut.** Notesheet's
  existing chart authoring pipeline only supports those four;
  importing a radar or scatter chart should fall back to bar
  with a warning (criterion 2). Adding new chart types is a
  follow-up M17.x.
- **Chart styling fidelity.** Notesheet's Chart.js renderer uses
  `CHART_PALETTE` (`src/charts/extractData.ts`), not Excel's per-
  chart style XML. The imported chart will look like a Notesheet
  chart, not pixel-identical to the source. Acceptable.
- **Chart titles with rich-text formatting.** Excel's chart titles
  can carry mixed font runs; M17 imports the plain text only.
  Punt M13/D-style rich-text title fidelity to a follow-up.
- **Multiple data label positions, trendlines, error bars, axis
  labels.** Trendlines exist in `10-bar-with-trendline.xlsx` —
  the import must not crash, but the trendline itself is dropped
  with a documented gap.
- **Chart drag/resize on import preservation beyond anchor cells.**
  Excel stores anchors as cell + EMU offset; Notesheet stores
  cell + col/row offset. We round to the nearest cell on import
  (drop the EMU sub-cell offset). Ditto on export — that's
  already M10 behaviour.
- **Pivot charts, sparklines, or any non-`xl/charts/chart*.xml`
  chart variant.** Those use different OOXML parts. Document
  as M18 candidates.
- **SVG chart styling parity with Chart.js.** The static SVG
  rendered for M16 export uses the same `CHART_PALETTE` colours
  as the live Chart.js renderer, but is NOT pixel-identical —
  no hover state, no animations, no gradient effects, simpler
  axis labels. Acceptable; the goal is a recognisable chart in
  PDF, not a perfect screenshot.
- **SVG accessibility extras** (ARIA roles, `<title>`/`<desc>`
  text descriptions of the data). Could be a follow-up; M17
  ships `<title>` with the chart name as the minimum.
- **README update beyond the M17 milestone row + the chart-
  unsupported known-shortcoming entry** — punted to a follow-up
  docs PR per precedent.

## Suggested fixtures

Primary anchor (each gets criterion-1 + criterion-2 coverage):
- `tests/fixtures/charts/01-bar-simple.xlsx` — single bar chart,
  one series, simplest case
- `tests/fixtures/charts/02-line-multi-series.xlsx` — line, three
  series (tests multi-dataset import)
- `tests/fixtures/charts/03-pie-single.xlsx` — pie
- `tests/fixtures/charts/04-doughnut.xlsx` — doughnut
- `tests/fixtures/charts/05-bar-special-chars.xlsx` — XML escape
  on category names with `&`, `<`, `>`
- `tests/fixtures/charts/06-two-charts-one-sheet.xlsx` —
  multi-chart-per-sheet (validates `chartId` uniqueness across
  charts in same sheet)
- `tests/fixtures/charts/07-chart-cross-sheet.xlsx` — chart on
  Sheet2 referencing data on Sheet1 (cross-sheet ref resolution)
- `tests/fixtures/charts/08-drag-resized.xlsx` — non-default
  anchor position (tests anchor parsing)
- `tests/fixtures/charts/09-bar-percent-axis.xlsx` — percent
  number format on axis (the format is dropped; bars still render)
- `tests/fixtures/charts/10-bar-with-trendline.xlsx` — trendline
  is dropped (out of scope) but bars render

Smoke fixture (criterion 3 + 8):
- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx` —
  the original "this crashes import" fixture, already shipped.
  The PGE harness exercises this end-to-end against BOTH the
  Univer canvas (criterion 3) AND the M16 preview pane
  (criterion 8). Two screenshot-grades from one fixture import.

Programmatic round-trip pack (criterion 6):
- 5–7 in-memory snapshots authored directly in test code, each
  carrying one chart in `SHEET_DRAWING_PLUGIN`, exercised through
  `snapshotToXlsxBuffer()` → `xlsxBufferToSnapshot()` to verify
  Notesheet's own emit-and-import round-trip is lossless. NOT
  authored via `new ExcelJS.Workbook()` (exceljs's chart write
  API is thin and uneven; M10 explicitly bypasses it). The point
  is "Notesheet emit ↔ Notesheet import," parallel to the
  "Excel emit ↔ Notesheet import" gate from criterion 1.

## Related risks

- **The exceljs reconcile crash happens BEFORE we get to read the
  zip directly.** Two viable architectures: (A) catch the throw,
  parse charts from the zip, retry exceljs's load with a
  drawing-stripped buffer; (B) extend the existing pre-load
  zip-direct readers (`readTablesFromXlsxZip`, `readThemeFont`,
  etc.) with `readChartsFromXlsxZip`, then strip the chart parts
  in-memory before passing to exceljs. Approach B is cleaner and
  matches existing conventions; A is easier but introduces a
  retry path that could mask other errors. **Generator picks one
  in the BUILD_PLAN.md and documents the choice.**
- **The chart's source range may reference cells that don't exist
  in the snapshot.** Excel allows charts to point at any range,
  including ranges with formulas that haven't been computed.
  Notesheet's new `extractDataFromSnapshot(snapshot, range)`
  helper (M17 adds it) reads `cell.v`
  (cached value) and tolerates missing cells — but if the chart
  references an entire column (`Sheet1!A:A`), that's a different
  shape we don't currently handle. M10's export rejects it; M17's
  import should reject it the same way (with a console.warn and
  fallback to dropping the chart).
- **`SHEET_DRAWING_PLUGIN` snapshot resource shape**. Notesheet's
  in-app chart command writes to this resource via Univer's
  `ICommandService.executeCommand('sheet.command.add-floating-dom',...)`.
  The import path needs to write the resource directly (we're
  outside Univer's command bus during snapshot construction).
  Document the exact JSON shape in BUILD_PLAN.md so the generator
  isn't reverse-engineering it during implementation. Reference:
  the in-app authoring command's output, captured by the M10
  export tests via `snapshot.resources[?].name === 'SHEET_DRAWING_PLUGIN'`.
- **Univer 0.23 chart float-DOM lifecycle.** Univer's drawing
  service binds the float-DOM to a sheet's drawingsManager;
  imported charts must register through the SAME service so
  drag/resize/persist behaves identically to authored charts.
  M10's export pipeline already documents the resource shape
  (`src/charts/xlsxChart.ts:25`); M17 should reuse that shape
  verbatim.
- **Test gap warning** (`feedback_pge_fidelity_test_gap.md`).
  M17 has two distinct authoring paths and they need different
  upstream anchors:
  - **Excel-authored fixture tests (criteria 1, 2, 5)**: anchor
    to the SOURCE Excel XML inside the `.xlsx`, NOT to what
    `xlsxBufferToSnapshot()` produces. Labels, dataset values,
    source range, type, anchor — all parseable independently
    from `xl/charts/chart*.xml` via stdlib XML reading. The
    test's expected values come from the source XML, never
    from "what our import emits" or "what we typed last time."
    Asserting `snapshot.charts[0].labels === ['Q1','Q2','Q3','Q4']`
    without first independently reading those labels from the
    XML is the M13/E mistake.
  - **Programmatic round-trip tests (criterion 6)**: anchor to
    the SOURCE snapshot we authored, NOT to whatever the
    re-import produces. Build the snapshot, export, re-import,
    compare RE-IMPORTED against ORIGINAL. The test's expected
    values come from the original snapshot. This is fundamentally
    different from criterion 1 — it tests that Notesheet's own
    emit + import are inverses, not that they match Excel.
  - **HTML / SVG export tests (criterion 7)**: anchor to the
    SVG element shape (count, fill colours, stroke widths)
    derived structurally from the input snapshot, NOT to the
    exact rendered byte-string. The renderer can change emit
    style without breaking the test as long as the count of
    `<rect>` / `<polyline>` / `<path>` matches what the data
    requires. Same M16-test discipline.
- **Chart import + table import interactions.** A workbook can
  have both a chart and a named table on the same sheet
  (`tests/fixtures/charts/06-two-charts-one-sheet.xlsx` does
  not, but `MultiSheet.xlsx` may). The existing
  `readTablesFromXlsxZip` and the new `readChartsFromXlsxZip`
  must be order-independent. Run the M9 + M13/E + M15 test
  suites after wiring; flag any fixture-level regression.
- **Don't symptom-patch test failures post-rebuild.** If a test
  regresses after wiring chart import, run
  `git diff package-lock.json` first. The new chart import path
  uses only stdlib + JSZip + the existing chart utilities — no
  new dependency should appear. Reference:
  `feedback_dependency_hygiene.md`. Do NOT downgrade exceljs or
  any other dep to make a symptom go away.
- **The PGE harness needs a chart region helper.** Until M17
  every screenshotted note had Univer rendering the entire
  visible content as a canvas. Charts are float-DOM (HTML
  overlays on top of the canvas). The harness's existing
  canvas-screenshot path will capture the canvas underneath
  the float-DOM but NOT the float-DOM itself unless we
  screenshot the editor frame as a whole. Add a `floatDomChart`
  regionKind to `eval-screenshot.js` that screenshots the
  outer Univer container instead of just the canvas. Reference:
  M16's `previewPane` regionKind (`scripts/pge/eval-screenshot.js`)
  for the precedent of a non-canvas region.
- **`scripts/pge/import-fixture.sh` may need updating** to
  accept fixtures from `tests/fixtures/charts/` (currently
  hardcodes `tests/ExcelBaseTestData/formatting-testdata/`).
  The harness change is part of the M17 cycle; expand the
  fixture path search rather than copying chart fixtures over.
- **M16 content-script bundle size discipline.** The M16 bundle
  shipped at ~8.4 KB (commonjs2, stdlib + the renderer's CF
  evaluator). Adding chart-to-SVG rendering will grow it. Don't
  pull Chart.js, D3, or any visualization library into the
  content script — Chart.js is already in the editor bundle but
  the renderer is a separate output target. Hand-author the SVG
  primitives (rect for bars, polyline/path for lines, path with
  arc commands for pie/doughnut slices). Target stays under
  ~20 KB after M17 lands. If it creeps higher, the planner
  should flag it.
- **SVG colour parity with Chart.js editor render.** The static
  SVG MUST use the same `CHART_PALETTE` array
  (`src/charts/extractData.ts`) that the live Chart.js renderer
  uses, so a single chart is the same colour in the editor and
  in the exported PDF. Don't define a second palette inside the
  content script. The palette either gets imported from
  `src/charts/extractData.ts` (webpack bundles it) or duplicated
  with a comment pointing at the source — the planner picks one
  approach in BUILD_PLAN.md.
- **SVG render of zero/negative-value edge cases.** Bar charts
  with negative values need axis-zero handling (bars below the
  axis line). Pie charts with all-zero data need a degenerate
  "no data to plot" rendering, not a divide-by-zero crash.
  The programmatic round-trip pack (criterion 6) exercises both;
  the renderer handles them without throwing.
- **PGE evaluator gate for criterion 8 (preview-pane chart
  screenshot)** depends on M16's `previewPane` regionKind
  working correctly. If the harness's `ensurePreviewPaneVisible()`
  has rotted since M16 shipped, fix it as part of the M17
  cycle. Same precedent as the M13/E `:variant` suffix work
  and M16's preview-pane introduction.
