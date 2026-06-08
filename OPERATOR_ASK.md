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
2. Each imported chart renders live: drag a source-range cell, the
   chart re-renders. Same `subscribeChartUpdate` pubsub the
   in-app charts use — there is no second renderer.
3. Round-trip: import an Excel chart, export the same notebook
   back to `.xlsx`, the resulting file opens in Excel with the
   chart visible. M10's existing export pipeline picks the chart
   up from the snapshot and writes it back — M17 only adds the
   import side; export is already shipped.
4. Existing tests stay green. `m12ImportRecovery.test.ts` flips
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
   - When `extractData(snapshot, sourceRange)` is called against
     the imported source range, it returns the same labels and
     dataset values the source XML declared
   This proves the import-to-bus wiring is the SAME as the
   authoring-to-bus wiring; the chart isn't a separate code path.
5. **Bidirectional round-trip.** A Jest test:
   - Imports `01-bar-simple.xlsx` to snapshot
   - Calls `snapshotToXlsxBuffer()` (the existing M10 export)
   - Reloads the resulting buffer via `xlsxBufferToSnapshot()`
   - Asserts the second snapshot has the same chart drawing as
     the first (same type, source range, labels, dataset values)
   This pins that import + export are inverse operations on the
   chart subset, just like they are on cells + tables today.
6. **`xlsx-charts-unsupported` error class is no longer thrown
   for the existing fixtures.** `tests/m12ImportRecovery.test.ts:59`
   "MultiSheet.xlsx → xlsx-charts-unsupported with a friendly
   message" flips to "MultiSheet.xlsx → snapshot with N charts"
   (positive pin-down). Same for `tests/m12ImportRecovery.test.ts:84`
   ("LargeWorkbook.xlsx → xlsx-charts-unsupported"). The error
   class itself stays defined for any FUTURE crash class we
   haven't classified yet — don't remove it.
7. **`npm run dist` succeeds and `npm test` count moves from 267
   to ≥ 280** (10+ new Jest tests across criteria 1, 2, 4, 5,
   plus the two flipped pin-downs).
8. **README's "Known shortcomings" entry referencing
   `xlsx-charts-unsupported` is removed.** The README line at
   `README.md:61` ("Workbooks with chart drawings... M17 addresses
   this") is replaced with whatever residual gap survives M17 —
   if a chart type isn't supported on import (radar, scatter,
   stock, etc.), document THAT residual gap, not the existing
   "any chart" gap.

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
- **Charts in HTML export (M16).** M16's content script doesn't
  render charts; that gap stays open. M17's HTML export
  follow-up is a separate cycle.
- **Pivot charts, sparklines, or any non-`xl/charts/chart*.xml`
  chart variant.** Those use different OOXML parts. Document
  as M18 candidates.
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

Smoke fixture (criterion 3):
- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx` —
  the original "this crashes import" fixture, already shipped.
  The PGE harness exercises this end-to-end.

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
  Notesheet's `extractData(snapshot, range)` reads `cell.v`
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
  M17's tests must anchor to the SOURCE Excel XML (the actual
  chart definitions in `tests/fixtures/charts/*.xlsx`), NOT to
  what `xlsxBufferToSnapshot()` produces. The chart's labels,
  dataset values, source range, type, anchor — all of these
  exist in the source XML and can be parsed independently as
  the upstream truth. Asserting `snapshot.charts[0].labels ===
  ['Q1','Q2','Q3','Q4']` without first independently reading
  those labels from the XML is the M13/E mistake.
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
