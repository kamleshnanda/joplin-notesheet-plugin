# Operator ask — M16: Snapshot → HTML for Joplin's PDF/HTML export

## Why this matters

Joplin's right-click → **Export → PDF** and **Export → HTML** menus
currently dump the raw fenced JSON for Notesheet notes:

```
```notesheet v=1
{ "id": "workbook-...", "sheetOrder": [...], ... }
```
```

The user gets a 4-line JSON blob in their PDF instead of a rendered
table. README's "Joplin's Export → PDF / HTML menu" entry under
"Known shortcomings" pins this. The current workaround is the
in-editor **Export .xlsx** button.

M16 ships proper rendering. When a Notesheet note is exported via
Joplin's PDF or HTML export, the resulting document carries a
human-readable HTML table with cell values and formatting (fills,
font colours, borders, merged cells, basic alignment). The .xlsx
export button stays as the primary "send to Excel" path; M16 is the
"send to PDF / HTML" path.

## The mechanism

Joplin's PDF and HTML export both go through the markdown renderer.
The plugin API exposes `joplin.contentScripts.register(
ContentScriptType.MarkdownItPlugin, ...)` which lets us intercept
fenced code blocks. We register a content script that recognises the
`notesheet` fence tag, parses the snapshot via the existing
`extractSnapshot()`, and emits HTML.

Single content script, multiple consumers — Joplin's editor preview,
Joplin's PDF export, and Joplin's HTML export all pass note body
through the same renderer pipeline. M16 ships the same HTML for all
three.

## The feature

When a user right-clicks a Notesheet note and picks **Export →
PDF** or **Export → HTML**:

1. The exported file shows a rendered HTML table for each sheet in
   the snapshot, NOT the raw JSON fence.
2. Cell values are visible: numbers, strings, formula results
   (using the cached value `cell.v` exceljs evaluated at last
   save — Univer doesn't re-evaluate during render).
3. Cell formatting is preserved as inline-styled HTML:
   - Background colour (`bg.rgb` → `background-color`)
   - Foreground colour (`cl.rgb` → `color`)
   - Font weight (`bl: 1` → `font-weight: bold`)
   - Italic (`it: 1` → `font-style: italic`)
   - Underline (`un.s: 1` → `text-decoration: underline`)
   - Horizontal alignment (`ht` → `text-align`)
   - Vertical alignment (`vt` → `vertical-align`)
   - Per-side borders (`bd.t/r/b/l` → `border-top/right/bottom/left`)
4. Merged cells render as `<td colspan=N rowspan=M>` per the
   snapshot's `mergeData`.
5. Multi-sheet workbooks render each sheet under its name as a
   heading.
6. The fenced JSON does NOT appear in the output anywhere.

The renderer also activates in Joplin's editor preview pane (the
markdown-rendered preview), but that's a side-effect of the same
hook. The Custom Editor still owns the active editing experience;
the preview just stops showing JSON.

## Acceptance criteria

The evaluator must verify ALL of:

1. **HTML output contains rendered table.** A new Jest test
   `tests/m16NotesheetMarkdownRender.test.ts` (or similar) imports
   the M16 content script, calls its render function on a fenced
   notesheet body containing a known snapshot, and asserts the
   output contains:
   - One `<table>` per sheet
   - `<td>` elements with cell values
   - At least one cell with inline `style="background-color: #..."`
     matching the snapshot's `styles[].bg.rgb`
   - At least one cell with inline `style="font-weight: bold"`
     matching `bl: 1`
   - Merged cells rendered as `<td colspan>` / `<td rowspan>`
   - The raw JSON does NOT appear anywhere in the output
2. **Multi-sheet workbooks render each sheet.** A second Jest test
   asserts `MultiSheet.xlsx` (existing fixture) renders all sheets,
   each with its sheet name as a heading.
3. **Fixture round-trip — Aptos FormattingSmorgasboard.xlsx renders
   recognisably.** A third Jest test imports
   `FormattingSmorgasboard.xlsx`, runs it through the M16 renderer,
   and asserts the output contains the expected `ProjectTracker`
   column headers (`Project`, `Website`, `Budget`, etc.) and the
   table-style banding colours (`#34692E` header bg, `#CAEFCB`
   banded rows per M13/E).
4. **Conditional formatting renders into the export.** A fourth
   Jest test imports `ConditionalFormatting-Variants.xlsx`, runs it
   through the M16 renderer, and asserts the output contains:
   - At least one cell with the colorScale red/yellow/green hue
     (sample any column-A cell's bg-color attribute)
   - At least one cell with the cellIs > 50 pink fill (`#FFC7CE`)
   - At least one cell with the top-3 light-green fill (`#C6EFCE`)
   The dataBar and iconSet rules can be punted to M16-followup —
   they require dynamic ranking / glyph rendering that's harder in
   static HTML. Document explicitly.
5. **Visual — open the rendered HTML in a real browser.** PGE
   evaluator captures Joplin's editor preview pane (which runs the
   same content script as the export pipeline) for the
   FormattingSmorgasboard fixture and confirms:
   - The `ProjectTracker` table is visible with column headers
   - Cell values are present
   - Header row carries the green table-style fill
   - Banded data rows are visible
   The PGE harness needs a region helper for the preview pane (a
   Joplin DOM region, not the Univer canvas). Add as needed.
6. **Content script is registered in `src/index.ts`.** New imports
   + a `joplin.contentScripts.register(...)` call alongside the
   existing editor registration.
7. **Joplin's right-click → Export PDF works without crashing on a
   Notesheet note.** The PGE harness or a manual smoke confirms the
   PDF generation completes (don't deeply validate the PDF
   pixel-level — the content-script test bed covers HTML output).
8. **`npm run dist` succeeds and `npm test` count moves from 220
   to ≥ 224** (4+ new Jest tests).

## Out of scope

- **Charts in HTML export.** Notesheet's anchored Chart.js charts
  don't survive into the static HTML output. Document as an
  M16-followup or M17 dependency. Listed cell values still render.
- **Live formula evaluation in the renderer.** HTML export uses the
  cached `cell.v` value from the snapshot. If a formula's cached
  value is stale (e.g. user typed in Univer but didn't save before
  exporting), the stale value renders. Acceptable.
- **Univer's own `IDocumentBody.dataStream` rich-text per-run
  formatting** — i.e., M13/D's bold word + plain word in one cell.
  M16 renders the cell's plain text. Per-run formatting in HTML is
  M16-followup. Document.
- **Theme palette resolution beyond what's already baked into the
  snapshot.** Colours come from the snapshot's resolved `bg.rgb`
  and `cl.rgb` fields; we don't re-walk theme references in the
  renderer.
- **Print stylesheets / page breaks.** Joplin's PDF export decides
  page sizing; we don't override.
- **iconSet glyphs as inline SVG.** Static glyph rendering for CF
  icon sets is a non-trivial feature unto itself. M16-followup.
- **DataBar gradient rendering.** Same — needs CSS gradient + per-
  cell bar width math. M16-followup.
- **README "Known shortcomings — Joplin's Export → PDF / HTML"
  entry update.** Punted to a follow-up docs PR per precedent.

## Suggested fixtures

- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard.xlsx`
  — Aptos workbook with table styling. Tests the M13/E banding
  fidelity flowing through to HTML.
- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx`
  — multi-sheet test for the per-sheet rendering.
- `tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
  — exercises the M15 CF rules through the renderer.

The PGE harness invokes `import-fixture.sh FormattingSmorgasboard.xlsx`,
opens the resulting note, and screenshots the editor preview pane
(NOT the Univer canvas — we want the markdown-rendered preview, which
is where the new content script's HTML lands).

## Related risks

- **Joplin's content script API runs in a sandboxed worker.** The
  content script entry exports a `default function(context)` that
  returns `{ plugin: (markdownIt, pluginOptions) => {...},
  assets: {...} }`. The function runs in Joplin's renderer, not
  the plugin process — so it can't import from the plugin's main
  bundle directly. The renderer needs to be self-contained or
  load via webpack's content-script entry point.
- **`extractSnapshot()` lives in `src/snapshot.ts` and is shared.**
  The content script can either bundle a copy or duplicate the
  fence-parsing logic. Bundling is cleaner. webpack's existing
  contentScript build target pattern (see `webpack.config.js`)
  should handle this.
- **The renderer is called for every note**, not just Notesheet
  notes. Make the fence-tag check (`notesheet v=1`) the first
  thing — if the body doesn't match, return undefined and let
  markdown-it render the code block normally. Don't break
  non-Notesheet notes.
- **Cell value rendering precision.** `cell.v` may be a number
  with full float precision. If a cell carries `numFmt` (e.g.
  `"0%"` or `"$"#,##0`), we should format the value through that
  pattern before emitting. exceljs's `numFmt` syntax is the
  source. Out-of-scope to fully replicate Excel's number-format
  engine; ship something reasonable for the common cases (percent,
  currency, dates) and document gaps.
- **CF rules don't paint cells in the snapshot — Univer's CF
  engine paints them at render time.** For the M16 renderer to
  show CF colours, we have to evaluate CF rules ourselves: walk
  the snapshot's `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource,
  apply per-rule logic (cellIs > 50, top-N rank, color scale
  interpolation by percentile), and bake the result into per-cell
  inline styles. Color scale and cellIs and top10 are doable;
  iconSet and dataBar are harder (need glyph / bar element
  rendering). Document the iconSet and dataBar gaps.
- **Merged cells affect table layout significantly.** A merge
  spans rows/cols; the renderer must skip the cells inside the
  merge range when emitting subsequent `<td>` elements (else the
  HTML ends up with extra cells pushing the layout). The snapshot's
  `mergeData` array provides the ranges.
- **Don't symptom-patch test failures post-rebuild.** If a test
  regresses after registering the content script, run
  `git diff package-lock.json` first. The content-script entry
  may have pulled new transitives. Reference:
  `feedback_dependency_hygiene.md`.
- **Test gap warning.** Per `feedback_pge_fidelity_test_gap.md`,
  pin-downs anchored to our own emit are no test. The M16 tests
  should assert what's IN the rendered HTML against the source
  fixture's expected content (cell values, colour values from the
  snapshot's already-validated styles), not against whatever
  format we decide to emit. The HTML output format CAN change
  without test failure as long as the rendered values + colours
  match.
