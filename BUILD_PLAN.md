# Notesheet PGE — build plan (M16: Snapshot → HTML for Joplin's PDF/HTML export)

> **Cycle context.** Fifth real-feature cycle through the PGE harness
> and the first post-M15 milestone. M13/A–E shipped via this loop
> (PRs #19, #20, #21, #22, #23), M14 was an explicit NO-GO documented
> at `docs/m14-sheetjs-spike.md` (PR #25), and M15 conditional
> formatting shipped via PRs #26 + #27 with all five rule types
> round-tripping cleanly. The harness is mature: scripts under
> `scripts/pge/` cover launch, install, fixture import, window prep,
> selection moves, column-width resets, and per-feature region-aware
> eval screenshots with pixel sidecars.
>
> **The two-layer fidelity test pattern remains the project default
> for "match Excel" features.** M13/E established it; M15 extended
> both layers (5 tests per layer for the 5 CF columns). M16 is
> structurally different — it does NOT render onto the Univer canvas,
> so canvas-vs-Excel pixel parity is not the gate. Instead the gate
> is HTML-content fidelity: the rendered HTML carries the same cell
> values and the same colour/border/alignment cues the snapshot
> already validated for prior milestones. The acceptance criteria
> below assert against that HTML.
>
> **Single-feature cycle, big feature.** The whole renderer ships in
> one PR. It's a new content-script entry point — not previously
> exercised by the codebase — plus the per-cell HTML emit logic plus
> CF-evaluation logic plus a webpack build target plus a harness
> region helper for the editor preview pane. Decomposing into
> sub-features would force several round-trips through the harness
> for what is structurally one piece of work.

## Operator ask

(See `OPERATOR_ASK.md`.) Joplin's right-click → **Export → PDF** and
**Export → HTML** currently dump the raw fenced JSON for Notesheet
notes — a 4-line JSON blob in the PDF instead of a rendered table.
The README's "Joplin's Export → PDF / HTML menu" entry under "Known
shortcomings" pins this. M16 ships proper rendering. When a Notesheet
note is exported via Joplin's PDF or HTML export, the resulting
document carries a human-readable HTML table with cell values and
formatting (fills, font colours, borders, merged cells, basic
alignment, multi-sheet layout). The .xlsx export button stays as the
primary "send to Excel" path; M16 is the "send to PDF / HTML" path.

The same renderer activates in Joplin's editor preview pane (the
markdown-rendered preview), since Joplin runs every note body
through the same markdown-it pipeline before export. That gives the
PGE evaluator a renderable surface to capture without spawning a
PDF round-trip.

## Harness extension chosen for this cycle

The PGE harness's `eval-screenshot.js` until now targets the Univer
main canvas inside the editor's `UserWebviewIndex.html` frame. M16
needs to capture **the editor preview pane** — a Joplin-side DOM
region containing the markdown-rendered HTML, NOT a Univer canvas.

The generator owns this work item as part of the cycle:

- **Add a new region kind** `previewPane` to `REGION_BY_FEATURE` and
  the corresponding sampler path in `eval-screenshot.js`.
- **Capture target.** Instead of `webview.locator(canvasSel)
  .screenshot()`, the new path locates Joplin's preview iframe
  (`iframe.note-viewer-iframe`, or whichever stable selector exists
  in current Joplin builds — the generator verifies via DevTools
  before hardcoding) and screenshots its body. Generator may need
  to query the parent Joplin page (NOT the editor page that hosts
  Univer) — the preview pane is part of the main Joplin shell.
- **Pixel sidecar.** The dominant-colour histogram still applies, but
  the relevant ink aggregates are the table fill colours: `greenInk`
  (Aptos header bg `#34692E` family) plus `pinkInk` / `lightGreenInk`
  for CF visualisation. No new aggregate definitions needed — reuse
  the M13/E + M15 ones.
- **Title-prefix lookup.** `TITLE_PREFIX_BY_FEATURE
  ['feature-1-m16-snapshot-to-html'] = 'PGE M16 HTML eval '`. The
  fixture-import path stays unchanged (`import-fixture.sh` lands the
  fixture as a Notesheet note); the editor preview pane renders it
  via the new content script automatically.
- **Window prep.** The existing `prep-joplin-window.sh` already
  toggles sidebar / note list off and fills the window. M16
  ADDITIONALLY needs the editor pane and preview pane both visible
  — the toggle (Cmd+Shift+L on macOS / Joplin's "Toggle editor /
  preview" command) needs to land in the "show both" state. If
  `prep-joplin-window.sh` ends in editor-only mode, extend it to
  detect `.note-viewer-iframe` visibility and toggle until both are
  shown.

The single-feature key is `feature-1-m16-snapshot-to-html`. No
variant suffixes this cycle — the FormattingSmorgasboard fixture is
the central capture, with the Multi-Sheet and CF-Variants fixtures
exercised via Jest only (not via additional eval screenshots).

## Features

### feature-1-m16-snapshot-to-html

**Spec**

When a user opens a Notesheet note, Joplin's editor preview pane no
longer shows the raw fenced JSON — it shows a rendered HTML table
per sheet, with cell values, fills, borders, alignment, font weight
and colour, merged cells, and CF rule colouring (colorScale, cellIs,
top10) baked into per-cell inline styles. The same content script
runs when the user invokes right-click → Export → PDF or
Export → HTML, so the exported document carries the same rendered
table instead of the JSON blob.

The renderer is registered as a Joplin `MarkdownItPlugin` content
script in `src/index.ts`. It recognises the `notesheet` fenced-code
tag, parses the snapshot via `extractSnapshot()` (bundled into the
content-script entry — content scripts run in a sandboxed worker and
can't import from the plugin process), and emits HTML. Non-Notesheet
fenced blocks fall through to markdown-it's default rendering.

The Custom Editor still owns the active editing experience; the
preview pane just stops showing JSON, and the export pipeline stops
shipping JSON to PDF/HTML.

The KNOWN SHORTCOMING entry "Joplin's Export → PDF / HTML menu shows
the raw JSON fence" (README line in the Known shortcomings table)
flips from "yes" to "no" — but the README docs cleanup is explicitly
deferred to a follow-up docs PR per the M13/C, M13/D, M13/E, M15
precedent.

**Acceptance criteria**

The evaluator must verify ALL of the following from a fresh context:

1. **HTML output contains a rendered table — base shape.** A new
   Jest test `tests/m16NotesheetMarkdownRender.test.ts` (or
   equivalent) imports the M16 content-script's render function,
   feeds it a fenced notesheet body containing a known minimal
   snapshot (1 sheet, 3×3 cells, one bold cell, one bg-coloured
   cell, one merge range), and asserts the returned HTML string:
   - Contains exactly one `<table>` element.
   - Contains `<td>` elements for every non-merged-interior cell.
   - For the bold cell, the matching `<td>`'s `style` attribute
     contains `font-weight: bold` (case-insensitive substring).
   - For the bg-coloured cell, the matching `<td>`'s `style`
     attribute contains `background-color: #` followed by the
     RGB hex from the snapshot's `styles[<id>].bg.rgb`.
   - For the merged cell, the resulting `<td>` carries `colspan`
     and/or `rowspan` matching the merge range; cells INSIDE the
     merge range are skipped (not emitted as additional `<td>`s).
   - The string `"sheetOrder"` (a key in the JSON snapshot) does
     NOT appear anywhere in the HTML.
   - The string `"workbook-"` (the snapshot's id prefix) does NOT
     appear in the HTML.

   The test asserts what's IN the rendered HTML against the source
   snapshot's expected content, NOT against our own emit format. Per
   `feedback_pge_fidelity_test_gap.md`: tests that pin "we emit
   `<table class='notesheet'>`" are no test — change the class name
   and the test still passes. Tests that pin "the bold cell carries
   `font-weight: bold` in its style" hold the renderer accountable
   to the snapshot.

2. **Multi-sheet workbooks render each sheet under its name.** A
   second Jest test imports `MultiSheet.xlsx` (existing fixture),
   converts it via `xlsxBufferToSnapshot()`, wraps in a fenced
   notesheet body, and runs the content-script render. The HTML
   output:
   - Contains TWO or more `<table>` elements (one per sheet — the
     fixture has at least 2 sheets; the test reads the actual count
     from the snapshot's `sheetOrder` length and asserts equality).
   - For each sheet, the heading text immediately preceding its
     table contains the sheet name from the snapshot's
     `sheets[id].name`. The HTML element used for the heading is
     not pinned — `<h2>`, `<h3>`, or a `<caption>` inside the
     `<table>` all qualify; the test asserts the sheet-name string
     appears within 200 characters before the corresponding `<table>`
     opening tag.

3. **FormattingSmorgasboard.xlsx renders recognisably.** A third
   Jest test imports the Aptos `FormattingSmorgasboard.xlsx`, runs
   it through the renderer, and asserts the output contains:
   - The known column headers from the `ProjectTracker` table:
     `Project`, `Website`, `Budget`, `Spent`, `Discount` (or the
     literal subset that the fixture defines — the test reads from
     the snapshot to source the expected values, NOT a hardcoded
     list — but each one MUST appear once in the HTML output).
   - At least one `<td>` whose inline style contains
     `background-color: #34692E` (the Aptos M13/E table-style header
     fill — already validated by the existing M13/E pin-downs as
     the synthesizer's emit colour).
   - At least one `<td>` whose inline style contains
     `background-color: #CAEFCB` (Aptos banded-row fill — also
     validated by M13/E).

   The expected RGBs come from the snapshot the test creates by
   running `xlsxBufferToSnapshot()` — anchored to the snapshot, NOT
   to a hardcoded literal. The HTML rendering is the outer test;
   the snapshot's correctness is already pinned by M13/E.

4. **Conditional formatting renders into the export.** A fourth
   Jest test imports `ConditionalFormatting-Variants.xlsx`, runs it
   through the renderer, and asserts the output contains:
   - At least one `<td>` whose inline style contains a
     background-color from the colorScale red/yellow/green palette
     (`#F8696B` family OR `#FFEB84` family OR `#63BE7B` family —
     the test asserts at least one of the three appears in any
     column-A `<td>`'s style).
   - At least one `<td>` whose inline style contains
     `background-color: #FFC7CE` (cellIs > 50 pink fill — M15
     spec).
   - At least one `<td>` whose inline style contains
     `background-color: #C6EFCE` (top-3 light-green fill — M15
     spec).

   `dataBar` and `iconSet` rules are explicitly OUT OF SCOPE for
   the Jest assertion — they require dynamic bar-width rendering or
   glyph rendering that isn't well-served by static HTML. The test
   does NOT assert anything about column C or column I. Document
   the gap in OUT OF SCOPE below and in `## Notes` for the next
   cycle.

5. **Visual — Joplin editor preview pane shows the rendered table.**
   Single evaluator-captured screenshot under
   `screenshots/feature-1-m16-snapshot-to-html/`, taken via
   `bash scripts/pge/eval-screenshot.sh feature-1-m16-snapshot-to-html`
   after `import-fixture.sh FormattingSmorgasboard.xlsx` lands the
   note and `prep-joplin-window.sh` fills the window with both
   editor and preview panes visible. The screenshot targets Joplin's
   note-viewer iframe (NOT the Univer canvas). The evaluator's
   verdict text MUST explicitly call out:
   - The `ProjectTracker` table renders as an HTML table with
     visible column headers.
   - The header row carries the green table-style fill (no
     pixel-parity required against Excel — the gate is "the user
     would recognise this as the same table they had in Excel,
     coloured greenish").
   - At least one banded-row colour difference is visible in the
     data rows.
   - The raw JSON `{"id":...}` text does NOT appear anywhere in
     the screenshot. (If JSON appears, the content script didn't
     register, the fence-tag check failed, or the renderer
     errored out — the visible-JSON regression is the M13-style
     failure mode for M16.)

   The pixel sidecar
   `screenshots/feature-1-m16-snapshot-to-html/eval-*.pixels.json`
   carries `greenInk ≥ 30` (Aptos header fill is visible) and the
   `dominant` histogram top entry is NOT a near-white background
   colour like `rgb(255,255,255)` (which would mean no table
   rendered at all). The harness work item — adding a `previewPane`
   region kind to `eval-screenshot.js` — is part of THIS feature's
   acceptance gate, not a separate row.

6. **Content script is registered in `src/index.ts`.** The fixed
   plugin entry point now invokes
   `joplin.contentScripts.register(ContentScriptType.MarkdownItPlugin,
   '<id>', './notesheetRenderer.js')` (or equivalent) alongside the
   existing editor registration. The script ID is stable; if a
   future cycle changes it, that's a breaking change requiring
   plugin re-install. `ContentScriptType.MarkdownItPlugin` is
   imported from `api/types`.

   Verification: a Jest test (criterion 1) imports the renderer
   directly to run its render path; a runtime smoke confirms the
   editor preview pane shows the rendered HTML for the
   FormattingSmorgasboard fixture (criterion 5). Both gates fail
   loudly if the registration is absent.

7. **Joplin's right-click → Export → PDF works without crashing on a
   Notesheet note.** PGE harness or manual smoke confirms the PDF
   generation completes for the FormattingSmorgasboard note. The
   evaluator does NOT pixel-validate the PDF — the HTML test bed
   covers HTML output, and PDF rendering is downstream of HTML.
   The gate here is:
   - The PDF file exists on disk after the export action.
   - The first page of the PDF is non-empty (file size > some
     reasonable lower bound, e.g. > 5 KB; an empty / errored PDF
     is typically < 2 KB).
   - The Joplin renderer process did NOT log
     "Cannot read properties of undefined" or similar error
     during the export. (Joplin's renderer log lives at
     `~/Library/Logs/Joplin/log.txt` on macOS; the evaluator
     greps it for the export window.)

   This criterion is intentionally loose because PDF pixel parity is
   a brittle gate (Electron's Chromium version, system fonts,
   page-break decisions all shift outputs); the value here is
   "exporting doesn't crash" and "the rendered HTML the user sees
   in preview is what gets baked into the PDF."

8. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** Generator captures own evidence under
   `screenshots/feature-1-m16-snapshot-to-html/`
   (`generator-evidence.png` + `.pixels.json` sidecar), reads via
   the Read tool to satisfy the `verify-gate` hook.

9. **`npm test` — all existing tests stay green plus new ones.**
   Baseline at session start (verify with `npm test`): **220**. M16
   adds AT LEAST 4 new tests:
   - 1 test for the base-shape render (criterion 1).
   - 1 test for multi-sheet (criterion 2).
   - 1 test for FormattingSmorgasboard render (criterion 3).
   - 1 test for CF render (criterion 4).
   Target: **220 → ≥ 224**. The full Jest suite is run before
   flipping the row, NOT just the new tests. M16 touches
   `src/index.ts` (content-script registration), introduces
   `src/contentScripts/notesheetRenderer.ts` (the new entry), and
   updates `plugin.config.json` (extraScripts entry list) and
   `webpack.config.js` (if the existing `buildExtraScripts` config
   doesn't already cover the new entry shape). Nothing else SHOULD
   regress; the renderer is read-only on the snapshot and does NOT
   touch the import/export pipeline (`src/xlsx.ts`).

**Out of scope**

- **Charts in HTML export.** Notesheet's anchored Chart.js charts
  don't survive into static HTML output. The renderer emits cell
  values for cells under the chart anchor, but the chart canvas is
  dropped. Listed cell values still render. Document as M16-followup
  or M17 dependency.
- **Live formula evaluation in the renderer.** HTML export uses the
  cached `cell.v` value from the snapshot. If a formula's cached
  value is stale (e.g. user typed in Univer but didn't save before
  exporting), the stale value renders. Acceptable.
- **Univer's `IDocumentBody.dataStream` rich-text per-run
  formatting** — i.e., M13/D's bold word + plain word in one cell.
  M16 renders the cell's plain text concatenation. Per-run
  formatting in HTML is M16-followup. Document.
- **Theme palette resolution beyond what's already baked into the
  snapshot.** Colours come from the snapshot's resolved `bg.rgb`
  and `cl.rgb` fields; the renderer does NOT re-walk theme
  references. M13/E's synthesizer already bakes the theme-resolved
  colours into per-cell `bg`/`cl` so this works correctly for
  table-styled fixtures.
- **Print stylesheets / page breaks.** Joplin's PDF export decides
  page sizing and break points; the renderer does NOT override.
- **iconSet glyphs as inline SVG.** Static glyph rendering for CF
  icon sets is a non-trivial feature unto itself. M16-followup.
- **DataBar gradient rendering.** Same — needs CSS gradient + per-
  cell bar width math. M16-followup.
- **Column widths and row heights.** The M16 renderer emits a
  default-sized `<table>` and lets the browser flow it. Excel's
  exact column-width replication is not pursued. If needed,
  `<col style="width: ...">` could be added later from the
  snapshot's `columnData` — out of scope for M16.
- **Number-format precision parity with Excel.** `cell.v` may be
  a number with full float precision. The renderer ships a
  reasonable `numFmt` formatter for the common cases (percent,
  currency, dates) but does NOT replicate every Excel format
  string. Document gaps.
- **Stop-If-True CF semantics in the static render.** If the M15
  fixture or any other adds a `stopIfTrue: true` rule, the M16
  renderer applies all matching CF rules in order regardless. M15
  already documented this gap on the Univer-side; M16 inherits it
  for HTML.
- **CF rules with operators not in the M15 set** (text-based,
  time-period, unique/duplicate, formula-based). Same M15 gap;
  the renderer emits no fill for those rules.
- **CF dataBar and iconSet rendering in HTML.** Static HTML can
  approximate dataBar via inline `<div style="width:N%">` background
  fill, and iconSet via inline SVG glyphs, but both require careful
  per-row math (dataBar's value-to-width interpolation; iconSet's
  cfvo-band lookup). Punt to M16-followup. The visible cell value
  still renders for those columns.
- **Cell hyperlinks rendering as `<a>`.** M12 ships per-cell
  hyperlinks via Univer's customRange shape; emitting them as `<a
  href="...">` in the HTML output is straightforward but punted to
  M16-followup unless trivial. The visible cell value still
  renders.
- **README "Known shortcomings — Joplin's Export → PDF / HTML"
  entry update.** Per the M13/C, M13/D, M13/E, M15 precedent,
  punted to a follow-up docs PR. The evaluator does NOT penalise
  if the docs edit is deferred.

**Suggested fixtures**

- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard.xlsx`
  — Aptos workbook with table styling. The central fixture for
  criterion 3 and the visual gate (criterion 5). Tests M13/E's
  banding fidelity flowing through to HTML.
- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx`
  — multi-sheet test for criterion 2's per-sheet rendering.
- `tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
  — exercises M15's CF rules through the renderer for criterion 4.

NO new fixtures are needed. The M16 acceptance surface is fully
covered by the existing catalog under
`tests/ExcelBaseTestData/formatting-testdata/`.

The PGE harness invokes `import-fixture.sh
FormattingSmorgasboard.xlsx`, lands the note, then
`eval-screenshot.sh feature-1-m16-snapshot-to-html` for the
preview-pane-targeted PNG and pixel sidecar.

**Related risks**

- **Joplin's content script API runs in a sandboxed renderer
  worker.** Per `api/JoplinContentScripts.d.ts` and Joplin's
  plugin samples, the entry exports a `default function(context)`
  that returns `{ plugin: (markdownIt, pluginOptions) => {...},
  assets: {...} }`. The function executes in Joplin's renderer,
  NOT the plugin process. It cannot import from the plugin's main
  bundle directly. The renderer must be self-contained — bundle
  `extractSnapshot()` (from `src/snapshot.ts`) into the content-
  script entry. Webpack handles this via the existing
  `buildExtraScripts` config; the new entry is added to
  `plugin.config.json`'s `extraScripts` list.

- **Webpack target for content scripts.** The existing
  `buildExtraScripts` config produces `commonjs2` bundles by
  default and `IIFE` bundles for browser-bound scripts (see
  `webpack.config.js:345`). Joplin's content scripts run in a
  Node-like context within Joplin's renderer process; they expect
  the `module.exports = function(context) { ... }` shape — i.e.
  `commonjs2`. The generator should NOT add the new entry to the
  `browserExtraScripts` set. If the renderer module fails to load
  with a "no default export" error, that's the symptom — confirm
  the bundle output uses `module.exports = ...` (commonjs2), not
  `(function() { ... })()` (IIFE).

- **The renderer is called for every note**, not just Notesheet
  notes. The first thing the fence handler does is check the
  fence info string for `notesheet v=1`; if it doesn't match,
  return undefined (or call markdown-it's default fence handler)
  to let other notes render normally. Don't break non-Notesheet
  notes.

- **Cell value formatting via `numFmt`.** `cell.v` is the cached
  value (number, string, boolean). For cells with a `numFmt`
  (e.g. `"0%"`, `"$"#,##0`, `"yyyy-mm-dd"`), the renderer should
  format the value through the pattern before emitting. exceljs's
  `numFmt` syntax is the source. Out-of-scope to fully replicate
  Excel's number-format engine; ship something reasonable for
  common cases (percent, currency, dates) and document gaps. The
  built-in numFmt code 0 (`General`) just stringifies the value.

- **CF rules don't paint cells in the snapshot — Univer's CF
  engine paints them at canvas-render time.** For the M16
  renderer to show CF colours, it has to evaluate CF rules
  itself: walk the snapshot's `SHEET_CONDITIONAL_FORMATTING_PLUGIN`
  resource, apply per-rule logic (cellIs > 50, top-N rank, color
  scale interpolation by percentile), and bake the result into
  per-cell inline styles. `colorScale`, `cellIs`, and `top10`
  are doable in a static pass; `iconSet` and `dataBar` are
  harder (need glyph / bar element rendering) and punted as
  M16-followup. The renderer code lives in
  `src/contentScripts/notesheetRenderer.ts` (or `cfEvaluator.ts`
  pulled in from there); split internal logic so the next cycle
  can extend.

- **Merged cells affect table layout significantly.** A merge
  spans rows/cols; the renderer must skip the cells inside the
  merge range when emitting subsequent `<td>` elements (else the
  HTML ends up with extra cells pushing the layout sideways).
  The snapshot's `mergeData` array provides the ranges as
  `{startRow, endRow, startColumn, endColumn}` (zero-based
  inclusive). Algorithm: build a `Set` of "cells inside a merge
  range but NOT the top-left anchor" before emitting; skip those
  cells; for the top-left anchor of each merge, emit `<td colspan
  rowspan>`.

- **The PGE harness adds a `previewPane` region helper this
  cycle.** Until M16, the harness eval-screenshot path was
  Univer-canvas-only. The preview pane is part of the main
  Joplin shell (NOT inside `UserWebviewIndex.html`), so the CDP
  page picker should attach to Joplin's main page. Selector:
  generator verifies via DevTools at session start; expected
  candidates include `iframe.note-viewer-iframe` (recent Joplin
  versions), `.rendered-md`, or `#rendered-md`. The picker
  scoring needs adjusting (the current heuristic favours the
  editor page for Notesheet notes; we want the SHELL page for
  this cycle). The harness fix lands as part of the cycle.

- **Window prep includes preview-visible toggle.** The existing
  `prep-joplin-window.sh` covers sidebar/note-list hiding and
  fill-window. M16 needs the preview pane visible — Joplin's
  default state is editor-only after toggling sidebar off. The
  generator extends `prep-joplin-window.sh` to detect preview
  pane visibility and toggle until both panes show. macOS
  shortcut: Cmd+L toggles editor only / split / preview only;
  the script needs state-aware toggling per the M13/C precedent
  for `prep-joplin-panes.js`.

- **Don't symptom-patch test failures post-build.** If a test
  regresses after registering the content script, run
  `git diff package-lock.json` first. The new content-script
  entry may have pulled new transitives (unlikely — the renderer
  uses only stdlib + existing snapshot helpers — but always
  audit). Reference: `feedback_dependency_hygiene.md`. Do NOT
  downgrade exceljs or any other dep to make a symptom go away.

- **The feature touches `src/index.ts` (plugin entry registrations)
  and adds `src/contentScripts/notesheetRenderer.ts` (new file).
  It does NOT touch `src/xlsx.ts` (import/export pipeline) or
  `src/editorView.tsx` (Univer wiring).** The full Jest suite must
  stay green — borders, hyperlinks, table styles, chart export,
  alignment, rotation (M13/C), rich text (M13/D), theme banding
  (M13/E), conditional formatting (M15) all share the import
  pipeline. M16 SHOULD be orthogonal — but run the whole suite
  anyway.

- **Test gap warning** (`feedback_pge_fidelity_test_gap.md`). The
  M16 Jest tests must assert what's IN the rendered HTML against
  source-fixture expected content (cell values from the snapshot,
  colour values from the snapshot's already-validated styles
  — those styles have their own M13/E pin-downs upstream), not
  against whatever format the M16 renderer happens to emit. The
  HTML output format CAN change without test failure as long as
  the rendered values + colours match. Concretely: pin "the bold
  cell carries `font-weight: bold`", not "the renderer emits
  `<td class='b'>`". Pin "the header `<td>` carries
  `background-color: #34692E`", not "the renderer emits a
  particular wrapping span shape". The HTML emit style is
  implementation detail — the cell-by-cell rendered content is
  the contract.

- **CF colour evaluation is duplicated logic.** The Univer CF
  preset (M15) already evaluates rules at canvas-render time;
  the M16 renderer evaluates them again in HTML-render time.
  This is not deduplicatable without restructuring the snapshot
  → renderer pipeline through Univer's headless engine, which
  is itself a multi-cycle project. M16 ships its own evaluator
  for cellIs, top10, and colorScale; the evaluator code lives
  inside the content-script bundle and is unit-tested via the
  Jest tests in criterion 4. If a cycle later wants to factor
  out the evaluator into a shared module, that's its own
  refactor.

- **Editor preview pane DOM stability.** Joplin's preview
  iframe DOM is not part of Joplin's plugin API contract —
  selectors can change between Joplin versions. The harness
  region helper should query for multiple candidate selectors
  in order and take the first hit. M13/E's `tableHeaderRowRegion`
  is the template (graceful fallback). If Joplin changes its
  preview-pane DOM, the eval screenshot fails loudly with "no
  preview-pane element found" — the generator updates the
  selector list, doesn't paper over with a hardcoded screenshot
  region.

- **Generator MUST audit any npm install.** If adding a new
  dev dep is needed for the content-script entry (markdown-it
  type stubs, etc.), the generator runs `npm install --save-dev`
  for the specific package, then surfaces the full
  `git diff package.json package-lock.json` to the operator
  per `feedback_dependency_hygiene.md`. Major-version transitive
  shifts are flagged; downgrades are blocked unless the
  operator explicitly approves.

- **The `extractSnapshot()` import path.** The content-script
  entry can `import { extractSnapshot } from '../snapshot';`
  via a relative path; webpack bundles `src/snapshot.ts` into
  the renderer output. Avoid pulling `src/xlsx.ts` — it's heavy
  (exceljs, JSZip) and the renderer doesn't need it. The
  fence-parsing logic in `extractSnapshot()` is sufficient.

- **Smoke note leak risk** (`feedback_known_shortcomings_over_bugs.md`).
  The PR #17 smoke seed in `emptySnapshot()` was a debug-state
  leak that shipped to production for ~6 weeks. M16 introduces a
  NEW entry point — make sure the renderer doesn't carry any
  test-bed defaults (e.g. a "render this if the snapshot is
  malformed" placeholder string that the user might see in
  their PDF). If the snapshot is malformed, the renderer should
  fall through to markdown-it's default fence rendering (which
  shows the JSON, the existing pre-M16 behaviour) — explicit
  failure beats silent injection.

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file (and the prior M13/C, M13/D, M13/E, M15 BUILD_PLAN.md
files preserved in git history at `dc80505`, `420d583`, `fca1cbc`,
`88297a1`) as the template. Each feature gets:

- `### feature-N-<kebab-id>` heading (stable; never renamed once
  written).
- **Spec**: one paragraph naming the user-observable change.
- **Acceptance criteria**: numbered list of observable evidence
  (HTML content, pixel-sidecar, Jest, runtime). Tests anchored
  UPSTREAM of our code (source fixture content for HTML; the
  operator-captured Excel reference PNG for canvas-vs-Excel
  features; the source XML for snapshot fidelity). NO assertions
  pinned to our own emit; NO "code does X" — only outcomes the
  evaluator can inspect from a fresh context.
- **Out of scope**: explicit non-goals so the evaluator does not
  penalise for them.
- **Suggested fixture(s)**: which file(s) under
  `tests/ExcelBaseTestData/formatting-testdata/` exercise this.
- **Related risks**: regression hot-spots and prior-bug pointers,
  including pointers into prior PR / commit hashes when prior
  work is the starting point.

The lowest-numbered feature with `passes: false` in
`test-results.json` is what the generator works on next.
