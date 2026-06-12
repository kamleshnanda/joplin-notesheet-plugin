# Notesheet — Spreadsheets for Joplin

Notesheet turns a Joplin note into a real spreadsheet. Powered by the [Univer SDK](https://github.com/dream-num/univer), it gives Joplin first-class support for formulas, formatting, sorting, filtering, named tables, conditional formatting, anchored charts, and `.xlsx` import / export — all inside the note editor pane you already use.

## Features

### Spreadsheets in any note

A new **"New Spreadsheet"** command (Tools menu, toolbar button, or `Cmd/Ctrl+Shift+S`) creates a note that opens directly in a Univer-powered spreadsheet editor. Persistence, sync, and full-text search all use Joplin's normal note storage. Notes that aren't spreadsheets open in the regular markdown editor — Notesheet adds zero overhead to non-spreadsheet notes.

### Formulas

- Univer's full Excel-compatible formula engine (`@univerjs/engine-formula`) runs inside the editor. Formulas recalculate on every cell edit and on snapshot load.
- Excel structured-reference formulas like `=Table1[[#This Row],[Investment]]` resolve natively because table definitions are preserved in the snapshot and registered with Univer's formula engine on load.

### Formatting

- Fonts (theme-default workbook fonts like Aptos Narrow / Calibri preserved), fills, alignment, rotated text, number formats, borders, merged cells.
- Rich text within a single cell (multi-run bold / colour / italic).
- Conditional formatting: color scale, data bar, cell-is, top-N, icon set — all five types round-trip.
- Theme-aware named-table styling: TableStyleMedium variants paint correctly under the workbook's own `<a:clrScheme>` (Aptos accent3 paints green, Office Classic accent3 paints grey, etc.) instead of being baked to a single hardcoded palette.

### Tables

- **Insert Table** from the Data ribbon; right-click inside a table for row / column insert / remove.
- Built-in styles (TableStyleMedium2, TableStyleMedium4, etc.) are preserved on round-trip.

### Anchored charts

Insert ribbon → **Insert Chart** opens a docked panel that mirrors your live cell selection. Charts are drag/resizable, pinned to the grid, and update live when source cells change. Bar / line / pie / doughnut via [Chart.js](https://www.chartjs.org/) (MIT). Charts export to native OOXML (`xl/charts/chart*.xml`) — opening the exported file in Excel produces a real Excel chart, not a screenshot — and **charts authored in Excel import back into Notesheet** and render live in the editor (M17). The round-trip preserves chart type, source range (including cross-sheet references), titles, legend position, axis styling, data labels, bar orientation / grouping / gap width, doughnut hole size, and per-series **trendlines** (linear fit drawn over the series with its equation + R² label). Pie / doughnut slice labels that don't fit inside a slice are pushed outside with leader lines, the way Excel lays them out.

> Univer's own chart packages (`@univerjs-pro/sheets-chart`) are commercial / require a license server. Notesheet ships a custom Chart.js + Univer drawing-preset integration instead so the floating overlay stays open-source.

### `.xlsx` import / export

- **Import / Export** buttons in the editor view, plus a Tools menu command **"Import .xlsx as Notesheet"** that creates a new note from a `.xlsx` file.
- Round-trips: values, formulas (including structured references), fonts, fills, alignment, rotation, number formats, borders, merged cells, named tables with built-in style, hyperlinks (both `{text, hyperlink}` cell values and the named-Hyperlink cell style), workbook theme palette (`<a:clrScheme>`), conditional formatting, rich-text runs, **charts** (bar / line / pie / doughnut with their definitions, anchors, and trendlines), workbook default font size, and per-sheet default column width.
- Workbook theme is preserved on round-trip so the same `TableStyleMediumN` resolves the same accent in the exported file.

### PDF / HTML export of rendered spreadsheet

A Markdown-It content script (`src/contentScripts/notesheetRenderer.ts`) renders Notesheet fenced bodies as inline-styled HTML tables for Joplin's preview pane and PDF / HTML export. Common Excel number formats render correctly; conditional formatting (cellIs / top-N / colorScale) bakes into the static HTML.

## Compatibility

- **Joplin 3.5+** — desktop only. Joplin Mobile's plugin model isn't yet ready for the Custom Editor API.
- **`.xlsx`**: tested against fixtures saved by Microsoft Excel and openpyxl.

## Known gaps

These are pinned by `KNOWN SHORTCOMING` Jest tests so accidental changes are caught.

### Editor

- **Left arrow in column A / Up arrow in row 1** jumps the cursor to the bottom-right of the sheet (upstream Univer bug, [dream-num/univer#6988](https://github.com/dream-num/univer/issues/6988)). Workaround: navigate with the mouse.
- **Cmd/Ctrl+K** opens Joplin's markdown link dialog instead of Univer's link dialog. Use the Univer toolbar's Insert → Link from inside the spreadsheet. Imported `.xlsx` hyperlinks are preserved and clickable; this affects only typing a new link via the keyboard shortcut.

### `.xlsx` import / export

- **Theme-tinted borders** (`{theme: N, tint: T}`) resolve against whichever `<a:clrScheme>` is loaded at import time. After round-trip the resolved RGB is fixed in the snapshot, so a later host-side theme change won't update the rendering.
- **Unsupported chart types** (radar, scatter, area, bubble, 3-D, etc.) import as a `bar` fallback with `meta.unsupportedSourceType` recording the original type; bar / line / pie / doughnut import faithfully.
- **Images / non-chart drawings don't round-trip through `.xlsx` (M18).** You _can_ insert an image via the editor's Insert ribbon (Univer's built-in image tool) and it persists in the Notesheet note across save/reload — it's stored in the snapshot's `SHEET_DRAWING_PLUGIN` as a base64 image drawing. But the `.xlsx` layer does not yet read or write image (or shape) drawings: importing an image-bearing workbook drops the picture (the cells import fine, no crash), and exporting a note that contains an inserted image omits it from the `.xlsx`. Chart drawings are the only drawing type that round-trips today. Full image/shape `.xlsx` round-trip is tracked for M18.
- **Multi-sheet workbooks with a named table on each sheet** trip exceljs's table-reduce — `xlsx-multi-table-unsupported`. Workaround: move all tables onto one sheet.
- **Other exceljs reconcile failures** surface as `xlsx-import-failed` with the original error preserved in `.cause`.

### PDF / HTML export

The Markdown-It renderer reads the snapshot directly without booting Univer. A few features don't have static-HTML equivalents:

- **dataBar conditional-formatting rules**: cells render with the value but no bar fragment. cellIs / top-N / colorScale do bake into HTML.
- **iconSet conditional-formatting rules**: cells render with the value but no glyph (arrows, traffic lights, etc.).
- **Anchored charts**: Chart.js canvases don't survive into static HTML. Cell values referenced by the chart still render; the chart visual does not.
- **Per-run rich text**: a cell with bold + plain runs in one cell renders as plain text in HTML.
- **Krona-pattern accounting symbol position**: Excel's `_-* #,##0.00 "kr"_-` puts the symbol after the number; Notesheet's accounting formatter places it before, so the layout differs from Excel for this specific pattern.
- **Accounting `_X` underscore-fill** (variable-width column-alignment construct) renders as a single space; numeric columns won't align as tightly as Excel's column-alignment fill.
- **Unsupported numFmt patterns** fall through to the raw stringified value.
- **Formula re-evaluation**: the renderer reads `cell.v` (cached value) directly. `cell.v` is kept fresh by Univer's formula engine in the editor — recalc happens on every cell edit and on snapshot load, then `workbook.save()` writes the result back. Every code path that persists a snapshot to a Joplin note goes through Univer, so a note opened in the editor at least once has up-to-date formula values. The narrow case where the renderer can show stale values: someone hand-edits the JSON inside a `notesheet v=1` markdown fence and views the preview before opening the editor. Opening the note triggers a save and the next preview is correct. We deliberately do not ship a second formula engine inside the renderer (Univer's is 1.6MB minified; reimplementing 470+ Excel functions would mean two engines drifting apart over time and a multi-MB bundle on every Joplin note open).

## Developer notes

### How a Notesheet note is stored

A Spreadsheet note is a regular Joplin note whose body is a Univer snapshot wrapped in a fenced markdown code block tagged ` ```notesheet `:

````
```notesheet v=1
{ "id": "...", "sheetOrder": [...], "sheets": {...}, ... }
```
````

When the active note's body matches that shape, Joplin's editor pane shows the Univer editor instead of the markdown editor (via Joplin's [Custom Editor API](https://joplinapp.org/api/references/plugin_api/classes/joplinviewseditor.html)). The fence's `v=1` marker is a forward-compatibility hook so a future on-disk format change can dispatch by version.

### Building

Requires Node.js 20+ (CI runs 20.x and 22.x).

```bash
npm install
npm run dist     # builds the .jpl into publish/
npm test         # runs Jest unit tests
```

The build produces `publish/com.kamleshnanda.joplin-notesheet.jpl`, installable in Joplin via **Tools → Options → Plugins → ⚙ → Install from file**.

### Milestones

|     | Milestone                                                                                                                                                                                                                                                              | PR                                                                                                                                                                                                                                                                                          |
| --- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| ✅  | M0 — Rebrand to Notesheet                                                                                                                                                                                                                                              | [#1](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/1)                                                                                                                                                                                                                        |
| ✅  | M1 — New Spreadsheet command + snapshot fence                                                                                                                                                                                                                          | [#2](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/2)                                                                                                                                                                                                                        |
| ✅  | M2 — Univer Custom Editor in Joplin's editor pane                                                                                                                                                                                                                      | [#3](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/3)                                                                                                                                                                                                                        |
| ✅  | M3 — Formatting (Univer core preset)                                                                                                                                                                                                                                   | —                                                                                                                                                                                                                                                                                           |
| ✅  | M4 — Sort & Filter                                                                                                                                                                                                                                                     | [#4](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/4)                                                                                                                                                                                                                        |
| ✅  | M5 — `.xlsx` import / export                                                                                                                                                                                                                                           | [#5](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/5)                                                                                                                                                                                                                        |
| ✅  | M6 — Named tables                                                                                                                                                                                                                                                      | [#6](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/6)                                                                                                                                                                                                                        |
| ✅  | M7 + M8 — Anchored Chart.js charts                                                                                                                                                                                                                                     | [#8](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/8)                                                                                                                                                                                                                        |
| ✅  | M9 — Excel structured-references + table fidelity + borders                                                                                                                                                                                                            | [#9](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/9)                                                                                                                                                                                                                        |
| ✅  | M10 — Chart export to `.xlsx` (native OOXML)                                                                                                                                                                                                                           | [#13](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/13)                                                                                                                                                                                                                      |
| ✅  | M11 — Dependency hygiene                                                                                                                                                                                                                                               | [#12](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/12)                                                                                                                                                                                                                      |
| ✅  | M12 — Formatting fidelity polish                                                                                                                                                                                                                                       | [#14](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/14) [#15](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/15)                                                                                                                                               |
| ✅  | PGE — Planner-Generator-Evaluator harness                                                                                                                                                                                                                              | [#17](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/17)                                                                                                                                                                                                                      |
| ✅  | M13/C — Rotated text round-trip                                                                                                                                                                                                                                        | [#19](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/19)                                                                                                                                                                                                                      |
| ✅  | M13/D — Rich-text within a single cell                                                                                                                                                                                                                                 | [#20](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/20)                                                                                                                                                                                                                      |
| ✅  | M13/E — Theme-aware banding accuracy                                                                                                                                                                                                                                   | [#22](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/22)                                                                                                                                                                                                                      |
| ❌  | M14 — SheetJS Community migration spike (NO-GO; see [`docs/m14-sheetjs-spike.md`](./docs/m14-sheetjs-spike.md))                                                                                                                                                        | [#24](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/24)                                                                                                                                                                                                                      |
| ✅  | M15 — Conditional formatting full round-trip                                                                                                                                                                                                                           | [#26](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/26)                                                                                                                                                                                                                      |
| ✅  | M16 — Snapshot → HTML for Joplin's PDF / HTML export                                                                                                                                                                                                                   | [#28](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/28) [#29](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/29) [#30](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/30) [#31](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/31) |
| ✅  | M17 — Chart import from `.xlsx` (chart definitions + trendlines + pie leader-line labels)                                                                                                                                                                              | [#32](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/32) [#33](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/33)                                                                                                                                               |
| ✅  | Roadmap & backlog grooming — collect M0→M17 pending scope ([`BACKLOG.md`](./BACKLOG.md)); split into M18–M20                                                                                                                                                           | [#34](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/34)                                                                                                                                                                                                                      |
| ✅  | Dev guidelines & repo-rules cleanup — consolidate contributor/agent guidelines ([`CONTRIBUTING.md`](./CONTRIBUTING.md), [`AGENTS.md`](./AGENTS.md)) and add enforcement (ESLint + Prettier + typecheck + dependency-downgrade guard, wired to pre-commit hooks and CI) | [#PENDING](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/PENDING)                                                                                                                                                                                                            |
| ⏳  | M18 — Drawings & charts: image / shape `.xlsx` round-trip, charts in HTML / PDF export, chart fidelity (per-series colours, rich-text titles, percentStacked, more types, markers, error bars) — backlog groups A–C                                                    | planned                                                                                                                                                                                                                                                                                     |
| ⏳  | M19 — Static-render & import gaps: dataBar / iconSet / rich-text in HTML export, accounting number formats, multi-table workbooks, theme-tinted borders, Univer keybinding bugs — backlog groups D–E                                                                   | planned                                                                                                                                                                                                                                                                                     |
| ⏳  | M20 — Codebase health: `uuid` CVE, transitive deprecation cleanup, exceljs watch-item — backlog group F                                                                                                                                                                | planned                                                                                                                                                                                                                                                                                     |

### Dependency hygiene

`npm install` and `npm audit` print warnings for transitive packages buried under our direct deps. We document rather than mask:

- **`uuid@8.3.2` moderate CVE** ([GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq)) — missing buffer bounds check in `uuid.v3`/`v5`/`v6`. Pulled in by exceljs. Not reachable: exceljs only calls the CVE-free random `uuid()` (v4). `npm audit fix --force` would downgrade exceljs to 3.4.0 (major-version downgrade), unacceptable. The replacement path was evaluated in the M14 spike and ruled NO-GO.
- **Transitive deprecation noise** (`inflight@1`, `rimraf@2`, `lodash.isequal`, `glob@7.x` × 4, `fstream@1`, `glob@10.x` × 3) — every entry is buried under exceljs (`>archiver`, `>unzipper`, `>fast-csv`) or jest internals. M11 already bumped jest to 30 to drop the deprecated transitive globs we could; the rest are upstream noise we can't act on from `package.json` without making something worse.
- **`glob@11.1.0`** (our direct devDep) — npm warns about "old versions of glob" for any glob; `glob@11.1.0` IS the current major. Ignore.

### `.xlsx` parser

`exceljs@^4.4.0` is the foundation of `src/xlsx.ts`. The M14 spike evaluated migrating to SheetJS Community and ruled NO-GO; full evidence is at [`docs/m14-sheetjs-spike.md`](./docs/m14-sheetjs-spike.md). The migration becomes worth re-evaluating only if (a) SheetJS Community proper (`xlsx`) adds first-class cell-style read support upstream, or (b) `exceljs` is publicly archived. Until then, the `xlsx-js-style` fork's parser-level loss of borders, alignment, rotation, and most font formatting on Excel-imported workbooks would directly regress shipped features.

### PGE harness (visual regression gate)

Jest unit tests cover the snapshot-data shape but cannot catch the failure mode where the data is correct and Univer's renderer ignores it. The **planner-generator-evaluator** harness adds a runtime visual gate on top of Jest:

- A **planner** agent translates an operator ask in `OPERATOR_ASK.md` into `BUILD_PLAN.md` with per-feature acceptance criteria phrased as user-observable outcomes.
- A **generator** agent picks the lowest-numbered `passes: false` row from `test-results.json`, builds the feature, captures a screenshot from the running Joplin desktop app, and only then flips its row.
- A **fresh-context evaluator** subprocess runs after the generator. It captures its OWN screenshot (Playwright over CDP, attached to the running Joplin) and grades PASS or NEEDS_WORK from the bytes — no memory of what the generator did, no plausibility bias.

Each evaluator screenshot is paired with a `<screenshot>.pixels.json` sidecar holding the top non-background colours sampled from the Univer canvas, so colour-sensitive assertions can be machine-checkable.

#### Running a cycle

```bash
# Joplin dev profile only — never touches your main profile.
./scripts/pge/launch-joplin.sh        # starts Joplin --env dev with CDP on :8315
./scripts/pge/install-plugin.sh       # builds + copies the .jpl into the dev profile
./scripts/pge/run-cycle.sh            # one PGE cycle: generator → evaluator
```

#### Files

- `.claude/CLAUDE.md` — generator runtime contract
- `.claude/agents/{planner,generator,evaluator}.md` — agent briefs
- `.claude/hooks/verify-gate.sh` — denies Write to `test-results.json` until evidence is Read
- `scripts/pge/eval-screenshot.{sh,js}` — Playwright + CDP attach, frame drop into the editor's `UserWebviewIndex.html` iframe (where Univer mounts), wait on `canvas[id^="univer-sheet-main-canvas"]`, screenshot + pixel sidecar
- `scripts/pge/run-cycle.sh` — orchestrator (one cycle per invocation, no auto-loop)
- `BUILD_PLAN.md` / `OPERATOR_ASK.md` / `PROGRESS.md` / `AUDIT.md` / `test-results.json` — operator-readable harness state
- `screenshots/<feature-id>/` — committed visual evidence per feature (both the generator's drop and the evaluator's authoritative captures)

#### Hard rules baked into the harness

- Joplin **dev profile** only (`--env dev` → `~/.config/joplindev-desktop/`). The harness never touches the operator's main Joplin profile.
- `test-results.json` defaults every row to `{ passes: false }`. A `verify-gate` hook denies any Write to that file until the agent has first Read a screenshot.
- The evaluator runs in a separate `claude` subprocess and has only `Read` / `Glob` / `Grep` / `Bash` — no `Write`, no `Edit`, so it cannot quietly fix problems instead of reporting them.

Pattern adapted from [anthropics/cwc-long-running-agents](https://github.com/anthropics/cwc-long-running-agents).

## License

MIT. Bundled libraries: Univer SDK is Apache-2.0, [exceljs](https://github.com/exceljs/exceljs) is MIT, [Chart.js](https://www.chartjs.org/) is MIT, [JSZip](https://stuk.github.io/jszip/) is MIT-or-GPLv3 (used under MIT).
