# Notesheet — Spreadsheets for Joplin

Notesheet turns a Joplin note into a real spreadsheet. Powered by the [Univer SDK](https://github.com/dream-num/univer), it gives Joplin first-class support for formulas, formatting, sorting, filtering, named tables, anchored charts, and `.xlsx` import/export — all inside the note editor pane you already use.

## Features

- **Spreadsheet inside any note.** A new "New Spreadsheet" command (Tools menu, toolbar button, or `Cmd/Ctrl+Shift+S`) creates a note that opens directly in a Univer-powered spreadsheet editor. Persistence, sync, and full-text search all use Joplin's normal note storage.
- **Formulas, formatting, sort, filter** — out of the box via Univer's standard toolbar.
- **Excel structured-reference formulas** — formulas like `=Table1[[#This Row],[Investment]]` from imported `.xlsx` files resolve natively because the table definition is preserved in the snapshot and registered with Univer's formula engine on load.
- **Named tables** — Insert Table from the Data ribbon; right-click inside a table for row/column insert/remove.
- **Anchored charts** — Insert ribbon → "Insert Chart" opens a docked panel that mirrors your live cell selection. Charts are drag/resizable, pinned to the grid, and update live when source cells change. Bar / line / pie / doughnut via Chart.js.
- **`.xlsx` import/export** — Import/Export buttons in the editor view, plus a Tools menu command "Import .xlsx as Notesheet" that creates a new note from a `.xlsx` file. Round-trips values, formulas (including structured references), fonts (theme-default workbook fonts like Aptos Narrow / Calibri preserved), fills, alignment (including text rotation), number formats, borders, merged cells, named tables with their built-in style (TableStyleMedium2 etc.), per-run rich-text formatting inside a single cell (bold + plain, multi-color, etc.), and hyperlinks (both `{text, hyperlink}` cell values and the named-Hyperlink cell style). Workbook theme palette (`<a:clrScheme>`) is preserved on round-trip, AND the in-Joplin renderer resolves table-style banding against the workbook's own accent palette so what Joplin shows matches what Excel renders. Workbooks with chart drawings or unusual zip layouts that previously crashed import are now pre-processed before they reach exceljs.

## Milestones

| | Milestone | PR |
|---|---|---|
| ✅ | M0 — Rebrand to Notesheet, strip v0 popup model | [#1](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/1) |
| ✅ | M1 — New Spreadsheet command + snapshot fence helpers | [#2](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/2) |
| ✅ | M2 — Univer Custom Editor renders in Joplin's editor pane | [#3](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/3) |
| ✅ | M3 — Formatting (already in Univer's core preset, no work needed) | — |
| ✅ | M4 — Sort & Filter via Univer presets | [#4](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/4) |
| ✅ | M5 — `.xlsx` import/export via exceljs | [#5](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/5) |
| ✅ | M6 — Named tables via Univer preset | [#6](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/6) |
| ✅ | M7 + M8 — Anchored Chart.js charts with live updates | [#8](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/8) |
| ✅ | M9 — Excel structured-references + table import/export fidelity + borders | [#9](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/9) |
| ✅ | M10 — Chart export to `.xlsx` (native OOXML) | [#13](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/13) |
| ✅ | M11 — Dependency hygiene: Jest 29 → 30 to drop deprecated transitive `glob@7` | [#12](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/12) |
| ✅ | M12 — Formatting fidelity polish: theme fonts, named-style banding, hyperlinks (Pattern A `{text, hyperlink}` + Pattern B named cell style), workbook theme palette round-trip, table grid border synthesis, friendly errors for unsupported `.xlsx` shapes | (this PR) |
| ✅ | M13 — Theme-aware banding accuracy, rotated text, rich-text within a single cell, import recovery (chart drawings + absolute rel Targets stripped before exceljs) | (this PR) |
| ⏳ | M14 — Snapshot → HTML for Joplin's PDF/HTML export menu | planned |
| ⏳ | M15 — Chart import from `.xlsx` | planned |
| ⏳ | M16 — Conditional formatting (color scale / data bar / cell-is / top-N / icon set) | planned |

> **Note on charts:** Univer's chart packages (`@univerjs-pro/sheets-chart`) are commercial / require a license server, so M7 + M8 ship a custom integration with [Chart.js](https://www.chartjs.org/) (MIT) and Univer's open-source drawing preset for the floating overlay.

## How a Notesheet note is stored

A Spreadsheet note is a regular Joplin note whose body is a Univer snapshot wrapped in a fenced markdown code block tagged ```` ```notesheet ````:

```
```notesheet v=1
{ "id": "...", "sheetOrder": [...], "sheets": {...}, ... }
```
```

When the active note's body matches that shape, Joplin's editor pane shows the Univer editor instead of the markdown editor (via Joplin's [Custom Editor API](https://joplinapp.org/api/references/plugin_api/classes/joplinviewseditor.html)). For any other note, the markdown editor opens normally — Notesheet adds zero overhead to non-spreadsheet notes.

The fence's `v=1` marker is a forward-compatibility hook so a future on-disk format change can dispatch by version.

## Development

```bash
npm install
npm run dist     # builds the .jpl into publish/
npm test         # runs Jest unit tests
```

Requires Node.js 18+. The build produces `publish/com.kamleshnanda.joplin-notesheet.jpl`, which can be installed in Joplin via **Tools → Options → Plugins → ⚙ → Install from file**.

## Compatibility

- Joplin 3.5+
- Desktop only (Joplin Mobile's plugin model isn't yet ready for the Custom Editor API)

## Known issues

### Editor

- Pressing **Left arrow** in column A (or **Up arrow** in row 1) jumps the
  cursor to the bottom-right corner of the sheet. This is an upstream Univer
  bug — track at [dream-num/univer#6988](https://github.com/dream-num/univer/issues/6988).
  Workaround: navigate to the edge with the mouse instead.
- Joplin's **Export → PDF / HTML** menu (right-click on a note) currently
  exports the raw fenced JSON for Notesheet notes instead of a rendered
  table. To save a Notesheet as an Excel file, use the in-editor
  **Export .xlsx** button. PDF/HTML export of rendered spreadsheet content
  is planned (M14).
- **Cmd/Ctrl+K** opens Joplin's markdown link dialog instead of Univer's
  link insertion UI for the active cell. Use the Univer toolbar's Insert →
  Link option from inside the spreadsheet. (Imported `.xlsx` hyperlinks are
  preserved and clickable; this affects only typing a new link via the
  keyboard shortcut.)

### `.xlsx` import — known shortcomings

The shape of these is "things that survive a round-trip cleanly through
the in-Joplin editor but render differently between Excel and Joplin
itself, OR features Excel supports that Notesheet doesn't yet model."
Each is pinned by a `KNOWN SHORTCOMING` test in
`tests/m12FixtureRoundTrip.test.ts` so we notice if the behavior
accidentally changes.

- **Conditional formatting**: color scale / data bar / cell-is /
  top-N / icon-set rules are dropped on import and not re-emitted on
  export. Cell values themselves survive. → M16.
- **Theme-tinted borders**: `{theme: N, tint: T}` border colors are
  resolved against whichever `<a:clrScheme>` is loaded at import time.
  After round-trip the resolved RGB is fixed in the snapshot, so a
  later theme change in the host won't update the rendering.

### `.xlsx` import — unforeseen crashes (handled with friendly errors)

If exceljs's reconcile pipeline ever throws on a workbook shape we
haven't pre-processed (M13 strips chart drawings + rewrites absolute
rel Targets to relative form before exceljs sees the buffer, which
unblocked all known crash classes), the catch wrapper produces a
`NotesheetImportError` with code `xlsx-import-failed` and preserves
the original error in `.cause`. The host UI shows a single-line
explanation instead of a raw stack trace.

The two M12-era classifications (`xlsx-charts-unsupported`,
`xlsx-multi-table-unsupported`) are no longer reached by any known
fixture but remain wired up — they fire if a future workbook variant
re-exposes the underlying exceljs bug. Charts in imported workbooks
are silently dropped on import; chart-as-data import is M15.

## License

MIT. Bundled libraries: Univer SDK is Apache-2.0, [exceljs](https://github.com/exceljs/exceljs) is MIT, [Chart.js](https://www.chartjs.org/) is MIT, [JSZip](https://stuk.github.io/jszip/) is MIT-or-GPLv3 (we use it under the MIT terms).
