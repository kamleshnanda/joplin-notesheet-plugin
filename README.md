# Notesheet — Spreadsheets for Joplin

Notesheet turns a Joplin note into a real spreadsheet. Powered by the [Univer SDK](https://github.com/dream-num/univer), it gives Joplin first-class support for formulas, formatting, sorting, filtering, named tables, anchored charts, and `.xlsx` import/export — all inside the note editor pane you already use.

## Features

- **Spreadsheet inside any note.** A new "New Spreadsheet" command (Tools menu, toolbar button, or `Cmd/Ctrl+Shift+S`) creates a note that opens directly in a Univer-powered spreadsheet editor. Persistence, sync, and full-text search all use Joplin's normal note storage.
- **Formulas, formatting, sort, filter** — out of the box via Univer's standard toolbar.
- **Excel structured-reference formulas** — formulas like `=Table1[[#This Row],[Investment]]` from imported `.xlsx` files resolve natively because the table definition is preserved in the snapshot and registered with Univer's formula engine on load.
- **Named tables** — Insert Table from the Data ribbon; right-click inside a table for row/column insert/remove.
- **Anchored charts** — Insert ribbon → "Insert Chart" opens a docked panel that mirrors your live cell selection. Charts are drag/resizable, pinned to the grid, and update live when source cells change. Bar / line / pie / doughnut via Chart.js.
- **`.xlsx` import/export** — Import/Export buttons in the editor view, plus a Tools menu command "Import .xlsx as Notesheet" that creates a new note from a `.xlsx` file. Round-trips values, formulas (including structured references), fonts, fills, alignment, number formats, borders, merged cells, and named tables with their style.

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
| ⏳ | M10 — Chart export to `.xlsx` (native OOXML; planned, [plan](M10-plan.md) drafted) | — |
| ⏳ | M11 — Dependency hygiene: Jest 29 → 30 to drop deprecated transitive `glob@7` (planned) | — |
| ⏳ | M12 — Formatting fidelity polish: theme fonts, named-style banding, hyperlinks (planned) | — |
| ⏳ | M13 — Snapshot → HTML for Joplin's PDF/HTML export menu (planned) | — |
| ⏳ | M14 — Chart import from `.xlsx` (planned) | — |

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

- Pressing **Left arrow** in column A (or **Up arrow** in row 1) jumps the
  cursor to the bottom-right corner of the sheet. This is an upstream Univer
  bug — track at [dream-num/univer#6988](https://github.com/dream-num/univer/issues/6988).
  Workaround: navigate to the edge with the mouse instead.
- Joplin's **Export → PDF / HTML** menu (right-click on a note) currently
  exports the raw fenced JSON for Notesheet notes instead of a rendered
  table. To save a Notesheet as an Excel file, use the in-editor
  **Export .xlsx** button. PDF/HTML export of rendered spreadsheet content
  is planned for a future release.
- **Cmd/Ctrl+K** opens Joplin's markdown link dialog instead of inserting a
  link into the active spreadsheet cell. Use the Univer toolbar's Insert →
  Link option from inside the spreadsheet.
- **Hyperlinks** in imported `.xlsx` cells are flattened to their visible
  text. Round-trip preservation is planned for a future release.
- **Font fidelity**: workbook-default fonts (e.g. Aptos Narrow declared via
  the Excel theme rather than per-cell) are not preserved on import. Cells
  with explicit font names round-trip correctly.

## License

MIT. Bundled libraries: Univer SDK is Apache-2.0, [exceljs](https://github.com/exceljs/exceljs) is MIT, [Chart.js](https://www.chartjs.org/) is MIT, [JSZip](https://stuk.github.io/jszip/) is MIT-or-GPLv3 (we use it under the MIT terms).
