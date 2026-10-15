# Changelog

All notable changes will be documented here. Versioning follows [SemVer](https://semver.org/).

## Unreleased (M19)

- **Arrow keys stop at the sheet edge, as in Excel.** Left in column A, Up in row 1, Right in the last column and Down in the last row now leave the active cell where it is. Previously the cursor wrapped to the other side of the sheet (Left or Up at A1 jumped to the bottom-right cell). This also applies when an arrow key commits a cell edit, and hidden rows and columns at the edge are skipped. Shift+arrow, Cmd/Ctrl+arrow and arrow-key nudging of images are unchanged. Upstream Univer closed [dream-num/univer#6988](https://github.com/dream-num/univer/issues/6988) as by-design, so Notesheet overrides the behaviour itself.

## 1.0.1

Security and maintenance release (milestone M20). No change to how spreadsheets look or behave.

- **Security:** fixes every known vulnerability in the plugin's runtime dependencies (`npm audit --omit=dev`: 98 → 0).
    - `uuid` updated to 11.1.1 for the `.xlsx` engine ([GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq)), including inside the spreadsheet editor itself.
    - `nanoid` updated to 5.1.16 for the spreadsheet engine ([GHSA-28wg-ghj8-5hjv](https://github.com/advisories/GHSA-28wg-ghj8-5hjv)).
    - Assorted transitive patch updates.
- **Smaller download:** the plugin package shrank from 9.3 MB to 8.7 MB because the editor now bundles only the parts of the `.xlsx` engine it uses.
- `.xlsx` import and export were re-verified end-to-end in the editor: conditional formatting (data bars, color scales, icon sets) and embedded images round-trip unchanged.

## 1.0.0

First public release on the Joplin plugin marketplace (milestones M1–M18, M22; see the README Milestones table for per-PR detail).

- **Spreadsheet notes:** a note can be a full Univer spreadsheet (Tools → New Spreadsheet, `Cmd/Ctrl+Shift+S`), stored as a `notesheet v=1` fenced snapshot so sync, search and encryption work as usual.
- **Formulas, formatting, sort / filter, named tables, conditional formatting** (color scale, data bar, cell-is, top-N, icon set), hyperlinks, merged cells, rotated and rich text.
- **Charts:** bar / line / pie / doughnut, anchored to the grid and updating live, built on Chart.js.
- **`.xlsx` import / export** (Tools → Import .xlsx as Notesheet, plus editor buttons): values, formulas incl. structured references, styles, tables, theme palette, conditional formatting, charts, images, and preserve-only shapes.
- **Preview / PDF / HTML export** renders sheets as static HTML tables with SVG charts.
- Known gaps are listed under "Known gaps" in the README.

## Unreleased (pre-1.0 history)

### M0 — Notesheet rebrand and architectural reset

- Rebranded the plugin from "Univer Worksheet Plugin" to **Notesheet**.
- Plugin id changed to `com.kamleshnanda.joplin-notesheet`; npm package name to `joplin-plugin-notesheet`.
- Removed the v0 popup-spreadsheet-embedded-in-markdown model (custom URL scheme, fenced-markdown mirror, dialog UI, custom HTML/JS spreadsheet implementation).
- Added Jest + GitHub Actions CI scaffolding for v1.

This release contains no end-user spreadsheet functionality on its own. v1 is built incrementally over subsequent milestones (M1 reintroduces the Univer SDK as the engine and adds the Custom Editor view; M2+ adds inline preview, formatting, sort/filter, .xlsx import/export, charts, and tables).
