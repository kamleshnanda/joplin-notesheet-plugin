# Changelog

All notable changes will be documented here. Versioning follows [SemVer](https://semver.org/).

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
