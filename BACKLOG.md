# Notesheet — Pending scope backlog

Every punted feature, deferred fidelity item, and known gap collected from
M0 → M17, across README known-gaps, PROGRESS notes, BUILD_PLAN /
OPERATOR_ASK out-of-scope sections, the M14 spike doc, and
`KNOWN SHORTCOMING` markers in tests/source.

Milestone split (operator-decided 2026-06-12):

- **M18** — backlog groups **A, B, C** (drawings/images, charts in HTML/PDF, chart fidelity)
- **M19** — backlog groups **D, E** (static-render gaps, editor/import limitations)
- **M20** — backlog group **F** (codebase health)

These are the milestone _buckets_; exact per-item scoping within each
milestone is decided when that milestone is planned. The `[in M18 row]`
tags below are historical (the pre-split README row) — the authoritative
mapping is the group→milestone split above.

_Last collected: 2026-06-11 (after M17 + #33 merged); milestone split added 2026-06-12._

## A. Drawings & images — M18

- **A1 — Image drawings round-trip through `.xlsx`** **[in M18 row]**. Import reads `xl/media/*` + the drawing anchor; export writes them back. Today: images insert via the Insert ribbon and persist in the note (snapshot `SHEET_DRAWING_PLUGIN`, base64), but the `.xlsx` layer drops them both ways. _Source: README known-gaps + M18 row._
- **A2 — Shape drawings** (non-chart, non-image) round-trip **[in M18 row]**. _Source: README._
- **A3 — EMU sub-cell anchor offset** preserved on import/export. Today we round drawings to the nearest cell, dropping the sub-cell EMU offset. _Source: OPERATOR_ASK out-of-scope._

## B. Charts in static export (HTML / PDF) — M18

- **B1 — Charts render in HTML / PDF / preview-pane export** **[in M18 row]**. Hand-author inline SVG in the M16 content script (`src/contentScripts/notesheetRenderer.ts`) using the same `CHART_PALETTE`. Was M17 feature-7, never implemented. Today Chart.js canvases don't survive to static HTML; cell values referenced by the chart still render. _Source: README PDF/HTML gaps, m17 memory._

## C. Chart fidelity follow-ups — M18

- **C1 — Per-series chart colours** from Excel `<c:spPr>` **[in M18 row]**. We always use `CHART_PALETTE`; imported charts look like Notesheet charts, not pixel-identical to source. _Source: m17 memory, OPERATOR_ASK._
- **C2 — Rich-text chart titles** **[in M18 row]**. Per-run bold/colour/size on a chart title is dropped; plain concatenated text is kept. _Source: m17 memory, OPERATOR_ASK._
- **C3 — percentStacked normalisation to 100%.** Currently routes to Chart.js stacked but doesn't normalise series to 100%. _Source: m17 memory._
- **C4 — More chart types** (radar, scatter, area, bubble, 3-D). Currently import as a `bar` fallback with `meta.unsupportedSourceType`. _Source: README, OPERATOR_ASK (M17.x)._
- **C5 — Per-series line marker shapes.** We read the chart-level marker toggle only; per-series `<c:marker><c:symbol>` shapes are dropped. _Source: xlsxChartImport.ts:460._
- **C6 — Error bars, multiple data-label positions, axis titles.** Not imported/rendered. _Source: OPERATOR_ASK._

## D. PDF / HTML renderer gaps (from M16) — M19

The Markdown-It renderer reads the snapshot without booting Univer; these have no static-HTML equivalent yet.

- **D1 — dataBar CF rules**: value renders, no bar fragment. _Source: README._
- **D2 — iconSet CF rules**: value renders, no glyph. _Source: README._
- **D3 — Per-run rich text in HTML**: bold + plain runs in one cell render as plain text. _Source: README._
- **D4 — Krona-pattern accounting symbol position**: `_-* #,##0.00 "kr"_-` — symbol placed before the number, Excel puts it after. _Source: README._
- **D5 — Accounting `_X` underscore-fill** (column-alignment construct) renders as a single space. _Source: README._

## E. Editor / `.xlsx` import limitations — M19

- **E1 — Multi-sheet workbooks with a named table on each sheet** trip exceljs's table-reduce (`xlsx-multi-table-unsupported`). Workaround: one sheet for all tables. A real import blocker for some workbooks. _Source: README._
- **E2 — Theme-tinted borders** (`{theme, tint}`) re-resolve against the active clrScheme on round-trip, not the source workbook theme. _Source: README._
- **E3 — Left-arrow in col A / Up-arrow in row 1** jumps cursor to bottom-right (upstream Univer bug [#6988](https://github.com/dream-num/univer/issues/6988)). Not ours to fix; track upstream. _Source: README._
- **E4 — Cmd/Ctrl+K** opens Joplin's markdown-link dialog instead of Univer's. Host keybinding conflict. _Source: README._

## F. Codebase health (non-functional — first-class delivery) — M20

- **F1 — `uuid` moderate CVE** ([GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq)), transitive via exceljs. Not reachable today (exceljs uses v4 random only), but a standing advisory. `npm audit fix --force` would downgrade exceljs to 3.4.0 (blocked major downgrade). _Source: README dep-hygiene._
- **F2 — Transitive deprecation noise** (`inflight@1`, `rimraf@2`, `glob@7.x`, `fstream@1`, `lodash.isequal`, etc.) buried under exceljs (`archiver`/`unzipper`/`fast-csv`) and jest. Mostly upstream; not actionable from `package.json` without making things worse. _Source: README dep-hygiene._
- **F3 — exceljs replacement (watch item).** SheetJS migration was NO-GO (M14). Revisit only if SheetJS Community adds first-class cell-style read upstream, OR exceljs is publicly archived. _Source: docs/m14-sheetjs-spike.md._

---

## README milestone rows

The README "Milestones" table now carries three planned rows matching the
split above: **M18** (groups A–C), **M19** (groups D–E), **M20** (group F).
Per-item scope within each milestone is finalised at that milestone's
planning time. Items E3/E4 are upstream Univer bugs (track, may not be
fixable by us); F2/F3 are watch-items more than actionable work.
