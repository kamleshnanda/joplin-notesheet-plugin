# Notesheet — Spreadsheets for Joplin

Notesheet turns a Joplin note into a real spreadsheet. Powered by the [Univer SDK](https://github.com/dream-num/univer), it gives Joplin first-class support for formulas, formatting, sorting, filtering, named tables, anchored charts, and `.xlsx` import/export — all inside the note editor pane you already use.

## Features

- **Spreadsheet inside any note.** A new "New Spreadsheet" command (Tools menu, toolbar button, or `Cmd/Ctrl+Shift+S`) creates a note that opens directly in a Univer-powered spreadsheet editor. Persistence, sync, and full-text search all use Joplin's normal note storage.
- **Formulas, formatting, sort, filter** — out of the box via Univer's standard toolbar.
- **Excel structured-reference formulas** — formulas like `=Table1[[#This Row],[Investment]]` from imported `.xlsx` files resolve natively because the table definition is preserved in the snapshot and registered with Univer's formula engine on load.
- **Named tables** — Insert Table from the Data ribbon; right-click inside a table for row/column insert/remove.
- **Anchored charts** — Insert ribbon → "Insert Chart" opens a docked panel that mirrors your live cell selection. Charts are drag/resizable, pinned to the grid, and update live when source cells change. Bar / line / pie / doughnut via Chart.js.
- **`.xlsx` import/export** — Import/Export buttons in the editor view, plus a Tools menu command "Import .xlsx as Notesheet" that creates a new note from a `.xlsx` file. Round-trips values, formulas (including structured references), fonts (theme-default workbook fonts like Aptos Narrow / Calibri preserved), fills, alignment, number formats, borders, merged cells, named tables with their built-in style (TableStyleMedium2 etc.), and hyperlinks (both `{text, hyperlink}` cell values and the named-Hyperlink cell style). Workbook theme palette (`<a:clrScheme>`) is preserved on round-trip so the same `TableStyleMediumN` resolves the same accent in the exported file.

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
| ✅ | M12 — Formatting fidelity polish: theme fonts, named-style banding, hyperlinks (Pattern A `{text, hyperlink}` + Pattern B named cell style), workbook theme palette round-trip, table grid border synthesis, friendly errors for unsupported `.xlsx` shapes | [#14](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/14) [#15](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/15) |
| ✅ | PGE — Planner-Generator-Evaluator harness for runtime visual gating (catches the M13 failure mode where Jest passes but Univer renders broken) | [#17](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/17) |
| ✅ | M13/C — Rotated text round-trip, validated via the PGE harness | [#19](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/19) |
| ✅ | M13/D — Rich-text within a single cell (multi-run bold / colour / italic), validated via the PGE harness | [#20](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/20) |
| ✅ | M13/E — Theme-aware banding accuracy: TableStyle synthesis driven by the source `<a:clrScheme>`, with a reference-anchored fidelity test bed against operator-captured Excel renders | [#22](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/22) |
| ❌ | M14 — SheetJS Community migration spike: investigated migrating `src/xlsx.ts` from `exceljs` to [`xlsx-js-style`](https://www.npmjs.com/package/xlsx-js-style) (the only SheetJS Community fork with cell-style support). **NO-GO** — the fork drops borders, alignment / rotation, and most font formatting from the OOXML indexed-cellXf path that every Microsoft-Excel-generated workbook uses; migrating would directly regress M12 / M13/C / M13/E features that already ship. Spike branch closed without merging; the decision document, capability matrix, and 14 golden-snapshot baselines are preserved at [`docs/m14-sheetjs-spike.md`](./docs/m14-sheetjs-spike.md). See "M14 — SheetJS evaluation outcome" below for the short version. | [#24 (closed)](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/24) |
| ✅ | M15 — Conditional formatting full round-trip (color scale / data bar / cell-is / top-N / icon set) on `exceljs` + Univer's CF preset wired into the editor; reference-anchored fidelity test bed against the operator-captured Excel render | [#26](https://github.com/kamleshnanda/joplin-notesheet-plugin/pull/26) |
| ⏳ | M16 — Snapshot → HTML for Joplin's PDF/HTML export menu. Independent of the `.xlsx` parser choice — operates on the in-memory snapshot. | planned |
| ⏳ | M17 — Chart import from `.xlsx` (drawings + chart definitions, currently `xlsx-charts-unsupported`). | planned |

> **Note on charts:** Univer's chart packages (`@univerjs-pro/sheets-chart`) are commercial / require a license server, so M7 + M8 ship a custom integration with [Chart.js](https://www.chartjs.org/) (MIT) and Univer's open-source drawing preset for the floating overlay.

## M14 — SheetJS evaluation outcome

`exceljs@^4.4.0` is the foundation of `src/xlsx.ts`. exceljs has gone quiet (last release December 2024) and ships stale transitives (`uuid@8` with a moderate CVE that we tolerate as not-reachable from our call sites; `glob@7` deprecated). M14 was a research spike to evaluate migrating to SheetJS Community.

**Result: NO-GO.** Notesheet stays on `exceljs`. The full evidence is in [`docs/m14-sheetjs-spike.md`](./docs/m14-sheetjs-spike.md); the short version:

- **Plain SheetJS Community (`xlsx`) does NOT support cell styles.** It's a non-starter for our use because every shipped feature from M9 onwards depends on per-cell styling.
- **The only community fork with styling support, [`xlsx-js-style@1.2.0`](https://www.npmjs.com/package/xlsx-js-style), drops most styling on Excel-imported workbooks.** The spike ran both parsers across all 14 fixtures under [`tests/ExcelBaseTestData/formatting-testdata/`](./tests/ExcelBaseTestData/formatting-testdata/) and recorded a per-dimension parity matrix:
  - **Borders** — 0 of 5 fixtures preserve borders on the SheetJS path. Direct M12 regression.
  - **Rotated text** — 0 of 1 fixtures preserve rotation. Direct M13/C regression.
  - **Style records** — 50–95% of indexed-cellXf style records lost on every Microsoft-Excel-generated workbook.
  - **Rich text per-run** — runs flattened; only recoverable via a hand-written raw-XML walker (which is what `xlsx-js-style` is supposed to do for us).
  - **Self-roundtrip works** — a workbook xlsx-js-style writes can be read back with borders intact. But that's interop-with-itself, not interop-with-Excel; the latter is what every operator-imported fixture exercises.
- **`xlsx-js-style` is a 2022 fork on a 2022 base.** It hasn't shipped a release since. Migrating would replace one quiet library (`exceljs`, last release Dec 2024) with one that has been silent for longer.
- **Phase 2 cost was estimated at 9.5–14.5 days**, most of which would be Notesheet building in-house replacements for what the styling fork is named after (a `xl/styles.xml` walker, table export, rich-text export, etc.). At that point we're maintaining a Notesheet-internal Excel parser, not benefiting from a third-party library — the dependency-hygiene argument inverts.

**Revisit conditions:** the migration becomes worth re-evaluating if either (a) SheetJS Community proper (`xlsx`) adds first-class cell-style read support upstream, or (b) `exceljs` is publicly archived, forcing the migration regardless of cost. Until either is true, NO-GO stands.

The spike's preserved artefacts at [`docs/m14-sheetjs-spike.md`](./docs/m14-sheetjs-spike.md) include a 24-fixture × 14-dimension capability matrix, an honest assessment of the fork's maintenance state, the exact migration cost breakdown, and a Phase-2 feature-smoke checklist for the 11 features the operator listed as MUST-NOT-REGRESS — useful inputs for any future revisit.

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

Requires Node.js 20+ (CI runs 20.x and 22.x). The build produces `publish/com.kamleshnanda.joplin-notesheet.jpl`, which can be installed in Joplin via **Tools → Options → Plugins → ⚙ → Install from file**.

### PGE harness (visual regression gate)

Jest unit tests cover the snapshot-data shape but cannot catch the failure mode where the data is correct and Univer's renderer ignores it (M13 shipped that bug twice). The **planner-generator-evaluator** harness adds a runtime visual gate on top of Jest:

- A **planner** agent translates an operator ask in `OPERATOR_ASK.md` into `BUILD_PLAN.md` (per-feature acceptance criteria phrased as user-observable outcomes, not data shape).
- A **generator** agent picks the lowest-numbered `passes: false` row from `test-results.json`, builds the feature, captures a screenshot from the running Joplin desktop app, and only then flips its row.
- A **fresh-context evaluator** subprocess runs after the generator. It captures its OWN screenshot (Playwright over CDP, attached to the running Joplin) and grades PASS or NEEDS_WORK from the bytes — no memory of what the generator did, no plausibility bias.

Each evaluator screenshot is paired with a `<screenshot>.pixels.json` sidecar holding the top non-background colours sampled from the Univer canvas, so colour-sensitive assertions ("A1 rendered red") can be machine-checkable instead of human-eyeball-only.

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
- `scripts/pge/eval-screenshot.{sh,js}` — Playwright + CDP attach, frame drop into the editor's `UserWebviewIndex.html` iframe (where Univer actually mounts), wait on `canvas[id^="univer-sheet-main-canvas"]`, screenshot + pixel sidecar
- `scripts/pge/run-cycle.sh` — orchestrator (one cycle per invocation, no auto-loop)
- `BUILD_PLAN.md` / `OPERATOR_ASK.md` / `PROGRESS.md` / `AUDIT.md` / `test-results.json` — operator-readable harness state
- `screenshots/<feature-id>/` — committed visual evidence per feature (both the generator's drop and the evaluator's authoritative captures)

#### Hard rules baked into the harness

- Joplin **dev profile** only (`--env dev` → `~/.config/joplindev-desktop/`). The harness never touches the operator's main Joplin profile.
- `test-results.json` defaults every row to `{ passes: false }`. A `verify-gate` hook denies any Write to that file until the agent has first Read a screenshot.
- The evaluator runs in a separate `claude` subprocess and has only `Read` / `Glob` / `Grep` / `Bash` — no `Write`, no `Edit`, so it cannot quietly fix problems instead of reporting them.

Pattern adapted from [anthropics/cwc-long-running-agents](https://github.com/anthropics/cwc-long-running-agents).

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
  is planned (M16).
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

- **Theme-aware banding**: Joplin renders synthesized table-style
  banding from a hardcoded catalog (Office 2007 RGBs). When a workbook
  uses a different theme — for example Aptos's `accent3 = #196B24`
  vs Office Classic's `#9BBB59` — the in-Joplin display can disagree
  with what the same file renders in Excel. The exported `.xlsx`
  preserves the source's `<a:clrScheme>`, so opening the round-tripped
  file back in Excel renders it correctly. M13/E (PR #22) made the
  in-Joplin paint clrScheme-aware via an empirical-override map keyed
  by `(styleName, accentHex)` for the two project-owned fixtures
  (Aptos `TableStyleMedium4` accent3 and Classic `TableStyleMedium4`
  accent3). Workbooks whose `(styleName, accentHex)` pair isn't in
  the override map fall through to an HSL-L tint formula that
  approximates the right hue but can drift by ~Δ18 RGB units. The
  M13/E rework (PR #22) shipped with a known Univer renderer gap
  where the totals row's bottom-border accent strip rendered in the
  header colour rather than the lighter accent strip Excel uses; a
  re-capture during the M16 cycle (2026-06-07) confirmed both top
  and bottom totals strips now render correctly in `#72D068` — the
  gap appears to have resolved through downstream snapshot-shape
  changes between M13/E and M16. Pinned by
  `tests/excelCanvasFidelity.test.ts` and the
  `screenshots/feature-1-m13-theme-aware-banding/eval-aptos-*.png`
  evidence.
- **Theme-tinted borders**: `{theme: N, tint: T}` border colors are
  resolved against whichever `<a:clrScheme>` is loaded at import time.
  After round-trip the resolved RGB is fixed in the snapshot, so a
  later theme change in the host won't update the rendering.

### `.xlsx` Markdown export — known shortcomings

The Markdown-It content script (`src/contentScripts/notesheetRenderer.ts`)
shipped in M16 renders Notesheet fenced bodies as HTML tables for
Joplin's PDF / HTML / preview-pane export. Coverage is intentional:
common Excel formats render correctly; a small set of complex patterns
ship with documented approximations. Each shortcoming carries a
`KNOWN SHORTCOMING` Jest test in
`tests/m16NotesheetMarkdownRender.test.ts` so a future change can't
silently regress.

- **Krona-pattern accounting symbol position**: Excel's
  `_-* #,##0.00 "kr"_-` pattern positions the currency symbol AFTER
  the number (`1,234.56 kr`). Notesheet's generic accounting formatter
  shares the layout with US Accounting (`"$"*`) where the symbol is a
  prefix, so the krona variant renders with the symbol in a different
  position than Excel. Pinned down at
  `tests/m16NotesheetMarkdownRender.test.ts:'Krona accounting pattern'`.
- **Accounting `_(`/`_)`/`* ` underscore-fill rendered as single space**:
  Excel's `_X` directive means "fill with as much space as character
  X would take" — a variable-width column-alignment construct.
  HTML doesn't have a column-alignment fill character; the M16
  renderer emits a single space per fill marker. Result: numbers
  won't align in vertical columns the way Excel would. Pinned down at
  `tests/m16NotesheetMarkdownRender.test.ts:'Excel Accounting positive'`
  and the negative / zero variants beneath it.
- **dataBar (CF rule) rendering**: Excel's data bars render
  proportional horizontal bars inside cells. The M16 HTML renderer
  doesn't synthesize the bar fragments — those cells render with
  their cell value only, no bar. Pinned down implicitly: the CF
  evaluator at `formatNumberWithPattern` documents which CF types
  bake into HTML.
- **iconSet (CF rule) rendering**: Excel's icon sets render glyphs
  (arrows, traffic lights, etc.) inside cells. The M16 HTML renderer
  doesn't synthesize the glyphs. Same scope rationale as dataBar.
- **Charts in HTML export**: Notesheet's anchored Chart.js charts
  don't survive into static HTML. Cell values referenced by the
  chart still render; the chart canvas itself does not.
- **Live formula re-evaluation**: HTML export uses the cached `cell.v`
  value from the snapshot (the value exceljs evaluated at last
  Notesheet save). If a formula's cached value is stale, the stale
  value renders.
- **Per-run rich text**: M13/D's bold-word + plain-word in one cell
  renders as plain text in HTML. Per-run formatting in HTML export
  is M16-followup.
- **Unsupported numFmt patterns**: any pattern not in the supported
  set (Tier 1+2+3 above) falls through to the raw stringified value.
  Pinned down at
  `tests/m16NotesheetMarkdownRender.test.ts:'unknown pattern returns raw stringified value'`.

### Tolerated transitive deprecations + audit warnings

`npm install` and `npm audit` print warnings for transitive packages
buried under our direct dependencies. The summary below documents
which we tolerate and why; we do not paper over them with `npm
audit fix --force` or `overrides` blocks because both fixes
introduce silent regression risk worse than the warnings themselves.

- **`uuid@8.3.2` → moderate CVE
  [GHSA-w5hq-g745-h8pq](https://github.com/advisories/GHSA-w5hq-g745-h8pq)**
  (missing buffer bounds check in `uuid.v3`/`v5`/`v6` when `buf`
  arg is supplied). Pulled in by exceljs. **Not reachable** —
  exceljs only calls the CVE-free `uuid()` (random v4) for
  identifier generation; `v3`/`v5`/`v6` are never invoked.
  `npm audit fix --force` would downgrade exceljs to 3.4.0
  (a major-version DOWNGRADE), which is unacceptable. Real fix
  is upstream in exceljs (migrating off was evaluated in M14 and
  ruled NO-GO — see [`docs/m14-sheetjs-spike.md`](./docs/m14-sheetjs-spike.md)).
- **Transitive deprecation noise** (`inflight@1`, `rimraf@2`,
  `lodash.isequal`, `glob@7.x` × 4, `fstream@1`, `glob@10.x` × 3):
  every entry is buried under exceljs (`>archiver`, `>unzipper`,
  `>fast-csv`) or jest@30 internals. We are already on jest@30
  (M11 bumped specifically to drop deprecated transitive globs);
  no further direct-dep change can clear these. Only an exceljs
  replacement does — that path was evaluated in M14 and ruled NO-GO
  (see [`docs/m14-sheetjs-spike.md`](./docs/m14-sheetjs-spike.md));
  the noise is accepted indefinitely until the revisit conditions
  in that document are met.
- **`glob@11.1.0`** (our direct devDep) — npm prints a blanket
  "old versions of glob" warning for any glob it sees, but
  `glob@11.1.0` IS the current major. Ignore.

These warnings are real upstream signals; we just can't act on them
from `package.json` without making something worse.

### `.xlsx` import — unsupported shapes (handled with friendly errors)

These trip exceljs's internal reconcile pipeline. Notesheet wraps the
crash with a `NotesheetImportError` carrying a stable `code` so the
host UI can show an actionable message instead of a raw stack trace.

- **Workbooks with chart drawings** → `xlsx-charts-unsupported`. exceljs's
  drawing-reconcile crashes on chart anchor structures that openpyxl
  and modern Excel emit. M17 will address this by reading the chart
  structure directly with a lightweight OOXML reader similar to the
  M10 export path (and pre-stripping drawings before they reach exceljs
  on the read path).
- **Multi-sheet workbooks where each sheet has its own named table**
  → `xlsx-multi-table-unsupported`. exceljs's table-reduce crashes when
  more than one sheet has a `<tableParts>` block.
- **Other reconcile failures** → `xlsx-import-failed` (generic wrapper,
  preserves the original error message in `.cause`).

You can still import the same workbooks if you remove the offending
content (delete charts, or move all tables onto one sheet) and
re-save in Excel before importing.

## License

MIT. Bundled libraries: Univer SDK is Apache-2.0, [exceljs](https://github.com/exceljs/exceljs) is MIT, [Chart.js](https://www.chartjs.org/) is MIT, [JSZip](https://stuk.github.io/jszip/) is MIT-or-GPLv3 (we use it under the MIT terms).
