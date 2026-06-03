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
| ⏳ | M13/E — Theme-aware banding accuracy (remaining M13 workstream) | planned |
| ⏳ | M14 — Snapshot → HTML for Joplin's PDF/HTML export menu | planned |
| ⏳ | M15 — Chart import from `.xlsx` | planned |
| ⏳ | M16 — Conditional formatting (color scale / data bar / cell-is / top-N / icon set) | planned |
| ⏳ | M17 — SheetJS Community migration spike: port `src/xlsx.ts` against [SheetJS Community](https://sheetjs.com) (Apache-2.0, active upstream), validate against the M12 + M13 fixture suite, decide whether to migrate. exceljs has gone quiet (last release Dec 2024) and ships stale transitives (uuid@8, glob@7). SheetJS reads conditional formatting upstream-correctly out of the box, and its sparse-dict cell model lines up structurally with Univer's snapshot. | planned |

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

- **Theme-aware banding**: Joplin renders synthesized table-style
  banding from a hardcoded catalog (Office 2007 RGBs). When a workbook
  uses a different theme — for example Aptos's `accent3 = #196B24`
  vs Office Classic's `#9BBB59` — the in-Joplin display can disagree
  with what the same file renders in Excel. The exported `.xlsx`
  preserves the source's `<a:clrScheme>`, so opening the round-tripped
  file back in Excel renders it correctly; only the in-Joplin paint
  is hardcoded. → M13/E.
- **Conditional formatting**: color scale / data bar / cell-is /
  top-N / icon-set rules are dropped on import and not re-emitted on
  export. Cell values themselves survive. → M16.
- **Theme-tinted borders**: `{theme: N, tint: T}` border colors are
  resolved against whichever `<a:clrScheme>` is loaded at import time.
  After round-trip the resolved RGB is fixed in the snapshot, so a
  later theme change in the host won't update the rendering.

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
  is upstream in exceljs (or migrate off — see M17).
- **Transitive deprecation noise** (`inflight@1`, `rimraf@2`,
  `lodash.isequal`, `glob@7.x` × 4, `fstream@1`, `glob@10.x` × 3):
  every entry is buried under exceljs (`>archiver`, `>unzipper`,
  `>fast-csv`) or jest@30 internals. We are already on jest@30
  (M11 bumped specifically to drop deprecated transitive globs);
  no further direct-dep change can clear these. Only an exceljs
  replacement does — see M17.
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
  and modern Excel emit. M15 will address this by reading the chart
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
