# M14 — SheetJS Community migration spike (Phase 1)

**Status:** Research-only spike. Phase 2 (the actual production swap) is a
separate PR and ships only if this document's recommendation is GO and the
golden snapshots in `tests/golden-snapshots/` stay structurally identical
under the new parser.

**Recommendation: NO-GO at `xlsx-js-style@1.2.0` as it ships today.**
The fork's import path silently drops borders, alignment / rotation, and
font formatting from the OOXML `xl/styles.xml` indexed-cellXf path that
every Microsoft-Excel-generated workbook uses. Empirically verified across
14 formatting fixtures — see the capability matrix below. A CONDITIONAL
recommendation is achievable IF Phase 2 ports a raw-XML styles.xml walker
(replacing what xlsx-js-style fails to do); see "Conditional GO" at the end.

---

## Why this matters

`exceljs@^4.4.0` is the entire foundation of `src/xlsx.ts` (1959 lines of
import/export logic). exceljs has gone quiet (last release December 2024)
and ships stale transitives:

- `uuid@8` — moderate CVE (`GHSA-w5hq-g745-h8pq`); not reachable from our
  call sites but flagged by `npm audit`.
- `glob@7` — deprecated.

The plan was to migrate to SheetJS Community via the styling-fork
`xlsx-js-style@^1.2.0`, with the operator-explicit constraint _"I DO NOT
WANT FEATURE REGRESSION ON FEATURES THAT ARE ALREADY WORKING."_ The spike
proves whether the migration is feasible.

---

## What ships in this spike PR

1. **`src/xlsxSheetJS.ts`** — parallel parser module (dead code at runtime;
   not imported by `src/index.ts` or `src/editorView.tsx`). Public surface
   matches `src/xlsx.ts`: `xlsxBufferToSnapshot`, `snapshotToXlsxBuffer`,
   `class NotesheetImportError`, and the two resource-name constants
   (`SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN`,
   `SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN`).
2. **`tests/xlsxParserParity.test.ts`** — runs both parsers across all 24
   `.xlsx` fixtures (14 formatting + 10 chart) and writes a per-dimension
   matrix to `tests/golden-snapshots/parity-matrix.json`. Test exits 0
   regardless of divergences — divergences are the artefact, not failures.
3. **`tests/goldenSnapshots.test.ts`** + **`tests/golden-snapshots/*.json`**
   — 14 baseline goldens captured from the current `src/xlsx.ts` output
   (volatile ids — `workbook-<ts>`, `tbl-…-<rand>`, `tblcol-N-<rand>`,
   `lnk-<rand>` — scrubbed to `…-STABLE` so re-runs don't churn). Phase 2
   must reproduce these byte-for-byte (or document each divergence and get
   operator approval) before merging.
4. **`devDependencies` update**: `xlsx-js-style@^1.2.0` is added; production
   `dependencies` is byte-identical to main. The fork pulls 14 transitive
   `dev: true` packages (`adler-32`, `cfb`, `codepage`, `commander`,
   `crc-32`, `exit-on-epipe`, `fflate`, `frac`, `printj`, `ssf`, `wmf`,
   `word`, plus the package itself + its inner duplicate `adler-32@1.3.1`).
   `npm audit` reports two pre-existing moderate vulns (the same exceljs /
   uuid pair) — **no new advisories** introduced by `xlsx-js-style`.

The webpack production build is byte-identical (`13412864` bytes both
before and after the dependency add) — the spike module and the
xlsx-js-style runtime are tree-shaken out because nothing in the
production tree imports them. Verified via `grep -c 'xlsxSheetJS\|xlsx-js-style'`
on `dist/{contentScript,editorView,index}.js` → 0 hits.

---

## Capability matrix

Per-dimension parity for the 24 fixtures, derived from
`tests/golden-snapshots/parity-matrix.json`. Status legend:

- **match** — both parsers produce structurally equivalent output.
- **divergence** — values present on both sides, differ in shape or count.
- **sheetjs-blocked** — xlsx-js-style returns nothing where exceljs does.
- **exceljs-blocked** — exceljs throws / drops where xlsx-js-style does not.
- **n/a** — fixture doesn't exercise this dimension.

| Dimension                                           | exceljs (current)                                               | xlsx-js-style@1.2.0                                                                                                                                                                                               |                                             Match | Diverge | sjs-blocked | n/a | Severity                                                                                                                                                       |
| --------------------------------------------------- | --------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------: | ------: | ----------: | --: | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| sheetSizes                                          | shape correct                                                   | divergence on workbooks where sjs imports more cells than exc throws on                                                                                                                                           |                                                17 |       4 |           0 |   0 | low — both correct on success                                                                                                                                  |
| cellValues                                          | string/number/boolean/date                                      | match on most; date / inline-rich text differ                                                                                                                                                                     |                                                16 |       5 |           0 |   0 | low — divergence is date encoding (Date object vs ISO string) and sjs-flattened rich text                                                                      |
| formulas                                            | structured refs (`Table1[#All]`) preserved                      | preserved as raw string; cached `v` set to 0 (drops exceljs's cached result)                                                                                                                                      |                                                 2 |       0 |           0 |  19 | LOW — formulas still evaluate in Univer at editor open                                                                                                         |
| mergedCells                                         | full                                                            | full                                                                                                                                                                                                              |                                                 1 |       0 |           0 |  20 | none                                                                                                                                                           |
| styleRecords (count)                                | 1 record per unique cell style                                  | massively under-populated for Microsoft-generated files                                                                                                                                                           |                                                 0 |       6 |          10 |   5 | **CRITICAL**                                                                                                                                                   |
| borders                                             | per-side style + colour, theme-tinted resolution                | **dropped** — `wb.Styles.Borders` is `[{}, {}, {}, ...]` empty for every fixture                                                                                                                                  |                                                 0 |       0 |           5 |  16 | **CRITICAL — direct M12 regression**                                                                                                                           |
| rotatedText (`alignment.textRotation`)              | full (M13/C)                                                    | **dropped** — alignment field absent on indexed-style cells                                                                                                                                                       |                                                 0 |       0 |           1 |  20 | **CRITICAL — direct M13/C regression**                                                                                                                         |
| numberFormats                                       | preserved                                                       | preserved                                                                                                                                                                                                         |                                                 5 |       0 |           0 |  16 | none                                                                                                                                                           |
| hyperlinks Pattern A (`{text, hyperlink}`)          | full                                                            | full (`c.l.Target`)                                                                                                                                                                                               |                                                 4 |       0 |           0 |  17 | none                                                                                                                                                           |
| hyperlinks Pattern B (named cellStyle="Hyperlink")  | parser-agnostic via `readNamedHyperlinkCells` (raw XML + JSZip) | parser-agnostic — same helper reused verbatim in spike                                                                                                                                                            | n/a (covered as part of styleRecords / Pattern A) |         |             |     | none                                                                                                                                                           |
| richText per-run (M13/D)                            | full via exceljs's `cell.value.richText`                        | **flattened** — runs collapsed into a single `c.v` string + a `c.h` HTML span with empty styles. Spike WORKED AROUND this with a raw-XML `<is><r><rPr>` walker (`readInlineRichTextRuns` in `src/xlsxSheetJS.ts`) |                                                 1 |       0 |           0 |  20 | medium — workaround required, not native                                                                                                                       |
| themeClrScheme (`<a:clrScheme>`)                    | parser-agnostic via `readThemeClrScheme` (raw XML + JSZip)      | parser-agnostic — same helper reused verbatim in spike                                                                                                                                                            |                                                21 |       0 |           0 |   0 | none                                                                                                                                                           |
| defaultStyle (theme minorFont)                      | filled from theme1.xml `minorFont`                              | not filled — sjs doesn't expose theme's font scheme                                                                                                                                                               |                                                 0 |       0 |          21 |   0 | medium — Phase-2 fix is a `readThemeMinorFont` helper, parser-agnostic via JSZip                                                                               |
| named tables (`SHEET_TABLE_PLUGIN`)                 | full table walk via raw XML + exceljs                           | not implemented in spike                                                                                                                                                                                          |                                                 0 |       0 |           3 |  18 | medium — Phase 2 cost; the existing `readTablesFromXlsxZip` is JSZip-driven and parser-agnostic, so a port is mechanical                                       |
| conditional formatting                              | dropped on import (existing shortcoming)                        | dropped (`!cf` undefined)                                                                                                                                                                                         |                           n/a — neither preserves |         |             |     | none — same behaviour                                                                                                                                          |
| charts (drawings + chart parts)                     | passed through; export injects via `injectChartsIntoZip`        | passed through opaquely (raw zip parts retained in `wb.files`)                                                                                                                                                    |                                                10 |       0 |           0 |  11 | UNKNOWN — chart-export round-trip with sjs-emitted zip is a Phase-2 question; matrix lists pass-through but does NOT verify the post-processor's compatibility |
| defined names                                       | exceljs surfaces them on `wb.definedNames`                      | sjs surfaces `wb.Workbook.Names` (empty array on every fixture we tested — fixtures don't use defined names)                                                                                                      |                   n/a — neither fixture exercises |         |             |     | medium — port is mechanical                                                                                                                                    |
| import error surface (typed `NotesheetImportError`) | catches exceljs reconcile crashes (charts / multi-table)        | sjs imports succeed where exceljs throws (3 fixtures: `ConditionalFormatting-Variants`, `LargeWorkbook`, `MultiSheet`)                                                                                            |                                               n/a |         |             |     | LOW — sjs is more lenient; would IMPROVE Notesheet's error surface, not regress it                                                                             |
| performance (LargeWorkbook 18,876 cells)            | (not measured this cycle)                                       | parses in 31 ms cold                                                                                                                                                                                              |                                               n/a |         |             |     | none                                                                                                                                                           |

**Imports succeeded**: exceljs 21/24, sjs 24/24. The 3 fixtures sjs imports
that exceljs throws on (`ConditionalFormatting-Variants.xlsx`,
`LargeWorkbook.xlsx`, `MultiSheet.xlsx`) are wins for sjs in isolation —
but per the operator's "no regression" constraint, the typed
`xlsx-charts-unsupported` / `xlsx-multi-table-unsupported` errors that
exceljs raises today are part of the contract: `src/index.ts` shows them
as a user-actionable dialog. A more-lenient sjs path would need to either
preserve the same errors (via post-import validation) OR a deliberate
operator decision to relax the contract.

---

## What xlsx-js-style cannot do (the NO-GO drivers)

Three findings from the spike, each independently sufficient to block GO:

### 1. Borders are completely dropped from indexed-cellXf cells

For every Microsoft-Excel-generated fixture the spike ran (the entire
`tests/fixtures/formatting-testdata/` set), `wb.Styles.Borders` is
an array of empty `{}` objects, and the per-cell `c.s` field carries no
border information. This was verified by reading
`xl/styles.xml` directly: the file DOES contain
`<border><left style="thin"><color rgb="FF000000"/></left>...</border>`
records, but xlsx-js-style's parser fills the cell-style register with `{}`
for every entry.

This is a direct M12 / M13/E feature regression:

- `BordersAndCellColors.xlsx` (32 cells, 11 unique border styles in
  source) — the spike's snapshot has `bd: {…}` only for 0 cells.
- `border-isolation.xlsx` (the operator-built M13/E fixture with explicit
  DOUBLE / THIN borders) — same: 11 borders in source, 0 in sjs snapshot.
- `FormattingSmorgasboard.xlsx`, `FormattingSmorgasboard-NonAptosClassic…`
  — the entire table-decoration round-trip goes blank because synthesizer
  inputs (raw cell border records) are missing.

The fork's own self-roundtrip works: a workbook xlsx-js-style writes can
be read back with borders intact. But that's interop-with-itself, not
interop-with-Excel — the latter is what every operator-imported fixture
exercises.

### 2. Alignment / rotation / wrap is completely dropped

`MergedCellsAndAlignment.xlsx` (M13/C's primary fixture) — the source has
three rotated cells (`alignment.textRotation = 45`, `90`, `-45`). After
sjs read, `c.s` is `{patternType: 'none'}` for all three. The
`alignment` field is NEVER populated on indexed-cellXf cells. Direct
M13/C regression: the rotation Univer renders as up-right diagonal /
vertical / down-right diagonal in Notesheet today would render flat
under a SheetJS-driven snapshot.

### 3. Rich-text per-run is dropped (workaround required)

`RichTextInOneCell.xlsx` (M13/D) — A1 has `<r><rPr><b/></rPr><t>Hello</t></r>
<r><t> world</t></r>`. After sjs read, `c.v` is the string
`"Hello world"` and `c.r` is `<t>Hello world</t>` (the bold is gone).
A2 has three colour runs (red / black / blue); after sjs read, `c.h` is
`<span style="">Red</span> and <span style="">Blue</span> text` — empty
style spans, colours stripped.

The spike works around this with a raw-XML walker
(`readInlineRichTextRuns` + `parseInlineRichRuns` in
`src/xlsxSheetJS.ts`, ~80 lines) that re-parses `<is><r><rPr>...<t>...</t></r></is>`
directly and reconstructs the runs. With the workaround the parity matrix
shows `richText: match (2 multi-run cells on both)`. **But that's the
spike doing what xlsx-js-style refuses to do**, not a capability of the
fork itself. Phase 2 would have to ship and maintain this walker.

---

## Honest assessment of `xlsx-js-style` maintenance state

`xlsx-js-style` is a **fork** of SheetJS Community (`xlsx@0.18.5`). Per the
fork's npm page (https://www.npmjs.com/package/xlsx-js-style), the latest
version `1.2.0` was published in 2022 and the package has not had a
release since. The base SheetJS Community version it forks (`xlsx@0.18.5`)
was published in 2022 as well.

The fork's selling point — "SheetJS Community plus styling" — is
critically dependent on the styling implementation matching what
Microsoft Excel produces. The spike proves it does NOT for the most basic
property in OOXML: cell borders. A 2022 abandonment date strongly
suggests this gap will not close upstream.

**Phase 2's trade**: replace one quiet library (`exceljs`, last release
December 2024) with another (`xlsx-js-style`, last release 2022).
`exceljs` is currently in a maintenance lull but its codebase is wider —
it has cell-style, border, alignment, rich-text, and table-style support
that work against Excel-generated files today, and a reachable user base.
`xlsx-js-style` works against itself but loses 60–80% of the styling
information for any Excel-imported file. The migration would not be a
strict improvement; it would be a different set of trade-offs, and the
trade goes the wrong direction for Notesheet's "match Excel" promise.

The plain `xlsx` (SheetJS Community proper) is actively maintained and
still ships from sheetjs.com — but it explicitly does NOT support cell
styles, so it's a non-starter for our use.

---

## Migration cost estimate (if we went ahead anyway)

This is the cost to bring xlsx-js-style up to functional parity for the
MUST-NOT-REGRESS feature list. Each item is the additional work Phase 2
would need to land beyond what the spike module already does.

| Task                                                                                                             | Estimate              | Notes                                                                                                                                                                                                                                                                                                                                                          |
| ---------------------------------------------------------------------------------------------------------------- | --------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Raw-XML styles.xml walker (per-cell border / alignment / font reconstruction)                                    | 2–3 days              | This is the work xlsx-js-style is supposed to do for us. Walks `xl/styles.xml`'s `<cellXfs>` and resolves `<font>`, `<fill>`, `<border>`, and `<alignment>` records into per-cell style objects.                                                                                                                                                               |
| Theme minorFont reader                                                                                           | 0.5 day               | Parser-agnostic JSZip + regex over `xl/theme/theme1.xml`. Direct port of `readThemeFont` from `src/xlsx.ts`.                                                                                                                                                                                                                                                   |
| Tables port (`readTablesFromXlsxZip`, `buildTableJsonForSheet`, `synthesizeTableStyleAssignments`)               | 2 days                | The existing helpers are mostly parser-agnostic (raw XML through JSZip); just need re-glueing to the SheetJS workbook model.                                                                                                                                                                                                                                   |
| Theme palette splice on export (`patchThemeClrScheme`)                                                           | 0.5 day               | Direct port — the helper is already parser-agnostic, just needs to be wired into the SheetJS-emitted zip post-processing pipeline.                                                                                                                                                                                                                             |
| Theme font splice on export (`patchThemeFont`)                                                                   | 0.5 day               | Same as above.                                                                                                                                                                                                                                                                                                                                                 |
| Chart-export post-processor compatibility (`injectChartsIntoZip`)                                                | 1–2 days (UNKNOWN)    | `src/charts/xlsxChart.ts` (572 lines) patches the zip exceljs writes. The SheetJS-emitted zip has a different internal structure (different `[Content_Types].xml`, different relationship ordering, different empty-defaults). Concrete test: write a SheetJS-emitted chart fixture through `injectChartsIntoZip` and see what breaks. Not done in this spike. |
| Rich-text per-run walker (port `readInlineRichTextRuns` from spike → production)                                 | already done in spike | The spike's `parseInlineRichRuns` is reusable; just needs to be promoted out of `src/xlsxSheetJS.ts`.                                                                                                                                                                                                                                                          |
| Pattern B hyperlink port (`readNamedHyperlinkCells`)                                                             | 0 (verbatim reuse)    | Already parser-agnostic.                                                                                                                                                                                                                                                                                                                                       |
| Tables export (write `xl/tables/tableN.xml` + sheet rels)                                                        | 1–2 days              | xlsx-js-style does not have an `addTable` equivalent. Phase 2 has to write the table XML directly into the JSZip-emitted output.                                                                                                                                                                                                                               |
| Rich-text export (`<is><r><rPr>` for multi-run cells)                                                            | 1 day                 | Inverse of the import walker; xlsx-js-style does not handle this.                                                                                                                                                                                                                                                                                              |
| Round-trip soak (every fixture → sjs export → re-import via both parsers, assert byte-or-structural equivalence) | 1–2 days              | Most likely surface area for surprises.                                                                                                                                                                                                                                                                                                                        |
| **Estimated total**                                                                                              | **9.5–14.5 days**     | A meaningful chunk of work, much of which is rebuilding what xlsx-js-style is supposed to provide.                                                                                                                                                                                                                                                             |

For comparison: the work to keep `exceljs` is **0 days**. exceljs's
quietness is not currently breaking anything; the moderate uuid CVE is
documented as not-reachable from our call sites.

---

## Recommendation: **NO-GO**

The spike's central finding is that `xlsx-js-style@1.2.0` does NOT carry
enough of M12's and M13/C's surface area to migrate without a meaningful
in-house port of the parts the fork is supposed to provide. Phase 2's
9.5–14.5-day scope is essentially "build what the styling fork is named
after, ourselves." That work is not free, and once it's done, the
maintenance burden moves from `exceljs` (a third-party library that gets
sporadic releases) to a Notesheet-internal styles.xml parser that we
have to keep current as Excel ships new OOXML constructs.

The dependency-hygiene ROI does not justify the swap. exceljs's
transitive CVEs are flagged but unreachable; its quiet-release cadence
is concerning but not blocking. A future cycle should revisit if either:

1. SheetJS Community proper (`xlsx`) adds first-class cell-style read
   support upstream (eliminating the need for the abandoned fork), OR
2. exceljs is publicly archived / repo-deleted (forcing the migration
   regardless of cost).

Until either condition is met, NO-GO.

---

## Conditional GO (alternate framing)

If the operator decides the dependency hygiene is worth the migration cost
and is willing to absorb the in-house parser work, the recommendation
becomes **CONDITIONAL GO with the following conditions**:

1. **In-house styles.xml walker.** Phase 2 ships a `readStylesXml` helper
   that parses `xl/styles.xml`'s `<cellXfs>` + `<fonts>` + `<fills>` +
   `<borders>` + `<numFmts>` directly via JSZip + regex (or `fast-xml-parser`).
   This replaces what xlsx-js-style fails to do for indexed-cellXf cells.
2. **In-house rich-text walker promoted to production.** The spike's
   `parseInlineRichRuns` (currently in `src/xlsxSheetJS.ts`) is moved to
   a shared helper module and tested against every multi-run fixture.
3. **Chart-export round-trip MUST be proven before Phase 2 merges.** Write
   `tests/charts/sheetjsRoundTrip.test.ts` that runs every fixture under
   `tests/fixtures/charts/` through
   `xlsxBufferToSnapshot` + `snapshotToXlsxBuffer` on the SheetJS path and
   re-imports via `src/xlsx.ts` to verify the post-processor's output
   stays parseable. If `injectChartsIntoZip` doesn't survive the structural
   change in SheetJS-emitted zips, Phase 2 forks the post-processor too.
4. **Golden snapshots stay green.** Every JSON file in
   `tests/golden-snapshots/` must continue to match what the new parser
   produces — every divergence requires explicit operator approval and a
   recorded note in the matrix. No symptom-patching the goldens to make
   the new parser look correct.
5. **Operator review of the in-house parser before Phase 2 ships.** Once
   the `readStylesXml` walker exists, the operator inspects what it does
   and what edge cases it punts on. If the punted edge cases overlap with
   shipped features, Phase 2 stops.

---

## Phase 2 feature-smoke test list (regression protection)

The 11 features on the operator's MUST-NOT-REGRESS list, with how Phase 2
will validate each:

| #   | Feature                                                       | Validation strategy                                                                                                                                                                                                                                                                                   |
| --- | ------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 1   | Snapshot creation via `New Spreadsheet` command               | Existing Jest test (`tests/snapshot.test.ts`'s `emptySnapshot` exercises) — Phase 2 runs unchanged.                                                                                                                                                                                                   |
| 2   | Snapshot editing in Univer (cell value changes round-trip)    | Existing Jest test (`tests/snapshot.test.ts:wrapSnapshot`/`extractSnapshot` round-trip pin-down) — Phase 2 runs unchanged.                                                                                                                                                                            |
| 3   | Univer toolbar formula bar                                    | Manual verification — open a fixture, click into a cell, confirm the formula bar shows the formula. No automated test today; Phase 2 should leave this as-is or add a PGE harness cycle.                                                                                                              |
| 4   | Univer formula evaluation (basic + structured refs)           | New Jest: load each fixture under `tests/fixtures/formatting-testdata/`, walk every cell with a formula, assert `cell.f` round-trips through the new parser. The `tests/golden-snapshots/FormulasAndStructuredRefs.json` is the regression baseline — if any formula text changes, golden test fails. |
| 5   | Named tables (insert / right-click row+col operations)        | Existing tests `tests/exportTableRoundTrip.test.ts` + `tests/m12FixtureRoundTrip.test.ts`. Phase 2 must keep these green.                                                                                                                                                                             |
| 6   | Chart insertion + live updates                                | Existing tests `tests/xlsxChart.test.ts` + the `tests/fixtures/charts/` fixtures' golden snapshots.                                                                                                                                                                                                   |
| 7   | Anchored chart drag/resize                                    | Manual verification on `08-drag-resized.xlsx`. PGE harness cycle is the right place if it becomes a regression hot-spot.                                                                                                                                                                              |
| 8   | .xlsx import via Tools menu                                   | Existing `tests/m12FixtureRoundTrip.test.ts` + the new `tests/xlsxParserParity.test.ts` — every fixture roundtrips through the new path.                                                                                                                                                              |
| 9   | .xlsx export via editor button                                | Existing `tests/exportTableRoundTrip.test.ts` + `tests/exportDebug.test.ts` — Phase 2 keeps these green.                                                                                                                                                                                              |
| 10  | Note navigation (open note, switch sheets, sidebar nav)       | Manual / PGE harness — out of parser scope.                                                                                                                                                                                                                                                           |
| 11  | Univer rendering pixel correctness (M13/E reference fidelity) | Existing `tests/excelReferenceFidelity.test.ts` + `tests/excelCanvasFidelity.test.ts`. Phase 2 must keep both green. The render is downstream of the snapshot, so a parser change that produces an equivalent snapshot leaves rendering untouched.                                                    |

For numbers 1, 2, 3, 7, 10, 11: **the parser is downstream of the
feature**. As long as the new parser produces a snapshot
deep-equal to the current parser's output (modulo documented divergences),
these features can't regress. The golden snapshot tests are the gate.

For numbers 4, 5, 6, 8, 9: **the parser is the feature**. The
parser-parity test, the round-trip tests, and the chart-export tests
together cover the surface area. Each must stay green for Phase 2 to ship.

---

## Spike artefacts inventory

| Artefact               | Path                             | Description                                                                                                    |
| ---------------------- | -------------------------------- | -------------------------------------------------------------------------------------------------------------- |
| Parallel parser module | `src/xlsxSheetJS.ts`             | Public surface mirrors `src/xlsx.ts`; dead code at runtime; exercised only by the parity test bed.             |
| Parser parity test     | `tests/xlsxParserParity.test.ts` | 24 fixtures × 14 dimensions matrix; writes `tests/golden-snapshots/parity-matrix.json`.                        |
| Golden snapshots       | `tests/golden-snapshots/*.json`  | 14 baseline goldens (one per fixture under `tests/fixtures/formatting-testdata/`) — Phase 2 regression target. |
| Decision document      | `docs/m14-sheetjs-spike.md`      | This file.                                                                                                     |

---

## Appendix A — Reproduction commands

```bash
# Install xlsx-js-style as devDep (already done):
npm install --save-dev xlsx-js-style@^1.2.0

# Run the spike's tests:
npx jest tests/xlsxParserParity.test.ts tests/goldenSnapshots.test.ts --verbose

# Inspect the parity matrix:
cat tests/golden-snapshots/parity-matrix.json | python3 -m json.tool | head -110
# (or any JSON viewer)

# Confirm baseline tests still pass:
npm test  # 209 baseline + 38 spike-added = 247 passed, 1 skipped

# Confirm production build is unaffected by the dep add:
npm run dist
ls -la publish/com.kamleshnanda.joplin-notesheet.jpl  # 13412864 bytes (same as main)
```

---

## Appendix B — exceljs failure modes preserved by spike's NotesheetImportError

These three fixtures throw `NotesheetImportError` on the exceljs path and
are recorded in their goldens as `{__importError: {name, code}}`:

| Fixture                          | code                           | exceljs cause                                                                                                                                                                                     |
| -------------------------------- | ------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `LargeWorkbook.xlsx`             | `xlsx-charts-unsupported`      | exceljs throws `Cannot read properties of undefined (reading 'anchors')` in `xl/xlsx/xlsx.js:100` during chart drawing reconcile                                                                  |
| `MultiSheet.xlsx`                | `xlsx-charts-unsupported`      | same anchors crash (one of the sheets has a chart)                                                                                                                                                |
| `FormulasAndStructuredRefs.xlsx` | `xlsx-multi-table-unsupported` | exceljs throws `Cannot read properties of undefined (reading 'name')` in `xl/doc/worksheet.js:920` during the table-reduce; symptomatic of multi-sheet workbooks each with their own named tables |

(`ConditionalFormatting-Variants.xlsx` imports cleanly on the exceljs side
— it appears on the parity matrix as a normal fixture; CF rules are
dropped on import as a documented existing shortcoming.)

`xlsx-js-style` does NOT throw on any of these. Phase 2 would either need
to replicate the typed-error contract via post-import validation OR get
operator approval to relax it (sjs is in fact more correct here — these
are valid workbooks Excel itself opens fine).
