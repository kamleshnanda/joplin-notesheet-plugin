# Notesheet PGE — build plan (M14: SheetJS Community migration spike — Phase 1)

> **Cycle context.** M13 is shipped end-to-end (M13/A through E +
> smoke + harness, see `PROGRESS.md` `## Done`). M14 is a **research
> spike**, not a feature in the usual user-observable sense. Phase 1
> (this PR) produces FOUR artefacts — a parallel-but-unused parser
> module, a parser-parity test bed, golden snapshots for Phase 2, and
> a decision document with a GO / NO-GO / CONDITIONAL recommendation.
> Phase 2 (the actual migration of `src/xlsx.ts` from `exceljs` to
> SheetJS / `xlsx-js-style`) is a SEPARATE PR and ships only if
> Phase 1's recommendation is GO and Phase 2's golden-snapshot tests
> stay green.
>
> **No PGE harness involvement.** The spike compares snapshot **data
> shape** between two parser implementations. There is no rendered
> canvas to screenshot — the parallel module is dead code at runtime,
> never imported by `src/index.ts` or `src/editorView.tsx`. The
> evaluator grades from committed files (`docs/m14-sheetjs-spike.md`,
> `tests/xlsxParserParity.test.ts`, `tests/golden-snapshots/*.json`,
> `src/xlsxSheetJS.ts`) plus `npm test` and `npm run dist` outputs,
> not from screenshots. The harness scripts under `scripts/pge/` keep
> working but are not exercised by this cycle.
>
> **Operator-flagged regression-protection list.** OPERATOR_ASK.md
> Acceptance criterion #5 enumerates the Notesheet features Phase 2
> MUST NOT regress: editor commands (`New Spreadsheet`), Univer
> rendering / formula bar / formula evaluation, named-table
> right-click ops, chart insertion + live updates + anchored chart
> drag/resize, .xlsx import (Tools menu) + export (editor button),
> note navigation. Phase 1 doesn't touch those code paths so they
> can't regress in this PR; the spike's job is to leave behind the
> test-bed Phase 2 will use to gate its merge. The
> "Feature-smoke test list" inside `docs/m14-sheetjs-spike.md` is
> where each of those gets a Phase-2 validation strategy.

## Operator ask

(See `OPERATOR_ASK.md`.) `exceljs@^4.4.0` is the entire foundation of
`src/xlsx.ts` — 1959 lines of import/export logic powering
`xlsxBufferToSnapshot`, `snapshotToXlsxBuffer`, `NotesheetImportError`,
the two resource-name constants
(`NOTESHEET_SYNTH_STYLES_RESOURCE`,
`NOTESHEET_THEME_CLR_SCHEME_RESOURCE`),
`readThemeClrScheme`, `readNamedHyperlinkCells`, and the
`synthesizeTableStyleAssignments` pipeline. exceljs has gone quiet
(last release Dec 2024) and ships stale transitives (uuid@8 with a
moderate CVE we tolerate as not-reachable; glob@7 deprecated). The
plan is to migrate to SheetJS Community via the styling-fork
`xlsx-js-style@1.2.0`. The spike proves whether that migration is
feasible without losing the formatting fidelity we shipped across M9,
M10, M12, and M13/A through E — or surfaces the gap that makes the
migration NO-GO.

The operator's directive is verbatim: *"I DO NOT WANT FEATURE
REGRESSION ON FEATURES THAT ARE ALREADY WORKING."* That is the bar
for Phase 2. Phase 1's bar is "produce the artefacts that let us
make a defensible GO / NO-GO decision."

## What the spike does NOT do

- **No production swap.** `src/xlsx.ts` stays as the production
  importer/exporter. `src/index.ts` and `src/editorView.tsx` stay
  unchanged. `src/xlsxSheetJS.ts` is dead code at runtime — only the
  parity test bed exercises it.
- **No `xlsx` direct dependency.** `package.json` `dependencies` stays
  unchanged. Only `devDependencies` adds `xlsx-js-style`. Phase 2
  (separate PR) is where `dependencies` gets swapped.
- **No README "Known shortcomings" edit.** Punt to Phase 2.
- **No Univer renderer follow-up** for M13/E's `bd.b` colour mismatch
  (rendered `#34692E` vs synthesized `#72D068`). Separate cycle.
- **No inter-banded-row strips for `TableStyleMedium4`.** Separate
  cycle.
- **No performance benchmarking** unless the spike surfaces a 10x+
  delta. Both libraries parse a 1 MB workbook in well under a second;
  the migration's value is dependency hygiene + longevity + better
  CF/style coverage, not speed.
- **No Univer rendering pixel parity.** That's a Phase-2 concern.
  The spike compares snapshot data shape only.

## Features

### feature-1-m14-sheetjs-spike-decision

**Spec**

When the operator reads this spike PR, they find FOUR committed
artefacts that together let them decide whether to start Phase 2
(the actual migration) or close it out as NO-GO:

1. **`src/xlsxSheetJS.ts`** — a parallel module implementing the same
   public surface as `src/xlsx.ts` (`xlsxBufferToSnapshot`,
   `snapshotToXlsxBuffer`, `NotesheetImportError`, the two resource
   constants), built on `xlsx-js-style@1.2.0`. It tries to handle
   every code-path category `src/xlsx.ts` handles. Where SheetJS
   can't carry a capability we depend on (e.g. theme palette
   extraction), it either implements a parser-agnostic workaround
   (raw-XML read via JSZip, the same shape `readThemeClrScheme` uses
   today) or stubs the path with a `TODO: blocked — see
   docs/m14-sheetjs-spike.md` marker. The module is NOT imported by
   `src/index.ts` or `src/editorView.tsx`; it is dead code at
   runtime.

2. **`tests/xlsxParserParity.test.ts`** — a Jest test that, for every
   `.xlsx` fixture under `tests/ExcelBaseTestData/formatting-testdata/`
   and `tests/ExcelBaseTestData/chart-testdata/`, runs both
   `src/xlsx.ts:xlsxBufferToSnapshot` and
   `src/xlsxSheetJS.ts:xlsxBufferToSnapshot` over the same input and
   produces a parity matrix. The matrix is committed (either as a
   JSON artefact under `tests/__snapshots__/` or printed in test
   output AND copy-pasted into the decision doc) and lists, per
   fixture, which fields match exactly, which diverge with a
   documented reason, and which fail outright. The test exits 0
   even when divergences exist — divergences become input to the
   decision doc, not test failures.

3. **`tests/golden-snapshots/`** — a directory of JSON files, one
   per fixture, capturing the current `src/xlsx.ts` output as the
   baseline Phase 2 must match. Each golden snapshot is asserted
   stable by a Jest test (`tests/goldenSnapshots.test.ts` or as a
   describe block in the parity test) — if Phase 2 changes how
   `src/xlsx.ts` emits any field, the assertion fails and the
   diff is auditable. Phase 2 will run the same goldens against
   `src/xlsxSheetJS.ts` and merge only when every divergence is
   either zero or documented + accepted.

4. **`docs/m14-sheetjs-spike.md`** — the decision document. Three
   sections:
   - **Capability matrix** — one row per .xlsx feature dimension
     Notesheet uses today, columns "exceljs (current)", "SheetJS
     (xlsx-js-style)", "gap notes". The dimensions to cover at
     minimum: workbook value read/write, cell styles (bg/fg/font),
     borders, merged cells, named tables + table styles, hyperlinks
     Pattern A (cell-level `hyperlink`), hyperlinks Pattern B
     (named-range emulation), rotated text, rich text per-run,
     theme palette `<a:clrScheme>`, conditional formatting rules,
     custom `<dxf>` records, drawings/charts, formulas + structured
     refs, defined names.
   - **Migration cost estimate** — for each gap, a 1-3-day estimate
     of "SheetJS API supports it, just port the code", "needs a
     workaround (raw-XML read like our existing `readThemeClrScheme`
     pattern)", or "blocked, no path".
   - **Recommendation** — explicit GO / NO-GO / CONDITIONAL with
     conditions named. Plus a "Phase 2 feature-smoke test list"
     enumerating every feature on the operator's
     MUST-NOT-REGRESS list (snapshot creation via `New Spreadsheet`
     command, snapshot editing in Univer, formula bar, formula
     evaluation incl. structured refs, named tables right-click
     ops, chart insertion + live updates, anchored chart drag /
     resize, .xlsx import via Tools menu, .xlsx export via editor
     button, note navigation, Univer rendering pixel correctness)
     with how Phase 2 will validate each (existing Jest, new Jest,
     PGE harness, manual).

The spike's user-observable outcome is "the operator now has enough
data to decide whether to start Phase 2." It is NOT user-visible to
Notesheet end-users; production behaviour is unchanged.

**Acceptance criteria**

The evaluator must verify ALL of the following from a fresh context.
Each criterion references either a committed file, a CLI exit, or a
text excerpt — not a screenshot:

1. **`src/xlsxSheetJS.ts` exists and exposes the full public
   surface.** The evaluator runs:
   ```
   grep -nE '^export ' src/xlsxSheetJS.ts
   ```
   and confirms presence of: `xlsxBufferToSnapshot` (async, takes
   `ArrayBuffer | Uint8Array | Buffer`, returns `Promise<UniverSnapshot>`),
   `snapshotToXlsxBuffer` (async, returns `Promise<ArrayBuffer>`),
   `class NotesheetImportError extends Error` (with the same
   `code` / `cause` fields the existing class carries — the
   evaluator cross-checks against `src/xlsx.ts:113`), and the two
   resource-name string constants
   `NOTESHEET_SYNTH_STYLES_RESOURCE = 'SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN'`
   and `NOTESHEET_THEME_CLR_SCHEME_RESOURCE = 'SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN'`
   (exact string values; pin-down sentinels at
   `tests/m12FixturePinDowns.test.ts` line ~226 and ~239 do an
   exact `r.name ===` lookup so the strings must match byte-for-byte).
   The evaluator additionally confirms the module is not imported
   by production code:
   ```
   grep -nE 'xlsxSheetJS' src/index.ts src/editorView.tsx
   ```
   returns no matches.

2. **`xlsx-js-style` is added as devDependency only, no production
   dep change.** The evaluator runs:
   ```
   git diff main -- package.json package-lock.json
   ```
   and confirms (a) `package.json` `devDependencies` contains
   `"xlsx-js-style": "^1.2.0"` (or the version the generator
   selected at spike time, called out in the decision doc),
   (b) `package.json` `dependencies` is byte-identical to main
   except for any explicit operator-approved changes, (c) the
   transitive deprecation noise from `npm install` is not
   noticeably worse than what `exceljs` already carries (i.e.
   `npm install` does not surface new high-severity advisories
   beyond the documented exceljs ones).

3. **`tests/xlsxParserParity.test.ts` runs and produces a matrix
   covering every fixture.** The evaluator runs:
   ```
   npx jest tests/xlsxParserParity.test.ts --verbose
   ```
   and confirms the test exits 0. The output (or a committed
   sidecar — generator's choice between `tests/__snapshots__/`,
   stdout, or a markdown table inlined into
   `docs/m14-sheetjs-spike.md`) lists at minimum one row per
   fixture under `tests/ExcelBaseTestData/formatting-testdata/`
   (15 fixtures) AND `tests/ExcelBaseTestData/chart-testdata/`
   (10 fixtures), 25 fixtures total, with each row indicating
   match / divergence / failure for the parser-pair comparison.
   For each divergence the matrix names the field and links to
   a row in the decision doc's capability matrix.

4. **`tests/golden-snapshots/` exists with one JSON per
   formatting-testdata fixture (15 files minimum).** The
   evaluator runs:
   ```
   ls tests/golden-snapshots/*.json | wc -l
   ```
   and confirms ≥ 15 files. (Chart fixtures are optional for
   golden coverage at the spike's discretion — the chart-export
   round-trip already has its own coverage in
   `tests/xlsxChart.test.ts` and goldens may double up; the
   generator decides.) A Jest test asserts `src/xlsx.ts:
   xlsxBufferToSnapshot` of each fixture matches its golden
   (`expect(snapshot).toEqual(golden)` or a per-field deep-equal
   loop). The evaluator runs:
   ```
   npx jest tests/golden-snapshots
   ```
   (or the named test file) and confirms exit 0.

5. **`docs/m14-sheetjs-spike.md` is committed and contains all
   three sections.** The evaluator opens the file and confirms:
   - A heading or table titled "Capability matrix" (or
     equivalent) with at least the dimensions enumerated in the
     spec (workbook value, cell styles, borders, merged cells,
     named tables, hyperlinks A, hyperlinks B, rotated text,
     rich text, theme palette, conditional formatting, dxf
     records, drawings/charts, formulas, defined names).
     Each row has a column for exceljs (current behaviour),
     SheetJS / xlsx-js-style (capability or gap), and notes.
   - A heading or table titled "Migration cost estimate" with
     a 1-3-day estimate per identified gap, plus a tally / sum.
   - A heading "Recommendation" containing the literal text
     "GO", "NO-GO", or "CONDITIONAL" — exactly one of the three —
     with the justification immediately following. If
     CONDITIONAL, the conditions are named explicitly (e.g.
     "GO if `xlsx-js-style` adds rich-text per-run support
     before Phase 2 lands").
   - A heading "Phase 2 feature-smoke test list" enumerating
     all 11 features from the operator's MUST-NOT-REGRESS list
     (see Spec point 4 above) with a one-line validation
     strategy per feature (existing Jest test name / new Jest
     to add / PGE harness `feature-N-...` cycle / manual
     verification with steps).
   - An honest paragraph addressing **`xlsx-js-style` maintenance
     state**. The fork last shipped on npm in 2022; if Phase 1
     proves the migration is technically feasible, Phase 2 trades
     one quiet library (`exceljs`, last 2024) for another
     (`xlsx-js-style`, last 2022). The decision doc surfaces
     this honestly — the recommendation accounts for it rather
     than papering over it.

6. **All 209 existing Jest tests stay green.** The evaluator runs:
   ```
   npm test
   ```
   and confirms `Tests: 209 passed` (plus however many new
   parity / golden tests the spike adds — those count as
   additional passes, not part of the 209 baseline). The 1
   skipped test in the M13/E baseline (`smoke.test.ts` —
   `xfail` placeholder) stays skipped. **No existing test
   assertion is altered to make a SheetJS-side symptom go
   away.** If a parity test reveals a divergence, the
   divergence is documented in the matrix and the spec assertion
   on the existing test stays unchanged. (Cf. operator memory
   `feedback_dependency_hygiene.md` — symptom-patching tests
   to make a dependency change "work" is the M11 anti-pattern;
   apply the same discipline here.)

7. **`npm run dist` builds clean.** The evaluator runs:
   ```
   npm run dist
   ```
   and confirms exit 0. The webpack TypeScript checker passes
   despite the new module — including any types `xlsx-js-style`
   pulls in. If `xlsx-js-style` lacks adequate `.d.ts` shipping,
   the spike adds a minimal `src/types/xlsx-js-style.d.ts`
   declaration file rather than disabling typecheck. The
   resulting `publish/com.kamleshnanda.joplin-notesheet.jpl`
   exists and is bytewise no larger than +50 KB versus main
   (the parallel module is dead code; webpack tree-shakes the
   xlsx-js-style import out of production bundle if production
   never references it — verify by extracting the .jpl and
   grep-ing for `xlsx-js-style` in `dist/index.js`; it should
   not appear, or if it appears the extra weight is justified
   in the decision doc).

**Out of scope**

- **Production swap of `src/xlsx.ts`.** Phase 2, separate PR.
- **Adding `xlsx-js-style` as a runtime dependency** (in
  `dependencies`). Phase 2.
- **Univer rendering pixel parity** between exceljs- and
  SheetJS-driven snapshots. The spike is data-shape parity only.
- **Performance benchmarking.** Out unless a 10x+ delta surfaces.
- **README edits** documenting the dependency change. Phase 2.
- **Hooking the parallel module into the harness so PGE can
  screenshot it.** The harness is not exercised this cycle.
  Adding screenshot wiring would conflate the spike with a
  feature ship, which is the wrong shape.
- **Resolving M13/E's `bd.b` Univer renderer-side colour
  mismatch.** Documented as a known gap (PROGRESS.md, Notes);
  unrelated to parser migration.
- **Inter-banded-row strip decoration on TableStyleMedium4.**
  Documented as a cross-feature follow-up (PROGRESS.md, Notes);
  unrelated to parser migration.

**Suggested fixtures**

Every existing `.xlsx` fixture is in scope — the spike's value comes
from running BOTH parsers across the WHOLE matrix, not cherry-picking.
Trip-points to highlight in the matrix:

- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard.xlsx`
  (Aptos M12+M13 sentinel — table styles, theme palette, banded
  rows, totals row, cell-level hand-edited overlays).
- `tests/ExcelBaseTestData/formatting-testdata/FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`
  (Classic theme — non-Aptos clrScheme, two `<conditionalFormatting>`
  blocks that are dropped on import today; spike must confirm
  `xlsx-js-style` doesn't inadvertently start preserving them in
  a way that diverges from current dropped-on-import behaviour).
- `tests/ExcelBaseTestData/formatting-testdata/RichTextInOneCell.xlsx`
  (M13/D — per-run formatting in a single cell, A1 = bold
  "**Hello**" + plain " world", A2 = three colours, A3 = blue
  underlined hyperlink. The biggest single open question for
  `xlsx-js-style` rich-text support).
- `tests/ExcelBaseTestData/formatting-testdata/MergedCellsAndAlignment.xlsx`
  (M13/C — rotated text at 45° / 90° / -45°. SheetJS's `alignment`
  shape is documented; spike must confirm angle round-trip parity).
- `tests/ExcelBaseTestData/formatting-testdata/Hyperlinks-Variants.xlsx`
  (M12 — Pattern A AND Pattern B hyperlinks. Pattern B is our
  raw-XML emulation via JSZip; spike must confirm parser-agnostic).
- `tests/ExcelBaseTestData/formatting-testdata/BordersAndCellColors.xlsx`
  (theme-tinted borders — `{theme: N, tint: T}` in the source
  XML; M12's `resolveExceljsColor` resolves these against
  whichever clrScheme is loaded).
- `tests/ExcelBaseTestData/formatting-testdata/FormulasAndStructuredRefs.xlsx`
  (M9 — structured references like `Table1[#Headers]`,
  `Table1[Column]`. Univer's formula engine must be fed these
  unchanged after the round-trip).
- `tests/ExcelBaseTestData/formatting-testdata/MultiSheet.xlsx`
  (multiple `xl/worksheets/sheetN.xml` files — sheet ordering,
  `sheetOrder` array shape).
- `tests/ExcelBaseTestData/formatting-testdata/NumberFormats.xlsx`
  (date format codes, accounting / currency, percent — exceljs
  preserves these; SheetJS handles them differently and the
  matrix must call out the divergence).
- `tests/ExcelBaseTestData/formatting-testdata/EmptyAndDegenerate.xlsx`
  (empty workbook, blank cells, single-sheet null state — failure
  modes are usually here, not in the rich fixtures).
- `tests/ExcelBaseTestData/formatting-testdata/MaliciousValues.xlsx`
  (`=cmd`, `\t`, leading-quote text, formula injection — the
  importer's safety guards live in `NotesheetImportError`; spike
  confirms `xlsx-js-style`'s error surface is at least as strict).
- `tests/ExcelBaseTestData/formatting-testdata/LargeWorkbook.xlsx`
  (size, cellRanges spanning thousands of rows; spike confirms
  no perf cliff worse than 10x).
- `tests/ExcelBaseTestData/formatting-testdata/ConditionalFormatting-Variants.xlsx`
  (CF rules — currently dropped on import; spike confirms
  `xlsx-js-style` parity).
- `tests/ExcelBaseTestData/formatting-testdata/border-isolation.xlsx`
  (M13/E rework #3 — operator-built fixture with explicit
  DOUBLE / THIN borders; spike must round-trip border-style
  enums correctly).
- All 10 chart-testdata fixtures (charts are imported as
  `xl/drawings/*.xml` + chart parts; spike must confirm
  `xlsx-js-style` either preserves charts unchanged on round-trip
  OR documents that charts pass through as opaque parts that
  Notesheet's chart-export post-processor in
  `src/charts/xlsxChart.ts` can still patch).

**Related risks**

- **`xlsx-js-style` may not carry our M12/M13 styling needs.**
  Plain `xlsx` (SheetJS Community) explicitly does not support
  cell styles. The fork claims to add basic styling on top, but
  "basic" is the open question. The spike must concretely test
  whether it can carry: theme palette round-trip (`<a:clrScheme>`
  preservation through import → snapshot → export → re-import),
  named-style banding (`TableStyleMedium4` synthesis with theme
  accents — the M13/E feature surface area), per-run rich text
  in a single cell (M13/D), rotated text angles (M13/C), Pattern
  A inline hyperlinks (M12), Pattern B named-range hyperlinks
  (M12 follow-up), theme-tinted borders (M12 outstanding short-
  coming, parser-agnostic via raw-XML read). If `xlsx-js-style`
  cannot carry rich text per-run OR theme palette extraction,
  the recommendation MUST be NO-GO or CONDITIONAL, not GO.

- **`xlsx-js-style` is a fork, last shipped 2022 per npm.** The
  fork's `xlsx@0.18.5` base is two years old at spike time. If
  the recommendation is GO, Phase 2 trades one quiet library
  (`exceljs`, last release 2024) for another (`xlsx-js-style`,
  last 2022). The decision doc must address this honestly under
  "Recommendation" — not bury it in a footnote, and not imply
  the migration is "obviously safer". A CONDITIONAL recommendation
  that says "GO if SheetJS Community itself adds styling
  upstream so we can drop the fork; otherwise NO-GO" is a
  legitimate outcome.

- **Round-trip equivalence is hard.** `xlsxBufferToSnapshot` then
  `snapshotToXlsxBuffer` then `xlsxBufferToSnapshot` should
  produce a snapshot equivalent to the first one — but cell IDs,
  style IDs, and resource ordering may differ between exceljs
  and SheetJS. The parity test must be tolerant of those benign
  differences (e.g. compare style records by deep-equal of the
  resolved style, not by string ID) while still catching real
  fidelity drift (e.g. a missing `bd.b` colour or a wrong
  `bg.rgb` value). The matrix must distinguish "style-ID
  remapped (benign, expected)" from "field value diverged
  (real drift, must be in capability matrix)".

- **The chart-export post-processor patches the zip exceljs
  writes.** `src/charts/xlsxChart.ts` (572 lines) post-processes
  the zip exceljs emits to splice in chart parts under
  `xl/charts/`, `xl/drawings/`, and the worksheet-relationship
  files. If `xlsx-js-style` writes a zip with a different
  internal structure (e.g. different order of zip entries,
  different `xl/_rels/workbook.xml.rels` shape, different
  empty-defaults), the patcher might break under SheetJS. The
  spike must include AT LEAST ONE chart-fixture round-trip in
  `tests/xlsxParserParity.test.ts` and document the patcher's
  compatibility status in the matrix. If incompatible, Phase 2
  must port the patcher too — flagged as a 1-3-day estimate
  in the Migration cost section.

- **Theme palette and Pattern B hyperlinks are parser-agnostic
  by design.** `readThemeClrScheme` (line 683 of `src/xlsx.ts`)
  and `readNamedHyperlinkCells` (line 744) both read raw XML via
  JSZip — they don't go through exceljs at all. The spike
  module should reuse these verbatim (extract them into a shared
  helper file or import them directly from `src/xlsx.ts`,
  generator's choice), confirming in the decision doc that
  these paths survive the migration unchanged. **If the spike
  finds itself reimplementing them, that's a smell — the
  parser-agnostic code shouldn't need a SheetJS variant.**

- **Don't symptom-patch the parity test.** If a parity test
  fails because SheetJS emits a different `cellXf` ID or a
  different style-record key order, do NOT change the test to
  accept the SheetJS output as "the new truth". Either:
  (a) find the equivalent SheetJS API that produces the
  exceljs shape, or (b) document the divergence in the
  capability matrix as a Phase-2 cost. Test parity is the
  whole point of the spike — relaxing it loses the spike's
  diagnostic value.

- **Don't rewrite all 1959 lines of `src/xlsx.ts` in the
  spike.** The spike's `src/xlsxSheetJS.ts` is allowed (and
  encouraged) to be incomplete. Aim for 100% capability
  coverage only on the operator's MUST-NOT-REGRESS list
  (snapshot creation, editing, formula bar, formulas, named
  tables, charts, anchored chart drag/resize, .xlsx import,
  .xlsx export, navigation, rendering — but rendering is
  Phase 2 anyway). For everything else (e.g. obscure CF rule
  variants, custom dxf records), a TODO marker with a pointer
  into the decision doc's "blocked, no path" row is the right
  shape. The spike's value is the matrix and recommendation,
  NOT a complete drop-in replacement.

- **Dependency drift discipline (operator memory
  `feedback_dependency_hygiene.md`).** After `npm install
  xlsx-js-style`, the generator MUST run `git diff package.json
  package-lock.json` and surface every version change in the
  PR description — not just the explicit add. Transitive
  resolution decisions are not implicitly approved. Downgrades
  are blocked by default. A typecheck error or test failure
  right after the install is hypothesis-1 "the dependency
  changed something I didn't expect", not "edit the test".

- **`tests/exportTableRoundTrip.test.ts` line 334 / 349 fix
  must stay.** Smoke session changed `'dashed'` to
  `'mediumDashed'` to satisfy exceljs 4's `BorderStyle` enum.
  Reverting it would re-break `npm run dist`. Unrelated to the
  spike but easy to accidentally undo if a generator runs
  `git restore` over the test directory.

- **The 1 currently-skipped test (`smoke.test.ts` — `xfail`
  placeholder)** stays skipped. The spike does not flip its
  state.

- **The new parity test must run inside the existing Jest
  config** (`jest.config.*` or the `package.json` `jest`
  block). Don't add a separate test runner. ts-jest handles
  the TypeScript imports already. `xlsx-js-style` must be
  resolvable via Node's standard `node_modules/` lookup — no
  webpack alias, no special loader.

- **The decision doc is markdown, committed to `docs/`.** The
  repo's only existing `docs/` content is the README's
  milestones table — `docs/` may not exist yet. The spike
  creates it. Use plain GitHub-flavoured markdown (no MDX, no
  diagrams that require a build step). Tables are markdown
  pipe-tables.

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file (and the prior M13 plans preserved in git history) as
the template. Each feature gets:

- `### feature-N-<kebab-id>` heading (stable; never renamed once
  written).
- **Spec**: one paragraph naming the user-observable change. For
  research-spike features (like M14), the "user" is the operator
  reading the PR; the observable is committed artefacts.
- **Acceptance criteria**: numbered list of observable evidence
  (visual screenshot, pixel-sidecar, Jest, CLI exit, file content,
  `grep` excerpt). NO data-shape-only assertions; NO "code does X"
  — only outcomes the evaluator can inspect from a fresh context.
- **Out of scope**: explicit non-goals so the evaluator does not
  penalise for them.
- **Suggested fixture(s)**: which file(s) under
  `tests/ExcelBaseTestData/` exercise this.
- **Related risks**: regression hot-spots, prior-bug pointers, and
  dependency-discipline reminders. Include pointers into prior
  PR / commit hashes when prior work is the starting point.

The lowest-numbered feature with `passes: false` in
`test-results.json` is what the generator works on next.
