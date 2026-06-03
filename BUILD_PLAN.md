# Notesheet PGE — build plan (M13/D: rich text within a single cell)

> **Cycle context.** This is the second real-feature cycle through the
> PGE harness. The first real cycle (M13/C — rotated text round-trip,
> PR #19) just shipped; its row sits in `PROGRESS.md` `## Done` and is
> intentionally absent from this `test-results.json` (the harness keeps
> only the in-flight cycle's rows). Harness scripts under
> `scripts/pge/` (launch-joplin.sh, install-plugin.sh,
> import-fixture.{ts,sh}, eval-screenshot.{js,sh},
> prep-joplin-window.sh, create-seeded-notesheet.js) are present and
> proven on M13/C. The pixel-sidecar (`<screenshot>.pixels.json`) and
> region-by-feature lookup (`REGION_BY_FEATURE` /
> `TITLE_PREFIX_BY_FEATURE` in `eval-screenshot.js`) are the
> machine-checkable signal alongside the visual screenshot.

## Operator ask

(See `OPERATOR_ASK.md`.) Restore Workstream D from the reverted M13
PR #16 — multi-run rich-text formatting inside one cell — and prove it
through the harness the same way M13/C just did. PR #16's commit
`6f33f3a` added the import/export plumbing plus 8 Jest tests
(test count moved 199 → 207 in that PR). All Jest tests passed when
PR #16 was open; the PR was reverted because workstream C broke
visually and workstream D was collateral damage. The PGE harness now
exists to gate exactly this class of "Jest passes, Univer renders
broken." We restore D and prove the per-run formatting renders in
real pixels, while keeping Pattern A hyperlinks intact.

## Features

### feature-1-m13-rich-text-renders

**Spec**

When the user imports
`tests/ExcelBaseTestData/formatting-testdata/RichTextInOneCell.xlsx`
into a Notesheet note (via the Tools menu **Import .xlsx as
Notesheet** command, or the headless harness equivalent
`scripts/pge/import-fixture.sh RichTextInOneCell.xlsx`), the resulting
note opens in the Univer editor with the "RichText" sheet's three
sentinel cells displaying **per-run formatting visibly preserved in
the rendered canvas**:

- **A1** "Hello world" — `Hello` is rendered visibly **bolder** (heavier
  glyph stroke / wider weight) than the trailing ` world` in the same
  cell. Not flat plain text.
- **A2** "Red and Blue text" — `Red` is **red**, ` and ` is the default
  text colour, `Blue` is **blue**, ` text` is the default text colour.
  Three distinct foreground colours in one cell.
- **A3** "Visit example.com for more info" — single-format hyperlink.
  Renders as a **hyperlink** (Pattern A — `cell.p` with one uniform
  run + a `customRange` link target). Does NOT regress to plain text
  and is NOT downgraded to a two-run rich-text shape.

Exporting the same note back to `.xlsx` via `snapshotToXlsxBuffer` (or
the editor's Export button) must round-trip the per-run formatting:
exceljs reads `cell.value = { richText: [{font, text}, ...] }` for A1
and A2 with the per-run font preserved; A3 round-trips as a Pattern A
`{ text, hyperlink }` plain-value cell — the rich-text emitter
explicitly skips when a hyperlink customRange is present.

**Acceptance criteria**

The evaluator must verify ALL of the following:

1. **Visual — per-run formatting is visible in pixels.** The
   evaluator's Playwright-captured screenshot of the imported note
   (via `scripts/pge/eval-screenshot.js` against the running Joplin
   dev profile, after `prep-joplin-window.sh` has filled the window
   and hidden the side panes) shows three sentinel cells whose
   per-run formatting is visually distinguishable:
   - **A1**: the leading `Hello` glyphs are **visibly heavier** (bolder
     stroke, wider weight) than the trailing ` world` glyphs in the
     same cell. The evaluator's verdict text must explicitly call out
     the weight contrast — "the cell is bold" or "the cell renders"
     is not sufficient.
   - **A2**: the cell carries at least three distinct foreground
     colours in the visible glyphs — **red** on `Red`, default on
     ` and `, **blue** on `Blue`, default on ` text`. The evaluator's
     verdict text must explicitly name the red and blue runs.
   - **A3**: renders as a **single-format hyperlink** — typically
     blue underlined text spanning the whole cell. NOT plain black
     text and NOT a two-tone rich-text shape. (Pattern A regression
     sentinel.)
2. **Pixel sidecar — per-run colours land in the histogram.** The
   `<screenshot>.pixels.json` sidecar emitted alongside the
   evaluator's screenshot, sampled over a region helper covering the
   A1+A2 vertical band of the Univer main canvas (and ideally an
   A2-only sub-region for cleaner colour signal — A3's hyperlink also
   contributes blue, which would alias the A2 `Blue` run if A1+A2+A3
   are sampled together), contains:
   - **Red ink** (R≥200, G≤80, B≤80) with at least 30 hits — proves
     A2's `Red` run rendered.
   - **Blue ink** (R≤80, G≤80, B≥200) with at least 30 hits — proves
     A2's `Blue` run rendered. (If the sample region is A1+A2 only,
     this signal is unambiguous; if A2 is sampled in isolation, even
     better.)
   - The harness adds a region helper analogous to
     `rotatedRowRegion` from M13/C (e.g. `richTextA1A2Region` and/or
     `richTextA2Region`) plus entries in `REGION_BY_FEATURE` and
     `TITLE_PREFIX_BY_FEATURE` keyed off this feature ID.
3. **Jest — reverted PR #16 D-tests are restored and green.** All of
   the following pass under `npm test`:
   - `tests/m13RichText.test.ts` is restored byte-equivalent to commit
     `6f33f3a` (8 tests: bold + plain run, multi-colour run,
     hyperlink + plain stays Pattern A, mixed bold+italic+colour
     run, underline run is NOT mistaken for a hyperlink, single-run
     collapse to plain string, plain-text round-trip stays plain,
     full rich-text round-trip).
   - `tests/m12FixtureRoundTrip.test.ts`'s two `KNOWN SHORTCOMING —
     rich-text` tests are flipped to **positive pin-downs** per
     `6f33f3a`'s diff: cells with rich text on import carry
     `cell.p.body.textRuns` with the right shape (per-run `ts`
     carrying bold/italic/colour as appropriate); export reproduces
     `cell.value = { richText: [{font, text}, ...] }` with the
     per-run font preserved.
   - The hyperlink+plain-in-one-cell assertion continues to assert
     **Pattern A wins** for that cell — the rich-text emitter must
     explicitly skip when a hyperlink customRange is present.
   - The total Jest passing count moves from the current baseline
     (187 after M13/C; verify via `npm test` at the start of the
     session) to **at least 195** (baseline + 8 from
     `m13RichText.test.ts`), with no `KNOWN SHORTCOMING — rich-text`
     test left referencing M13.
4. **Round-trip — multi-run formatting survives export → re-import.**
   A Jest test (in the restored `m13RichText.test.ts`, or as a
   restored block in `m12FixtureRoundTrip.test.ts`) imports
   `RichTextInOneCell.xlsx`, calls `snapshotToXlsxBuffer` on the
   resulting snapshot, loads the buffer with a fresh exceljs
   `Workbook`, and asserts:
   - **A1** value is `{ richText: [{ font: { bold: true }, text: 'Hello' },
     { text: ' world' }] }` (or structurally equivalent — `bold: true`
     must survive on the first run; the second run must NOT carry
     `bold`).
   - **A2** value is `{ richText: [{ font: { color: { argb: ...FF0000 } },
     text: 'Red' }, { text: ' and ' }, { font: { color: { argb:
     ...0000FF } }, text: 'Blue' }, { text: ' text' }] }` (or
     structurally equivalent — the red and blue colours must survive
     on the right runs; the two default-colour runs must NOT carry a
     colour).
   - **A3** value is `{ text: 'Visit example.com for more info',
     hyperlink: 'https://example.com' }` (Pattern A — NOT a richText
     value). This is the regression sentinel against the rich-text
     emitter winning for single-format hyperlinks.
5. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** The generator captures its own screenshot via
   `scripts/pge/install-plugin.sh` + `eval-screenshot.js` as evidence
   under `screenshots/feature-1-m13-rich-text-renders/`, then opens
   that screenshot via the Read tool so the `verify-gate` hook
   unlocks the `test-results.json` write. The evaluator captures its
   OWN screenshot independently — the generator's drop is
   corroborating, not authoritative.

**Out of scope**

- **Rich text combined with rotation.** OOXML rotation is cell-level
  only, not per-run; nothing to combine. The fixture doesn't exercise
  it.
- **Rich text with size variations.** The `RichTextInOneCell.xlsx`
  fixture uses default font size on every run. Per-run font-size
  handling can fall out of the same pipeline incidentally, but the
  visual gate doesn't grade for it. Don't add a size assertion.
- **Rich text inside merged cells.** The fixture has no merges. If
  rich-text incidentally renders correctly inside a merge, fine; if
  it doesn't, that's a separate follow-up, not a blocker.
- **README "Known shortcomings — Rich-text within a single cell"
  edit.** PR #16's `6f33f3a` drops that bullet (currently lines
  146–149 of `README.md`); restoring that drop is acceptable but the
  evaluator does NOT penalise if the docs edit is punted to a
  follow-up. Visual + Jest + round-trip evidence is the gate.
- **Conditional formatting rendered through rich-text.** Excel
  conditional formats can paint colours on top of runs — that's M16,
  not M13/D.
- **Stale single-run `cell.p` cleanup pass.** Do NOT retroactively
  rewrite prior snapshots to strip single-run `cell.p` entries. The
  import side simply must not emit `cell.p` for a single uniformly
  styled run going forward — covered by the single-run-collapse Jest
  case.
- **Per-run underline that is NOT a hyperlink.** The Jest test
  ensures an underline-only run is recognised as styling rather than
  a hyperlink; the visual gate does not grade an underline-only
  cell because the fixture doesn't carry one.

**Suggested fixture**

`tests/ExcelBaseTestData/formatting-testdata/RichTextInOneCell.xlsx`,
sheet `RichText`:

- **A1** — `Hello` (bold) + ` world` (plain). Bold-vs-plain pin-down.
- **A2** — `Red` (red) + ` and ` (default) + `Blue` (blue) + ` text`
  (default). Multi-colour pin-down with two default-colour runs to
  prove the resolver doesn't smear a colour across all runs.
- **A3** — `Visit example.com for more info` with `hyperlink:
  https://example.com`. Pattern A non-regression sentinel.

The harness invokes the headless import path
(`scripts/pge/import-fixture.sh RichTextInOneCell.xlsx`) and the
canvas-targeted screenshot via `eval-screenshot.js`. The screenshot
is taken against the Univer main canvas (selector
`canvas[id^="univer-sheet-main-canvas"]`) inside the editor frame,
NOT against the plugin sandbox page (the plugin sandbox is
`<body></body>` per the M13/C harness notes).

**Related risks**

- **The reverted PR #16's `src/xlsx.ts` changes are the right
  starting point.** Commit `6f33f3a` adds 4 helpers
  (`buildTextStyleFromExceljsFont`, `buildRichTextCellP`,
  `buildExceljsFontFromTextStyle`, `extractRichTextRunsFromCellP`)
  and wires them into `extractCellValue` and `snapshotToXlsxBuffer`.
  The Jest tests passing tells you the data plumbing is sound — what
  is unverified is whether **Univer's resolver actually paints
  per-run formatting from `cell.p.body.textRuns`**. Don't rewrite the
  helpers prematurely; investigate the renderer first if the visual
  gate fails. The lesson from M13/C is that the reverted code
  worked first try once the build and cache state were clean.

- **`cell.p` shape must be the documentSkeleton-finite shape.** The
  hyperlink path uses `buildHyperlinkCellP` with `pageSize`,
  paragraphs, and sectionBreaks. `6f33f3a`'s `buildRichTextCellP`
  reuses that shape — a malformed `cell.p` (missing `pageSize`,
  missing `sectionBreaks`) will Jest-pass on a textRuns shape match
  but **crash Univer's layout pipeline on render or hover**. If the
  visual gate shows nothing rendering, blank cells, or text
  disappearing on hover, this is the prime suspect.

- **Univer `IDocumentBody.textRuns` `ts` shape — exact key names
  matter.** Univer 0.23 reads runs as `{st, ed, ts}` with `ts` being
  `ITextStyle`. Bold is `ts.bl` (1 = on, NOT `true`); italic is
  `ts.it`; underline is `ts.un.s` (1 = on); colour is `ts.cl.rgb`
  (string, including `#`). If runs render flat in the visual gate
  but Jest passes, the bug is most likely in how `ts` is constructed
  by `buildTextStyleFromExceljsFont` — wrong key names or value
  shapes (`bl: true` instead of `bl: 1`, missing leading `#` on
  `cl.rgb`, etc).

- **Pattern A precedence must hold for single-format hyperlinks.**
  The export side now has TWO `cell.p` consumers — the hyperlink
  emitter and the rich-text emitter. They cannot both fire. Per
  `6f33f3a`, `extractRichTextRunsFromCellP` explicitly skips when a
  hyperlink customRange is present on `cell.p`. If A3 round-trips as
  a 1-element `richText` (one run, one font, no `hyperlink`), the
  ordering is wrong and the round-trip Jest assertion will fail.
  Confirm by inspecting the A3 round-trip assertion before flipping
  the row.

- **Per-run colours go through the theme+tint resolver.** exceljs
  surfaces `font.color` as either `{argb}` or `{theme, tint}`. The
  reverted commit's `buildTextStyleFromExceljsFont` resolves both via
  the same resolver `extractCellValue` already uses for cell-level
  fonts. If A2's `Red` and `Blue` render as default-colour, check
  that the resolver is invoked on **the run's font**, not the parent
  cell's font.

- **Univer style lookup is by `s` reference for cell-level styles —
  but rich-text runs are inline on `cell.p.body.textRuns`, NOT in
  `styles[id]`.** This is a different code path from rotation
  (M13/C) and the smoke (cell-level colour). Don't try to "fix"
  rich-text by extracting per-run styles into the `styles` map; runs
  carry their `ts` inline by design. The smoke / M13/C lesson
  ("style only renders when on `styles[id]` and the cell carries
  `s: <id>`") applies only to cell-level style — not to text-run
  style inside `cell.p`.

- **Region-by-feature pixel sampling.** Add a region helper covering
  A1:A2 (and optionally A2-only) to `eval-screenshot.js`, plus
  entries in `REGION_BY_FEATURE` and `TITLE_PREFIX_BY_FEATURE` keyed
  off `feature-1-m13-rich-text-renders`. M13/C's `rotatedRowRegion`
  is the template. The A2-only region produces the cleanest colour
  signal because A3's hyperlink also contributes blue.

- **Don't symptom-patch the test on a Jest failure post-rebuild.** If
  a Jest assertion regresses after restoring the file, run
  `git diff package-lock.json` first. exceljs's `richText` and `font`
  shapes drifted between 3.x and 4.x. We are intentionally on 4.4.0;
  do NOT edit the test to make a silent downgrade pass — fix the
  dependency drift instead. Reference:
  `feedback_dependency_hygiene.md` in operator memory.

- **The feature touches `src/xlsx.ts`, the same file as
  M9/M10/M12/M13/C.** Run the **full** Jest suite before flipping
  the row, not just the rich-text tests. Borders, hyperlinks
  (Pattern A and Pattern B), table-style synthesis, chart export,
  alignment, and rotation all share the import/export pipeline. The
  m12 fixture round-trip and m13/C rotation tests must stay green.

- **Window prep is mandatory before evaluator screenshots.**
  `scripts/pge/prep-joplin-window.sh` (added in M13/C) fills the
  Joplin window, hides the sidebar and note list, and closes
  DevTools. A small / panes-up window crops the Univer canvas
  horizontally, which can make A2 partially offscreen and produce
  false-negative pixel-sidecar readings. The harness wires this in
  ahead of `eval-screenshot.js` automatically; do not bypass it.

- **Pre-existing typecheck fix on `tests/exportTableRoundTrip.test.ts`
  must stay.** Smoke session changed `'dashed'` → `'mediumDashed'`
  on line 334 (with matching assertion on line 349) to satisfy
  exceljs 4's `BorderStyle` enum; reverting it would break
  `npm run dist`. Unrelated to rich-text but blocks the .jpl build.

## How the planner agent should fill out future BUILD_PLAN.md files

Use this file (and the prior M13/C BUILD_PLAN.md preserved in git
history at `dc80505`) as the template. Each feature gets:

- `### feature-N-<kebab-id>` heading (stable; never renamed once
  written)
- **Spec**: one paragraph naming the user-observable change
- **Acceptance criteria**: numbered list of observable evidence
  (visual, pixel-sidecar, Jest, runtime). NO data-shape-only
  assertions; NO "code does X" — only outcomes the evaluator can
  inspect from a fresh context.
- **Out of scope**: explicit non-goals so the evaluator does not
  penalise for them.
- **Suggested fixture(s)**: which file(s) under
  `tests/ExcelBaseTestData/formatting-testdata/` exercise this.
- **Related risks**: regression hot-spots and prior-bug pointers,
  including pointers into the reverted PR's commit hashes when the
  reverted code is the starting point.

The lowest-numbered feature with `passes: false` in
`test-results.json` is what the generator works on next.
