# Operator ask — M13/D: rich-text within a single cell

Second real-feature cycle through the PGE harness. Workstream D from
the reverted M13 PR #16 — multi-run rich-text formatting inside one
cell — is the target. M13/C (rotated text, PR #19) is shipped.

## Why this matters

PR #16's commit `6f33f3a` added rich-text import/export plus 8 new
Jest tests (test count moved 199 → 207) and dropped the README
"Rich-text within a single cell" known-shortcoming entry. All Jest
tests passed. The PR was reverted because workstream C broke
visually — workstream D was collateral damage. The PGE harness now
exists to gate exactly this class of "Jest passes, Univer renders
broken." We restore D and prove it the same way C just did.

## The feature

When a user imports
`tests/ExcelBaseTestData/formatting-testdata/RichTextInOneCell.xlsx`
into a Notesheet note, the rich-text cells on the "RichText" sheet
must render in Univer with **per-run formatting visibly preserved**:

- **A1** "Hello world" — the word `Hello` renders **bolder** than
  ` world` in the same cell. Visibly heavier stroke, not flat plain
  text.
- **A2** "Red and Blue text" — `Red` renders **red**, ` and ` renders
  in the default text colour, `Blue` renders **blue**, ` text`
  renders in the default text colour. Three distinct colours visible
  in the same cell.
- **A3** "Visit example.com for more info" — single-format hyperlink
  cell. Renders as a hyperlink (Pattern A — `cell.p` with one
  uniform run, `customRange` for the link target). Does NOT regress
  to plain text and does NOT get downgraded to two-run rich-text.

Plus: exporting the same note back to `.xlsx` via the existing
export command must produce a workbook where exceljs reads
`cell.value = { richText: [{font, text}, ...] }` for A1 and A2 with
the per-run font preserved. A3 must round-trip as a plain
`{ text, hyperlink }` cell value (Pattern A wins for single-format
hyperlinks; rich-text path explicitly skips when a hyperlink
customRange is present).

## Acceptance criteria

The evaluator must verify ALL of:

1. **Visual — per-run formatting is visible in pixels.** Evaluator's
   Playwright-captured screenshot of the imported note shows:
   - **A1**: the leading `Hello` glyphs are visibly heavier (bolder
     stroke / wider glyph weight) than the trailing ` world` glyphs
     in the same cell. The evaluator must explicitly call out the
     weight contrast.
   - **A2**: the cell carries at least three distinct foreground
     colours in the visible glyphs — red on `Red`, default on ` and `,
     blue on `Blue`, default on ` text`. The evaluator must explicitly
     name the red and blue runs.
   - **A3**: renders as a single-format hyperlink cell — typically
     blue underlined text. NOT plain black text and NOT two-tone
     rich-text. (Pattern A regression sentinel.)
2. **Pixel sidecar — per-run colours land in the histogram.** The
   `<screenshot>.pixels.json` sidecar emitted alongside the
   evaluator's screenshot, sampled over the A1:A2 vertical band of
   the canvas, contains:
   - red ink (R≥200, G≤80, B≤80) with at least 30 hits — proves A2's
     `Red` run rendered.
   - blue ink (R≤80, G≤80, B≥200) with at least 30 hits — proves A2's
     `Blue` run rendered. (A3's hyperlink also contributes blue, so
     a single-row sample of A2 alone is preferable; the harness can
     pick either A2-only or A1+A2 region as long as the colour signal
     is present.)
   - The evaluator may use a region helper analogous to
     `rotatedRowRegion` from M13/C.
3. **Jest — reverted PR #16 D-tests are restored and green.** All of:
   - `tests/m13RichText.test.ts` is restored byte-equivalent to commit
     `6f33f3a` (8 tests: bold+plain, multi-color, hyperlink+plain
     stays Pattern A, mixed bold+italic+color, underline run isn't
     mistaken for hyperlink, single-run collapse, plain-text round-trip
     stays plain, full round-trip).
   - `tests/m12FixtureRoundTrip.test.ts`'s two `KNOWN SHORTCOMING —
     rich-text` tests are flipped to positive pin-downs (per
     `6f33f3a`'s diff: cells with rich text on import carry
     `cell.p.body.textRuns` with the right shape; export reproduces
     `richText: [{font, text}, ...]`).
   - The hyperlink+plain-in-one-cell test continues to assert
     **Pattern A** wins for that cell — the rich-text path must
     explicitly skip when a hyperlink customRange is present.
   - Total Jest passing count moves from the current baseline (187
     after M13/C) to at least 195 (M13/C +8 from the new
     `m13RichText.test.ts`). No `KNOWN SHORTCOMING — rich-text` test
     is left referencing M13.
4. **Round-trip — multi-run formatting survives export → re-import.**
   A Jest test (in the restored `m13RichText.test.ts` or
   `m12FixtureRoundTrip.test.ts`) imports `RichTextInOneCell.xlsx`,
   calls `snapshotToXlsxBuffer`, loads the buffer with a fresh
   exceljs `Workbook`, and asserts:
   - A1's value is `{ richText: [{font: {bold: true}, text: 'Hello'},
     {text: ' world'}] }` (or structurally equivalent — the bold
     attribute must survive on the first run, the second run must
     have no bold).
   - A2's value is `{ richText: [{font: {color: {argb: ...FF0000}},
     text: 'Red'}, {text: ' and '}, {font: {color: {argb:
     ...0000FF}}, text: 'Blue'}, {text: ' text'}] }` (or structurally
     equivalent — the red and blue colours must survive on the right
     runs).
   - A3's value is `{ text: 'Visit example.com for more info',
     hyperlink: 'https://example.com' }` (Pattern A — NOT a richText
     value).
5. **`npm run dist` succeeds and the .jpl is installed in the dev
   profile.** The generator captures its own screenshot via
   `scripts/pge/install-plugin.sh` + `eval-screenshot.js` as evidence
   under `screenshots/feature-1-m13-rich-text-renders/`, then opens
   that screenshot via the Read tool so the `verify-gate` hook
   unlocks the `test-results.json` write. The evaluator captures its
   OWN screenshot independently.

## Out of scope

- **Rich text combined with rotation.** OOXML rotation is cell-level
  only, not per-run, so there is nothing to combine. The fixture
  doesn't exercise it.
- **Rich text with size variations.** The fixture uses default size
  on every run; per-run font-size handling can fall out of the same
  pipeline but the visual gate doesn't grade for it.
- **Rich text inside merged cells.** The RichTextInOneCell fixture has
  no merges. If it incidentally works, fine; not a blocker.
- **README "Known shortcomings — Rich-text within a single cell"
  edit.** PR #16's commit drops it; restoring that drop is acceptable
  but the evaluator does not penalize if the docs edit is punted to a
  follow-up — the visual + Jest + round-trip evidence is the gate.
- **Conditional formatting rendered through rich-text.** Excel
  conditional formats can paint colours over runs; that's M16, not
  M13/D.
- **Stale single-run `cell.p` cleanup pass.** Do not retroactively
  rewrite prior snapshots. The import side simply must not emit
  `cell.p` for a single uniformly-styled run — covered by the
  single-run-collapse Jest case.

## Suggested fixture

`tests/ExcelBaseTestData/formatting-testdata/RichTextInOneCell.xlsx`,
sheet `RichText`:
- A1: `Hello` (bold) + ` world` (plain) — the bold pin-down.
- A2: `Red` (red) + ` and ` (default) + `Blue` (blue) + ` text`
  (default) — the multi-colour pin-down.
- A3: `Visit example.com for more info` with `hyperlink:
  https://example.com` — the Pattern A non-regression sentinel.

The harness invokes the headless import path
(`scripts/pge/import-fixture.sh RichTextInOneCell.xlsx`) and the
canvas-targeted screenshot via `eval-screenshot.js`. A region helper
covering A1+A2 (and optionally A2-only for cleaner colour signal) is
needed in `eval-screenshot.js`'s `REGION_BY_FEATURE` table.

## Related risks

- **The reverted PR #16's `src/xlsx.ts` changes are the right starting
  point.** Commit `6f33f3a` added 4 helpers (`buildTextStyleFromExceljsFont`,
  `buildRichTextCellP`, `buildExceljsFontFromTextStyle`,
  `extractRichTextRunsFromCellP`) and wired them into
  `extractCellValue` and `snapshotToXlsxBuffer`. The Jest tests
  passing tells you the data plumbing is sound — what's unverified is
  whether Univer's resolver actually paints per-run formatting from
  `cell.p.body.textRuns`. Don't rewrite the helpers; investigate the
  renderer first if visuals fail.

- **`cell.p` shape must be the documentSkeleton-finite shape.** The
  hyperlink path uses `buildHyperlinkCellP` with `pageSize`,
  paragraphs, sectionBreaks. The reverted commit's
  `buildRichTextCellP` reuses that shape — a malformed `cell.p`
  (missing pageSize, missing sectionBreaks) will Jest-pass on a shape
  match but crash Univer's layout pipeline on render or hover. If the
  visual gate shows nothing rendering or text disappearing, this is
  the prime suspect.

- **Pattern A precedence must hold for single-format hyperlinks.**
  The export side has two `cell.p` consumers — the hyperlink emitter
  and the rich-text emitter. They can't both fire. Per `6f33f3a`,
  rich-text export explicitly skips when a hyperlink customRange is
  present. If A3 round-trips as a 1-element richText (one run, one
  font, no hyperlink), that ordering is wrong and the test will fail.
  Confirm by inspecting the A3 round-trip assertion before flipping.

- **Per-run colours go through the theme+tint resolver.** exceljs
  surfaces `font.color` as `{argb}` or `{theme, tint}`. The reverted
  commit's `buildTextStyleFromExceljsFont` resolves both via the same
  resolver `extractCellValue` already uses for cell-level fonts. If
  A2's red/blue render as default-colour, check that the resolver is
  invoked on the run's font, not the parent cell's.

- **Univer `IDocumentBody.textRuns` is the snapshot shape.** The cell
  carries `cell.p.body.dataStream` (the concatenated text +
  paragraph terminator) and `cell.p.body.textRuns` (an array of
  `{st, ed, ts}` with `ts` being `ITextStyle`). Univer 0.23 reads
  `ts.bl` for bold (1=on), `ts.it` for italic, `ts.cl.rgb` for
  colour, `ts.un.s` for underline (1=on). If runs render flat, the
  bug is likely in how the `ts` object is constructed (wrong key
  names, value shapes off — `bl: true` instead of `bl: 1`).

- **Don't symptom-patch the test on a Jest failure post-rebuild.** If
  a Jest assertion regresses, run `git diff package-lock.json` first.
  exceljs's `richText`/`font` shapes drifted between 3.x and 4.x. We
  are intentionally on 4.4.0; do not edit the test to make a silent
  downgrade pass — fix the dependency drift instead. Reference:
  `feedback_dependency_hygiene.md` in operator memory.

- **The feature touches `src/xlsx.ts`, the same file as M9/M10/M12/M13/C.**
  Run the full Jest suite, not just the rich-text tests, before
  flipping the row. Borders, hyperlinks, table styles, chart export,
  alignment, and rotation all share the import/export pipeline. The
  m12 and m13/C round-trip tests must stay green.

- **Region-by-feature pixel sampling.** The eval-screenshot script
  needs a region helper for the A1:A2 (or A2-only) vertical band,
  plus entries in `REGION_BY_FEATURE` and `TITLE_PREFIX_BY_FEATURE`
  keyed off the new feature ID. M13/C's `rotatedRowRegion` is the
  template.
