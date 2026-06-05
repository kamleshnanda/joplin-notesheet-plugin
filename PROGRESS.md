# Notesheet PGE — progress log

This file is the generator's session-to-session handoff. Update after
every feature.

## Done

- **feature-1-smoke-red-cell** (2026-06-02) — Generator inline-implemented
  the smoke seed in `src/snapshot.ts:emptySnapshot()`. A1 = "harness-smoke-OK"
  styled via `styles['pge-smoke-red'] = { cl: { rgb: '#FF0000' } }`,
  cell carries `s: 'pge-smoke-red'` reference. Built .jpl, installed in
  dev profile, captured generator-evidence screenshot showing red text
  rendered in Univer at A1. Fresh-context evaluator subprocess graded
  PASS (cd8bf51).
- **harness-hardening** (2026-06-02) — `eval-screenshot.js` now drops
  into the `UserWebviewIndex.html` frame inside the editor page (where
  Univer actually mounts) and waits on the real Univer canvas selector
  `canvas[id^="univer-sheet-main-canvas"]` instead of a 5s sleep.
  Emits a `<screenshot>.pixels.json` sidecar with the top non-background
  colours sampled from the row-0 canvas slab — gives evaluators a
  machine-checkable signal alongside the visual screenshot. Confirmed
  on the smoke note: dominant `rgb(234,237,249)` (header band),
  `rgb(255,0,0)` appears in top-3 with 353 hits, proving the red is
  real pixels not just snapshot data.
- **feature-1-m13-rotated-text-renders** (2026-06-03) — Cherry-picked
  the reverted PR #16 (`415b4a4`) rotation import/export in
  `src/xlsx.ts`, plus `tests/m13RotatedText.test.ts` and the
  `m12FixtureRoundTrip.test.ts` flips. README docs edit explicitly
  out-of-scope. Built .jpl, installed in dev profile, imported
  `MergedCellsAndAlignment.xlsx` headlessly via the new
  `scripts/pge/import-fixture.{ts,sh}`, captured canvas-targeted
  screenshot showing A6 up-right diagonal, B6 vertical, C6 down-right
  diagonal. Pixel sidecar over the rotated row band reports
  `inkRowSpread=1.000` (text ink occupies every sampled y-row in the
  slab) — strong non-horizontal signal independent of colour. Jest
  187/187 (was 181 baseline + 6 new rotation tests). The reverted
  code worked first try in Univer 0.23 — earlier visual failure was
  almost certainly a stale-build issue, not a rendering gap. Evaluator
  graded PASS (PR #19, dc80505).
- **feature-1-m13-rich-text-renders** (2026-06-03) — Cherry-picked PR
  #16 commit `6f33f3a` rich-text import/export in `src/xlsx.ts`
  (4 new helpers: `buildTextStyleFromExceljsFont`,
  `buildRichTextCellP`, `buildExceljsFontFromTextStyle`,
  `extractRichTextRunsFromCellP`), plus `tests/m13RichText.test.ts`
  (8 new tests) and the two flipped `m12FixtureRoundTrip.test.ts`
  pin-downs. README docs drop explicitly out-of-scope. Imported
  `RichTextInOneCell.xlsx` headlessly via
  `scripts/pge/import-fixture.sh`. Canvas screenshot shows
  A1 = bold "**Hello**" + plain " world", A2 = "Red"(red) + " and "
  + "Blue"(blue) + " text" with three distinct foreground colours,
  A3 = blue underlined "Visit example.com for more info" hyperlink
  (Pattern A). Pixel sidecar over the A2-only band reports
  `redInk=67`, `blueInk=76` — both well above the spec's ≥30
  threshold; dominant histogram bucket `rgb(0,0,255)` 76 hits
  followed by `rgb(255,0,0)` 67 hits. Jest 195/195 (was 187 baseline
  + 8 new rich-text tests). Reverted code worked first try in Univer
  0.23 — same hypothesis as M13/C confirmed: the original revert
  was almost certainly a build/cache issue, not a renderer gap.
  Evaluator graded PASS (PR #20). README cleanup followed in PR #21
  (5aaf690).
- **feature-1-m13-theme-aware-banding** (2026-06-03 → 2026-06-04,
  three-rework cycle) — Routed `synthesizeTableStyleAssignments`
  through a new `EXCEL_TABLE_STYLE_RECIPE_BY_NAME` parallel table
  (`src/charts/excelTableStyleRecipes.ts`) that names each TableStyle
  slot's accent index + tint. The synthesizer reads the source
  workbook's `<a:clrScheme>` (already captured by
  `readThemeClrScheme`) and resolves the recipe via the ECMA-376
  HSL-L tint formula. Aptos fixture (`FormattingSmorgasboard.xlsx`)
  paints green; Classic fixture
  (`FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`)
  paints grey instead of green — same `EXCEL_TABLE_STYLE_BY_NAME[Medium4]`
  lookup, two distinct rendered outputs driven by source clrScheme.
  Achromatic styles (Light1/8/15, Medium1/8/15/22, Dark1/8) keep
  their literal greys.

  **First multi-screenshot cycle.** Added `:variant` suffix support
  to `eval-screenshot.js` (`feature-1-m13-theme-aware-banding:aptos` +
  `:classic`), `tableHeaderRowRegion` helper sampling cols B+ to dodge
  A1 active-cell selection blue, `greyInk` aggregate (R,G,B∈[140,180]
  AND `abs(R-G)≤10` AND `abs(G-B)≤10`), and broadened the existing
  `greenInk` aggregate so it catches dark Aptos green `#196B24` (G=107)
  too.

  **Rework #2 (canvas fidelity, PR #22).** Operator caught a
  render-side bug the snapshot-fidelity test couldn't see: the
  `totalsTopBorder` slot was emitting `bd.t.s = 7` (DOUBLE) on the
  totals row. Univer 0.23's `_renderDoubleBorder` paints DOUBLE as
  two 1px strips with a 1-px white gap, which reads as anti-aliased
  `#89CE74` rather than the pure `#72D068` Excel paints. AND Excel's
  render is actually a single 2px strip, not a true double-line — so
  DOUBLE was the wrong style code in the first place. Phase 1 added
  `tests/excelCanvasFidelity.test.ts` (361 lines, pure-stdlib via
  `tests/util/pngSampler.ts`) — region-finding heuristics align Joplin
  canvas screenshot with Excel reference at structural regions
  (header / banded / totals-top), asserts dominant-colour parity for
  tall regions (Δ ≤ 8) and looser for 1-2px strips (Δ ≤ 32). Phase 2
  switched `BORDER_STYLE_TO_UNIVER.medium` (style 8, lineWidth=2)
  instead of `.double` (style 7) for the totals-top slot. Phase 3
  re-captured eval screenshots; Phase 4 raised counts 204 → 206.

  **Rework #3 (totals-row BOTTOM border).** Operator's eyeball +
  side-by-side pixel-probe revealed Excel paints TWO accent strips
  framing top AND bottom of the totals body. Recipe extended with
  `totalsBottomBorder` slot, parallel to `totalsTopBorder`. Empirical
  overrides set both Aptos (`#72D068`) and Classic (`#C9C9C9`) to
  the lighter accent. `synthesizeTableStyleAssignments` emits `bd.b`
  on every totals cell, REPLACING the table outline's thin frame.
  Diagnostic asset:
  `tests/ExcelBaseTestData/formatting-testdata/border-isolation.xlsx`
  (operator-built fixture with explicit border combinations,
  pixel-probed against Excel to establish ground truth).

  **KNOWN GAP — Univer renders `bd.b` with the wrong colour.** The
  synthesized snapshot is correct (`bd.b.cl.rgb === '#72D068'`
  verified via direct `xlsxBufferToSnapshot` introspection on every
  totals cell). When the .jpl is loaded in Joplin and pixel-probed,
  the rendered bottom strip at y=436 shows `rgb(52,106,46) = #34692E`
  — the header's dark green, NOT the lighter `#72D068` from `bd.b`.
  Top strip at y=398 renders correctly. The mismatch is in Univer's
  renderer, not in our synthesis. Filed as renderer-side follow-up;
  the synthesis change ships as it improves snapshot fidelity even
  before the renderer is fixed.

  Final test totals: 209 (was 197 pre-cycle). Three new pin-downs in
  `tests/m12FixturePinDowns.test.ts` (M13/E describe block), 6 new
  tests in `tests/excelReferenceFidelity.test.ts` (snapshot fidelity),
  4 new tests in `tests/excelCanvasFidelity.test.ts` (canvas
  fidelity, includes bottom-border tests added in rework #3), 4 leak
  pin-downs in `tests/m13RedoSmokeRedCell.test.ts` (smoke seed leak
  fix). Generator-evidence:
  `screenshots/feature-1-m13-theme-aware-banding/generator-evidence-{aptos,classic}.png`.
  Operator-captured Excel references:
  `screenshots/excel-reference/FormattingSmorgasboard-{Aptos,Classic}.png`.
  Evaluator graded PASS across all reworks. Shipped via PR #22 (commit
  `fca1cbc`); README cleanup PR #23 marked M13/E shipped.
- **M14 NO-GO** (2026-06-04) — `xlsx-js-style` (the SheetJS-style
  parser explored as an `exceljs` replacement) ALSO has `!cf`
  undefined on indexed-cellXf import. The "wait for SheetJS to make
  CF cheaper" rationale evaporates; M15 ships on `exceljs`. Decision
  document preserved at `docs/m14-sheetjs-spike.md` and merged via
  PR #25 (commit `f423d5a`). Roadmap renumbered: M15 is now the next
  active milestone (was already so before but the README marker was
  off-by-one).
- **feature-1-m15-conditional-formatting** (2026-06-05) — Full
  round-trip on the 5 CF rule types in
  `ConditionalFormatting-Variants.xlsx`: colorScale, dataBar,
  cellIs/highlightCell.number, top10/highlightCell.rank, iconSet.
  Phase 1 wired `UniverSheetsConditionalFormattingPreset` into
  `src/editorView.tsx` (the single-point-of-failure step — Phase 1
  build smoke confirmed canvas renders). Phase 2 added 6 fidelity
  tests in `excelReferenceFidelity.test.ts` anchored to the source
  XML (parsed via JSZip + regex), authored failing first per the
  fidelity-test-gap discipline; per-type translators in `src/xlsx.ts`
  (`translateExceljsCfRuleToUniver`) flipped them green. Phase 3
  extended `eval-screenshot.js` with `cfAllColumns` regionKind +
  `cfColumnRegion(col)` per-column samplers + new
  `pinkInk`/`lightGreenInk`/`yellowInk` aggregates plus broadened
  `redInk` (g/b ≤ 80 → ≤ 140 to catch #F8696B colorScale red end)
  and `blueInk` (b clearly dominant + ≥ 150, catches #638EC6 dataBar
  blue). Captured generator-evidence screenshot showing all 5 CF
  columns rendering correctly: A red→yellow→green gradient, C
  proportional blue bars, E pink fills on >50, G light-green on top-3
  cells, I red-down/yellow-flat/green-up arrows. Phase 4 added
  `excelCanvasFidelity.test.ts` describe block (1 test + 4 todos)
  gated `describe.skip` when the operator-captured Excel reference is
  absent at `screenshots/excel-reference/ConditionalFormatting-Variants.png`.
  Phase 5 added `translateUniverCfRuleToExceljs` + assigns
  `worksheet.conditionalFormattings` in `snapshotToXlsxBuffer`;
  flipped the KNOWN SHORTCOMING test at
  `tests/m12FixtureRoundTrip.test.ts:206` to a positive 'round-trip:
  5 conditional-formatting rules survive export → re-import' pin-down
  asserting all 5 source rules carry through structurally. Test
  total: 209 → 215 passed (6 new fidelity tests + 1 flipped, 6
  skipped including the canvas-fidelity todos). Generator-evidence:
  `screenshots/feature-1-m15-conditional-formatting/generator-evidence.png`
  (+ `.pixels.json` sidecar). Excel reference screenshot at
  `screenshots/excel-reference/ConditionalFormatting-Variants.png`
  still pending operator capture.

## In progress

(Empty — M15 done; awaiting evaluator verdict. Harness sampler bug
fix shipped 2026-06-05 — see `## Notes` entry "M15 harness sampler:
content-discovering geometry via FUniver".)

## Next

(Empty for now. M16 / M17 are post-M15 and will be specced after
M15 ships.)

## Notes

- **`emptySnapshot()` is the seam.** It's the single function that
  produces a fresh workbook for both the New Spreadsheet command
  (src/index.ts:103) and any "load empty fence" path. Putting the
  smoke seed here meant Jest tests could pin it down without mocking
  the Joplin Data API.
- **Univer's style resolver requires the colour on `styles[id]`,
  not inline on the cell.** The cell carries `s: <id>` (not
  `s: { cl: ... }`) and `cl.rgb` must include the leading `#`. We
  documented this in the BUILD_PLAN risks block; M13 lessons confirm
  it. **Important caveat for M13/D:** this rule is for cell-level
  style. Rich-text runs carry their `ts` inline on
  `cell.p.body.textRuns[i].ts` by design — they do NOT live in
  `styles[id]`. Don't conflate the two paths.
- **Evaluator needs a SEEDED note, not a `{}` note.** The plugin's
  `extractSnapshot('{}'...)` succeeds on `{}` (valid JSON, valid object)
  and the editor loads an empty Univer workbook — bypassing
  `emptySnapshot()` entirely. The harness needs a note whose body is
  `wrapSnapshot(emptySnapshot())`. We added
  `scripts/pge/create-seeded-notesheet.js` for that; it duplicates
  the seed shape, which is intentional coupling (the harness is
  allowed to know what the runtime does).
- **CDP page picker.** Joplin's CDP exposes 4-5 pages: a DevTools
  page, plugin sandboxes (one per loaded plugin including ours), and
  the editor. eval-screenshot.js scores by URL/title; editor wins
  unambiguously (16 vs 11/0/0). Plugin sandbox pages are
  `<body></body>` — they host plugin process logic, NOT the editor
  view. Don't try to attach to them.
- **Univer mounts in `UserWebviewIndex.html`** (a frame of the editor
  page), not the plugin sandbox. Stable selectors (Univer 0.23):
    - `canvas[id^="univer-sheet-main-canvas"]` — the main spreadsheet
      canvas. Id is `univer-sheet-main-canvas_<workbookId>`.
    - `[class*="univer-flex"]` — the toolbar wrapper (appears slightly
      before the canvas).
    - `#joplin-plugin-content` — the Joplin webview wrapper (always
      present once the plugin loads).
- **Pixel sidecar (`.pixels.json`)** — the harness samples the
  Univer main canvas's top-80px row-0 slab (or a feature-specific
  region) and writes a histogram of non-background colours alongside
  the screenshot. Use it for machine-checkable assertions like
  "top contains rgb(255,0,0) > 50 hits" instead of "I saw red."
  Sampling is stride-2 to keep cost cheap. Background (>235 in all
  channels) and gridline ink (<30) are filtered out.
- **Pre-existing `tests/exportTableRoundTrip.test.ts:334` typecheck
  bug** — `'dashed'` is not in exceljs `BorderStyle` enum. Changed
  to `'mediumDashed'` (with matching assertion update on line 349)
  during the smoke session because webpack's TS check blocks .jpl
  build. This was unrelated to smoke; the smoke didn't introduce it.
  Reverting it would re-break `npm run dist` for any subsequent
  cycle.
- **`scripts/pge/import-fixture.{ts,sh}`** — headless equivalent of
  the plugin's "Import .xlsx as Notesheet" command. The .ts calls
  `xlsxBufferToSnapshot()` from `src/xlsx.ts` directly so the harness
  exercises the SAME conversion the plugin runs at runtime; the .sh
  wrapper compiles via `node_modules/.bin/tsc` to a temp dir and runs
  the JS through Node with `NODE_PATH` pointing at the repo's
  `node_modules`. `PGE_REPO_ROOT` env var is required because
  `__dirname` after compile lives inside the temp dir.
- **Joplin window pane crops the Univer canvas.** `page.screenshot()`
  with `fullPage:false` captures the whole Joplin window — but at
  Joplin's default pane sizes the editor pane is narrower than the
  Univer canvas, so the canvas is cropped or partially offscreen.
  Fixed in `eval-screenshot.js` by screenshotting the canvas element
  directly via `webview.locator(canvasSel).screenshot()` when a
  Notesheet note is opened. The whole-page screenshot is the
  fallback for smoke / verification mode.
- **`inkRowSpread` metric.** Added to the pixel sidecar:
  `inkRows / ceil(regionHeight/2)` — the fraction of sampled y-rows
  in the region that carry text-coloured ink. Horizontal text
  concentrates ink on a narrow band (low spread); rotated/stacked
  text spreads across the band (high spread, near 1.0). Use this
  alongside the colour histogram for rotation-style features. For
  M13/D, the relevant metric is per-colour histogram hits within the
  A1+A2 (or A2-only) region — `inkRowSpread` is not the right signal
  for rich-text.
- **`feature-1-m13-rotated-text-renders` worked first try in Univer
  0.23.** PR #16's `style.tr = { a: <angle> }` (and `{ a: 0, v: 1 }`
  for stacked) maps directly to Univer's `ITextRotation` and the
  resolver honours it without extra plugin registration. The earlier
  visual failure that prompted the revert was almost certainly a
  stale-build / cache-not-wiped issue, not a renderer-vs-shape
  mismatch. The harness's `install-plugin.sh` cache wipe + Joplin-
  quit gate is what makes the difference. **Apply the same
  hypothesis to M13/D first** — restore `6f33f3a` verbatim, do a
  full rebuild + cache wipe + Joplin re-launch via the harness
  scripts, then judge from real pixels before suspecting the
  helpers.
- **Region-by-feature pixel sampling.** `eval-screenshot.js` now
  consults `REGION_BY_FEATURE` and `TITLE_PREFIX_BY_FEATURE` tables
  to pick the right canvas region and note title prefix per feature.
  When adding a new feature whose evidence isn't on row 0, add a
  region helper (like `rotatedRowRegion`) and an entry in both
  tables. M13/D needs an A1+A2 region (and ideally an A2-only
  variant for cleaner colour signal — A3's hyperlink also
  contributes blue ink).
- **Window prep is mandatory before evaluator screenshots.** Three
  pre-conditions had to be enforced after operator hit them:
    1. Joplin window must fill its display (not 800×568 or whatever
       the operator left it at) — Univer sizes its canvas to the
       editor pane, so a small window means fewer columns rendered
       and the visual gate may miss content.
    2. Sidebar (`.rli-sideBar`) and note list (`.note-list`) must be
       hidden — they consume ~500-700px of horizontal real estate
       the editor could otherwise use.
    3. Any DevTools window must be closed — it shrinks the renderer
       AND can confuse `eval-screenshot.js`'s CDP page picker.
  `scripts/pge/prep-joplin-window.sh` handles all three. AppleScript
  invokes `Window > Fill` and sends `Cmd+Alt+S` / `Cmd+Alt+L` for
  pane toggles (Playwright's `keyboard.press()` does NOT reach
  Joplin's Electron-accelerator-routed shortcuts — the renderer
  receives the keydown but the command never fires; OS-level
  System Events keystroke does work). State-aware: queries renderer
  DOM via `element.offsetParent !== null` (canonical for "any
  ancestor has display:none") and only toggles when needed. Wired
  into `eval-screenshot.sh` ahead of the screenshot. Requires
  Accessibility permission for the terminal app in System Settings.
- **Pattern A hyperlink emitter has precedence over the rich-text
  emitter (M13/D pre-context).** `src/xlsx.ts` already emits Pattern
  A hyperlinks via `buildHyperlinkCellP` for single-format hyperlink
  cells (M12 work). `6f33f3a`'s rich-text export path adds a second
  consumer of `cell.p` — `extractRichTextRunsFromCellP` — which MUST
  explicitly skip when a hyperlink customRange is present on the
  cell, otherwise A3 of `RichTextInOneCell.xlsx` regresses to a
  1-element richText value with no hyperlink. The Jest test
  `hyperlink + plain stays Pattern A` is the regression sentinel;
  the round-trip assertion on A3 (`{text, hyperlink}` not richText)
  is the user-visible sentinel.
- **Pixel histogram fragments under anti-aliasing — added per-spec
  band aggregates to `samplePixelsAt`.** During M13/D, even with
  clean A2-text rendering, the `top` histogram showed `rgb(255,0,0)`
  at only 17 hits and `rgb(0,0,255)` at 12 hits when sampled at
  stride-2. Anti-aliased glyph edges produce many near-but-not-exact
  RGB tuples, fragmenting the per-bucket counts even when the
  feature is rendering fine. We added stride-1 sampling plus three
  inequality-band aggregates that match the spec verbatim:
  `redInk` (R≥200, G≤80, B≤80), `blueInk` (R≤80, G≤80, B≥200),
  `greenInk` (R≤80, G≥150, B≤80). The aggregates are robust to
  anti-aliasing because they sum every pixel that satisfies the
  inequality, regardless of exact RGB. M13/D's final reading was
  `redInk=67`, `blueInk=76` — well above the ≥30 threshold. Future
  per-run colour features should use these aggregates rather than
  the `top` histogram for gating. **For M13/E**, add a `greyInk`
  aggregate using the same template — recommended thresholds
  `R∈[140,180]`, `G∈[140,180]`, `B∈[140,180]` with `abs(R−G) ≤ 10`
  AND `abs(G−B) ≤ 10` (the equality between channels is what
  distinguishes grey from a tinted hue at similar luminance).
  **For M15**, add `pinkInk` (`R≥220 AND G∈[180,220] AND
  B∈[180,220]`), `lightGreenInk` (`R∈[180,220] AND G≥220 AND
  B∈[180,220]`), and `yellowInk` (`R≥200 AND G≥200 AND B≤120`)
  using the same template. The CF columns each have a different
  signature colour band; the gates are per-column.
- **The Univer cell-selection blue border at `rgb(44,83,241)` will
  saturate any region that includes A1 + cell-border y-band.** When
  a cell is the active selection (typical state on a freshly opened
  note), Univer paints a ~1px blue border around it. That border
  contributes ~150+ blue pixels to a slab covering the cell, even
  when the cell text itself is pure black. For colour-band
  features, sample a y-region that **excludes the active-cell
  border** — empirically on the M13/D fixture the borders sat at
  y≈20–21 (top of A1) and y≈38–40 (bottom of A1 / top of A2). We
  put the A2-only region at y=41–58 to land inside A2's text band
  cleanly. If a future feature needs an A1 colour gate (e.g. a red
  font on A1), either move selection off A1 first or carve the
  border-y rows out of the region. **For M13/E** the table header
  row is row 0 of the table data area = row 1 of the visible
  worksheet — the active-cell selection is on A1 by default and
  lands inside the header band. Either click off A1 before
  sampling (programmatic Univer click via the webview frame's
  command bus) OR sample a header column that is NOT col A
  (operator suggests a mid-table column like the Spent / Discount
  column where the green-vs-grey signal is unambiguous).
- **Rich-text rendering worked first try in Univer 0.23 once the
  build/cache state was clean.** Same lesson as M13/C: cherry-pick
  `6f33f3a` verbatim, full `npm run dist` + `install-plugin.sh`
  (cache-wipe gated on Joplin being quit) + relaunch + headless
  import via `import-fixture.sh` — the per-run formatting renders
  correctly without any helper rewrite. The original PR #16 revert
  was almost certainly a stale-build artefact, not a renderer gap.
  The hypothesis from M13/C generalises.
- **M13/E is the first multi-screenshot cycle.** Until M13/E every
  cycle had exactly one note → one screenshot → one row in
  `test-results.json`. M13/E needs TWO independent screenshots
  (Aptos fixture + Classic fixture) because the failure mode is
  "same catalog entry must produce DIFFERENT colours under
  different source clrSchemes." The plan picks the
  **variant-suffix approach** for the harness extension: feature
  ID becomes `feature-1-m13-theme-aware-banding:aptos` and
  `feature-1-m13-theme-aware-banding:classic`, and
  `eval-screenshot.js` looks up the suffixed key in
  `REGION_BY_FEATURE` / `TITLE_PREFIX_BY_FEATURE` (falling back to
  the plain key for prior-cycle compatibility). The generator MUST
  add the suffixed entries WITHOUT renaming or removing the
  M13/C and M13/D entries — those are evidence-bearing for the
  prior-cycle rows already in `## Done`. Document the suffix
  convention here once the M13/E session lands so the next planner
  knows it exists.
- **`tableHeaderRowRegion` y-band hint.** The header row of the
  imported `ProjectTracker` / `ProductCatalog` table sits
  immediately below Univer's column header (~y=0–18 at default
  zoom). At default row height (19px) the header band is roughly
  y=19–37. Use a slab y=18–40 (h≈22) to absorb default row-height
  variation. The new region helper goes in
  `scripts/pge/eval-screenshot.js` next to `rotatedRowRegion` and
  `richTextA1A2Region`.
- **Header colour gate column choice.** The top-left active-cell
  selection (`rgb(44,83,241)` border) on A1 will pollute a header-
  row sample that includes col A. Two valid mitigations: (1)
  programmatically click off A1 to dismiss the selection before
  sampling, or (2) restrict `tableHeaderRowRegion` to start
  several columns in (e.g. starting at `x=2*colWidth`). The
  generator picks one and documents it. **M13/E chose option 2**:
  `tableHeaderRowRegion` starts at `x=80` (col B onward, col width
  ~73px), excluding A1's active-cell selection border entirely.
  The selection-blue still shows up at low counts (~136 hits of
  `rgb(150,169,248)` in the histogram — bleed from the cell-border
  anti-aliasing along the right edge of A1 still visible at x≈74)
  but doesn't trigger redInk/blueInk/greenInk/greyInk gates.
- **Variant-suffix harness extension.** `eval-screenshot.js` now
  splits `FEATURE_ID` on `:` into `BASE_FEATURE_ID` + `VARIANT`.
  Lookups try the suffixed key first, then the plain base key —
  prior single-screenshot features keep working unchanged. Output
  filenames bake the variant: `eval-aptos-...png` /
  `eval-classic-...png`. Both screenshots land in the same
  `screenshots/<base-feature-id>/` directory. The `verify-gate`
  hook reads the directory; both PNGs must be Read before the
  test-results.json flip is allowed. **For future multi-variant
  cycles**: add a `:<variant>` row in BOTH `TITLE_PREFIX_BY_FEATURE`
  AND `REGION_BY_FEATURE` keyed off the suffixed id, run
  `eval-screenshot.sh feature-id:variant` once per variant. Note
  list and import order do NOT matter — the title-prefix lookup
  picks the latest match by `updated_time`.
- **Theme-aware synth approach: parallel recipe table.**
  `src/charts/excelTableStyleRecipes.ts` is a parallel mirror of
  `excelTableStyles.ts` that names each slot's accent index +
  HSL-L tint amount. At synthesis time
  `resolveTableStylePalette(styleName, catalog, themeRgb)` looks
  up the recipe and resolves each slot via `tintRgb(accent, tint)`
  using the source workbook's clrScheme accent values. The
  Aptos baseline catalog is preserved unchanged (legacy fallback
  when `themeRgb` is null) and verified exact against the new
  HSL-L formula by direct comparison: `tint('#196B24', 0.6)` =
  `'#84E291'` (catalog Medium4 even-row), within zero RGB units.
  Achromatic styles use `accent: null, rgb: '#...'`. **Don't add
  hardcoded Classic-Medium4 entries** — Excel ships ~thousands of
  accent permutations and the catalog can't enumerate them; the
  recipe table is the source of truth.
- **In-process ExcelJS test workbooks ship the Office 2007 default
  theme.** When tests build a workbook via
  `new ExcelJS.Workbook(); ws.addTable({...style: 'TableStyleMedium2'...})`
  and then round-trip it through xlsxBufferToSnapshot, the
  resulting snapshot's per-cell `bg` is `#4F81BD` (exceljs
  accent1) NOT `#156082` (Aptos accent1). M13/E updated three
  pin-downs in `tests/formattingFidelity.test.ts` and one in
  `tests/m12FixtureRoundTrip.test.ts` to assert the actual
  exceljs-default colour. The Aptos and Classic project-owned
  fixtures pin the Aptos and Classic palettes specifically (see
  the M13/E describe block in `m12FixturePinDowns.test.ts`).
- **`greenInk` aggregate threshold relaxed for M13/E.** The M13/D
  threshold (`R≤80 AND G≥150 AND B≤80`) only matches pure greens
  like `rgb(0,255,0)` — Aptos accent3 `#196B24` (R=25,G=107,B=36)
  has G < 150 and fails. Relaxed to `G > R+30 AND G > B+30 AND
  G ≥ 80` so any green-channel-dominant pixel ≥ 80 luminance
  qualifies. Catches Aptos accent3 (107 vs 25/36), Aptos pastel
  banded `#84E291` (132,226,145), AND pure `rgb(0,255,0)`.
  redInk/blueInk thresholds stay tight (M13/D's gate values
  unchanged). The post-Phase-3 measured Aptos header `#34692E`
  (52,105,46) also qualifies (`g=105 > r+30=82 && g > b+30=76 &&
  g >= 80`).
- **PR #17 smoke seed leak — production code carried debug state
  for ~6 weeks before catch.** `emptySnapshot()` was the simplest
  seam to put the smoke seed in (one function, pure-testable via
  Jest, no Joplin Data API mock needed). It was the wrong place
  long-term: every "New Spreadsheet" command in the deployed plugin
  shipped with cell A1 pre-filled. Lesson: **harness-only state
  belongs in `scripts/pge/`, not in `src/`.** The PGE
  `create-seeded-notesheet.js` had ALWAYS inlined the seed shape
  itself — the production seed in `emptySnapshot()` was redundant.
  Fix landed Phase 1 of M13/E rework session.
- **Why "asserting our own emit" let PR #22 ship wrong colours.**
  The original M13/E pin-downs in `m12FixturePinDowns.test.ts`
  asserted `bg.rgb === '#196B24'` for the Aptos header — that's
  what our code emitted (raw accent3), NOT what Excel actually
  paints (`#34692E`). The acceptance criterion "dark green" passed
  trivially because both `#196B24` and `#34692E` ARE dark greens.
  The pixel-sidecar threshold `greenInk ≥ 30` passed for both too.
  Anchoring tests to your own behaviour gives you confidence your
  code is consistent, NOT that it's correct. Phase 2's
  `excelReferenceFidelity.test.ts` is the corrective: it samples
  the operator-captured `screenshots/excel-reference/*.png` and
  asserts the synthesizer's output matches what Excel actually
  paints, within Δ ≤ 8 per channel. **Future "match Excel" features
  should establish a reference-anchored gate FIRST (Phase 2 before
  Phase 3) so the recipe lands against ground truth, not against
  hand-derivation.**
- **Excel TableStyle algorithm not crackable from a single accent.**
  Phase 3 investigation tested HSL-L tint, HSV scale, satMod+lumMod,
  RGB mix toward grey, lumMod+lumOff in HSL space — none reproduce
  all four target RGBs (Aptos header #34692E, Aptos band #CAEFCB,
  Classic band #EDEDED, Aptos totals top #72D068) from a single
  accent. Excel's built-in TableStyle definitions live in Office's
  installed assets (NOT in the workbook's `xl/styles.xml`) and the
  actual transformation isn't documented in OOXML. Operator's
  allowed empirical-lookup path was the right escape: ship measured
  RGBs in `EXCEL_TABLE_STYLE_EMPIRICAL_OVERRIDES`, keyed by
  `(styleName, accentHex)`. The HSL-L formula remains as fallback
  for accents we haven't measured. **Adding a third fixture means
  adding a third entry to the override map** — measure with
  `tests/util/pngSampler.ts:dominantColor` against the new Excel
  reference PNG.
- **Recipe shape gained `totalsTopBorder`.** PR #22's shape only
  modelled `borderColor` (table outline). Excel paints a separate
  border above the totals row in a lighter shade of the accent —
  `#72D068` for Aptos accent3, `#C9C9C9` for Classic accent3. Added
  to the recipe interface; consumed by
  `synthesizeTableStyleAssignments` which emits the totals row's top
  side in that colour. Falls back to the existing thin border when
  the slot is missing. **Style code is MEDIUM (8), not DOUBLE (7) —
  see "DOUBLE on top renders as anti-aliased two-strip" note below.**
  M13/E pin-down `Aptos fixture: totals-row top border carries the
  Excel separator (#72D068 green, MEDIUM)` is the sentinel.

- **Univer 0.23 DOUBLE-on-top renders as anti-aliased two-strip,
  NOT a true double-line.** `getLineWidth(BorderStyleTypes.DOUBLE)=1`
  in `node_modules/@univerjs/engine-render/lib/es/index.js:1872`,
  and `_renderDoubleBorder` (line 5102) places inner+outer strokes
  at `±lineWidth/2 = ±0.5px` from centre. Net result: two 1px strips
  separated by 1px of white, which reads as `#89CE74` (anti-aliased
  green) rather than the pure target colour. **Excel doesn't even
  paint a true double-line for `TableStyleMedium4` totals-top — pixel
  sampling of the operator-captured reference shows a single 2px
  strip in the lighter accent**, identical to the strips Excel paints
  at every banded-row boundary. So MEDIUM (style 8, lineWidth=2) is
  both: (a) closer to Excel's actual render, and (b) more visually
  reliable than DOUBLE in Univer 0.23. The visual "double-line"
  perception of the totals-top in Excel comes from the strip pairing
  with the banded-row decoration above it — see the cross-feature
  follow-up below.

- **Cross-feature follow-up: banded-row boundary decoration.**
  Excel paints `#72D068` (Aptos) / `#C9C9C9` (Classic) 2px strips at
  EVERY banded-row boundary inside `TableStyleMedium4`, not just the
  totals-top. Notesheet currently relies on Univer's default
  `#D7D8DB` thin gridline for inter-row separation, so our render
  shows the right cell-fill colours but the wrong inter-row borders
  — only the totals-top gets the accent-shade decoration. Out of
  scope for this rework; should be its own feature in BUILD_PLAN.md.
  The synthesizer would need to emit MEDIUM borders on every banded-
  data-row boundary (top of rows 2-9 in the Aptos/Classic fixtures)
  in `palette.totalsTopBorder` colour.

- **Canvas-vs-Excel fidelity test layer (`tests/excelCanvasFidelity.test.ts`).**
  Second gate added in M13/E rework #2. The original
  `excelReferenceFidelity.test.ts` compares Excel reference PNG →
  `xlsxBufferToSnapshot` output (catches import-side bugs); the
  canvas-fidelity test compares Joplin canvas screenshot → Excel
  reference PNG (catches render-side bugs). Two PNGs at different
  DPRs and table positions: aligned by structural regions, NOT by
  raw pixel coordinates. `findColouredStrip` scans for the dominant-
  coloured strip whose hex is within tolerance of the spec target.
  Per-region tolerance:
    - Region-FINDING: Δ ≤ 24 per channel (wide for anti-aliased
      strip edges).
    - Header / banded ASSERTION: Δ ≤ 8 (tall regions, anti-aliasing
      doesn't dominate).
    - Totals-top STRIP ASSERTION: Δ ≤ 32 (1-2px strip at DPR=2
      pulls in sub-pixel anti-aliasing; pure-pixel parity isn't
      achievable without a different sampling scheme).
  Structural sentinel: "Joplin's totals-top is a SINGLE strip, not
  two strips with a ≤4px gap" — directly catches the DOUBLE-render
  artifact. The test reads the LATEST `eval-{aptos,classic}-*.png`
  by mtime, so re-capturing screenshots automatically refreshes
  what's tested.
- **`tests/util/pngSampler.ts` is test-only.** Pure-stdlib `zlib`
  PNG decoder. Supports 8-bit RGB / RGBA, non-interlaced — the only
  shapes our reference PNGs use. **Don't promote to a runtime
  dependency.** If runtime PNG decoding is ever needed, that's a
  separate design conversation; pulling `pngjs` or `sharp` into the
  bundle is not free.
- **`tableHeaderRowRegion` was Retina-broken until Phase 4.** The
  hardcoded y=22 / h=13 region values assumed a non-Retina Univer
  canvas (DPR=1). On macOS Retina the canvas backing store is at
  DPR=2, so `canvas.width = 3008` instead of `1504`, and y=22
  landed inside the Univer column-letter strip instead of the
  table header row. Fix: scale x/y/w/h by
  `canvas.width / canvas.clientWidth`. clientWidth is the CSS box;
  the ratio is exactly devicePixelRatio. **Future region helpers
  should use this same scaling pattern** — there's no clean way to
  hardcode pixel offsets that work on both Retina and standard
  displays. The simplest invariant: write the offsets in CSS px and
  multiply by `(canvas.width / canvas.clientWidth)` at sample time.
- **Two-layer fidelity test pattern is now the project default for
  "match Excel" features.** Snapshot-data fidelity
  (`tests/excelReferenceFidelity.test.ts`) anchored to the Excel
  reference PNG via `tests/util/pngSampler.ts:dominantColor` —
  catches import-side bugs. Canvas-vs-Excel fidelity
  (`tests/excelCanvasFidelity.test.ts`) anchored to the same
  reference PNG, samples Joplin's `eval-*.png` capture — catches
  render-side bugs. Both anchored UPSTREAM of our code, never
  against our own emit. M13/E rework #2 + #3 is the precedent;
  M15 extends both layers (5 tests per layer for the 5 CF
  columns).
- **Operator captures the Excel reference screenshot for a new
  fixture as part of the cycle.** M13/E precedent: when a new
  fixture is introduced (Aptos and Classic Smorgasboard PNGs were
  captured during the M13/E cycle, not before), the generator
  proceeds with own evidence and snapshot-fidelity work that
  doesn't need the reference, then the reference lands and the
  canvas-fidelity layer can be unskipped. M15's
  `screenshots/excel-reference/ConditionalFormatting-Variants.png`
  follows the same pattern. The canvas-fidelity tests should
  use `test.skip` (or the M13/E "read latest by mtime, fail
  clearly if missing" pattern) when the reference is absent.
- **M15: Univer CF resource shape — `{ [subUnitId]:
  IConditionFormattingRule[] }`** stringified as the snapshot's
  `SHEET_CONDITIONAL_FORMATTING_PLUGIN` resource. Each rule has
  `{cfId, ranges, stopIfTrue, rule: <type-specific>}`. The CF
  preset's `parseJson` does `JSON.parse(json)` and walks subunits;
  ranges use Univer's `IRange` shape `{startRow, endRow,
  startColumn, endColumn}` (zero-based). Multi-rule import works
  via flattening exceljs's `worksheet.conditionalFormattings`
  (which is `Array<{ref, rules: Array<...>}>`) into the per-subunit
  flat array. The CF preset registers via `UniverInstanceType.UNIVER_SHEET`
  business — no per-instance hook setup required.
- **`@univerjs/sheets-conditional-formatting`'s CJS bundle pulls
  `lodash-es`** (ESM-only), which jest-runtime can't parse without
  a transformer override. Tests that need
  `SHEET_CONDITIONAL_FORMATTING_PLUGIN` hard-code the string
  literal locally. Runtime preset (`createUniver` in
  `src/editorView.tsx`) reads the const from the package directly
  (webpack happily bundles the ESM dep). If Univer ever renames the
  constant, the runtime gets the new value but our hard-coded
  string stays the old; the snapshot/runtime mismatch trips loudly.
  Future cycle could add a Jest transformer override
  (`transformIgnorePatterns: ['node_modules/(?!lodash-es)']`) to
  let tests import the const directly — out of scope for M15.
- **iconSet shape mismatch: Univer descending + catch-all vs
  Excel ascending.** Univer's `IconSetCalculateUnit` walks the
  config from index 0 forward, returning the first item whose
  value/operator matches; the standard layout puts the HIGH icon
  at index 0 (e.g. `3Arrows = [up-green, right-gold, down-red]`).
  To match Excel's iconSet semantics where cfvo[0] = lowest band /
  cfvo[N-1] = highest, the import-side translator emits config in
  descending threshold order and uses MAX_SAFE_INTEGER for the
  catch-all entry that handles the lowest band. The export-side
  inverse drops the catch-all and reverses to ascending — Excel
  cfvo[0] is synthesized as `percent=0` since Univer's catch-all
  doesn't carry the lowest-band threshold.
- **CF colour aggregates** (`pinkInk` / `lightGreenInk` /
  `yellowInk`) added to `samplePixelsAt` plus a parallel
  per-column variant in `sampleCfColumns`. The thresholds match
  the operator-ask spec exactly:
    pinkInk:        R≥220 AND G∈[180,220] AND B∈[180,220] (#FFC7CE family)
    lightGreenInk:  R∈[180,220] AND G≥220 AND B∈[180,220] (#C6EFCE family)
    yellowInk:      R≥200 AND G≥200 AND B≤120 (gold arrow / yellow flat icon)
  Two existing aggregates also broadened (monotonic — no prior
  gate regresses):
    redInk:   g/b ≤ 80 → ≤ 140 (catches #F8696B = (248,105,107))
    blueInk:  r/g ≤ 80 → 'b clearly dominant by ≥30 over r AND g, AND b ≥ 150'
              (catches #638EC6 = (99,142,198))
- **`cfAllColumns` regionKind** routes through `sampleCfColumns()`
  instead of `samplePixelsAt(regionFn)`. The sidecar carries a
  top-level `cfColumns` object keyed by column letter (A/C/E/G/I)
  with per-column dominant + ink-aggregate sub-summaries; the
  legacy top-level fields (dominant / sampled / redInk / etc.)
  are filled from an aggregate sample over the whole A2:I11 band
  for backward compat. Future single-fixture multi-region cycles
  can mirror this pattern — name a dedicated `cfXxxYyy` regionKind
  and add a parallel sampler.
- **CF + table-style synthesizer interaction.** The fixture
  doesn't ship a table, so CF-on-top-of-table-style precedence is
  hypothetical for the M15 spec. Verified the
  `FormattingSmorgasboard-NonAptosClassicTheme...` fixture (which
  has both a table AND CF rules) imports without regression — the
  M13/E test suite stays green and Aptos/Classic header colours
  stay correct. The `cfBound` synth-styles sidecar tagging fix is
  out of scope for M15; if a future fixture exercises both
  surfaces and CF colour gets masked by synthesizer fill, mark
  CF-bound cells in `synthStyleSidecar` so the synthesizer skips
  them.
- **iconSet glyphs render correctly** via the preset's bundled
  SVG icon font. No missing-glyph squares observed for `3Arrows`
  on the M15 fixture. The `@univerjs/preset-sheets-conditional-formatting/lib/index.css`
  CSS import is what pulls the fonts; deferring it would have
  rendered as `?`/`□` placeholders. Other iconSets (4Arrows,
  3TrafficLights1, etc.) probably work via the same import but
  aren't under test in M15.
- **The CF round-trip is fully closed via exceljs.** No CF parts
  are dropped on import or export — the `<conditionalFormatting>`
  blocks in the re-emitted xml carry every source rule's type /
  ref / cfvo / colour / operator / rank / iconSet. exceljs
  preserves the dxf `bgColor.argb` for cellIs and top10 styles via
  its `style.fill.bgColor` field. `dataBar` blocks gain extra
  exceljs metadata on round-trip (x14Id, minLength, maxLength,
  axisPosition, etc.) — those are exceljs's defaults and don't
  break Excel's render.
- **Initial title-prefix mismatch caught at eval-screenshot
  time** — `findLatestNoteByTitle` requires the note title to
  start with the prefix INCLUDING the trailing space. Initial
  test note title `"PGE M15 CF eval"` (no trailing space) didn't
  match `"PGE M15 CF eval "` (with space). The fix is to use
  `--title "PGE M15 CF eval $(date -u +...)"` so the timestamp
  appended after the space supplies the matching prefix. Future
  cycles should remember: `import-fixture.sh --title "<PREFIX> ..."`
  must echo the title-prefix verbatim followed by a separator.
- **M15 harness sampler: content-discovering geometry via FUniver
  (Option A).** The first M15 cycle hardcoded
  `COL_W_CSS = 73` in `eval-screenshot.js:sampleCfColumns` and
  computed each CF column's x-origin as
  `ROW_HEADER_W_CSS + colIndex * COL_W_CSS`. When the operator
  widened columns to fit content (the long header labels "Color
  Scale (3-stop)" / "Highlight (>50 = Red)" / etc.), the sampler's
  per-column regions drifted sideways and bucket-leaked across
  columns: `cfColumns.C.blueInk = 0`, `cfColumns.E.pinkInk = 0`,
  `cfColumns.G.lightGreenInk = 0` while `cfColumns.G.pinkInk =
  20790` (col E's pink leaked into col G's region) and
  `cfColumns.I.lightGreenInk = 14706` (col G's green leaked into
  col I's region).

  Fix: `src/editorView.tsx:bootUniver()` exposes the FUniver facade
  on `window.__notesheetUniverAPI` (read-only, diagnostic; harmless
  in production — just an extra global property pointer). The
  sampler now reads live per-column widths and per-row heights via
  `ws.getColumnWidth(c)` / `ws.getRowHeight(r)` and computes
  cumulative x/y origins from those. The pixel sidecar carries a
  new `geometrySource` field — `fUniver` when the facade was
  reachable, `default-fallback` (with reason) when not, so future
  evaluators can tell which path was taken at a glance.

  Other geometry options considered:
    - **Option B** (scan canvas column-header strip for borders):
      doable but brittle — the column-header gridline is anti-aliased
      and faint at DPR=2.
    - **Option C** (OCR column letters): too brittle for harness use.
    - **Option D** (sample whole CF area, bucket by colour signature
      only): drops per-column resolution; would weaken the spec.

  Trade-off vs Option B: Option A requires a live Univer instance.
  If the plugin fails to load, `geometrySource=default-fallback`
  kicks in and the sampler degrades to the (broken) hardcoded math.
  That's acceptable because if the plugin can't load, the screenshot
  itself would also fail to capture CF rendering — the sampler not
  working is the least of the problems.

- **`scripts/pge/widen-columns.js` — utility to align live Joplin
  column widths to the canvas-fidelity test reference geometry.**
  The harness sampler is now width-agnostic, but
  `tests/excelCanvasFidelity.test.ts` (the M15 canvas-vs-Excel
  parity tests) still hardcodes device-pixel x-bands for col G (
  `1290..1480`) and col I (`1635..1655`). Those bands assume all 9
  columns A..I at uniform CSS width 95. After any session that
  changes column widths (operator interaction, snapshot save, etc.)
  re-run `node scripts/pge/widen-columns.js` to reset all 9
  columns to 95 CSS px before invoking `eval-screenshot.sh`. A
  follow-up cleanup could re-tune the canvas-fidelity test
  coordinates to use FUniver-discovered widths the same way the
  sampler does — out of scope for the harness fix cycle.

- **`scripts/pge/move-selection.js` — utility to dismiss A1 active
  cell selection so the screenshot doesn't show the edit cursor over
  cell content.** Univer's active-cell selection paints a `|` cursor
  over the cell's text-render position, which on a narrow-width A1
  visually obscures the header text (you see only `i` when the cell
  contains "Color Scale (3-stop)"). The actual cell value is
  unaffected — verified via `getValue()` — but the screenshot looks
  wrong. Calling `ws.getRange('N1').activate()` moves the selection
  to a far-off blank cell. Future evaluators should run this BEFORE
  `eval-screenshot.sh` if active-cell selection ends up landing
  somewhere that pollutes the visual gate. The existing
  `tableHeaderRowRegion` mitigates this for M13/E by sampling
  past col A; the M15 fixture has multi-column visual evidence
  including col A so we need to actually move the selection.
