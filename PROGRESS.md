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
- **feature-1-m13-theme-aware-banding** (2026-06-03) — Routed
  `synthesizeTableStyleAssignments` through a new `EXCEL_TABLE_STYLE_RECIPE_BY_NAME`
  parallel table (`src/charts/excelTableStyleRecipes.ts`) that names each
  TableStyle slot's accent index + tint. The synthesizer reads the source
  workbook's `<a:clrScheme>` (already captured by `readThemeClrScheme`)
  and resolves the recipe via the ECMA-376 HSL-L tint formula. Aptos
  fixture (`FormattingSmorgasboard.xlsx`) still paints green
  (`#196B24` header + `#84E291` band); Classic fixture
  (`FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`)
  now paints grey (`#A5A5A5` header + `#DBDBDB` band) instead of green.
  Same `EXCEL_TABLE_STYLE_BY_NAME[Medium4]` lookup; two distinct rendered
  outputs driven by source clrScheme. Achromatic styles
  (Light1/8/15, Medium1/8/15/22, Dark1/8) keep their literal greys. Tint
  maths verified exact against the existing Aptos catalog
  (`tint('#196B24', 0.6) = '#84E291'`, `tint('#156082', 0.6) = '#83CBEB'`,
  `tint('#196B24', -0.25) = '#13501B'`). First multi-screenshot cycle:
  added `:variant` suffix support to `eval-screenshot.js`
  (`feature-1-m13-theme-aware-banding:aptos` + `:classic`),
  `tableHeaderRowRegion` helper sampling (x=80..480, y=22..35) cols B+
  to dodge A1 active-cell selection blue, `greyInk` aggregate
  (R,G,B∈[140,180] AND `abs(R-G)≤10` AND `abs(G-B)≤10`), and broadened
  the existing `greenInk` aggregate so it catches dark Aptos green
  `#196B24` (G=107) too. Pixel sidecar over header band:
  Aptos `dominant=rgb(25,107,36) greenInk=4113 greyInk=0`,
  Classic `dominant=rgb(165,165,165) greenInk=0 greyInk=3914`. Three
  new pin-downs in `tests/m12FixturePinDowns.test.ts`
  (M13/E describe block) bring Jest from 195→198. Generator-evidence
  screenshots: `screenshots/feature-1-m13-theme-aware-banding/generator-evidence-{aptos,classic}.png`.
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

## In progress

(Empty between sessions.)

## Next

- **feature-1-m13-theme-aware-banding** — When a workbook ships its
  own non-Aptos `<a:clrScheme>`, the in-Joplin render of every named-
  style table (`TableStyleLight*`, `TableStyleMedium*`,
  `TableStyleDark*`) must derive per-cell `bg` / `cl` / `borderColor`
  from the source clrScheme rather than the hardcoded Aptos catalog
  in `src/charts/excelTableStyles.ts`. Two project-owned fixtures
  pin the two halves: `FormattingSmorgasboard.xlsx` (Aptos, accent3
  `#196B24` green — regression sentinel) and
  `FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx`
  (Classic, accent3 `#A5A5A5` grey — failure-mode sentinel). Both
  use `TableStyleMedium4`; the same catalog entry must paint green
  for one and grey for the other. **First multi-screenshot cycle:**
  the harness needs a `:variant` suffix on `FEATURE_ID`
  (`feature-1-m13-theme-aware-banding:aptos` /
  `feature-1-m13-theme-aware-banding:classic`) plus new entries in
  `REGION_BY_FEATURE` / `TITLE_PREFIX_BY_FEATURE`, a new
  `tableHeaderRowRegion` helper, and a new `greyInk` aggregate in
  `samplePixelsAt`. Don't rename or remove the prior cycles' single-
  key entries. Jest target: ≥197 (M13/D's 195 baseline + ≥2 new
  pin-downs for header colour). M12 invariants
  (`SHEET_NOTESHEET_SYNTH_STYLES_PLUGIN` shape,
  `SHEET_NOTESHEET_THEME_CLR_SCHEME_PLUGIN` shape, applyFont=1
  invariant on synth-only header cells, table-name round-trip) must
  stay green. README "Known shortcomings — Theme-aware banding"
  edit is out-of-scope (defer to follow-up like PR #21 did for
  M13/C and M13/D). See `BUILD_PLAN.md` for the full spec.

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
  unchanged).
