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

## In progress

(Empty between sessions.)

## Next

- **feature-1-m13-rotated-text-renders** — Restore reverted PR #16
  (`415b4a4`) rotation import/export in `src/xlsx.ts`, restore
  `tests/m13RotatedText.test.ts` and the
  `m12FixtureRoundTrip.test.ts` flips. Prove via PGE harness that
  `MergedCellsAndAlignment.xlsx` row 6 (A6 +45°, B6 +90°, C6 -45°)
  renders visibly rotated in Univer. Per `OPERATOR_ASK.md`, this
  is the first real-feature cycle post-smoke; the harness was built
  to catch this class of bug (Jest passes, Univer renders broken)
  and this is the regression test for the harness itself.

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
  it.
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
  Univer main canvas's top-80px row-0 slab and writes a histogram of
  non-background colours alongside the screenshot. Use it for
  machine-checkable assertions like "top contains rgb(255,0,0) > 50
  hits" instead of "I saw red." Sampling is stride-2 to keep cost
  cheap. Background (>235 in all channels) and gridline ink (<30)
  are filtered out.
- **Pre-existing `tests/exportTableRoundTrip.test.ts:334` typecheck
  bug** — `'dashed'` is not in exceljs `BorderStyle` enum. Changed
  to `'mediumDashed'` (with matching assertion update on line 349)
  during this session because webpack's TS check blocks .jpl build.
  This was unrelated to smoke; the smoke didn't introduce it.
