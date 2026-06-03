# Operator ask — M13/C: rotated text round-trip

This is the first real-feature cycle through the PGE harness, post-smoke.
Workstream C from the reverted M13 PR #16 — rotated-text round-trip —
is the target.

## Why this matters

PR #16 added rotated-text import/export code in `src/xlsx.ts` and 6
new Jest tests. All Jest tests passed. The PR was reverted because
the manual smoke caught the failure mode the PGE was built to catch:
**snapshot data correct, Univer didn't actually render the text rotated.**

The README's "Known shortcomings" table currently flags this as
M13-redo. The harness now exists; this cycle proves it can ship the
feature.

## The feature

When a user imports `tests/ExcelBaseTestData/formatting-testdata/MergedCellsAndAlignment.xlsx`
into a Notesheet note, the rotated cells on the "Merged and
Alignment" sheet must render **visibly rotated in Univer** — not as
horizontal text. Specifically:

- **A6** "Rotated 45 degrees" — text runs upward-to-the-right (CCW 45°)
- **B6** "Rotated 90 degrees" — text runs vertical, baseline pointing
  up (CCW 90°)
- **C6** "Rotated -45 degrees" — text runs downward-to-the-right (CW 45°)

Plus: exporting the same note back to .xlsx via the existing export
command must produce a workbook whose three cells carry
`alignment.textRotation` of 45, 90, -45 respectively (verified with
exceljs in a Jest test). Round-trip stability matters — it's what
prevents future regressions silently dropping rotation again.

Stacked-text mode (`textRotation: 'vertical'`) is acceptable to cover
in Jest only since the existing fixture doesn't exercise it.

## Acceptance criteria

The evaluator must verify ALL of:

1. **Visual** — Evaluator's Playwright-captured screenshot of the
   imported note shows A6, B6, C6 with text rendered at angles
   distinct from horizontal. Eyeball the angles: A6 leans up-right,
   B6 is vertical, C6 leans down-right. NOT three horizontal strings.
2. **Pixel sidecar** — `.pixels.json` for the captured screenshot shows
   the row-6 slab carries text-coloured pixels in a non-horizontal
   distribution. (The harness's existing colour-histogram check is fine
   here; rotated text breaks the dominant-row-of-text-ink pattern.)
3. **Jest pin-down** — The reverted PR #16's tests
   (`tests/m13RotatedText.test.ts` content + the
   `m12FixtureRoundTrip.test.ts` flips) are restored and pass. They
   pin down the snapshot shape: cells with rotation carry
   `style.tr = { a: <angle> }` or `{ a: 0, v: 1 }` for stacked.
4. **Round-trip** — A Jest test that imports
   MergedCellsAndAlignment.xlsx, immediately exports back via
   `snapshotToXlsxBuffer`, re-reads with exceljs, and asserts A6/B6/C6
   carry textRotation 45/90/-45.

## Out of scope

- Conditional rotation (e.g., per-row-style-block in OOXML xfs) is not
  exercised by the fixture. Don't add support for it in this feature.
- Rotated text inside merged cells is also not exercised by row 6.
  If it incidentally works, fine; if it doesn't, it's a separate
  follow-up.
- Rich text combined with rotation (per-run angles) is not a thing in
  OOXML — rotation is cell-level.
- README "Known shortcomings" entry: leave it for the next session.
  This cycle is one feature; the README edit can ride a follow-up.

## Suggested fixture

`tests/ExcelBaseTestData/formatting-testdata/MergedCellsAndAlignment.xlsx`
sheet "Merged and Alignment", row 6 (A6 +45°, B6 +90°, C6 -45°).

The harness's existing `eval-screenshot.js` captures the editor's
Univer canvas. The screenshot must show the rotated cells; the
evaluator picks the visual region of row 6 in its grading.

## Related risks

- **The reverted PR #16's `src/xlsx.ts` changes are the right starting
  point** — `style.tr = { a, v? }` is Univer's ITextRotation shape per
  Univer 0.23. The Jest tests passing isn't the failure mode; what's
  unverified is whether something downstream (snapshot serialization,
  Univer's resolver, the stylesheet flow) drops or no-ops the field.
- **Investigate before patching.** If the reverted code's snapshot
  carries `tr` correctly but Univer ignores it, the fix is in how the
  style is registered or referenced — not in the angle math. Pin down
  what Univer actually receives via a small test or DevTools inspection
  before changing the rotation pipeline.
- **Univer style lookup is by `s` reference.** Per the smoke (M13
  lesson), Univer reads styles from `styles[id]` on the snapshot, not
  inline on the cell. Any new style attribute (here `tr`) must be on
  the styles map, not stuffed onto cellData entries.
- **Don't symptom-patch the test.** If a Jest assertion fails post-
  rebuild, check `git diff package-lock.json` first — exceljs's
  rotation-API surface drifted between 3.x and 4.x. Don't change the
  test to make a downgrade pass.
