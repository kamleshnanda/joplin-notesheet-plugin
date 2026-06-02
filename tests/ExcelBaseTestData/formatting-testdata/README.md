# Formatting fidelity test fixtures

Two hand-saved Excel files that anchor M12's formatting fidelity test
suite. Both are real Excel saves (Mac, Excel 365, mid-2026) — checked
in so test assertions match what Excel actually emits, not what
exceljs guesses. **Do not edit these files** in code; if a feature gap
calls for new fixture content, save a new file and add it to this
table rather than mutating an existing one.

## Files

### `FormattingSmorgasboard.xlsx` — Aptos theme

Worksheet `Sheet1` carries one named table `ProjectTracker` over `A1:G10`
with 8 data rows + 1 totals row. Columns:

| Col | Name | numFmt | Notes |
|-----|------|--------|-------|
| A | Project | (none) | Plain string |
| B | Website | (none) | Each cell uses the **`Hyperlink` named cell style** (cellStyleXfs id=1, builtinId=8). Each row's URL points at `https://example.com/<project>`. |
| C | Budget | `"$"#,##0` (custom 164) | Currency |
| D | Spent | `"$"#,##0` | Currency |
| E | % Complete | `9` (built-in `0%`) | Percent |
| F | Start Date | `"yyyy\\-mm\\-dd"` (custom 165) | Date serial |
| G | Status | (none) | Plain string |

Notable structural elements:

- **Theme**: `Office` (Aptos), `<a:minorFont>` `Aptos Narrow`,
  `<a:majorFont>` `Aptos Display`. `clrScheme` uses Aptos accents
  (accent1 = `#156082`, accent3 = `#196B24`, hlink = `#467886`).
- **Table style**: `TableStyleMedium4` with `showRowStripes=1` and
  `totalsRowCount=1`. Has `headerRowDxfId`, `dataDxfId`, and
  `totalsRowDxfId` references; per-column dataDxfId/totalsRowDxfId
  references too. **17 dxf entries** in `xl/styles.xml` carry the
  numFmt assignments — none of them carry colors. Excel synthesizes
  the actual fill/font from the table style at render time.
- **Cell A2** uses cellXf id 2 which references `borderId=1`. Border
  id 1 is a **medium border in `#F4B183`** (theme-tinted orange) on
  all four sides — this is a hand-applied "highlight" border on the
  first project row. Tests should use this as the canonical
  "user-added cell border" case.
- **Cell B1 (header)** uses cellXf id 1 referencing **font 2** (`Aptos`,
  size 11) and `applyFont=1`. This is unusual — the header carries
  an explicit font. Importing the fixture should preserve that.
- **`cellStyles` element**: `Hyperlink` (`xfId=1, builtinId=8`) and
  `Normal` (`xfId=0, builtinId=0`). Built-in `Hyperlink` is what
  Excel uses to paint blue-underlined hyperlinks via `theme=10` and
  `<u/>`.
- **Table totals row** (row 10) uses `SUBTOTAL(109, …)` formulas in
  C10/D10/E10/G10 with structured table refs
  (`ProjectTracker[Budget]` etc.).
- **Row 1 has `thickBot=1`** on its `ht="17"` height — Excel's way of
  painting a slightly heavier bottom border under the header.

### `FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx` — Office Classic theme + conditional formatting

Worksheet `Sheet1` carries one named table `ProductCatalog` over
`A1:F10` with 8 data rows + 1 totals row. Columns:

| Col | Name | numFmt | Notes |
|-----|------|--------|-------|
| A | Product | (none) | Plain string |
| B | Website | (none) | Each row's cell carries the `Hyperlink` named style |
| C | Price | (currency) | |
| D | Discount | `9` (`0%`) | **Color-scale conditional formatting on `D2:D9`** |
| E | Launch Date | `m/d/yy` | Date |
| F | Revenue | `"$"#,##0` | **Color-scale conditional formatting on `F2:F9`** |

Notable structural elements:

- **Theme**: `Office` (Classic, NOT Aptos). `<a:minorFont>`
  `Calibri`. `clrScheme` uses 2007-era accents
  (accent1 = `#4F81BD`, accent3 = `#9BBB59`, hlink = `#0000FF`). This
  is the palette exceljs hard-codes by default — the round-trip test
  for theme palette fidelity uses BOTH fixtures so we can tell
  whether we're correctly swapping accent3 between Aptos green
  (`#196B24`) and Classic olive (`#9BBB59`).
- **Conditional formatting**: two `<conditionalFormatting>` blocks,
  both `colorScale` rules with three stops: red `#F8696B` at min,
  yellow `#FFEB84` at percentile 50, green `#63BE7B` at max. Anchored
  on `D2:D9` (Discount) and `F2:F9` (Revenue). **This is the only
  conditional formatting we ship in fixtures** — represents the
  baseline shape we need to support before declaring "preserves
  conditional formatting".
- **Table style**: `TableStyleMedium4`. `dxfs` count is `4` (smaller
  than the Aptos fixture's 17 because this file has fewer
  per-column custom formats).
- **Cell B1 (header)** uses `s=1` referencing font 1 (`Calibri`,
  underline). This is a quirk — the header was clicked into the
  Hyperlink-style xfId at some point, even though it has no
  hyperlink. Tests should NOT assert that headers never carry
  underline; some hand-edited Excel files do.

## Common patterns both fixtures preserve

- A named table with a totals row using `SUBTOTAL` structured refs.
- A column that uses Excel's built-in `Hyperlink` named cell style
  (`cellStyleXfs builtinId=8`) — this is the construct we currently
  fail to preserve on round-trip. **Pin-down test target.**
- A workbook theme distinct from exceljs's hardcoded Office 2007
  default. Round-trip must preserve `<a:clrScheme>` byte-for-byte
  (or at least equivalent palette values).
- `dxfs` (differential formats) referenced from a table's
  `headerRowDxfId`/`dataDxfId`/`totalsRowDxfId` and from each
  `tableColumn`. These carry per-column numFmt overrides that the
  table style can't express directly. Currently we drop these on
  import (we only read cell-level `cell.numFmt`), but Excel's
  rendering for the round-tripped file uses the cell-level numFmt
  copies that exceljs writes per-cell, so values still display
  correctly — just at the cost of a fatter `cellXfs`. **Documented
  shortcoming, not a bug.**

## What's intentionally NOT in these fixtures

- Charts (covered by `tests/ExcelBaseTestData/chart-testdata/`).
- Cross-sheet references.
- Pivot tables / slicers / timelines.
- Comments / notes.
- Data validation.
- Drawing shapes other than charts.
- Macros / VBA.

If a future fix needs one of these, ship a separate fixture rather
than expanding either of these two — keep each fixture focused on
the milestone it anchors.
