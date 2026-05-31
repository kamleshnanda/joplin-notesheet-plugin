# Chart sample xlsx files

Excel-authored `.xlsx` files used as the canonical reference for OOXML chart
schema during M10 (chart export). Each file was created in real Microsoft
Excel — they show what valid chart XML looks like for a given chart type or
edge case. Useful both as test fixtures and as a "ground truth" the export
implementation should converge toward.

| File | What it tests |
|---|---|
| `01-bar-simple.xlsx` | Single bar chart, single series. The simplest non-trivial reference. |
| `02-line-multi-series.xlsx` | Line chart with three series. Confirms the per-series `<c:idx>`/`<c:order>` pattern repeats. |
| `03-pie-single.xlsx` | Pie chart, single series. Confirms `<c:varyColors val="1"/>` + per-data-point `<c:dPt>` color blocks. |
| `04-doughnut.xlsx` | Doughnut variant — same as pie but with `<c:holeSize val="50"/>`. |
| `05-bar-special-chars.xlsx` | Title `Margins "FY26"`, columns `Q1 & Q2`, `Sales <USD>`. Reference for XML escape conventions: `&` `<` `>` only; `"` survives unescaped in element text. |
| `06-two-charts-one-sheet.xlsx` | Two charts on the same sheet. Reference for one-`drawing1.xml`-aggregates-multiple-anchors pattern. |
| `07-chart-cross-sheet.xlsx` | Chart on Sheet2 referencing Sheet1 data. Reference for sheet-qualified `<c:f>` and where the drawing rel lives. |
| `08-drag-resized.xlsx` | Bar chart manually dragged + resized. Reference for non-default `<xdr:colOff>`/`<xdr:rowOff>` EMU encoding. |
| `09-bar-percent-axis.xlsx` | Bar chart whose value axis is percent-formatted. Reference for `<c:numFmt formatCode="0%" sourceLinked="0"/>` on the val axis. |
| `10-bar-with-trendline.xlsx` | Bar chart with a linear trendline. Trendlines are out of M10 scope — kept for M-something-later. |

## Notes for test authors

- These files are **read-only in tests**. Do not modify them in CI runs;
  they are committed as the canonical shape.
- When asserting expected XML, prefer parsing the relevant element
  (`xl/charts/chart1.xml`) and comparing structurally. Excel-authored
  XML carries Microsoft-extension `<c:extLst>` / `c16:uniqueId` /
  `c14:style` blocks our exporter doesn't emit; structural comparison
  handles that gracefully.
- File sizes are 16–21 KB. Adding more fixtures is fine but factor into
  repo growth.
