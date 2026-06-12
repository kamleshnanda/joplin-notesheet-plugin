# M10 — Chart Export to `.xlsx` (Native OOXML)

> **Status:** Plan revised after Step-0 spike (2026-05-30). The original plan
> had several OOXML correctness gaps (zero-sized chart `xfrm` was a false
> positive, but the schema-completeness, drawing-aggregation, range-cache,
> rId allocation, and worksheet-element-ordering risks were real). All
> findings below are confirmed against either the spike's working
> `/tmp/inject-out.xlsx` (proven to open in Excel cleanly) or the canonical
> Excel-authored samples in `~/Desktop/notesheet-samples/`.

---

## Background — what we learned in the spike

### Where Univer stores chart drawings in the snapshot

The drawing preset writes to `IWorkbookData.resources` under name
`"SHEET_DRAWING_PLUGIN"`. The `data` field is a **JSON-stringified** map:

```ts
{ [subUnitId]: { data: { [drawingId]: ISheetDrawing }, order: string[] } }
```

Each per-drawing entry includes:

```ts
{
  unitId: string,
  subUnitId: string,
  drawingId: string,            // generated, NOT taken from data.chartId
  drawingType: number,          // DrawingTypeEnum.DRAWING_DOM = 8 for our charts
  componentKey: 'NotesheetChart',
  data: {                       // our chart payload, preserved whole
    chartId, type, sourceRange, title, labels, datasets,
  },
  allowTransform: true,
  transform: { left, top, width, height },                     // pixel rect — for the canvas
  sheetTransform: { from: {col, colOff, row, rowOff}, to: {...} },
  axisAlignSheetTransform: { from, to },                       // ← USE THIS for xlsx export
}
```

**Filter** for our charts on export: `entry.componentKey === 'NotesheetChart'`.
Other drawings (images, shapes) coexist in the same resource and must be left
alone.

### Where Excel stores charts in `.xlsx`

Confirmed across all 10 sample files:

| File                                    | Per                     | Notes                                                                                                  |
| --------------------------------------- | ----------------------- | ------------------------------------------------------------------------------------------------------ |
| `xl/charts/chart{N}.xml`                | each chart              | The `<c:chartSpace>` definition (~6.5 KB minimum for a real Excel-authored chart; ours can be smaller) |
| `xl/charts/style{N}.xml`                | each chart              | Modern Excel chart styling (~9.5 KB). Excel 2016+ expects this.                                        |
| `xl/charts/colors{N}.xml`               | each chart              | Color palette (~0.9 KB)                                                                                |
| `xl/charts/_rels/chart{N}.xml.rels`     | each chart              | Links chart to its style + colors via Microsoft-extension types                                        |
| `xl/drawings/drawing{M}.xml`            | **one per sheet**       | Aggregates **all** chart anchors for that sheet as `<xdr:twoCellAnchor>` siblings                      |
| `xl/drawings/_rels/drawing{M}.xml.rels` | one per drawing         | Maps each anchor's `rId` to its `chart{N}.xml`                                                         |
| `xl/worksheets/_rels/sheet{S}.xml.rels` | the chart-bearing sheet | Adds drawing rel; create if absent (exceljs omits rels file when no rels)                              |
| `xl/worksheets/sheet{S}.xml`            | the chart-bearing sheet | Add `<drawing r:id="..."/>` immediately before `</worksheet>`                                          |
| `[Content_Types].xml`                   | workbook-level          | Add four Override entries per chart family (drawing, chart, chartstyle, chartcolorstyle)               |

### Subtleties confirmed in the spike

- **`<xdr:xfrm><a:off x="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` is correct.**
  Excel itself emits zeros and uses the `<xdr:from>`/`<xdr:to>` cell anchors
  to size the chart. We do not need to compute EMU dimensions.
- **One drawing.xml per sheet, multiple charts inside.** Each chart gets a
  `<xdr:twoCellAnchor>` sibling with its own `cNvPr id="2", "3", ...` and
  its own `r:id="rId1", "rId2", ...` pointing into the drawing's rels.
- **Cross-sheet refs work as-is.** When the chart sits on Sheet2 referencing
  Sheet1 data, the drawing rel goes on Sheet2's `_rels` file but the
  `<c:f>` strings inside `chart.xml` use `Sheet1!$A$2:$A$5` — sheet-qualified.
- **XML escaping**: `&`, `<`, `>` in text content. Excel does NOT escape `"`
  or `'` in element text (they are XML-legal there). For attribute values
  (none of which carry user text in our case), all five would need escaping.
- **`<c:barDir>` distinguishes orientation**: `"bar"` = horizontal, `"col"` =
  vertical column. Our Chart.js `type: 'bar'` should map to `barDir="col"`
  to match what users expect from "Bar Chart". The spike copied
  `barDir="bar"` from a sample and produced a horizontal bar layout —
  the technique works either way; the mapping must just be deliberate.

### What works — `/tmp/inject-out.xlsx` proof

The spike injected a hand-authored chart into a vanilla exceljs-written
xlsx, and Excel opened the result with no error dialog and rendered the
chart correctly with title, axis labels, and bars matching source data.
This validates the entire post-processing approach.

---

## Architecture

```
snapshotToXlsxBuffer()                   // src/xlsx.ts — writes cells/styles/tables via exceljs
  → wb.xlsx.writeBuffer()                // produces ArrayBuffer (the zip)
  → maybeInjectCharts(buffer, snapshot)  // NEW post-processing step
  → return ArrayBuffer
```

`maybeInjectCharts` is a no-op when there are no `NotesheetChart` drawings
in the snapshot (graceful degradation, mirrors the M9 pattern for tables).

**New file: `src/charts/xlsxChart.ts`** owns the OOXML synthesis and zip
surgery. It does not import any Univer code at runtime — it works from
the snapshot JSON alone, like the rest of `xlsx.ts`.

---

## Step 1 — Read chart drawings from the snapshot

```ts
interface ChartDrawing {
    chartId: string;
    sheetId: string; // subUnitId from the resource map
    type: 'bar' | 'line' | 'pie' | 'doughnut';
    title: string;
    sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    labels: string[];
    datasets: Array<{ label?: string; data: number[] }>;
    anchor: {
        fromCol: number;
        fromColOff: number; // EMUs
        fromRow: number;
        fromRowOff: number;
        toCol: number;
        toColOff: number;
        toRow: number;
        toRowOff: number;
    };
}

function readChartsFromSnapshot(snapshot: UniverSnapshot): ChartDrawing[];
```

**Implementation details:**

- Walk `snapshot.resources[]`, find `entry.name === 'SHEET_DRAWING_PLUGIN'`.
  Return `[]` if absent.
- `JSON.parse(entry.data)` (with try/catch — fall through to `[]` on
  malformed JSON).
- For each subUnitId, iterate `subUnit.data`, filter to entries where
  `componentKey === 'NotesheetChart'`.
- Pull `data.chartId / type / sourceRange / title / labels / datasets`
  from the entry's `data` field (already in our shape — see M7+M8).
- Pull anchor bounds from `axisAlignSheetTransform.from/to`. The
  `column`/`columnOffset`/`row`/`rowOffset` fields map directly to OOXML
  `<xdr:col>`/`<xdr:colOff>`/`<xdr:row>`/`<xdr:rowOff>`. Offsets are
  already in EMU.
- Skip drawings whose `sheetId` doesn't appear in `snapshot.sheetOrder`.

---

## Step 2 — Build OOXML chart XML

**One generator function per chart type**, all returning a string:

```ts
function buildBarChartXml(c: ChartDrawing, opts: { sheetName: string }): string;
function buildLineChartXml(c: ChartDrawing, opts: { sheetName: string }): string;
function buildPieChartXml(c: ChartDrawing, opts: { sheetName: string }): string;
function buildDoughnutChartXml(c: ChartDrawing, opts: { sheetName: string }): string;
```

Each emits a `<c:chartSpace>` containing:

- Required namespaces: `c`, `a`, `r`, `c16r2` (the alternate-content
  fallback for older readers can be omitted — older readers see a slightly
  duller chart, not an error)
- `<c:chart>` → `<c:title>` (rich text, with our title) → `<c:autoTitleDeleted val="0"/>` → `<c:plotArea>`
- Chart-type-specific node: `<c:barChart>` / `<c:lineChart>` / `<c:pieChart>` / `<c:doughnutChart>`
- `<c:plotVisOnly val="1"/>` and `<c:dispBlanksAs val="gap"/>` after plotArea
- `<c:catAx>` + `<c:valAx>` for bar/line (axId pair must match), omitted for pie/doughnut
- `<c:printSettings>` block at the end (boilerplate)

### Per-series structure inside the chart-type node

```xml
<c:ser>
  <c:idx val="0"/>
  <c:order val="0"/>
  <c:tx>
    <c:strRef>
      <c:f>{sheet}!${col}${row}</c:f>      <!-- header cell -->
      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>{label}</c:v></c:pt></c:strCache>
    </c:strRef>
  </c:tx>
  <c:spPr>
    <a:solidFill><a:srgbClr val="{paletteHex}"/></a:solidFill>
    <a:ln><a:noFill/></a:ln>
  </c:spPr>
  <c:invertIfNegative val="0"/>
  <c:cat>                                  <!-- category labels (string ref) -->
    <c:strRef>
      <c:f>{sheet}!${labelRange}</c:f>
      <c:strCache>...labels...</c:strCache>
    </c:strRef>
  </c:cat>
  <c:val>                                  <!-- values (number ref + cache) -->
    <c:numRef>
      <c:f>{sheet}!${dataRange}</c:f>
      <c:numCache>
        <c:formatCode>General</c:formatCode>
        <c:ptCount val="{N}"/>
        <c:pt idx="0"><c:v>{n}</c:v></c:pt>
        ...
      </c:numCache>
    </c:numRef>
  </c:val>
</c:ser>
```

**Caches are NOT optional.** Without `<c:strCache>` and `<c:numCache>`,
Excel often opens the chart blank until the user manually triggers recalc.
Cache contents come from the ChartDrawing's `labels` and `datasets`
arrays (which the editor already populates from the live sheet at
chart-insert time).

### Per-type variations

| Chart                     | `<c:barChart>` / `<c:lineChart>` / `<c:pieChart>` / `<c:doughnutChart>`                             |
| ------------------------- | --------------------------------------------------------------------------------------------------- |
| **bar (vertical column)** | `<c:barDir val="col"/>`, `<c:grouping val="clustered"/>`, `<c:gapWidth val="182"/>`, two `<c:axId>` |
| **line**                  | `<c:grouping val="standard"/>`, `<c:smooth val="0"/>`, two `<c:axId>`                               |
| **pie**                   | `<c:varyColors val="1"/>`, single `<c:ser>` only, no axIds; per-data-point colors via `<c:dPt>`     |
| **doughnut**              | Same as pie + `<c:holeSize val="50"/>`                                                              |

### Color encoding

Use `CHART_PALETTE` from `src/charts/extractData.ts:11` (already imports).

- **bar/line**: per-series `<c:spPr><a:solidFill><a:srgbClr val="{hex}"/>` —
  one solid color per series, picks from palette by series index.
- **pie/doughnut**: `<c:varyColors val="1"/>` and per-data-point `<c:dPt>`
  override blocks:
    ```xml
    <c:dPt>
      <c:idx val="0"/>
      <c:bubble3D val="0"/>
      <c:spPr><a:solidFill><a:srgbClr val="3B82F6"/></a:solidFill></c:spPr>
    </c:dPt>
    ```
    (Strip leading `#` from `CHART_PALETTE[i]` since OOXML expects a
    hex string without prefix.)

### Range-ref builder

Reuse the existing `colLetters()` helper at `src/xlsx.ts:714` and add a
sibling:

```ts
function rangeRef(sheetName: string, startRow: number, endRow: number, col: number): string {
    // Single column, multiple rows
    return `${escapeSheetName(sheetName)}!$${colLetters(col)}$${startRow + 1}:$${colLetters(col)}$${endRow + 1}`;
}
function cellRef(sheetName: string, row: number, col: number): string {
    return `${escapeSheetName(sheetName)}!$${colLetters(col)}$${row + 1}`;
}
function escapeSheetName(name: string): string {
    // OOXML wraps sheet names containing special chars in single quotes
    // and doubles internal apostrophes.
    return /^[A-Za-z_][A-Za-z0-9_]*$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}
```

### XML escape helper

```ts
function escapeXml(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
```

Required for: chart `title`, every `<c:v>` text node, every `<c:f>` if
the sheet name contains `<` `>` (rare). Numbers don't need escaping.

---

## Step 3 — Build drawing XML

**One drawing.xml per sheet that has at least one chart.** Aggregate all of
that sheet's `ChartDrawing`s into sibling `<xdr:twoCellAnchor>` blocks.

```ts
function buildDrawingXml(charts: ChartDrawing[], rIdBase: number): string;
```

Each anchor:

```xml
<xdr:twoCellAnchor>
  <xdr:from>
    <xdr:col>{fromCol}</xdr:col><xdr:colOff>{fromColOff}</xdr:colOff>
    <xdr:row>{fromRow}</xdr:row><xdr:rowOff>{fromRowOff}</xdr:rowOff>
  </xdr:from>
  <xdr:to>...</xdr:to>
  <xdr:graphicFrame macro="">
    <xdr:nvGraphicFramePr>
      <xdr:cNvPr id="{2 + index}" name="Chart {index + 1}"/>
      <xdr:cNvGraphicFramePr/>
    </xdr:nvGraphicFramePr>
    <xdr:xfrm>
      <a:off x="0" y="0"/>
      <a:ext cx="0" cy="0"/>
    </xdr:xfrm>
    <a:graphic>
      <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
        <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                 r:id="rId{rIdBase + index}"/>
      </a:graphicData>
    </a:graphic>
  </xdr:graphicFrame>
  <xdr:clientData/>
</xdr:twoCellAnchor>
```

Skip the `cNvPr/extLst/creationId` GUID — it's metadata Excel uses for
revision tracking. Spike showed Excel doesn't require it.

---

## Step 4 — Reuse style and color XML across charts

**Open question for the spike**: do all charts share one `style1.xml`/`colors1.xml`,
or does each chart need its own copy?

In all sample files, every chart had its own style and colors files even
when they were identical bytes. We have two cleanup options:

| Option                   | Files emitted per chart                                               | Risk                                                              |
| ------------------------ | --------------------------------------------------------------------- | ----------------------------------------------------------------- |
| **A: per-chart copies**  | `styleN.xml` + `colorsN.xml` per chart                                | Larger zip, identical bytes — bloat. But matches what Excel does. |
| **B: shared singletons** | `style1.xml` + `colors1.xml` once total, all chart rels point to them | Smaller, cleaner. Untested — Excel may complain.                  |

**Decision: ship Option A in v1**, clean up to Option B later if it works.
The bytes are tiny relative to a chart-bearing xlsx anyway.

The style and colors XML content can be a **canonical hard-coded constant**
copied verbatim from `01-bar-simple.xlsx` (since they describe Excel's
default modern theme — universal across our chart types).

---

## Step 5 — Zip surgery

```ts
async function injectChartsIntoZip(
    buffer: ArrayBuffer,
    snapshot: UniverSnapshot,
): Promise<ArrayBuffer> {
    const charts = readChartsFromSnapshot(snapshot);
    if (charts.length === 0) return buffer;

    const zip = await JSZip.loadAsync(buffer);
    // group charts by sheetId
    const bySheet = groupBy(charts, (c) => c.sheetId);

    let chartCounter = 1; // chart{N}.xml
    let drawingCounter = 1; // drawing{M}.xml

    for (const [sheetId, sheetCharts] of bySheet) {
        const sheetIndex = lookupSheetIndex(snapshot, sheetId); // 1-based
        const sheetName = lookupSheetName(snapshot, sheetId);

        // 1. Allocate rIds for the drawing rel + each chart-in-drawing rel
        const sheetRelsPath = `xl/worksheets/_rels/sheet${sheetIndex}.xml.rels`;
        const existingRels = zip.files[sheetRelsPath];
        const startRId = existingRels ? maxRId(await existingRels.async('string')) + 1 : 1;

        // 2. Patch sheet rels (or create if missing)
        upsertSheetRels(zip, sheetRelsPath, startRId, drawingCounter);

        // 3. Patch sheet xml — insert <drawing r:id="rId..."/> before </worksheet>
        const sheetXmlPath = `xl/worksheets/sheet${sheetIndex}.xml`;
        patchSheetXml(zip, sheetXmlPath, startRId);

        // 4. Build and insert drawing.xml + drawing rels
        const drawingPath = `xl/drawings/drawing${drawingCounter}.xml`;
        const drawingRelsPath = `xl/drawings/_rels/drawing${drawingCounter}.xml.rels`;
        const drawingChartIdStart = chartCounter;
        zip.file(drawingPath, buildDrawingXml(sheetCharts, 1 /* rIdBase inside drawing */));
        zip.file(drawingRelsPath, buildDrawingRelsXml(sheetCharts.length, drawingChartIdStart));

        // 5. For each chart on this sheet, emit chart{N}, style{N}, colors{N}, chart rels
        for (let i = 0; i < sheetCharts.length; i++) {
            const c = sheetCharts[i];
            const n = chartCounter;
            zip.file(`xl/charts/chart${n}.xml`, buildChartXml(c, { sheetName }));
            zip.file(`xl/charts/style${n}.xml`, CHART_STYLE_XML); // canonical constant
            zip.file(`xl/charts/colors${n}.xml`, CHART_COLORS_XML); // canonical constant
            zip.file(`xl/charts/_rels/chart${n}.xml.rels`, buildChartRelsXml(n));
            chartCounter++;
        }

        drawingCounter++;
    }

    // 6. Patch [Content_Types].xml — add overrides for everything we wrote
    patchContentTypes(zip, chartCounter - 1);

    return zip.generateAsync({ type: 'arraybuffer' });
}
```

### Patch operations — all string-based, regex-anchored on closing tags

- **`<drawing r:id="rIdN"/>` insertion**: regex on `/<\/worksheet>/` to
  insert just before the closing tag. Schema allows drawing late in the
  worksheet element list (after pageSetup), so this is always safe.
- **Sheet rels patching**: parse the existing `<Relationships>` XML if
  present, append a new `<Relationship>` element, re-emit. Allocate the
  next free `rId` by parsing existing IDs.
- **Content_Types patching**: regex on `/<\/Types>/`, insert four
  `<Override>` per chart:
    ```xml
    <Override PartName="/xl/drawings/drawing{M}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
    <Override PartName="/xl/charts/chart{N}.xml"     ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
    <Override PartName="/xl/charts/style{N}.xml"     ContentType="application/vnd.ms-office.chartstyle+xml"/>
    <Override PartName="/xl/charts/colors{N}.xml"    ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
    ```

### Failure handling

Wrap the whole `injectChartsIntoZip` in try/catch. On any failure
(malformed snapshot, regex mismatch, JSZip error), log and **return the
original unpatched buffer** — better to ship a chart-less xlsx than a
corrupt one. Same fallback discipline as M9's table-export path.

---

## Step 6 — Wire into snapshotToXlsxBuffer

In `src/xlsx.ts`, replace the final `return buffer as ArrayBuffer;` with:

```ts
import { injectChartsIntoZip } from './charts/xlsxChart';

// ...
const buffer = await wb.xlsx.writeBuffer();
return injectChartsIntoZip(buffer as ArrayBuffer, snapshot);
```

---

## Step 7 — Tests

**New file: `tests/xlsxChart.test.ts`**

| Test                               | Asserts                                                                                                                                                                                                           |
| ---------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `rangeRef` builder                 | `Sheet1!$A$2:$A$5` for col=0, startRow=1, endRow=4, sheetName="Sheet1"                                                                                                                                            |
| Sheet name with spaces gets quoted | `'My Sheet'!$B$2:$B$5` for sheetName="My Sheet"                                                                                                                                                                   |
| `escapeXml` handles `&<>`          | `Foo & Bar <X>` → `Foo &amp; Bar &lt;X&gt;`                                                                                                                                                                       |
| `buildBarChartXml` shape           | Output contains `<c:barDir val="col"/>`, one `<c:ser>` per dataset, axId pair, cached values matching dataset                                                                                                     |
| `buildPieChartXml` shape           | Output contains `<c:varyColors val="1"/>`, one `<c:ser>`, no axIds, `<c:dPt>` blocks for each data point                                                                                                          |
| Empty charts → no zip surgery      | A snapshot with no `SHEET_DRAWING_PLUGIN` resource → exporter returns the original buffer unchanged                                                                                                               |
| Bar chart round-trip               | Build snapshot with one bar chart, run full export, unzip result, assert `xl/charts/chart1.xml`, `xl/drawings/drawing1.xml`, the four content-type overrides, and the patched `<drawing>` element in `sheet1.xml` |
| Multiple charts per sheet          | Two bar charts on Sheet1 → one drawing1.xml with two `<xdr:twoCellAnchor>` blocks, two chart files, two of each style/colors file                                                                                 |
| Cross-sheet chart                  | Chart on Sheet2 referencing Sheet1 data → drawing rel on Sheet2's rels, chart `<c:f>` strings carry `Sheet1!` prefix                                                                                              |
| Special chars in title             | Title `Margins "FY26" & <X>` correctly escapes `&` and `<>`, leaves `"` alone                                                                                                                                     |
| Existing M9 tests                  | All 42 existing tests still pass — regression check                                                                                                                                                               |

---

## Step 8 — Documentation

- Flip M10 to ✅ in README milestones table
- Update README "Chart export" hint (currently says it's planned)
- Update `xlsx.ts` header comment — remove "Charts [...] out of scope"

---

## Implementation gotchas (folded in from spike)

- **`<c:f>` AND caches must agree.** If we emit `<c:f>Sheet1!$B$2:$B$5</c:f>`
  but the `<c:numCache>` shows different numbers than what's in those
  cells, Excel may use either depending on calc state. Always derive the
  cache from the same `labels`/`datasets` arrays used for the formula
  ref's bounds.
- **`rId` allocation is dynamic.** M9 already added a tables rel as `rId1`
  to some sheets. Read the existing rels file (if any), find the highest
  `rId{N}`, allocate `N+1` for our drawing.
- **Sheet rels file may not exist.** A sheet with no rels means no rels
  file. Create it from scratch with the correct namespace declarations.
- **Worksheet element ordering.** Insert `<drawing r:id="..."/>` just
  before `</worksheet>` via regex. exceljs writes a single line of XML;
  schema allows drawing as the last element of the sheet content sequence
  so this is always correct.
- **Regenerate `c16:uniqueId` GUIDs per chart** — or omit them. Spike
  showed Excel accepts charts without these.
- **`barDir="col"` for vertical bar charts.** Our Chart.js `type: 'bar'`
  is what users call "Bar Chart" but visually means vertical columns
  (matches Excel's "Column Chart" UI). Map `'bar'` → `barDir="col"`.
  `barDir="bar"` would produce a horizontal layout — confirmed in spike.

---

## Effort estimate

**~2 days end-to-end**, broken down:

| Phase                                                                        | Effort    |
| ---------------------------------------------------------------------------- | --------- |
| Chart XML builders (4 types, escape, range refs, color encoding)             | ~half day |
| Drawing XML, drawing rels                                                    | ~2 hr     |
| Zip surgery (3 patch operations + content-types)                             | ~half day |
| Tests (ranges, escape, per-type shape, round-trip, multi-chart, cross-sheet) | ~half day |
| Manual validation in real Excel + iteration on rough edges                   | ~half day |

This is roughly twice what the original plan implied. The original "couple
of hours" framing missed that XML schema correctness, rId allocation, and
content-types patching are real work.

---

## Out of scope for M10

- **Chart import** from `.xlsx` (parking as M14 — symmetric inverse, bigger scope)
- **Trendlines** (sample 10 showed they're encoded as `<c:trendline>` inside
  `<c:ser>` — easy to add later but defer for v1)
- **Stacked / 100% bar variants** (only clustered for v1)
- **Custom colors per data point** in bar/line (only per-series fills for v1)
- **Multi-axis charts** (single value axis only for v1)
- **Chart legend customization** (rely on Excel's defaults)
- **Date / time axis formats** (only General format on the value axis for v1)

---

## Risk mitigation summary

| Risk                                                 | Mitigation                                                                                                               |
| ---------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------ |
| Generated XML rejected by Excel                      | Spike proved a hand-crafted chart opens cleanly. Build incrementally; test each chart type in real Excel before merging. |
| Real-world snapshots have unexpected drawing entries | Filter strictly on `componentKey === 'NotesheetChart'`. Other drawings (images from M11) coexist untouched.              |
| `rId` collision with M9 tables                       | Read existing sheet rels, allocate next free `rId`.                                                                      |
| Worksheet XML element-order violation                | Insert `<drawing>` immediately before `</worksheet>` — schema-compliant for any worksheet shape exceljs produces.        |
| Cache out of sync with formula ref                   | Always derive `<c:strCache>` and `<c:numCache>` from the SAME `labels`/`datasets` arrays the formula bounds reference.   |
| Failure mid-injection produces corrupt zip           | try/catch around `injectChartsIntoZip` — on error, return the original buffer (chart-less but valid).                    |
