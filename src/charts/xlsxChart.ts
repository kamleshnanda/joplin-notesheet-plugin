// M10: chart export to .xlsx via post-processing the zip exceljs writes.
//
// Architecture (see M10-plan.md for full background):
//   snapshotToXlsxBuffer()  → exceljs writes cells/styles/tables → buffer
//                          → injectChartsIntoZip(buffer, snapshot)
//                          → JSZip surgery: add chart XML parts, patch
//                            sheet rels + content-types
//                          → return new buffer
//
// We bypass exceljs's chart write API (it's thin, uneven across types,
// and produces brittle output). Instead we generate OOXML chart parts
// directly, modeled on the canonical Excel-authored fixtures committed
// at tests/fixtures/charts/. The Step-0 spike proved this technique
// produces files Excel opens without complaint.

import JSZip from 'jszip';

import type { UniverSnapshot } from '../snapshot';
import { CHART_PALETTE } from './extractData';
import { CHART_STYLE_XML, CHART_COLORS_XML } from './xlsxChartConstants';

// ─── Public types ──────────────────────────────────────────────────────────

export type ChartType = 'bar' | 'line' | 'pie' | 'doughnut';

// One chart, normalized for emission. Read out of the snapshot's
// SHEET_DRAWING_PLUGIN resource, stripped of Univer-internal noise.
export interface ChartDrawing {
    chartId: string;
    sheetId: string;
    type: ChartType;
    title: string;
    sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    // Display name of the sheet whose data the chart references. For a chart
    // anchored on the SAME sheet as its data, equals the host sheet's name.
    // For a cross-sheet chart (`07-chart-cross-sheet.xlsx`), points at the
    // DATA sheet's name, NOT the chart's host sheet. M17 plumbs this so the
    // M10 export rebuilds <c:f> formula refs against the original data sheet
    // rather than the chart-containing sheet (which would silently break
    // Excel's re-evaluation from cells on cross-sheet charts).
    sourceSheetName?: string;
    labels: string[];
    datasets: Array<{ label?: string; data: number[] }>;
    anchor: {
        fromCol: number;
        fromColOff: number;
        fromRow: number;
        fromRowOff: number;
        toCol: number;
        toColOff: number;
        toRow: number;
        toRowOff: number;
    };
    // M17 metadata that doesn't fit the rest of the shape. Mirrors
    // `ImportedChartDrawing.meta` so a round-trip preserves it.
    //   - legendPos: 'r'|'l'|'t'|'b'|'tr' (Excel <c:legendPos val="..."/>)
    //   - categoryAxisType: 'index' when the source chart had no <c:cat>
    //     element. Export emits no <c:cat> and treats every column in
    //     sourceRange as a separate values series; otherwise the
    //     existing "first column = labels, rest = series" convention.
    //   - barDir / unsupportedSourceType: surfaced for forward
    //     compatibility; M10 export currently honors barDir only on
    //     emit (out-of-scope for round-trip per BUILD_PLAN feature-5).
    meta?: {
        legendPos?: 'r' | 'l' | 't' | 'b' | 'tr';
        categoryAxisType?: 'index' | 'category';
        barDir?: 'bar' | 'col';
        barGrouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
        barGapWidth?: number;
        unsupportedSourceType?: string;
        holeSize?: number;
        lineSmooth?: boolean;
        lineMarkerOn?: boolean;
        dispBlanksAs?: 'gap' | 'zero' | 'span';
        crossBetween?: 'between' | 'midCat';
        tickMark?: 'none' | 'in' | 'out' | 'cross';
        // Axis-line styling. Excel's modern default chart template draws
        // the CATEGORY axis as a thin light-grey line and the VALUE axis
        // with NO line (every fixture in tests/fixtures/charts uses this).
        // When we omit the axis <c:spPr> on export Excel falls back to its
        // legacy DARK solid line on BOTH axes — which reads as "Joplin
        // introduced axes / a vertical line that wasn't in the source".
        // 'grey' = the light-grey modern line; 'none' = <a:ln><a:noFill/>.
        // Defaults on export (when meta is absent): catAxisLine 'grey',
        // valAxisLine 'none' — i.e. the modern template, NOT Excel's dark
        // legacy default.
        catAxisLine?: 'grey' | 'none';
        valAxisLine?: 'grey' | 'none';
        // Number format applied to value-axis tick labels at runtime.
        // Source XML <c:valAx><c:numFmt formatCode="..."/>. Routed to
        // Chart.js scales.y.ticks.callback so e.g. "0%" renders as
        // 5%, 10%, ... instead of 0.05, 0.1, ... (the percent-axis bug
        // observed on 09-bar-percent-axis).
        valAxisNumFmt?: string;
        // Per-chart data-label flags. Source <c:dLbls> ships these as
        // a set; we forward them so M10 export emits matching <c:dLbls>
        // and pie/doughnut slices show their values/percentages on
        // re-open in Excel.
        dLbls?: {
            showVal?: boolean;
            showCatName?: boolean;
            showPercent?: boolean;
            showSerName?: boolean;
            showLegendKey?: boolean;
            showBubbleSize?: boolean;
        };
        // Per-series trendline. Re-emitted as <c:ser><c:trendline> so a
        // chart imported with a trendline (10-bar-with-trendline) keeps it
        // on export. Attached to the FIRST series only (matches import,
        // which reads the first series' trendline).
        trendline?: {
            type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
            order?: number;
            period?: number;
            dispRSqr?: boolean;
            dispEq?: boolean;
        };
    };
}

// ─── XML primitives ────────────────────────────────────────────────────────

// Element-text escaping. OOXML follows XML rules: &, <, > must be escaped.
// " and ' are XML-legal in element text — Excel itself emits them raw, so
// we follow suit (verified in tests/fixtures/charts/05-bar-special-chars).
export function escapeXml(s: string): string {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// 0-based column index → A1 letters. Mirrors src/xlsx.ts:colLetters but kept
// local to avoid a cross-file circular import.
export function colLetters(idx: number): string {
    let n = idx + 1;
    let s = '';
    while (n > 0) {
        const r = (n - 1) % 26;
        s = String.fromCharCode(65 + r) + s;
        n = Math.floor((n - 1) / 26);
    }
    return s;
}

// Excel range/cell refs in OOXML use $A$1 form for absolute. Sheet names
// containing anything other than alphanumerics/underscores must be wrapped
// in single quotes with internal apostrophes doubled.
export function escapeSheetName(name: string): string {
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
    return `'${name.replace(/'/g, "''")}'`;
}

export function cellRef(sheetName: string, row: number, col: number): string {
    return `${escapeSheetName(sheetName)}!$${colLetters(col)}$${row + 1}`;
}

// Single-column range on one sheet.
export function rangeRefCol(
    sheetName: string,
    startRow: number,
    endRow: number,
    col: number,
): string {
    return `${escapeSheetName(sheetName)}!$${colLetters(col)}$${startRow + 1}:$${colLetters(col)}$${endRow + 1}`;
}

// ─── Color helper ──────────────────────────────────────────────────────────

// CHART_PALETTE values are #RRGGBB; OOXML wants RRGGBB without the hash.
function paletteHex(seriesIndex: number): string {
    const c = CHART_PALETTE[seriesIndex % CHART_PALETTE.length];
    return c.replace(/^#/, '').toUpperCase();
}

// Build the chart-level <c:dLbls> XML from meta.dLbls. Element order
// inside <c:dLbls> is strict (showLegendKey, showVal, showCatName,
// showSerName, showPercent, showBubbleSize) — Excel rejects out-of-order.
// Each <c:show*> defaults to "0" when meta omits it; we always emit all
// six children since Excel needs a complete shape (matches what every
// source fixture writes).
function buildDataLabelsXml(c: ChartDrawing): string {
    const flags = c.meta?.dLbls;
    if (!flags) return '';
    const v = (b: boolean | undefined): string => (b ? '1' : '0');
    return `<c:dLbls><c:showLegendKey val="${v(flags.showLegendKey)}"/><c:showVal val="${v(flags.showVal)}"/><c:showCatName val="${v(flags.showCatName)}"/><c:showSerName val="${v(flags.showSerName)}"/><c:showPercent val="${v(flags.showPercent)}"/><c:showBubbleSize val="${v(flags.showBubbleSize)}"/></c:dLbls>`;
}

// ─── chart{N}.xml builders ─────────────────────────────────────────────────

// All four builders return a complete <c:chartSpace> XML document including
// the XML prolog. Element ordering follows OOXML's strict EG_ChartContent
// sequence; deviating produces "We found a problem" in Excel.

interface BuildChartOpts {
    sheetName: string;
}

// Bar/column. <c:barDir val="col"/> = vertical (Excel "Column Chart"),
// val="bar" = horizontal (Excel "Bar Chart"). M17 plumbs source
// orientation through meta.barDir so a re-imported chart re-renders
// in the same direction as the original. Likewise meta.barGrouping
// surfaces 'clustered' (default), 'stacked', or 'percentStacked';
// 'standard' is line-only and would be invalid here, so we coerce.
// Stacked bars also need <c:overlap val="100"/> so segments stack
// flush instead of side-by-side — Excel's default for clustered is
// 'overlap=-27' (small gap) and for stacked is 'overlap=100'.
export function buildBarChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ true, () => {
        const barDir = c.meta?.barDir ?? 'col';
        const rawGrouping = c.meta?.barGrouping;
        const grouping =
            rawGrouping === 'stacked' || rawGrouping === 'percentStacked'
                ? rawGrouping
                : 'clustered';
        const overlapXml = grouping === 'clustered' ? '' : '<c:overlap val="100"/>';
        // Default 150 matches Excel's default for bar charts
        // (gap = 1.5x bar width). Source values from the import path
        // override.
        const gapWidth = c.meta?.barGapWidth ?? 150;
        const seriesXml = c.datasets
            .map((ds, i) => buildSeriesXml(c, opts, ds, i, /* solidFill */ paletteHex(i)))
            .join('');
        const dLblsXml = buildDataLabelsXml(c);
        return `<c:barChart><c:barDir val="${barDir}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${dLblsXml}<c:gapWidth val="${gapWidth}"/>${overlapXml}<c:axId val="111111"/><c:axId val="222222"/></c:barChart>${categoryAndValueAxes(c)}`;
    });
}

// Line. ECMA-376 line grouping values: 'standard' (default — series
// drawn separately), 'stacked' (cumulative), 'percentStacked'
// (cumulative normalised to 100%). Pie/doughnut have no grouping.
export function buildLineChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ true, () => {
        const rawGrouping = c.meta?.barGrouping;
        const grouping =
            rawGrouping === 'stacked' || rawGrouping === 'percentStacked'
                ? rawGrouping
                : 'standard';
        // Chart-level marker on/off. Excel default 0 (off); plumb via
        // meta.lineMarkerOn so a fixture with markers round-trips.
        const markerOn = c.meta?.lineMarkerOn ? '1' : '0';
        const seriesXml = c.datasets
            .map((ds, i) => buildSeriesXml(c, opts, ds, i, paletteHex(i), /* lineSeries */ true))
            .join('');
        const dLblsXml = buildDataLabelsXml(c);
        return `<c:lineChart><c:grouping val="${grouping}"/><c:varyColors val="0"/>${seriesXml}${dLblsXml}<c:marker val="${markerOn}"/><c:axId val="111111"/><c:axId val="222222"/></c:lineChart>${categoryAndValueAxes(c)}`;
    });
}

export function buildPieChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ false, () => {
        // Pie has exactly one series — datasets[0] is the data, labels are
        // the category points. Per-data-point colors via <c:dPt> overrides.
        const ds = c.datasets[0] ?? { data: [] };
        const dLblsXml = buildDataLabelsXml(c);
        return `<c:pieChart><c:varyColors val="1"/>${buildPieSeriesXml(c, opts, ds, /* doughnut */ false)}${dLblsXml}<c:firstSliceAng val="0"/></c:pieChart>`;
    });
}

export function buildDoughnutChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ false, () => {
        const ds = c.datasets[0] ?? { data: [] };
        // Hole size: Excel default 50 (half-radius hole); user range 1-90.
        // Source value preserved via meta.holeSize.
        const holeSize = c.meta?.holeSize ?? 50;
        const dLblsXml = buildDataLabelsXml(c);
        return `<c:doughnutChart><c:varyColors val="1"/>${buildPieSeriesXml(c, opts, ds, /* doughnut */ true)}${dLblsXml}<c:firstSliceAng val="0"/><c:holeSize val="${holeSize}"/></c:doughnutChart>`;
    });
}

// Top-level wrapper shared by all chart types. The element order inside
// <c:chart> is mandatory: title, autoTitleDeleted, plotArea, legend,
// plotVisOnly, dispBlanksAs. After </c:chart>, the chartSpace gets a
// default spPr.
//
// Legend visibility mirrors what NotesheetChart's runtime config shows:
// always for pie/doughnut (one slice per category), and for any
// multi-series chart. Single-series bar/line gets `autoTitleDeleted=1`
// equivalent for the legend (no <c:legend> element emitted, Excel
// hides it by default).
function chartSpaceWrap(
    c: ChartDrawing,
    _opts: BuildChartOpts,
    _hasAxes: boolean,
    plotAreaInner: () => string,
): string {
    const titleXml = c.title
        ? `<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" kern="1200"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(c.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
        : `<c:autoTitleDeleted val="1"/>`;

    const showLegend = c.type === 'pie' || c.type === 'doughnut' || c.datasets.length > 1;
    // Honor source-XML legend position when it survived round-trip via
    // meta.legendPos. Falls back to 'r' (Excel's default).
    const legendPos = c.meta?.legendPos ?? 'r';
    const legendXml = showLegend
        ? `<c:legend><c:legendPos val="${legendPos}"/><c:overlay val="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900" b="0" kern="1200"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:legend>`
        : '';

    // dispBlanksAs: 'gap' (default — leave a hole), 'zero' (plot 0),
    // 'span' (bridge across the blank). Source value via meta.
    const dispBlanksAs = c.meta?.dispBlanksAs ?? 'gap';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${titleXml}<c:plotArea><c:layout/>${plotAreaInner()}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legendXml}<c:plotVisOnly val="1"/><c:dispBlanksAs val="${dispBlanksAs}"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

// One <c:ser> for bar or line. Element order inside <c:ser> is strict:
// idx, order, tx, spPr, marker?, invertIfNegative?, cat, val.
//
// Two range conventions:
//   - 'category' (default): first column = labels, rest = series. Each
//     series picks dataCol = startColumn + 1 + seriesIndex; <c:cat>
//     points at labelCol.
//   - 'index' (no <c:cat> in source — see ImportedChartDrawing.meta.
//     categoryAxisType): every column in sourceRange is its own values
//     series. dataCol = startColumn + seriesIndex; <c:cat> is omitted.
//     dataStartRow is sourceRange.startRow itself (no header row to
//     skip; the header lives one row above and provides the series
//     name only).
function buildSeriesXml(
    c: ChartDrawing,
    opts: BuildChartOpts,
    ds: { label?: string; data: number[] },
    seriesIndex: number,
    fillHex: string,
    lineSeries = false,
): string {
    const isIndexAxis = c.meta?.categoryAxisType === 'index';
    const labelCol = c.sourceRange.startColumn;
    const dataCol = isIndexAxis
        ? c.sourceRange.startColumn + seriesIndex
        : c.sourceRange.startColumn + 1 + seriesIndex;
    const dataStartRow = isIndexAxis ? c.sourceRange.startRow : c.sourceRange.startRow + 1; // skip the header row
    const dataEndRow = c.sourceRange.endRow;

    // Header cell at the top of the series's column gives the series
    // label. For category-axis charts the header sits at sourceRange's
    // startRow (since dataStartRow = startRow + 1). For index-axis
    // charts the data starts at startRow itself, so the header lives
    // one row ABOVE — typically row 0 in our normalized snapshot. If
    // startRow is already 0 there's no header row available; use the
    // dataset's `label` directly without a sheet ref (Excel renders
    // the inline strCache value in the legend just fine).
    const headerRow = isIndexAxis ? c.sourceRange.startRow - 1 : c.sourceRange.startRow;
    const seriesName = ds.label ?? `Series ${seriesIndex + 1}`;
    const txXml =
        headerRow >= 0
            ? `<c:tx><c:strRef><c:f>${cellRef(opts.sheetName, headerRow, dataCol)}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(seriesName)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`
            : `<c:tx><c:v>${escapeXml(seriesName)}</c:v></c:tx>`;

    // Per-series fill (bar) or line stroke (line). Both reference the palette.
    const spPrXml = lineSeries
        ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>`
        : `<c:spPr><a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>`;

    // Per-series marker symbol. For lines: when meta.lineMarkerOn we
    // emit a generic 'circle' symbol (Chart.js renders dots; Excel
    // shows the same). When markers are off (default) we explicitly
    // emit val="none" so Excel doesn't fall back to its own per-style
    // default (which can be triangles or squares depending on series
    // index). For bars: <c:invertIfNegative val="0"/> matches Excel's
    // default — niche feature, no fixture exercises it today, would
    // be straightforward to plumb if needed.
    const markerXml = lineSeries
        ? `<c:marker><c:symbol val="${c.meta?.lineMarkerOn ? 'circle' : 'none'}"/></c:marker>`
        : `<c:invertIfNegative val="0"/>`;

    // Categories: the label column (string ref + cache). Skip entirely
    // for index-axis charts — Excel infers row index 1..N when <c:cat>
    // is absent, matching the source workbook's behaviour.
    let catXml = '';
    if (!isIndexAxis) {
        const labelsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, labelCol);
        const labelsCacheXml = c.labels
            .map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`)
            .join('');
        catXml = `<c:cat><c:strRef><c:f>${labelsRef}</c:f><c:strCache><c:ptCount val="${c.labels.length}"/>${labelsCacheXml}</c:strCache></c:strRef></c:cat>`;
    }

    // Values: this column's range (number ref + cache).
    const valsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, dataCol);
    const valsCacheXml = ds.data
        .map((n, i) => {
            // NaN/Infinity must be omitted — Excel won't load <c:v>NaN</c:v>.
            if (typeof n !== 'number' || !Number.isFinite(n)) return '';
            return `<c:pt idx="${i}"><c:v>${n}</c:v></c:pt>`;
        })
        .join('');
    const valXml = `<c:val><c:numRef><c:f>${valsRef}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${ds.data.length}"/>${valsCacheXml}</c:numCache></c:numRef></c:val>`;

    // Smoothing: meta.lineSmooth controls per-series. Bar charts
    // ignore <c:smooth> (it only affects line splines) but Excel
    // accepts the element for any chart type.
    const smoothVal = c.meta?.lineSmooth ? '1' : '0';
    // Trendline: emit on the FIRST series only (matches import, which reads
    // a single trendline). OOXML element order inside <c:ser> puts
    // <c:trendline> AFTER marker/dLbls but BEFORE <c:cat>/<c:val>.
    const trendlineXml = seriesIndex === 0 ? buildTrendlineXml(c) : '';
    return `<c:ser><c:idx val="${seriesIndex}"/><c:order val="${seriesIndex}"/>${txXml}${spPrXml}${markerXml}${trendlineXml}${catXml}${valXml}<c:smooth val="${smoothVal}"/></c:ser>`;
}

// Build <c:trendline> from meta.trendline. Element order inside is strict:
// spPr, trendlineType, order (poly), period (movingAvg), dispRSqr, dispEq.
// Returns '' when no trendline is set. The dashed accent line + R²/equation
// label flags mirror what Excel emits for a default linear trendline.
function buildTrendlineXml(c: ChartDrawing): string {
    const tl = c.meta?.trendline;
    if (!tl) return '';
    const spPr =
        '<c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="sysDot"/></a:ln><a:effectLst/></c:spPr>';
    const typeXml = `<c:trendlineType val="${tl.type}"/>`;
    // <c:order> is only meaningful for poly; <c:period> only for movingAvg.
    const orderXml =
        tl.type === 'poly' && typeof tl.order === 'number' ? `<c:order val="${tl.order}"/>` : '';
    const periodXml =
        tl.type === 'movingAvg' && typeof tl.period === 'number'
            ? `<c:period val="${tl.period}"/>`
            : '';
    const rsqrXml = tl.dispRSqr ? '<c:dispRSqr val="1"/>' : '';
    const eqXml = tl.dispEq ? '<c:dispEq val="1"/>' : '';
    return `<c:trendline>${spPr}${typeXml}${orderXml}${periodXml}${rsqrXml}${eqXml}</c:trendline>`;
}

// Pie/doughnut have a single series; per-slice color overrides via <c:dPt>.
function buildPieSeriesXml(
    c: ChartDrawing,
    opts: BuildChartOpts,
    ds: { label?: string; data: number[] },
    _doughnut: boolean,
): string {
    const labelCol = c.sourceRange.startColumn;
    const dataCol = c.sourceRange.startColumn + 1;
    const headerRow = c.sourceRange.startRow;
    const dataStartRow = c.sourceRange.startRow + 1;
    const dataEndRow = c.sourceRange.endRow;

    const seriesNameRef = cellRef(opts.sheetName, headerRow, dataCol);
    const seriesName = ds.label ?? 'Series 1';
    const txXml = `<c:tx><c:strRef><c:f>${seriesNameRef}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(seriesName)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;

    // Per-data-point fills via <c:dPt>. Element order within each dPt is
    // strict: idx, invertIfNegative, bubble3D, spPr.
    const dPtXml = ds.data
        .map((_, i) => {
            const hex = paletteHex(i);
            return `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:dPt>`;
        })
        .join('');

    const labelsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, labelCol);
    const labelsCacheXml = c.labels
        .map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`)
        .join('');
    const catXml = `<c:cat><c:strRef><c:f>${labelsRef}</c:f><c:strCache><c:ptCount val="${c.labels.length}"/>${labelsCacheXml}</c:strCache></c:strRef></c:cat>`;

    const valsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, dataCol);
    const valsCacheXml = ds.data
        .map((n, i) => {
            if (typeof n !== 'number' || !Number.isFinite(n)) return '';
            return `<c:pt idx="${i}"><c:v>${n}</c:v></c:pt>`;
        })
        .join('');
    const valXml = `<c:val><c:numRef><c:f>${valsRef}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${ds.data.length}"/>${valsCacheXml}</c:numCache></c:numRef></c:val>`;

    return `<c:ser><c:idx val="0"/><c:order val="0"/>${txXml}${dPtXml}${catXml}${valXml}</c:ser>`;
}

// Shared cat/val axes block for bar/line. axId pair must match the
// values used inside the chart-type element above (we reuse 111111 /
// 222222 — these are within-chart identifiers, not cross-chart, so
// safe to recycle).
//
// Axis position depends on bar orientation:
//   * Vertical bars/lines (barDir="col" or any line):
//       catAx (categories) at bottom, valAx (values) on left
//   * Horizontal bars (barDir="bar"):
//       catAx (categories) on left, valAx (values) at bottom
// Earlier we hardcoded the vertical-bar layout; horizontal-bar exports
// rendered with swapped gridlines and tick positions. Now driven by
// meta.barDir + the chart type.
//
// Gridline color: source workbooks ship explicit
// <a:lumMod val="15000"/><a:lumOff val="85000"/> (light grey,
// ~85% white on tx1). My emit previously dropped that styling, so
// Excel rendered its default darker gridlines. We now embed the same
// light-grey spPr Excel itself emits so the visible chart matches.
//
// Tick marks default to 'none' — every fixture surveyed uses 'none';
// the previous 'out' hardcode was wrong. Plumb through meta.tickMark
// for round-trip; default 'none' is also Excel's default for most
// modern bar/line styles.
function categoryAndValueAxes(c: ChartDrawing): string {
    const crossBetween = c.meta?.crossBetween ?? 'between';
    const isHorizontalBar = c.type === 'bar' && c.meta?.barDir === 'bar';
    const catAxPos = isHorizontalBar ? 'l' : 'b';
    const valAxPos = isHorizontalBar ? 'b' : 'l';
    const tickMark = c.meta?.tickMark ?? 'none';
    const lightGreyLine =
        '<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>';
    const majorGridlinesXml = `<c:majorGridlines><c:spPr>${lightGreyLine}<a:effectLst/></c:spPr></c:majorGridlines>`;
    // Axis-line <c:spPr>. WITHOUT this Excel paints its legacy DARK solid
    // line on both axes — the "introduced axes / vertical line" the user
    // reported on 01/02/07. Default to Excel's modern template (every
    // source fixture uses it): category axis = light-grey line, value axis
    // = no line. meta can override per-axis for round-trip faithfulness.
    const axisSpPr = (style: 'grey' | 'none'): string =>
        style === 'none'
            ? '<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>'
            : `<c:spPr><a:noFill/>${lightGreyLine}<a:effectLst/></c:spPr>`;
    const catAxisSpPr = axisSpPr(c.meta?.catAxisLine ?? 'grey');
    const valAxisSpPr = axisSpPr(c.meta?.valAxisLine ?? 'none');
    // Value-axis number format. When source provides a non-General
    // format (e.g. "0%" on 09-bar-percent-axis) we emit it with
    // sourceLinked="0" so Excel uses our explicit code instead of
    // pulling from the cell. catAx stays General — categorical
    // axis shows label strings, not numbers.
    const valNumFmt = c.meta?.valAxisNumFmt
        ? `<c:numFmt formatCode="${escapeXml(c.meta.valAxisNumFmt)}" sourceLinked="0"/>`
        : `<c:numFmt formatCode="General" sourceLinked="1"/>`;
    return `<c:catAx><c:axId val="111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catAxPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="${tickMark}"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${catAxisSpPr}<c:crossAx val="222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valAxPos}"/>${majorGridlinesXml}${valNumFmt}<c:majorTickMark val="${tickMark}"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${valAxisSpPr}<c:crossAx val="111111"/><c:crosses val="autoZero"/><c:crossBetween val="${crossBetween}"/></c:valAx>`;
}

// ─── drawing{N}.xml builder ────────────────────────────────────────────────

// Aggregates ALL of a sheet's chart anchors into one wsDr. The drawing-rels
// rIds inside the anchors are 1-based per drawing.xml (independent of the
// sheet rels rIds — those live in xl/worksheets/_rels/sheet{S}.xml.rels).
export function buildDrawingXml(charts: ChartDrawing[]): string {
    const anchors = charts
        .map((c, i) => {
            const a = c.anchor;
            const rId = i + 1; // 1-based rId inside drawing{N}.xml.rels
            const cNvPrId = i + 2; // Excel uses id=2,3,... for graphic frames
            return (
                `<xdr:twoCellAnchor>` +
                `<xdr:from><xdr:col>${a.fromCol}</xdr:col><xdr:colOff>${a.fromColOff}</xdr:colOff><xdr:row>${a.fromRow}</xdr:row><xdr:rowOff>${a.fromRowOff}</xdr:rowOff></xdr:from>` +
                `<xdr:to><xdr:col>${a.toCol}</xdr:col><xdr:colOff>${a.toColOff}</xdr:colOff><xdr:row>${a.toRow}</xdr:row><xdr:rowOff>${a.toRowOff}</xdr:rowOff></xdr:to>` +
                `<xdr:graphicFrame macro="">` +
                `<xdr:nvGraphicFramePr><xdr:cNvPr id="${cNvPrId}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
                // <a:off>/<a:ext> zeros are correct — anchor drives size, this xfrm is a placeholder. Spike confirmed.
                `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
                `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${rId}"/></a:graphicData></a:graphic>` +
                `</xdr:graphicFrame>` +
                `<xdr:clientData/>` +
                `</xdr:twoCellAnchor>`
            );
        })
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`;
}

// ─── rels builders ─────────────────────────────────────────────────────────

// drawing{N}.xml.rels — one Relationship per chart on this drawing's sheet.
// chartFileNumbers lines up 1:1 with the drawing's anchors; e.g. anchors
// pointing to chart3.xml + chart4.xml (because charts 1,2 are on a different
// sheet's drawing) → chartFileNumbers=[3,4].
export function buildDrawingRelsXml(chartFileNumbers: number[]): string {
    const rels = chartFileNumbers
        .map(
            (n, i) =>
                `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${n}.xml"/>`,
        )
        .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
}

// chart{N}.xml.rels — every chart points to its own style + colors files.
// rId numbers match the canonical Excel-authored layout (style=rId1,
// colors=rId2). The chartFileNumber here is just for naming the targets.
export function buildChartRelsXml(chartFileNumber: number): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style${chartFileNumber}.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors${chartFileNumber}.xml"/></Relationships>`;
}

// ─── Dispatch ──────────────────────────────────────────────────────────────

export function buildChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    switch (c.type) {
        case 'bar':
            return buildBarChartXml(c, opts);
        case 'line':
            return buildLineChartXml(c, opts);
        case 'pie':
            return buildPieChartXml(c, opts);
        case 'doughnut':
            return buildDoughnutChartXml(c, opts);
        default:
            return buildBarChartXml(c, opts);
    }
}

// ─── Read drawings from snapshot (Step 1 of M10) ───────────────────────────

// The SHEET_DRAWING_PLUGIN resource's data field is a JSON-stringified map:
//   { [subUnitId]: { data: { [drawingId]: ISheetDrawing }, order: string[] } }
// We filter to entries whose componentKey === 'NotesheetChart' (other
// drawings — images, shapes — coexist and must be left alone).
export function readChartsFromSnapshot(snapshot: UniverSnapshot): ChartDrawing[] {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> })
        .resources;
    if (!Array.isArray(resources)) return [];
    const entry = resources.find((r) => r?.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry || typeof entry.data !== 'string') return [];

    let parsed: Record<string, { data?: Record<string, unknown> }>;
    try {
        parsed = JSON.parse(entry.data);
    } catch {
        return [];
    }
    if (!parsed || typeof parsed !== 'object') return [];

    const out: ChartDrawing[] = [];
    for (const subUnitId of Object.keys(parsed)) {
        const subUnit = parsed[subUnitId];
        const drawings = subUnit?.data;
        if (!drawings || typeof drawings !== 'object') continue;

        for (const drawingId of Object.keys(drawings)) {
            const d = drawings[drawingId] as {
                componentKey?: string;
                data?: {
                    chartId?: string;
                    type?: string;
                    sourceRange?: {
                        startRow?: number;
                        endRow?: number;
                        startColumn?: number;
                        endColumn?: number;
                    };
                    sourceSheetName?: string;
                    title?: string;
                    labels?: unknown[];
                    datasets?: Array<{ label?: string; data?: unknown[] }>;
                    meta?: {
                        legendPos?: 'r' | 'l' | 't' | 'b' | 'tr';
                        categoryAxisType?: 'index' | 'category';
                        barDir?: 'bar' | 'col';
                        barGrouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
                        barGapWidth?: number;
                        unsupportedSourceType?: string;
                        holeSize?: number;
                        lineSmooth?: boolean;
                        lineMarkerOn?: boolean;
                        dispBlanksAs?: 'gap' | 'zero' | 'span';
                        crossBetween?: 'between' | 'midCat';
                        tickMark?: 'none' | 'in' | 'out' | 'cross';
                        catAxisLine?: 'grey' | 'none';
                        valAxisLine?: 'grey' | 'none';
                        valAxisNumFmt?: string;
                        dLbls?: {
                            showVal?: boolean;
                            showCatName?: boolean;
                            showPercent?: boolean;
                            showSerName?: boolean;
                            showLegendKey?: boolean;
                            showBubbleSize?: boolean;
                        };
                        trendline?: {
                            type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
                            order?: number;
                            period?: number;
                            dispRSqr?: boolean;
                            dispEq?: boolean;
                        };
                    };
                };
                axisAlignSheetTransform?: {
                    from?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                    to?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                };
                sheetTransform?: {
                    from?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                    to?: {
                        column?: number;
                        columnOffset?: number;
                        row?: number;
                        rowOffset?: number;
                    };
                };
            };

            if (d?.componentKey !== 'NotesheetChart') continue;
            const data = d.data;
            if (!data) continue;

            const type =
                data.type === 'bar' ||
                data.type === 'line' ||
                data.type === 'pie' ||
                data.type === 'doughnut'
                    ? data.type
                    : 'bar';

            const sr = data.sourceRange;
            if (
                !sr ||
                typeof sr.startRow !== 'number' ||
                typeof sr.endRow !== 'number' ||
                typeof sr.startColumn !== 'number' ||
                typeof sr.endColumn !== 'number'
            )
                continue;

            // Prefer axisAlignSheetTransform — it's xlsx-aligned per Univer's
            // service. Fall back to sheetTransform on the (rare) chance the
            // axisAlign one wasn't computed.
            const tx = d.axisAlignSheetTransform ?? d.sheetTransform;
            if (!tx?.from || !tx?.to) continue;

            const labels = Array.isArray(data.labels)
                ? data.labels.map((l) => String(l ?? ''))
                : [];
            const datasets = Array.isArray(data.datasets)
                ? data.datasets.map((ds) => ({
                      label: ds?.label,
                      data: Array.isArray(ds?.data) ? ds.data.map((v) => Number(v)) : [],
                  }))
                : [];

            out.push({
                chartId: typeof data.chartId === 'string' ? data.chartId : drawingId,
                sheetId: subUnitId,
                type,
                title: typeof data.title === 'string' ? data.title : '',
                sourceRange: {
                    startRow: sr.startRow,
                    endRow: sr.endRow,
                    startColumn: sr.startColumn,
                    endColumn: sr.endColumn,
                },
                ...(typeof data.sourceSheetName === 'string' && data.sourceSheetName
                    ? { sourceSheetName: data.sourceSheetName }
                    : {}),
                labels,
                datasets,
                anchor: {
                    fromCol: tx.from.column ?? 0,
                    fromColOff: tx.from.columnOffset ?? 0,
                    fromRow: tx.from.row ?? 0,
                    fromRowOff: tx.from.rowOffset ?? 0,
                    toCol: tx.to.column ?? 0,
                    toColOff: tx.to.columnOffset ?? 0,
                    toRow: tx.to.row ?? 0,
                    toRowOff: tx.to.rowOffset ?? 0,
                },
                ...(data.meta && Object.keys(data.meta).length > 0 ? { meta: data.meta } : {}),
            });
        }
    }
    return out;
}

// ─── Zip surgery (Step 5 of M10) ───────────────────────────────────────────

// Look up a sheet's 1-based index in the workbook (matches xl/worksheets/sheet{N}.xml)
// and its display name (used for <c:f> sheet-qualified ranges).
function lookupSheet(
    snapshot: UniverSnapshot,
    sheetId: string,
): { index: number; name: string } | null {
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const sheets = (snapshot as { sheets?: Record<string, { name?: string }> }).sheets ?? {};
    const idx0 = sheetOrder.indexOf(sheetId);
    if (idx0 < 0) return null;
    return { index: idx0 + 1, name: sheets[sheetId]?.name ?? `Sheet${idx0 + 1}` };
}

// Highest existing rId in a Relationships XML, or 0 if none. Used for
// allocating a new rId without colliding with existing ones (e.g. if M9
// added a rels entry for a table).
export function maxExistingRId(relsXml: string | null): number {
    if (!relsXml) return 0;
    const re = /Id="rId(\d+)"/g;
    let max = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(relsXml)) !== null) {
        const n = parseInt(m[1], 10);
        if (n > max) max = n;
    }
    return max;
}

// Add or update a Relationship in a sheet's rels XML.
// Returns the new rels XML (creates the file from scratch if missing).
export function upsertRelationship(existing: string | null, newRel: string): string {
    if (!existing) {
        return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${newRel}</Relationships>`;
    }
    // Insert just before </Relationships>. Safe on any well-formed input.
    return existing.replace(/<\/Relationships>\s*$/, `${newRel}</Relationships>`);
}

// Insert <drawing r:id="..."/> into worksheet XML at its schema-correct
// position. OOXML's CT_Worksheet element order requires drawing to come
// AFTER pageSetup/headerFooter but BEFORE tableParts. Putting <drawing>
// last (just before </worksheet>) corrupts the file when <tableParts> is
// already present (M9 adds it for named-table exports). Excel's load
// validator rejects with "We found a problem" in that case.
export function insertDrawingRefIntoSheet(sheetXml: string, rId: number): string {
    // Idempotency: if a drawing ref with the same rId already exists, skip.
    if (sheetXml.includes(`<drawing r:id="rId${rId}"/>`)) return sheetXml;
    const drawingTag = `<drawing r:id="rId${rId}"/>`;
    // Insert before <tableParts> if present.
    if (sheetXml.includes('<tableParts')) {
        return sheetXml.replace(/<tableParts\b/, `${drawingTag}<tableParts`);
    }
    // Otherwise insert just before </worksheet> (last-child slot).
    return sheetXml.replace(/<\/worksheet>\s*$/, `${drawingTag}</worksheet>`);
}

// Patch [Content_Types].xml with the four Override entries Excel needs per
// drawing + chart family. The drawing override takes a single drawingNum;
// the chart family takes the chart number + its style/colors siblings.
function patchContentTypes(
    contentTypesXml: string,
    drawingNums: number[],
    chartNums: number[],
): string {
    const toAdd: string[] = [];
    for (const n of drawingNums) {
        toAdd.push(
            `<Override PartName="/xl/drawings/drawing${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`,
        );
    }
    for (const n of chartNums) {
        toAdd.push(
            `<Override PartName="/xl/charts/chart${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`,
        );
        toAdd.push(
            `<Override PartName="/xl/charts/style${n}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>`,
        );
        toAdd.push(
            `<Override PartName="/xl/charts/colors${n}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>`,
        );
    }
    if (toAdd.length === 0) return contentTypesXml;
    return contentTypesXml.replace(/<\/Types>\s*$/, `${toAdd.join('')}</Types>`);
}

// ─── Public entry point ────────────────────────────────────────────────────

// Post-process the xlsx buffer exceljs produced. Adds chart parts for every
// 'NotesheetChart' drawing in the snapshot. On any failure, returns the
// original buffer unchanged — better to ship a chart-less but valid xlsx
// than a corrupt one. Mirrors the M9 fail-soft pattern for table import.
export async function injectChartsIntoZip(
    buffer: ArrayBuffer,
    snapshot: UniverSnapshot,
): Promise<ArrayBuffer> {
    let charts: ChartDrawing[];
    try {
        charts = readChartsFromSnapshot(snapshot);
    } catch (e) {
        console.warn('[Notesheet] readChartsFromSnapshot threw; skipping chart export', e);
        return buffer;
    }
    if (charts.length === 0) return buffer;

    try {
        const zip = await JSZip.loadAsync(buffer);

        // Group charts by sheet — one drawing.xml per sheet.
        const bySheet = new Map<string, ChartDrawing[]>();
        for (const c of charts) {
            const arr = bySheet.get(c.sheetId);
            if (arr) arr.push(c);
            else bySheet.set(c.sheetId, [c]);
        }

        let chartCounter = 1; // global chart{N} numbering across the whole workbook
        let drawingCounter = 1; // drawing{N} numbering
        const chartNumsAdded: number[] = [];
        const drawingNumsAdded: number[] = [];

        for (const [sheetId, sheetCharts] of bySheet) {
            const sheet = lookupSheet(snapshot, sheetId);
            if (!sheet) continue;

            const drawingNum = drawingCounter++;
            drawingNumsAdded.push(drawingNum);

            // Allocate the chart numbers for this sheet's charts (sequentially).
            const chartFileNumbers: number[] = [];
            for (let i = 0; i < sheetCharts.length; i++) {
                chartFileNumbers.push(chartCounter++);
            }
            chartNumsAdded.push(...chartFileNumbers);

            // 1) Patch sheet rels (or create from scratch) — add a drawing rel.
            const sheetRelsPath = `xl/worksheets/_rels/sheet${sheet.index}.xml.rels`;
            const existingRels = zip.file(sheetRelsPath);
            const relsXml = existingRels ? await existingRels.async('string') : null;
            const newRId = maxExistingRId(relsXml) + 1;
            const newRel = `<Relationship Id="rId${newRId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawingNum}.xml"/>`;
            zip.file(sheetRelsPath, upsertRelationship(relsXml, newRel));

            // 2) Patch worksheet xml — insert <drawing r:id="..."/>.
            const sheetXmlPath = `xl/worksheets/sheet${sheet.index}.xml`;
            const sheetXmlFile = zip.file(sheetXmlPath);
            if (!sheetXmlFile) continue; // unreachable in practice
            const sheetXml = await sheetXmlFile.async('string');
            zip.file(sheetXmlPath, insertDrawingRefIntoSheet(sheetXml, newRId));

            // 3) Add drawing.xml + drawing rels.
            zip.file(`xl/drawings/drawing${drawingNum}.xml`, buildDrawingXml(sheetCharts));
            zip.file(
                `xl/drawings/_rels/drawing${drawingNum}.xml.rels`,
                buildDrawingRelsXml(chartFileNumbers),
            );

            // 4) Add each chart's xml + style + colors + chart rels.
            // Cross-sheet (M17): when the chart drawing carries a
            // sourceSheetName that differs from the host sheet, the <c:f>
            // formula refs we emit must point at the DATA sheet, not the
            // chart-containing sheet — otherwise Excel re-evaluates against
            // the wrong sheet on open and the cached values silently drift.
            for (let i = 0; i < sheetCharts.length; i++) {
                const c = sheetCharts[i];
                const chartNum = chartFileNumbers[i];
                const refSheetName =
                    c.sourceSheetName && c.sourceSheetName !== sheet.name
                        ? c.sourceSheetName
                        : sheet.name;
                zip.file(
                    `xl/charts/chart${chartNum}.xml`,
                    buildChartXml(c, { sheetName: refSheetName }),
                );
                zip.file(`xl/charts/style${chartNum}.xml`, CHART_STYLE_XML);
                zip.file(`xl/charts/colors${chartNum}.xml`, CHART_COLORS_XML);
                zip.file(`xl/charts/_rels/chart${chartNum}.xml.rels`, buildChartRelsXml(chartNum));
            }
        }

        // 5) Patch [Content_Types].xml.
        const ctPath = '[Content_Types].xml';
        const ctFile = zip.file(ctPath);
        if (ctFile) {
            const ctXml = await ctFile.async('string');
            zip.file(ctPath, patchContentTypes(ctXml, drawingNumsAdded, chartNumsAdded));
        }

        const out = await zip.generateAsync({ type: 'arraybuffer' });
        return out as ArrayBuffer;
    } catch (e) {
        console.error('[Notesheet] injectChartsIntoZip failed; falling back to chart-less xlsx', e);
        return buffer;
    }
}
