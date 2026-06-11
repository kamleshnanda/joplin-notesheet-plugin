// M17: chart import from .xlsx — pure-stdlib zip+regex chart parser.
//
// Architecture:
//   xlsxBufferToSnapshot(buffer)
//     -> readChartsFromXlsxZip(buffer)   — runs BEFORE wb.xlsx.load
//         walks every sheet's drawing rels, every drawing's anchors in
//         document order, every chart{N}.xml referenced by an anchor.
//         Produces ImportedChartDrawing[] keyed by sheetIndex (1-based).
//     -> stripChartPartsFromZip(buffer)  — strips chart drawings from a
//         copy of the zip so exceljs's reconcile loop no longer crashes
//         on the `anchors` reference; non-chart workbooks pass through
//         byte-equal.
//     -> wb.xlsx.load(strippedBuffer)    — succeeds cleanly.
//     -> xlsxBufferToSnapshot continues; emits SHEET_DRAWING_PLUGIN
//       resource from the chart drawings collected above.
//
// We deliberately AVOID exceljs's chart parser. The `anchors` reconcile
// crash only ever surfaces inside exceljs's draw-side reconcile; reading
// the parts directly via JSZip + regex sidesteps it entirely. Same
// pattern as readTablesFromXlsxZip / readThemeClrScheme / readThemeFont
// (post-load readers) but inverted in call order.

import JSZip from 'jszip';

import type { ChartType } from './xlsxChart';

// One chart, after import. Keyed by sheetIndex (the 1-based exceljs
// worksheet index — matches xl/worksheets/sheet{N}.xml).
export interface ImportedChartDrawing {
    sheetIndex: number;
    chartId: string;
    type: ChartType;
    title: string;
    sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    sourceSheetName?: string;
    labels: string[];
    datasets: Array<{ label?: string; data: number[] }>;
    anchor: {
        fromCol: number; fromColOff: number;
        fromRow: number; fromRowOff: number;
        toCol: number; toColOff: number;
        toRow: number; toRowOff: number;
    };
    meta?: {
        unsupportedSourceType?: string;
        // Bar orientation. Excel's <c:barDir> attribute carries 'bar'
        // (horizontal) or 'col' (vertical). M17's ChartType union has
        // only one 'bar', so we surface direction via `meta.barDir` and
        // route NotesheetChart to set Chart.js's `options.indexAxis` at
        // render time. M10 export reads this back to emit the matching
        // <c:barDir val="..."/> (was hardcoded 'col' before round-trip
        // support landed).
        barDir?: 'bar' | 'col';
        // Bar/line grouping. Excel's <c:grouping val="..."/>:
        //   'clustered' (Excel default) — series side-by-side
        //   'stacked' — series stacked on the same X-tick
        //   'percentStacked' — stacked normalised to 100%
        //   'standard' (line charts only — equivalent to 'clustered')
        // Routed to Chart.js's `options.scales.x.stacked` /
        // `options.scales.y.stacked` for bars, and to dataset
        // `fill: 'origin' | '-1'` for stacked lines.
        barGrouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
        // Excel <c:gapWidth val="..."/> — gap between bar groups,
        // expressed as a percentage of bar width. Excel default is 150
        // (gap = 1.5x bar width). Larger values → thinner bars,
        // smaller → fatter. Routed to Chart.js dataset
        // categoryPercentage at render and re-emitted on export.
        barGapWidth?: number;
        // Legend position. Excel's <c:legendPos val="..."/> values: 'r'
        // (right, default), 'l', 't' (top), 'b' (bottom), 'tr'. Routed
        // to Chart.js options.plugins.legend.position at render time and
        // re-emitted by M10 export.
        legendPos?: 'r' | 'l' | 't' | 'b' | 'tr';
        // 'index' when the source chart had no <c:cat> element — Excel
        // shows row index 1..N as the X-axis. NotesheetChart synthesizes
        // labels and treats every column in sourceRange as a values
        // series (no separate label column). M10 export emits no
        // <c:cat> for these.
        categoryAxisType?: 'index' | 'category';
        // Doughnut hole diameter, percent of outer radius. Excel
        // default 50; user-tunable up to 90. <c:holeSize val="N"/>.
        holeSize?: number;
        // Line smoothing. <c:smooth val="0|1"/> — applied per series in
        // the source. We track only chart-level "any series smooth"
        // since Chart.js doesn't easily mix smooth + non-smooth series
        // in one chart; mixed cases collapse to "smooth" (true).
        lineSmooth?: boolean;
        // Whether the line chart shows data-point markers.
        // <c:marker val="0|1"/> at chart level. Per-series symbol info
        // is finer-grained (circle, square, diamond, ...); for now we
        // surface ON/OFF only. Chart.js per-dataset `pointRadius=0`
        // hides markers; non-zero shows them.
        lineMarkerOn?: boolean;
        // How blank cells render. <c:dispBlanksAs val="gap|zero|span"/>.
        // 'gap' (default) leaves a hole; 'zero' plots zero; 'span'
        // bridges the gap. Affects only sparse data.
        dispBlanksAs?: 'gap' | 'zero' | 'span';
        // Where the value axis crosses the category axis.
        // <c:crossBetween val="between|midCat"/>. 'between' (default)
        // crosses between ticks; 'midCat' crosses at tick centres.
        // Affects spacing of the leftmost/rightmost bars or line points.
        crossBetween?: 'between' | 'midCat';
        // <c:majorTickMark val="none|in|out|cross"/>. Excel default
        // 'none' for both axes on bar/line charts (every fixture
        // surveyed). Plumbed for fidelity round-trip.
        tickMark?: 'none' | 'in' | 'out' | 'cross';
        // Axis-line presence: 'grey' (light-grey line, Excel's modern
        // category-axis default) or 'none' (<a:ln><a:noFill/>, the
        // value-axis default). Re-emitted on export so Excel doesn't
        // fall back to its dark legacy axis line. Bar/line only.
        catAxisLine?: 'grey' | 'none';
        valAxisLine?: 'grey' | 'none';
        // Value-axis number format. <c:valAx><c:numFmt formatCode="..."
        // sourceLinked="0|1"/>. When sourceLinked="1" Excel reads the
        // format from the underlying cell; sourceLinked="0" uses the
        // explicit formatCode. We always preserve the formatCode so
        // round-trip + render are consistent.
        valAxisNumFmt?: string;
        // Chart-level data-label visibility flags from <c:dLbls>.
        // Each <c:show*> child carries val="0|1". Pie/doughnut slices
        // and bar/line series read from the chart-level dLbls when no
        // per-series override exists.
        dLbls?: {
            showVal?: boolean;
            showCatName?: boolean;
            showPercent?: boolean;
            showSerName?: boolean;
            showLegendKey?: boolean;
            showBubbleSize?: boolean;
        };
        // Per-series trendline (<c:ser><c:trendline>). Excel draws a fitted
        // line (linear/exp/log/poly/power/movingAvg) over the series, with
        // optional R²/equation labels. We support the common 'linear' case
        // for render (compute least-squares fit + draw a synthetic line);
        // others round-trip the type but render as a straight fit fallback.
        trendline?: {
            type: 'linear' | 'exp' | 'log' | 'poly' | 'power' | 'movingAvg';
            order?: number;   // poly order
            period?: number;  // movingAvg period
            dispRSqr?: boolean;
            dispEq?: boolean;
        };
    };
}

// ─── XML helpers (small, regex-only — no DOMParser) ────────────────────────

// Strip XML element text decoding for the limited entity set OOXML uses
// (&amp; &lt; &gt; &apos; &quot; plus numeric refs).
function decodeXmlEntities(s: string): string {
    return s
        .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
        .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)))
        .replace(/&apos;/g, "'")
        .replace(/&quot;/g, '"')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&amp;/g, '&');
}

function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ─── Anchor walker — drawing{N}.xml ────────────────────────────────────────

interface RawAnchor {
    rId: string;
    kind: 'twoCell' | 'oneCell' | 'absolute';
    from: { col: number; colOff: number; row: number; rowOff: number };
    to: { col: number; colOff: number; row: number; rowOff: number };
}

// Walks the three anchor element types in DOCUMENT ORDER. The 06 fixture
// packs both charts into ONE drawing1.xml as two consecutive
// <xdr:twoCellAnchor> blocks; walking by zip-key order would mis-align.
//
// Drawing XML can use either the canonical `xdr:` namespace prefix
// (Excel-authored) or a default namespace with NO prefix (some
// programmatically-generated workbooks like the project's MultiSheet.xlsx
// fixture). Match both shapes via an optional `(?:xdr:)?` prefix.
function walkAnchors(drawingXml: string): RawAnchor[] {
    const out: RawAnchor[] = [];
    const re = /<(?:xdr:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/(?:xdr:)?\1>/g;
    let match: RegExpExecArray | null;
    while ((match = re.exec(drawingXml)) !== null) {
        const kindRaw = match[1];
        const body = match[2];
        // Chart graphicFrame ref. The element may be `<c:chart>` or `<chart>`
        // depending on whether the chart namespace is the default; the
        // r:id attribute may be unprefixed in workbooks that omit the
        // standard `r` namespace alias.
        const chartRefMatch =
            body.match(/<(?:c:)?chart\b[^>]*\sr:id="([^"]+)"/) ??
            body.match(/<(?:c:)?chart\b[^>]*\sid="([^"]+)"/);
        if (!chartRefMatch) continue;
        const rId = chartRefMatch[1];

        const fromBody = body.match(/<(?:xdr:)?from>([\s\S]*?)<\/(?:xdr:)?from>/);
        if (!fromBody) continue;

        const from = parseAnchorPoint(fromBody[1]);
        if (!from) continue;

        let to: RawAnchor['to'];
        const toBody = body.match(/<(?:xdr:)?to>([\s\S]*?)<\/(?:xdr:)?to>/);
        if (toBody) {
            const parsed = parseAnchorPoint(toBody[1]);
            if (!parsed) continue;
            to = parsed;
        } else {
            // oneCellAnchor / absoluteAnchor: synthesize a `to` ~6 cells /
            // 14 rows past `from`. Excel renders these by laying down the
            // EMU ext at the from point; for our cell-anchored UI a
            // reasonable approximation is from + a chart-sized span.
            to = {
                col: from.col + 6,
                colOff: 0,
                row: from.row + 14,
                rowOff: 0,
            };
        }

        const kind: RawAnchor['kind'] =
            kindRaw === 'twoCellAnchor' ? 'twoCell'
            : kindRaw === 'oneCellAnchor' ? 'oneCell'
            : 'absolute';

        out.push({ rId, kind, from, to });
    }
    return out;
}

function parseAnchorPoint(body: string): { col: number; colOff: number; row: number; rowOff: number } | null {
    const col = body.match(/<(?:xdr:)?col>(\d+)<\/(?:xdr:)?col>/);
    const colOff = body.match(/<(?:xdr:)?colOff>(-?\d+)<\/(?:xdr:)?colOff>/);
    const row = body.match(/<(?:xdr:)?row>(\d+)<\/(?:xdr:)?row>/);
    const rowOff = body.match(/<(?:xdr:)?rowOff>(-?\d+)<\/(?:xdr:)?rowOff>/);
    if (!col || !row) return null;
    return {
        col: parseInt(col[1], 10),
        colOff: colOff ? parseInt(colOff[1], 10) : 0,
        row: parseInt(row[1], 10),
        rowOff: rowOff ? parseInt(rowOff[1], 10) : 0,
    };
}

// ─── Drawing rels — map rId → chart part path ─────────────────────────────

// Walks every <Relationship .../> element. Attribute strings can contain
// '/' (Type is a URL like .../officeDocument/.../drawing), so we match
// each Relationship terminator as the end of a self-closing tag — the
// element body is "everything up to the closing slash + greater-than"
// where the closing pair is preceded by a quote (the last attr's value
// terminator).
function parseDrawingRels(relsXml: string): Map<string, string> {
    const out = new Map<string, string>();
    const re = /<Relationship\s+([^>]*?)\/>/g;
    let match: RegExpExecArray | null;
    while ((match = re.exec(relsXml)) !== null) {
        const attrs = match[1];
        const idMatch = attrs.match(/\bId="([^"]+)"/);
        const targetMatch = attrs.match(/\bTarget="([^"]+)"/);
        if (idMatch && targetMatch) {
            out.set(idMatch[1], targetMatch[1]);
        }
    }
    return out;
}

// Resolve a relative target like "../charts/chart1.xml" against the
// drawing's own folder ("xl/drawings"). Returns the absolute zip key.
// Targets may also be absolute ("/xl/charts/chart1.xml" with a leading
// slash — programmatically-built fixtures use this shape) or
// already-relative-without-dotdot.
function resolveRelTarget(drawingFolder: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    const segs = drawingFolder.split('/').filter(Boolean);
    const targetSegs = target.split('/');
    for (const seg of targetSegs) {
        if (seg === '..') segs.pop();
        else if (seg !== '.' && seg !== '') segs.push(seg);
    }
    return segs.join('/');
}

// ─── Cell-ref decoder ─────────────────────────────────────────────────────

export interface DecodedRef {
    sheetName: string;
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}

// Decode a formula like "Sheet1!$A$2:$A$5" or "'My Sheet'!$A$2:$B$5"
// into a 0-based DecodedRef. Returns null for whole-column $A:$A,
// expressions, names, etc.
export function decodeCellRef(formula: string): DecodedRef | null {
    if (!formula) return null;
    const sheetMatch = formula.match(/^(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_]*))!(.+)$/);
    if (!sheetMatch) return null;
    const sheetName = (sheetMatch[1] ? sheetMatch[1].replace(/''/g, "'") : sheetMatch[2]);
    const rangePart = sheetMatch[3];
    const rangeRe = /^\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?$/;
    const m = rangePart.match(rangeRe);
    if (!m) return null;
    const startCol = colIndex(m[1]);
    const startRow = parseInt(m[2], 10) - 1;
    const endCol = m[3] ? colIndex(m[3]) : startCol;
    const endRow = m[4] ? parseInt(m[4], 10) - 1 : startRow;
    if (startCol < 0 || endCol < 0 || startRow < 0 || endRow < 0) return null;
    return {
        sheetName,
        startRow: Math.min(startRow, endRow),
        endRow: Math.max(startRow, endRow),
        startColumn: Math.min(startCol, endCol),
        endColumn: Math.max(startCol, endCol),
    };
}

function colIndex(letters: string): number {
    let n = 0;
    const upper = letters.toUpperCase();
    for (let i = 0; i < upper.length; i++) {
        const c = upper.charCodeAt(i);
        if (c < 65 || c > 90) return -1;
        n = n * 26 + (c - 64);
    }
    return n - 1;
}

// ─── Chart XML walker ─────────────────────────────────────────────────────

interface ParsedChart {
    type: ChartType;
    title: string;
    sourceSheetName?: string;
    sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    labels: string[];
    datasets: Array<{ label?: string; data: number[] }>;
    unsupportedSourceType?: string;
    barDir?: 'bar' | 'col';
    barGrouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard';
    barGapWidth?: number;
    legendPos?: 'r' | 'l' | 't' | 'b' | 'tr';
    categoryAxisType?: 'index' | 'category';
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
}

const SUPPORTED_TYPES: Record<string, ChartType> = {
    barChart: 'bar',
    lineChart: 'line',
    pieChart: 'pie',
    doughnutChart: 'doughnut',
};

// Match an element name with or without a `c:` namespace prefix. Some
// programmatically-built workbooks (e.g. MultiSheet.xlsx) declare the
// chart namespace as the default namespace and emit unprefixed element
// names; Excel-authored fixtures emit the canonical `c:` prefix.
const C = '(?:c:)?';

export function parseChartXml(chartXml: string): ParsedChart | null {
    // Chart-type element name. ECMA-376 chart-types live as direct
    // children of <c:plotArea>.
    const typeMatch = chartXml.match(/<(?:c:)?(bar|line|pie|doughnut|radar|scatter|area|bubble|surface|stock|ofPie|bar3D|line3D|pie3D|area3D|surface3D)Chart\b/);
    let type: ChartType = 'bar';
    let unsupportedSourceType: string | undefined;
    if (typeMatch) {
        const elemKey = `${typeMatch[1]}Chart`;
        if (elemKey in SUPPORTED_TYPES) {
            type = SUPPORTED_TYPES[elemKey];
        } else {
            unsupportedSourceType = typeMatch[1];
            type = 'bar';
        }
    }

    // Bar orientation. Only meaningful when type === 'bar'. ECMA-376
    // <c:barDir val="bar"/> = horizontal, val="col" = vertical (the
    // default Excel emits when the user picks "Column Chart"). When
    // absent, treat as 'col' since Excel's default chart-of-type-bar
    // is what Notesheet renders today.
    let barDir: 'bar' | 'col' | undefined;
    if (type === 'bar') {
        const barDirMatch = chartXml.match(/<(?:c:)?barDir\s+val="(bar|col)"/);
        barDir = barDirMatch ? (barDirMatch[1] as 'bar' | 'col') : 'col';
    }

    // Grouping. <c:grouping val="..."/>: 'clustered' (default for bar),
    // 'stacked', 'percentStacked', 'standard' (line charts).
    // Pie/doughnut have no grouping element; leave undefined.
    let barGrouping: 'clustered' | 'stacked' | 'percentStacked' | 'standard' | undefined;
    if (type === 'bar' || type === 'line') {
        const groupingMatch = chartXml.match(/<(?:c:)?grouping\s+val="(clustered|stacked|percentStacked|standard)"/);
        if (groupingMatch) barGrouping = groupingMatch[1] as 'clustered' | 'stacked' | 'percentStacked' | 'standard';
    }

    // Gap width. <c:gapWidth val="N"/> where N is a percentage of bar
    // width. Bar charts only.
    let barGapWidth: number | undefined;
    if (type === 'bar') {
        const gapMatch = chartXml.match(/<(?:c:)?gapWidth\s+val="(\d+)"/);
        if (gapMatch) {
            const n = parseInt(gapMatch[1], 10);
            if (Number.isFinite(n)) barGapWidth = n;
        }
    }

    // Doughnut hole size. <c:holeSize val="N"/> — percent of outer
    // radius (Excel default 50; user range 1-90).
    let holeSize: number | undefined;
    if (type === 'doughnut') {
        const holeMatch = chartXml.match(/<(?:c:)?holeSize\s+val="(\d+)"/);
        if (holeMatch) {
            const n = parseInt(holeMatch[1], 10);
            if (Number.isFinite(n)) holeSize = n;
        }
    }

    // Line smoothing. <c:smooth val="0|1"/> — appears per-series. We
    // record "any series smoothed" as the chart-level toggle; Chart.js
    // applies smoothing per dataset via `tension`.
    let lineSmooth: boolean | undefined;
    if (type === 'line') {
        const smoothMatches = [...chartXml.matchAll(/<(?:c:)?smooth\s+val="([01])"/g)];
        if (smoothMatches.length > 0) {
            lineSmooth = smoothMatches.some((m) => m[1] === '1');
        }
    }

    // Line marker on/off. Chart-level <c:marker val="0|1"/> applies to
    // every series unless a per-series <c:marker><c:symbol val="..."/>
    // overrides. For M17 we read the chart-level toggle only; per-series
    // symbol shapes are out of scope.
    let lineMarkerOn: boolean | undefined;
    if (type === 'line') {
        const markerMatch = chartXml.match(/<(?:c:)?marker\s+val="([01])"/);
        if (markerMatch) {
            lineMarkerOn = markerMatch[1] === '1';
        }
    }

    // Disp-blanks-as. <c:dispBlanksAs val="gap|zero|span"/>. Affects
    // how blank cells render across all chart types.
    let dispBlanksAs: 'gap' | 'zero' | 'span' | undefined;
    const dispBlanksMatch = chartXml.match(/<(?:c:)?dispBlanksAs\s+val="(gap|zero|span)"/);
    if (dispBlanksMatch) {
        dispBlanksAs = dispBlanksMatch[1] as 'gap' | 'zero' | 'span';
    }

    // Cross-between. <c:crossBetween val="between|midCat"/> on the
    // value axis. Bar/line only; pie/doughnut have no axes.
    let crossBetween: 'between' | 'midCat' | undefined;
    if (type === 'bar' || type === 'line') {
        const crossMatch = chartXml.match(/<(?:c:)?crossBetween\s+val="(between|midCat)"/);
        if (crossMatch) {
            crossBetween = crossMatch[1] as 'between' | 'midCat';
        }
    }

    // Tick marks. We sample <c:majorTickMark> from EITHER axis (they
    // typically match in source workbooks). Bar/line only.
    let tickMark: 'none' | 'in' | 'out' | 'cross' | undefined;
    if (type === 'bar' || type === 'line') {
        const tickMatch = chartXml.match(/<(?:c:)?majorTickMark\s+val="(none|in|out|cross)"/);
        if (tickMatch) {
            tickMark = tickMatch[1] as 'none' | 'in' | 'out' | 'cross';
        }
    }

    // Value-axis number format. We scope to the <c:valAx> block to avoid
    // grabbing the catAx's General format. Bar/line only.
    let valAxisNumFmt: string | undefined;
    if (type === 'bar' || type === 'line') {
        const valAxBlock = chartXml.match(/<(?:c:)?valAx>([\s\S]*?)<\/(?:c:)?valAx>/);
        if (valAxBlock) {
            const numFmtMatch = valAxBlock[1].match(/<(?:c:)?numFmt\s+formatCode="([^"]+)"/);
            if (numFmtMatch) valAxisNumFmt = numFmtMatch[1];
        }
    }

    // Axis-line presence. Excel's modern chart template draws the
    // category axis as a thin light-grey line and the value axis with NO
    // line. Older/custom charts may differ. We classify each axis's
    // OWN <c:spPr><a:ln> (the one NOT nested in <c:majorGridlines>) as
    // 'none' (<a:ln><a:noFill/>) or 'grey' (any visible line). M10 export
    // re-emits the matching style so the round-trip doesn't swap Excel's
    // intended light line for its dark legacy default. Bar/line only;
    // pie/doughnut have no axes.
    const classifyAxisLine = (axTag: 'catAx' | 'valAx'): 'grey' | 'none' | undefined => {
        const block = chartXml.match(new RegExp(`<(?:c:)?${axTag}>([\\s\\S]*?)</(?:c:)?${axTag}>`));
        if (!block) return undefined;
        // Drop the gridlines block so we read only the axis-line spPr.
        const body = block[1].replace(/<(?:c:)?majorGridlines>[\s\S]*?<\/(?:c:)?majorGridlines>/g, '');
        const spPr = body.match(/<(?:c:)?spPr>([\s\S]*?)<\/(?:c:)?spPr>/);
        if (!spPr) return undefined;
        // A no-line axis carries <a:ln><a:noFill/></a:ln>.
        if (/<a:ln>\s*<a:noFill\/>\s*<\/a:ln>/.test(spPr[1])) return 'none';
        if (/<a:ln\b/.test(spPr[1])) return 'grey';
        return undefined;
    };
    let catAxisLine: 'grey' | 'none' | undefined;
    let valAxisLine: 'grey' | 'none' | undefined;
    if (type === 'bar' || type === 'line') {
        catAxisLine = classifyAxisLine('catAx');
        valAxisLine = classifyAxisLine('valAx');
    }

    // Chart-level data-label flags from <c:dLbls>. Each <c:show*>
    // child carries val="0|1". Pie/doughnut especially rely on these
    // for slice labels; missing dLbls in our export is why 03/06-chart-2
    // exported without slice labels.
    let dLbls: ParsedChart['dLbls'] | undefined;
    // Data-label visibility lives in TWO possible places. Excel often
    // writes a chart-level <c:dLbls> with everything OFF *and* a
    // SERIES-level <c:ser><c:dLbls> that turns the actual labels on
    // (this is exactly how 03-pie-single / 06-chart-2 ship their pie
    // slice labels: chart-level all-zero, series-level showCatName=1 +
    // showPercent=1). Reading only the chart-level block — as we used to
    // — imported "no labels", so the re-exported pie had no slice labels
    // (issues 3 / 6). We now parse a dLbls body into flags, read the
    // chart-level block first, and FALL BACK to the first series-level
    // block when the chart-level one is absent or shows nothing.
    const parseDLblsBody = (body: string): ParsedChart['dLbls'] | undefined => {
        const flag = (name: string): boolean | undefined => {
            const m = body.match(new RegExp(`<(?:c:)?${name}\\s+val="([01])"`));
            return m ? m[1] === '1' : undefined;
        };
        const out = {
            ...(flag('showVal') !== undefined ? { showVal: flag('showVal') } : {}),
            ...(flag('showCatName') !== undefined ? { showCatName: flag('showCatName') } : {}),
            ...(flag('showPercent') !== undefined ? { showPercent: flag('showPercent') } : {}),
            ...(flag('showSerName') !== undefined ? { showSerName: flag('showSerName') } : {}),
            ...(flag('showLegendKey') !== undefined ? { showLegendKey: flag('showLegendKey') } : {}),
            ...(flag('showBubbleSize') !== undefined ? { showBubbleSize: flag('showBubbleSize') } : {}),
        };
        return Object.keys(out).length > 0 ? out : undefined;
    };
    // "Shows something" = at least one label-content flag is true. Used to
    // decide whether the chart-level block is meaningful or we should
    // prefer the series-level one.
    const dLblsShowsSomething = (d: ParsedChart['dLbls'] | undefined): boolean =>
        !!d && !!(d.showVal || d.showCatName || d.showPercent || d.showSerName);
    {
        // Chart-level: strip every <c:ser>...</c:ser> so we read only the
        // block OUTSIDE the series.
        const stripped = chartXml.replace(/<(?:c:)?ser\b[\s\S]*?<\/(?:c:)?ser>/g, '');
        const chartLevelBlock = stripped.match(/<(?:c:)?dLbls>([\s\S]*?)<\/(?:c:)?dLbls>/);
        const chartLevel = chartLevelBlock ? parseDLblsBody(chartLevelBlock[1]) : undefined;

        // Series-level: the FIRST <c:ser>'s own <c:dLbls> (pie/doughnut
        // have one series; bar/line share the per-series toggle). We scope
        // the search to inside a <c:ser> block to avoid re-grabbing the
        // chart-level one.
        const serBlock = chartXml.match(/<(?:c:)?ser\b[\s\S]*?<\/(?:c:)?ser>/);
        const serDLblsBlock = serBlock ? serBlock[0].match(/<(?:c:)?dLbls>([\s\S]*?)<\/(?:c:)?dLbls>/) : null;
        const seriesLevel = serDLblsBlock ? parseDLblsBody(serDLblsBlock[1]) : undefined;

        // Prefer whichever actually turns a label on. If the chart-level
        // block shows something, keep it; otherwise fall back to the
        // series-level block when IT shows something; else keep whatever
        // chart-level had (preserves explicit all-off intent).
        if (dLblsShowsSomething(chartLevel)) {
            dLbls = chartLevel;
        } else if (dLblsShowsSomething(seriesLevel)) {
            dLbls = seriesLevel;
        } else {
            dLbls = chartLevel;
        }
    }

    // Trendline (<c:ser><c:trendline>). Per-series; we read the FIRST
    // series' trendline (the common single-trendline case — 10-bar-with-
    // trendline). type defaults to 'linear' (Excel's default and the only
    // one we render exactly; others round-trip + render as a straight fit).
    let trendline: ParsedChart['trendline'] | undefined;
    {
        const tlBlock = chartXml.match(/<(?:c:)?trendline>([\s\S]*?)<\/(?:c:)?trendline>/);
        if (tlBlock) {
            const body = tlBlock[1];
            const typeMatch = body.match(/<(?:c:)?trendlineType\s+val="(linear|exp|log|poly|power|movingAvg)"/);
            const orderMatch = body.match(/<(?:c:)?order\s+val="(\d+)"/);
            const periodMatch = body.match(/<(?:c:)?period\s+val="(\d+)"/);
            const rsqr = /<(?:c:)?dispRSqr\s+val="1"/.test(body);
            const eq = /<(?:c:)?dispEq\s+val="1"/.test(body);
            trendline = {
                type: (typeMatch ? typeMatch[1] : 'linear') as NonNullable<ParsedChart['trendline']>['type'],
                ...(orderMatch ? { order: parseInt(orderMatch[1], 10) } : {}),
                ...(periodMatch ? { period: parseInt(periodMatch[1], 10) } : {}),
                ...(rsqr ? { dispRSqr: true } : {}),
                ...(eq ? { dispEq: true } : {}),
            };
        }
    }

    // Title. Excel's <c:title><c:tx><c:rich>...</c:rich></c:tx></c:title>
    // can contain multiple <a:r><a:t>...</a:t></a:r> runs (one per
    // formatting variation — e.g. "Investment" + " vs Balance so far"
    // each as its own run). The displayed title is all runs
    // concatenated. Earlier we matched only the first <a:t>, which
    // truncated multi-run titles. Walk every <a:t> inside the title
    // block and join them.
    let title = '';
    const titleBlock = chartXml.match(/<(?:c:)?title\b[\s\S]*?<\/(?:c:)?title>/);
    if (titleBlock) {
        const runs = [...titleBlock[0].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => decodeXmlEntities(m[1]));
        title = runs.join('');
    }

    // Legend position. <c:legend><c:legendPos val="b|l|r|t|tr"/></c:legend>.
    // When the legend element is absent, leave undefined (NotesheetChart
    // applies its own default visibility based on series count).
    let legendPos: 'r' | 'l' | 't' | 'b' | 'tr' | undefined;
    const legendPosMatch = chartXml.match(/<(?:c:)?legend\b[\s\S]*?<(?:c:)?legendPos\s+val="(r|l|t|b|tr)"/);
    if (legendPosMatch) legendPos = legendPosMatch[1] as 'r' | 'l' | 't' | 'b' | 'tr';

    // Walk every <c:ser>...</c:ser> in document order. Element name may
    // be `c:ser` or just `ser` depending on namespace shape.
    const serRe = new RegExp(`<${C}ser\\b[^>]*>([\\s\\S]*?)</${C}ser>`, 'g');
    const serBodies: string[] = [];
    let serMatch: RegExpExecArray | null;
    while ((serMatch = serRe.exec(chartXml)) !== null) {
        serBodies.push(serMatch[1]);
    }
    if (serBodies.length === 0) return null;

    let labels: string[] = [];
    let sourceSheetName: string | undefined;
    let labelsRange: DecodedRef | null = null;
    const valRanges: DecodedRef[] = [];
    const datasets: Array<{ label?: string; data: number[] }> = [];
    // Track whether any series provided a <c:cat>. When NONE do, Excel
    // shows row index 1..N as the X-axis and treats every column in
    // the value-refs as a separate values series. We surface this
    // distinction via meta.categoryAxisType so M10 export can re-emit
    // a no-<c:cat> chart and NotesheetChart can render synthetic 1..N
    // labels rather than misinterpreting column 0 as a label column.
    let anyCatRef = false;

    const reTx = new RegExp(`<${C}tx>([\\s\\S]*?)</${C}tx>`);
    const reStrCacheV = new RegExp(`<${C}strCache>[\\s\\S]*?<${C}pt\\b[^>]*>[\\s\\S]*?<${C}v>([\\s\\S]*?)</${C}v>`);
    const reInlineV = new RegExp(`<${C}v>([\\s\\S]*?)</${C}v>`);
    const reCat = new RegExp(`<${C}cat>([\\s\\S]*?)</${C}cat>`);
    const reF = new RegExp(`<${C}f>([\\s\\S]*?)</${C}f>`);
    const reVal = new RegExp(`<${C}val>([\\s\\S]*?)</${C}val>`);
    const reNumCache = new RegExp(`<${C}numCache>([\\s\\S]*?)</${C}numCache>`);

    for (const ser of serBodies) {
        // Series label.
        const txMatch = ser.match(reTx);
        let label: string | undefined;
        if (txMatch) {
            const strCacheVal = txMatch[1].match(reStrCacheV);
            const inlineV = !strCacheVal && txMatch[1].match(reInlineV);
            if (strCacheVal) label = decodeXmlEntities(strCacheVal[1]);
            else if (inlineV) label = decodeXmlEntities(inlineV[1]);
        }

        // Categories — only consume if the series has a <c:cat>.
        const catMatch = ser.match(reCat);
        if (catMatch) {
            anyCatRef = true;
            const catBody = catMatch[1];
            const catFormula = catBody.match(reF);
            if (catFormula) {
                const catRef = decodeCellRef(catFormula[1]);
                if (catRef) {
                    if (!labelsRange) labelsRange = catRef;
                    if (!sourceSheetName) sourceSheetName = catRef.sheetName;
                }
            }
            if (labels.length === 0) {
                labels = readPtCache(catBody);
            }
        }

        // Values — every series contributes a dataset.
        const valMatch = ser.match(reVal);
        const data: number[] = [];
        if (valMatch) {
            const valBody = valMatch[1];
            const valFormula = valBody.match(reF);
            if (valFormula) {
                const valRef = decodeCellRef(valFormula[1]);
                if (valRef) {
                    valRanges.push(valRef);
                    if (!sourceSheetName) sourceSheetName = valRef.sheetName;
                }
            }
            const numCacheBody = valBody.match(reNumCache);
            if (numCacheBody) {
                for (const v of readPtCache(numCacheBody[1])) {
                    const n = Number(v);
                    data.push(Number.isFinite(n) ? n : 0);
                }
            }
        }
        datasets.push({ label, data });
    }

    if (datasets.length === 0) return null;

    // Compose source range as the bounding box of cat-ref + val-refs.
    const refs: DecodedRef[] = [];
    if (labelsRange) refs.push(labelsRange);
    refs.push(...valRanges);
    if (refs.length === 0) {
        // No usable formulas — shape too irregular (numeric-only scatter
        // etc.). caller drops with a warn.
        return null;
    }

    let startRow = refs[0].startRow;
    let endRow = refs[0].endRow;
    let startColumn = refs[0].startColumn;
    let endColumn = refs[0].endColumn;
    // The series-name (header) cell sits at the row above data, so we
    // include that row in the bounding box.
    if (labelsRange) {
        startRow = Math.min(startRow, labelsRange.startRow - 1);
    }
    for (const r of refs) {
        if (r.startRow < startRow) startRow = r.startRow;
        if (r.endRow > endRow) endRow = r.endRow;
        if (r.startColumn < startColumn) startColumn = r.startColumn;
        if (r.endColumn > endColumn) endColumn = r.endColumn;
    }
    if (startRow < 0) startRow = 0;

    return {
        type,
        title,
        sourceSheetName,
        sourceRange: { startRow, endRow, startColumn, endColumn },
        labels,
        datasets,
        unsupportedSourceType,
        ...(barDir ? { barDir } : {}),
        ...(barGrouping ? { barGrouping } : {}),
        ...(barGapWidth !== undefined ? { barGapWidth } : {}),
        ...(legendPos ? { legendPos } : {}),
        categoryAxisType: anyCatRef ? 'category' : 'index',
        ...(holeSize !== undefined ? { holeSize } : {}),
        ...(lineSmooth !== undefined ? { lineSmooth } : {}),
        ...(lineMarkerOn !== undefined ? { lineMarkerOn } : {}),
        ...(dispBlanksAs ? { dispBlanksAs } : {}),
        ...(crossBetween ? { crossBetween } : {}),
        ...(tickMark ? { tickMark } : {}),
        ...(catAxisLine ? { catAxisLine } : {}),
        ...(valAxisLine ? { valAxisLine } : {}),
        ...(valAxisNumFmt ? { valAxisNumFmt } : {}),
        ...(dLbls ? { dLbls } : {}),
        ...(trendline ? { trendline } : {}),
    };
}

// Read a series of <c:pt idx="..."><c:v>...</c:v></c:pt> from a cache body.
function readPtCache(cacheBody: string): string[] {
    const re = /<(?:c:)?pt\b[^>]*idx="(\d+)"[^>]*>[\s\S]*?<(?:c:)?v>([\s\S]*?)<\/(?:c:)?v>[\s\S]*?<\/(?:c:)?pt>/g;
    const map = new Map<number, string>();
    let match: RegExpExecArray | null;
    let max = -1;
    while ((match = re.exec(cacheBody)) !== null) {
        const idx = parseInt(match[1], 10);
        const v = decodeXmlEntities(match[2]);
        map.set(idx, v);
        if (idx > max) max = idx;
    }
    const out: string[] = [];
    for (let i = 0; i <= max; i++) {
        out.push(map.get(i) ?? '');
    }
    return out;
}

// ─── Sheet → drawing path map ─────────────────────────────────────────────

interface SheetDrawingLink {
    sheetIndex: number;
    drawingPath: string;
}

async function findSheetDrawingLinks(zip: JSZip): Promise<SheetDrawingLink[]> {
    const out: SheetDrawingLink[] = [];
    const sheetRelsRe = /^xl\/worksheets\/_rels\/sheet(\d+)\.xml\.rels$/;
    for (const path of Object.keys(zip.files)) {
        const m = path.match(sheetRelsRe);
        if (!m) continue;
        const sheetIndex = parseInt(m[1], 10);
        const xml = await zip.files[path].async('string');
        // Walk every <Relationship .../> in this rels file, parse attrs
        // independently, and pick the one whose Type ends with "/drawing".
        // Attribute order within a Relationship element is not fixed across
        // workbook authors — a programmatic emitter may put Type before
        // Target before Id (MultiSheet.xlsx) while Excel emits Id-first.
        const relRe = /<Relationship\s+([^>]*?)\/>/g;
        let mr: RegExpExecArray | null;
        while ((mr = relRe.exec(xml)) !== null) {
            const attrs = mr[1];
            const typeMatch = attrs.match(/\bType="([^"]+)"/);
            const targetMatch = attrs.match(/\bTarget="([^"]+)"/);
            if (!typeMatch || !targetMatch) continue;
            if (!/\/drawing$/.test(typeMatch[1])) continue;
            const drawingPath = resolveRelTarget('xl/worksheets', targetMatch[1]);
            out.push({ sheetIndex, drawingPath });
            break;
        }
    }
    return out;
}

// ─── Entry point: read all chart drawings ─────────────────────────────────

export async function readChartsFromXlsxZip(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<ImportedChartDrawing[]> {
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return [];
    }

    const sheetLinks = await findSheetDrawingLinks(zip);
    if (sheetLinks.length === 0) return [];

    const out: ImportedChartDrawing[] = [];
    for (const link of sheetLinks) {
        const drawingFile = zip.file(link.drawingPath);
        if (!drawingFile) continue;
        const drawingXml = await drawingFile.async('string');

        const drawingFolder = link.drawingPath.split('/').slice(0, -1).join('/');
        const drawingFilename = link.drawingPath.split('/').pop() ?? '';
        const drawingRelsPath = `${drawingFolder}/_rels/${drawingFilename}.rels`;
        const drawingRelsFile = zip.file(drawingRelsPath);
        if (!drawingRelsFile) continue;
        const drawingRelsXml = await drawingRelsFile.async('string');
        const relMap = parseDrawingRels(drawingRelsXml);

        const anchors = walkAnchors(drawingXml);
        let anchorIndex = 0;
        for (const anchor of anchors) {
            anchorIndex++;
            const target = relMap.get(anchor.rId);
            if (!target) {
                console.warn(`[Notesheet] M17: drawing ${link.drawingPath} anchor ${anchorIndex} rId ${anchor.rId} has no rels target — dropping`);
                continue;
            }
            const chartPath = resolveRelTarget(drawingFolder, target);
            const chartFile = zip.file(chartPath);
            if (!chartFile) {
                console.warn(`[Notesheet] M17: drawing ${link.drawingPath} anchor ${anchorIndex} -> missing chart part ${chartPath} — dropping`);
                continue;
            }
            const chartXml = await chartFile.async('string');
            const parsed = parseChartXml(chartXml);
            if (!parsed) {
                console.warn(`[Notesheet] M17: chart ${chartPath} could not be parsed (no series or no usable refs) — dropping`);
                continue;
            }

            const chartId = `chart-imported-${link.sheetIndex}-${anchorIndex}-${chartPath.replace(/[^a-zA-Z0-9]/g, '_')}`;
            const drawing: ImportedChartDrawing = {
                sheetIndex: link.sheetIndex,
                chartId,
                type: parsed.type,
                title: parsed.title,
                sourceRange: parsed.sourceRange,
                labels: parsed.labels,
                datasets: parsed.datasets,
                anchor: {
                    fromCol: anchor.from.col,
                    fromColOff: anchor.from.colOff,
                    fromRow: anchor.from.row,
                    fromRowOff: anchor.from.rowOff,
                    toCol: anchor.to.col,
                    toColOff: anchor.to.colOff,
                    toRow: anchor.to.row,
                    toRowOff: anchor.to.rowOff,
                },
            };
            if (parsed.sourceSheetName) drawing.sourceSheetName = parsed.sourceSheetName;
            const hasMeta = parsed.unsupportedSourceType
                || parsed.barDir
                || parsed.barGrouping
                || parsed.barGapWidth !== undefined
                || parsed.legendPos
                || parsed.categoryAxisType
                || parsed.holeSize !== undefined
                || parsed.lineSmooth !== undefined
                || parsed.lineMarkerOn !== undefined
                || parsed.dispBlanksAs
                || parsed.crossBetween
                || parsed.tickMark
                || parsed.catAxisLine
                || parsed.valAxisLine
                || parsed.valAxisNumFmt
                || parsed.dLbls
                || parsed.trendline;
            if (hasMeta) {
                drawing.meta = {
                    ...(parsed.unsupportedSourceType ? { unsupportedSourceType: parsed.unsupportedSourceType } : {}),
                    ...(parsed.barDir ? { barDir: parsed.barDir } : {}),
                    ...(parsed.barGrouping ? { barGrouping: parsed.barGrouping } : {}),
                    ...(parsed.barGapWidth !== undefined ? { barGapWidth: parsed.barGapWidth } : {}),
                    ...(parsed.legendPos ? { legendPos: parsed.legendPos } : {}),
                    ...(parsed.categoryAxisType ? { categoryAxisType: parsed.categoryAxisType } : {}),
                    ...(parsed.holeSize !== undefined ? { holeSize: parsed.holeSize } : {}),
                    ...(parsed.lineSmooth !== undefined ? { lineSmooth: parsed.lineSmooth } : {}),
                    ...(parsed.lineMarkerOn !== undefined ? { lineMarkerOn: parsed.lineMarkerOn } : {}),
                    ...(parsed.dispBlanksAs ? { dispBlanksAs: parsed.dispBlanksAs } : {}),
                    ...(parsed.crossBetween ? { crossBetween: parsed.crossBetween } : {}),
                    ...(parsed.tickMark ? { tickMark: parsed.tickMark } : {}),
                    ...(parsed.catAxisLine ? { catAxisLine: parsed.catAxisLine } : {}),
                    ...(parsed.valAxisLine ? { valAxisLine: parsed.valAxisLine } : {}),
                    ...(parsed.valAxisNumFmt ? { valAxisNumFmt: parsed.valAxisNumFmt } : {}),
                    ...(parsed.dLbls ? { dLbls: parsed.dLbls } : {}),
                    ...(parsed.trendline ? { trendline: parsed.trendline } : {}),
                };
                if (parsed.unsupportedSourceType) {
                    console.warn(`[Notesheet] M17: chart type '${parsed.unsupportedSourceType}' is not supported; falling back to 'bar'`);
                }
            }
            out.push(drawing);
        }
    }
    return out;
}

// ─── Strip chart drawings — produce a buffer exceljs can load ─────────────

// exceljs's reconcile loop crashes on workbooks with chart drawings (the
// `anchors` reference). We pre-strip every chart-drawing-related part so
// the buffer that exceljs actually loads has no charts at all. The
// snapshot's chart resource is built from the ORIGINAL buffer's chart
// parts via readChartsFromXlsxZip; this stripped buffer is only used to
// drive exceljs's load.
//
// What gets stripped (idempotent on chart-less buffers):
//   * xl/drawings/drawing*.xml whose rels point at any chart part.
//   * xl/drawings/_rels/drawing*.xml.rels for the same drawings.
//   * xl/charts/* — every chart, style, colors part and their rels.
//   * <drawing r:id="..."/> elements + matching <Relationship> in the
//     sheet's rels for any sheet that referenced a chart-bearing drawing.
//   * Override entries in [Content_Types].xml for every part removed.
//
// Non-chart drawings (images, shapes) inside drawing*.xml aren't yet
// stripped; the M17 fixture set only carries chart drawings, and exceljs
// loads image-only drawings without crashing.
export async function stripChartPartsFromZip(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<ArrayBuffer> {
    const zip = await JSZip.loadAsync(buffer as ArrayBuffer);

    const chartBearingDrawings = new Set<string>();
    for (const path of Object.keys(zip.files)) {
        const m = path.match(/^xl\/drawings\/_rels\/(drawing\d+\.xml)\.rels$/);
        if (!m) continue;
        const xml = await zip.files[path].async('string');
        if (/Target="[^"]*\/charts\/chart\d+\.xml"/.test(xml)) {
            chartBearingDrawings.add(`xl/drawings/${m[1]}`);
        }
    }

    if (chartBearingDrawings.size === 0) {
        const out = await zip.generateAsync({ type: 'arraybuffer' });
        return out as ArrayBuffer;
    }

    // Walk sheet rels — for every drawing rel that points at a
    // chart-bearing drawing, drop the rel and the matching
    // <drawing r:id="..."/> in the worksheet xml.
    for (const path of Object.keys(zip.files)) {
        const m = path.match(/^xl\/worksheets\/_rels\/sheet(\d+)\.xml\.rels$/);
        if (!m) continue;
        const sheetIndex = m[1];
        let relsXml = await zip.files[path].async('string');

        // Walk every <Relationship .../> independently of attr order.
        const relWalkRe = /<Relationship\s+([^>]*?)\/>/g;
        const droppedRIds: string[] = [];
        let match: RegExpExecArray | null;
        while ((match = relWalkRe.exec(relsXml)) !== null) {
            const attrs = match[1];
            const idM = attrs.match(/\bId="([^"]+)"/);
            const typeM = attrs.match(/\bType="([^"]+)"/);
            const targetM = attrs.match(/\bTarget="([^"]+)"/);
            if (!idM || !typeM || !targetM) continue;
            if (!/\/drawing$/.test(typeM[1])) continue;
            const targetPath = resolveRelTarget('xl/worksheets', targetM[1]);
            if (chartBearingDrawings.has(targetPath)) {
                droppedRIds.push(idM[1]);
            }
        }
        if (droppedRIds.length === 0) continue;

        for (const rId of droppedRIds) {
            // Match a <Relationship ... Id="rId" ... /> regardless of attr order.
            const relRe = new RegExp(`<Relationship\\s+[^>]*?\\bId="${escapeRegex(rId)}"[^>]*?/>`, 'g');
            relsXml = relsXml.replace(relRe, '');
        }
        zip.file(path, relsXml);

        const sheetPath = `xl/worksheets/sheet${sheetIndex}.xml`;
        const sheetFile = zip.file(sheetPath);
        if (sheetFile) {
            let sheetXml = await sheetFile.async('string');
            for (const rId of droppedRIds) {
                // Sheet xml may carry the drawing element with extra attrs
                // (xmlns:r="http://..." in some programmatically-built
                // workbooks — note the URL contains '/' so [^/]* won't
                // span it). Use [^>]* up to the self-closing tag end.
                const tagRe = new RegExp(`<drawing\\s+[^>]*?\\br:id="${escapeRegex(rId)}"[^>]*?/>`, 'g');
                sheetXml = sheetXml.replace(tagRe, '');
            }
            zip.file(sheetPath, sheetXml);
        }
    }

    const removedParts: string[] = [];
    for (const drawingPath of chartBearingDrawings) {
        const folder = drawingPath.split('/').slice(0, -1).join('/');
        const filename = drawingPath.split('/').pop() ?? '';
        const relsPath = `${folder}/_rels/${filename}.rels`;
        zip.remove(drawingPath); removedParts.push(`/${drawingPath}`);
        zip.remove(relsPath);
    }
    for (const path of Object.keys(zip.files)) {
        if (path.startsWith('xl/charts/')) {
            zip.remove(path); removedParts.push(`/${path}`);
        }
    }

    const ctPath = '[Content_Types].xml';
    const ctFile = zip.file(ctPath);
    if (ctFile) {
        let ctXml = await ctFile.async('string');
        for (const partName of removedParts) {
            const re = new RegExp(`<Override\\b[^/]*PartName="${escapeRegex(partName)}"[^/]*/>`, 'g');
            ctXml = ctXml.replace(re, '');
        }
        zip.file(ctPath, ctXml);
    }

    const out = await zip.generateAsync({ type: 'arraybuffer' });
    return out as ArrayBuffer;
}
