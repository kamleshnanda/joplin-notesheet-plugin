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
    labels: string[];
    datasets: Array<{ label?: string; data: number[] }>;
    anchor: {
        fromCol: number; fromColOff: number;
        fromRow: number; fromRowOff: number;
        toCol: number; toColOff: number;
        toRow: number; toRowOff: number;
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
export function rangeRefCol(sheetName: string, startRow: number, endRow: number, col: number): string {
    return `${escapeSheetName(sheetName)}!$${colLetters(col)}$${startRow + 1}:$${colLetters(col)}$${endRow + 1}`;
}

// ─── Color helper ──────────────────────────────────────────────────────────

// CHART_PALETTE values are #RRGGBB; OOXML wants RRGGBB without the hash.
function paletteHex(seriesIndex: number): string {
    const c = CHART_PALETTE[seriesIndex % CHART_PALETTE.length];
    return c.replace(/^#/, '').toUpperCase();
}

// ─── chart{N}.xml builders ─────────────────────────────────────────────────

// All four builders return a complete <c:chartSpace> XML document including
// the XML prolog. Element ordering follows OOXML's strict EG_ChartContent
// sequence; deviating produces "We found a problem" in Excel.

interface BuildChartOpts {
    sheetName: string;
}

// Bar/column: <c:barChart> with barDir='col' for our 'bar' type. (barDir='bar'
// is horizontal — confirmed in spike. Our Chart.js 'bar' is visually a
// vertical column, matching Excel's "Column Chart" UI.)
export function buildBarChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ true, () => {
        const seriesXml = c.datasets.map((ds, i) =>
            buildSeriesXml(c, opts, ds, i, /* solidFill */ paletteHex(i))).join('');
        return `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml}<c:gapWidth val="182"/><c:axId val="111111"/><c:axId val="222222"/></c:barChart>${categoryAndValueAxes()}`;
    });
}

export function buildLineChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ true, () => {
        const seriesXml = c.datasets.map((ds, i) =>
            buildSeriesXml(c, opts, ds, i, paletteHex(i), /* lineSeries */ true)).join('');
        return `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml}<c:marker val="0"/><c:axId val="111111"/><c:axId val="222222"/></c:lineChart>${categoryAndValueAxes()}`;
    });
}

export function buildPieChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ false, () => {
        // Pie has exactly one series — datasets[0] is the data, labels are
        // the category points. Per-data-point colors via <c:dPt> overrides.
        const ds = c.datasets[0] ?? { data: [] };
        return `<c:pieChart><c:varyColors val="1"/>${buildPieSeriesXml(c, opts, ds, /* doughnut */ false)}<c:firstSliceAng val="0"/></c:pieChart>`;
    });
}

export function buildDoughnutChartXml(c: ChartDrawing, opts: BuildChartOpts): string {
    return chartSpaceWrap(c, opts, /* hasAxes */ false, () => {
        const ds = c.datasets[0] ?? { data: [] };
        return `<c:doughnutChart><c:varyColors val="1"/>${buildPieSeriesXml(c, opts, ds, /* doughnut */ true)}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
    });
}

// Top-level wrapper shared by all chart types. The element order inside
// <c:chart> is mandatory: title, autoTitleDeleted, plotArea, plotVisOnly,
// dispBlanksAs. After </c:chart>, the chartSpace gets a default spPr.
function chartSpaceWrap(
    c: ChartDrawing,
    _opts: BuildChartOpts,
    _hasAxes: boolean,
    plotAreaInner: () => string,
): string {
    const titleXml = c.title
        ? `<c:title><c:tx><c:rich><a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0" kern="1200"/></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(c.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
        : `<c:autoTitleDeleted val="1"/>`;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><c:chart>${titleXml}<c:plotArea><c:layout/>${plotAreaInner()}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr><c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`;
}

// One <c:ser> for bar or line. Element order inside <c:ser> is strict:
// idx, order, tx, spPr, marker?, invertIfNegative?, cat, val.
function buildSeriesXml(
    c: ChartDrawing,
    opts: BuildChartOpts,
    ds: { label?: string; data: number[] },
    seriesIndex: number,
    fillHex: string,
    lineSeries = false,
): string {
    const labelCol = c.sourceRange.startColumn;
    const dataCol = c.sourceRange.startColumn + 1 + seriesIndex;
    const dataStartRow = c.sourceRange.startRow + 1; // skip the header row
    const dataEndRow = c.sourceRange.endRow;

    // Header cell at the top of the series's column gives the series label.
    const headerRow = c.sourceRange.startRow;
    const seriesNameRef = cellRef(opts.sheetName, headerRow, dataCol);
    const seriesName = ds.label ?? `Series ${seriesIndex + 1}`;

    const txXml = `<c:tx><c:strRef><c:f>${seriesNameRef}</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${escapeXml(seriesName)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`;

    // Per-series fill (bar) or line stroke (line). Both reference the palette.
    const spPrXml = lineSeries
        ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill><a:round/></a:ln><a:effectLst/></c:spPr>`
        : `<c:spPr><a:solidFill><a:srgbClr val="${fillHex}"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr>`;

    const markerXml = lineSeries ? `<c:marker><c:symbol val="none"/></c:marker>` : `<c:invertIfNegative val="0"/>`;

    // Categories: the label column (string ref + cache).
    const labelsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, labelCol);
    const labelsCacheXml = c.labels.map((v, i) =>
        `<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`).join('');
    const catXml = `<c:cat><c:strRef><c:f>${labelsRef}</c:f><c:strCache><c:ptCount val="${c.labels.length}"/>${labelsCacheXml}</c:strCache></c:strRef></c:cat>`;

    // Values: this column's range (number ref + cache).
    const valsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, dataCol);
    const valsCacheXml = ds.data.map((n, i) => {
        // NaN/Infinity must be omitted — Excel won't load <c:v>NaN</c:v>.
        if (typeof n !== 'number' || !Number.isFinite(n)) return '';
        return `<c:pt idx="${i}"><c:v>${n}</c:v></c:pt>`;
    }).join('');
    const valXml = `<c:val><c:numRef><c:f>${valsRef}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${ds.data.length}"/>${valsCacheXml}</c:numCache></c:numRef></c:val>`;

    return `<c:ser><c:idx val="${seriesIndex}"/><c:order val="${seriesIndex}"/>${txXml}${spPrXml}${markerXml}${catXml}${valXml}<c:smooth val="0"/></c:ser>`;
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
    const dPtXml = ds.data.map((_, i) => {
        const hex = paletteHex(i);
        return `<c:dPt><c:idx val="${i}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${hex}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:dPt>`;
    }).join('');

    const labelsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, labelCol);
    const labelsCacheXml = c.labels.map((v, i) =>
        `<c:pt idx="${i}"><c:v>${escapeXml(String(v))}</c:v></c:pt>`).join('');
    const catXml = `<c:cat><c:strRef><c:f>${labelsRef}</c:f><c:strCache><c:ptCount val="${c.labels.length}"/>${labelsCacheXml}</c:strCache></c:strRef></c:cat>`;

    const valsRef = rangeRefCol(opts.sheetName, dataStartRow, dataEndRow, dataCol);
    const valsCacheXml = ds.data.map((n, i) => {
        if (typeof n !== 'number' || !Number.isFinite(n)) return '';
        return `<c:pt idx="${i}"><c:v>${n}</c:v></c:pt>`;
    }).join('');
    const valXml = `<c:val><c:numRef><c:f>${valsRef}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${ds.data.length}"/>${valsCacheXml}</c:numCache></c:numRef></c:val>`;

    return `<c:ser><c:idx val="0"/><c:order val="0"/>${txXml}${dPtXml}${catXml}${valXml}</c:ser>`;
}

// Shared cat/val axes block for bar/line. axId pair must match the values
// used inside the chart-type element above (we hardcode 111111 / 222222).
function categoryAndValueAxes(): string {
    return `<c:catAx><c:axId val="111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="111111"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
}

// ─── drawing{N}.xml builder ────────────────────────────────────────────────

// Aggregates ALL of a sheet's chart anchors into one wsDr. The drawing-rels
// rIds inside the anchors are 1-based per drawing.xml (independent of the
// sheet rels rIds — those live in xl/worksheets/_rels/sheet{S}.xml.rels).
export function buildDrawingXml(charts: ChartDrawing[]): string {
    const anchors = charts.map((c, i) => {
        const a = c.anchor;
        const rId = i + 1; // 1-based rId inside drawing{N}.xml.rels
        const cNvPrId = i + 2; // Excel uses id=2,3,... for graphic frames
        return `<xdr:twoCellAnchor>` +
            `<xdr:from><xdr:col>${a.fromCol}</xdr:col><xdr:colOff>${a.fromColOff}</xdr:colOff><xdr:row>${a.fromRow}</xdr:row><xdr:rowOff>${a.fromRowOff}</xdr:rowOff></xdr:from>` +
            `<xdr:to><xdr:col>${a.toCol}</xdr:col><xdr:colOff>${a.toColOff}</xdr:colOff><xdr:row>${a.toRow}</xdr:row><xdr:rowOff>${a.toRowOff}</xdr:rowOff></xdr:to>` +
            `<xdr:graphicFrame macro="">` +
            `<xdr:nvGraphicFramePr><xdr:cNvPr id="${cNvPrId}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
            // <a:off>/<a:ext> zeros are correct — anchor drives size, this xfrm is a placeholder. Spike confirmed.
            `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
            `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${rId}"/></a:graphicData></a:graphic>` +
            `</xdr:graphicFrame>` +
            `<xdr:clientData/>` +
            `</xdr:twoCellAnchor>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`;
}

// ─── rels builders ─────────────────────────────────────────────────────────

// drawing{N}.xml.rels — one Relationship per chart on this drawing's sheet.
// chartFileNumbers lines up 1:1 with the drawing's anchors; e.g. anchors
// pointing to chart3.xml + chart4.xml (because charts 1,2 are on a different
// sheet's drawing) → chartFileNumbers=[3,4].
export function buildDrawingRelsXml(chartFileNumbers: number[]): string {
    const rels = chartFileNumbers.map((n, i) =>
        `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${n}.xml"/>`).join('');
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
        case 'bar': return buildBarChartXml(c, opts);
        case 'line': return buildLineChartXml(c, opts);
        case 'pie': return buildPieChartXml(c, opts);
        case 'doughnut': return buildDoughnutChartXml(c, opts);
        default: return buildBarChartXml(c, opts);
    }
}

// ─── Read drawings from snapshot (Step 1 of M10) ───────────────────────────

// The SHEET_DRAWING_PLUGIN resource's data field is a JSON-stringified map:
//   { [subUnitId]: { data: { [drawingId]: ISheetDrawing }, order: string[] } }
// We filter to entries whose componentKey === 'NotesheetChart' (other
// drawings — images, shapes — coexist and must be left alone).
export function readChartsFromSnapshot(snapshot: UniverSnapshot): ChartDrawing[] {
    const resources = (snapshot as { resources?: Array<{ name?: string; data?: string }> }).resources;
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
                    sourceRange?: { startRow?: number; endRow?: number; startColumn?: number; endColumn?: number };
                    title?: string;
                    labels?: unknown[];
                    datasets?: Array<{ label?: string; data?: unknown[] }>;
                };
                axisAlignSheetTransform?: {
                    from?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number };
                    to?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number };
                };
                sheetTransform?: {
                    from?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number };
                    to?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number };
                };
            };

            if (d?.componentKey !== 'NotesheetChart') continue;
            const data = d.data;
            if (!data) continue;

            const type = (data.type === 'bar' || data.type === 'line' || data.type === 'pie' || data.type === 'doughnut')
                ? data.type
                : 'bar';

            const sr = data.sourceRange;
            if (!sr || typeof sr.startRow !== 'number' || typeof sr.endRow !== 'number'
                || typeof sr.startColumn !== 'number' || typeof sr.endColumn !== 'number') continue;

            // Prefer axisAlignSheetTransform — it's xlsx-aligned per Univer's
            // service. Fall back to sheetTransform on the (rare) chance the
            // axisAlign one wasn't computed.
            const tx = d.axisAlignSheetTransform ?? d.sheetTransform;
            if (!tx?.from || !tx?.to) continue;

            const labels = Array.isArray(data.labels) ? data.labels.map((l) => String(l ?? '')) : [];
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
            });
        }
    }
    return out;
}

// ─── Zip surgery (Step 5 of M10) ───────────────────────────────────────────

// Look up a sheet's 1-based index in the workbook (matches xl/worksheets/sheet{N}.xml)
// and its display name (used for <c:f> sheet-qualified ranges).
function lookupSheet(snapshot: UniverSnapshot, sheetId: string): { index: number; name: string } | null {
    const sheetOrder = (snapshot as { sheetOrder?: string[] }).sheetOrder ?? [];
    const sheets = (snapshot as { sheets?: Record<string, { name?: string }> }).sheets ?? {};
    const idx0 = sheetOrder.indexOf(sheetId);
    if (idx0 < 0) return null;
    return { index: idx0 + 1, name: sheets[sheetId]?.name ?? `Sheet${idx0 + 1}` };
}

// Highest existing rId in a Relationships XML, or 0 if none. Used for
// allocating a new rId without colliding with existing ones (e.g. if M9
// added a rels entry for a table).
function maxExistingRId(relsXml: string | null): number {
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
function upsertRelationship(existing: string | null, newRel: string): string {
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
function insertDrawingRefIntoSheet(sheetXml: string, rId: number): string {
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
function patchContentTypes(contentTypesXml: string, drawingNums: number[], chartNums: number[]): string {
    const toAdd: string[] = [];
    for (const n of drawingNums) {
        toAdd.push(`<Override PartName="/xl/drawings/drawing${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
    }
    for (const n of chartNums) {
        toAdd.push(`<Override PartName="/xl/charts/chart${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
        toAdd.push(`<Override PartName="/xl/charts/style${n}.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>`);
        toAdd.push(`<Override PartName="/xl/charts/colors${n}.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>`);
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
            zip.file(`xl/drawings/_rels/drawing${drawingNum}.xml.rels`, buildDrawingRelsXml(chartFileNumbers));

            // 4) Add each chart's xml + style + colors + chart rels.
            for (let i = 0; i < sheetCharts.length; i++) {
                const c = sheetCharts[i];
                const chartNum = chartFileNumbers[i];
                zip.file(`xl/charts/chart${chartNum}.xml`, buildChartXml(c, { sheetName: sheet.name }));
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
