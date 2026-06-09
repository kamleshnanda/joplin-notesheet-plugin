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

    // Title — first <a:t>...</a:t> inside <c:title><c:tx><c:rich>...
    let title = '';
    const titleMatch = chartXml.match(/<(?:c:)?title\b[\s\S]*?<a:t>([\s\S]*?)<\/a:t>/);
    if (titleMatch) title = decodeXmlEntities(titleMatch[1]);

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
            if (parsed.unsupportedSourceType) {
                drawing.meta = { unsupportedSourceType: parsed.unsupportedSourceType };
                console.warn(`[Notesheet] M17: chart type '${parsed.unsupportedSourceType}' is not supported; falling back to 'bar'`);
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
