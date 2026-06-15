// M18 A1: image drawing import from .xlsx — pure-stdlib zip+regex reader.
//
// Architecture mirrors src/charts/xlsxChartImport.ts (read BEFORE wb.xlsx.load
// from the ORIGINAL buffer), and for the SAME reason: the chart import path
// runs stripChartPartsFromZip before the workbook loads, and that strip
// removes an entire drawing part when its rels point at a chart — which would
// take any image anchors sharing that drawing down with it. Reading images
// zip-direct from the original buffer is robust regardless of the strip, and
// also sidesteps the loader's draw-side reconcile entirely. (The exceljs
// getImages() API works on plain image-only workbooks, but not on a
// chart+image workbook once the chart drawing has been stripped — see
// tests/m18ImageChartCoexist.)
//
// We walk every sheet's drawing rels, every drawing's image anchors in
// document order, resolve each anchor's r:embed against the drawing rels to
// find the xl/media/* part, and read the bytes. emf/wmf are skipped with a
// logged known-shortcoming warning (not browser-renderable).

import JSZip from 'jszip';

// One image, after import. Keyed by sheetIndex (1-based — matches
// xl/worksheets/sheet{N}.xml and the chart import shape).
export interface ImportedImageDrawing {
    sheetIndex: number;
    // Stable per-(sheet,drawing) identifier derived from the media part name.
    imageId: number;
    mime: string;
    // Raw image bytes; xlsx.ts base64-encodes into a data: URI source.
    // Typed as Uint8Array (a Node Buffer is a Uint8Array at runtime).
    buffer: Uint8Array;
    anchor: {
        fromCol: number;
        fromColOff: number; // EMU
        fromRow: number;
        fromRowOff: number; // EMU
        toCol: number;
        toColOff: number; // EMU
        toRow: number;
        toRowOff: number; // EMU
    };
    // Pixel extent from the source <xdr:ext>, when present. Used to derive the
    // Univer transform width/height (and to synthesize a to-cell for a
    // oneCellAnchor without an <xdr:to>).
    ext?: { width: number; height: number };
}

// Map a media extension to its MIME type. emf/wmf return null — callers skip
// them (vector metafiles, not browser-renderable). jpg aliases jpeg.
export function extensionToMime(extension: string | undefined): string | null {
    switch ((extension ?? '').toLowerCase()) {
        case 'png':
            return 'image/png';
        case 'jpg':
        case 'jpeg':
            return 'image/jpeg';
        case 'gif':
            return 'image/gif';
        case 'bmp':
            return 'image/bmp';
        case 'webp':
            return 'image/webp';
        case 'emf':
        case 'wmf':
            return null;
        default:
            return null;
    }
}

// EMU per pixel at 96 DPI (mirrors the chart import path's 9525 factor).
const EMU_PER_PX = 9525;

// Approximate cell span used to synthesize a to-cell from an <xdr:ext> pixel
// extent when the source anchor is a oneCellAnchor (no <xdr:to>). Mirrors the
// chart import path's DEFAULT_COL_W / DEFAULT_ROW_H constants.
const APPROX_COL_W_PX = 73;
const APPROX_ROW_H_PX = 19;

// --- zip path helpers (local copies of the chart import patterns) ---

function parseRels(relsXml: string): Map<string, { target: string; type: string }> {
    const out = new Map<string, { target: string; type: string }>();
    const re = /<Relationship\s+([^>]*?)\/>/g;
    let match: RegExpExecArray | null;
    while ((match = re.exec(relsXml)) !== null) {
        const attrs = match[1];
        const idM = attrs.match(/\bId="([^"]+)"/);
        const targetM = attrs.match(/\bTarget="([^"]+)"/);
        const typeM = attrs.match(/\bType="([^"]+)"/);
        if (idM && targetM) {
            out.set(idM[1], { target: targetM[1], type: typeM ? typeM[1] : '' });
        }
    }
    return out;
}

function resolveRelTarget(baseFolder: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    const segs = baseFolder.split('/').filter(Boolean);
    const targetSegs = target.split('/');
    for (const seg of targetSegs) {
        if (seg === '..') segs.pop();
        else if (seg !== '.' && seg !== '') segs.push(seg);
    }
    return segs.join('/');
}

// Map worksheet-xml filename number -> the workbook sheetId. The
// xl/worksheets/sheet{N}.xml filename number is NOT the workbook sheet id:
// the workbook may renumber (e.g. a deleted first sheet leaves sheet1.xml
// pointing at sheetId=2). xlsx.ts keys snapshot subUnitIds off exceljs's
// `ws.id`, which equals the workbook.xml <sheet sheetId="..."> attribute. To
// emit a subUnitId that matches, the importer must resolve filename -> sheetId
// via workbook.xml.rels (rId -> worksheet target) + workbook.xml (sheetId ->
// rId). Returns an empty map when the workbook parts are absent (callers then
// fall back to the filename number, which is correct for single-sheet files).
async function buildFilenameToSheetId(zip: JSZip): Promise<Map<number, number>> {
    const out = new Map<number, number>();
    const wbRelsFile = zip.file('xl/_rels/workbook.xml.rels');
    const wbFile = zip.file('xl/workbook.xml');
    if (!wbRelsFile || !wbFile) return out;
    // rId -> worksheet filename number.
    const rIdToFilenameNum = new Map<string, number>();
    const relsXml = await wbRelsFile.async('string');
    const relRe = /<Relationship\s+([^>]*?)\/>/g;
    let m: RegExpExecArray | null;
    while ((m = relRe.exec(relsXml)) !== null) {
        const attrs = m[1];
        const idM = attrs.match(/\bId="([^"]+)"/);
        const typeM = attrs.match(/\bType="([^"]+)"/);
        const targetM = attrs.match(/\bTarget="([^"]+)"/);
        if (!idM || !typeM || !targetM) continue;
        if (!/\/worksheet$/.test(typeM[1])) continue;
        const fnM = targetM[1].match(/sheet(\d+)\.xml$/);
        if (fnM) rIdToFilenameNum.set(idM[1], parseInt(fnM[1], 10));
    }
    // sheetId + r:id from workbook.xml <sheet .../>.
    const wbXml = await wbFile.async('string');
    const sheetRe = /<sheet\b([^>]*)\/>/g;
    while ((m = sheetRe.exec(wbXml)) !== null) {
        const attrs = m[1];
        const sheetIdM = attrs.match(/\bsheetId="(\d+)"/);
        const ridM = attrs.match(/\br:id="([^"]+)"/);
        if (!sheetIdM || !ridM) continue;
        const filenameNum = rIdToFilenameNum.get(ridM[1]);
        if (filenameNum !== undefined) {
            out.set(filenameNum, parseInt(sheetIdM[1], 10));
        }
    }
    return out;
}

interface SheetDrawingLink {
    // sheetIndex is the resolved workbook sheetId (matches exceljs ws.id and
    // the snapshot's `sheet-${id}` subUnitId), NOT the worksheet filename num.
    sheetIndex: number;
    drawingPath: string;
}

async function findSheetDrawingLinks(zip: JSZip): Promise<SheetDrawingLink[]> {
    const out: SheetDrawingLink[] = [];
    const filenameToSheetId = await buildFilenameToSheetId(zip);
    const sheetRelsRe = /^xl\/worksheets\/_rels\/sheet(\d+)\.xml\.rels$/;
    for (const path of Object.keys(zip.files)) {
        const m = path.match(sheetRelsRe);
        if (!m) continue;
        const filenameNum = parseInt(m[1], 10);
        const sheetIndex = filenameToSheetId.get(filenameNum) ?? filenameNum;
        const xml = await zip.files[path].async('string');
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

// --- anchor parsing ---

interface ParsedPoint {
    col: number;
    colOff: number;
    row: number;
    rowOff: number;
}

function parseAnchorPoint(body: string): ParsedPoint | null {
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

interface RawImageAnchor {
    rEmbed: string; // r:embed pointing at the media rel
    from: ParsedPoint;
    to: ParsedPoint | null;
    ext: { cx: number; cy: number } | null; // EMU extent (oneCellAnchor)
}

// Walk every anchor block that contains an <xdr:pic> in DOCUMENT order.
// Namespace-tolerant (xdr: prefix or default namespace), mirroring the chart
// anchor walker.
function walkImageAnchors(drawingXml: string): RawImageAnchor[] {
    const out: RawImageAnchor[] = [];
    const re =
        /<(?:xdr:)?(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/(?:xdr:)?\1>/g;
    let match: RegExpExecArray | null;
    while ((match = re.exec(drawingXml)) !== null) {
        const body = match[2];
        // Only image anchors: must contain a <pic>.
        if (!/<(?:xdr:)?pic\b/.test(body)) continue;
        // The blip embed ref. r: prefix is conventional but may be absent.
        const embedM =
            body.match(/<(?:a:)?blip\b[^>]*\sr:embed="([^"]+)"/) ??
            body.match(/<(?:a:)?blip\b[^>]*\sembed="([^"]+)"/);
        if (!embedM) continue;
        const rEmbed = embedM[1];

        const fromBody = body.match(/<(?:xdr:)?from>([\s\S]*?)<\/(?:xdr:)?from>/);
        if (!fromBody) continue;
        const from = parseAnchorPoint(fromBody[1]);
        if (!from) continue;

        let to: ParsedPoint | null = null;
        const toBody = body.match(/<(?:xdr:)?to>([\s\S]*?)<\/(?:xdr:)?to>/);
        if (toBody) to = parseAnchorPoint(toBody[1]);

        // oneCellAnchor extent (EMU). <xdr:ext cx="..." cy="..."/>.
        let ext: { cx: number; cy: number } | null = null;
        const extM = body.match(/<(?:xdr:)?ext\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"/);
        if (extM) ext = { cx: parseInt(extM[1], 10), cy: parseInt(extM[2], 10) };

        out.push({ rEmbed, from, to, ext });
    }
    return out;
}

// --- entry point ---

// Read every image drawing out of the ORIGINAL xlsx buffer (call BEFORE the
// chart strip + workbook load). Returns one ImportedImageDrawing per
// renderable image; emf/wmf are skipped with a warn.
export async function readImagesFromXlsxZip(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<ImportedImageDrawing[]> {
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return [];
    }

    const sheetLinks = await findSheetDrawingLinks(zip);
    if (sheetLinks.length === 0) return [];

    const out: ImportedImageDrawing[] = [];
    for (const link of sheetLinks) {
        const drawingFile = zip.file(link.drawingPath);
        if (!drawingFile) continue;
        const drawingXml = await drawingFile.async('string');

        const drawingFolder = link.drawingPath.split('/').slice(0, -1).join('/');
        const drawingFilename = link.drawingPath.split('/').pop() ?? '';
        const drawingRelsPath = `${drawingFolder}/_rels/${drawingFilename}.rels`;
        const drawingRelsFile = zip.file(drawingRelsPath);
        if (!drawingRelsFile) continue;
        const relMap = parseRels(await drawingRelsFile.async('string'));

        const anchors = walkImageAnchors(drawingXml);
        let perSheetSeq = 0;
        for (const anchor of anchors) {
            const rel = relMap.get(anchor.rEmbed);
            if (!rel) continue;
            // Only follow image-type rels (skip chart/hyperlink etc.).
            if (rel.type && !/\/image$/.test(rel.type)) continue;
            const mediaPath = resolveRelTarget(drawingFolder, rel.target);
            const mediaFile = zip.file(mediaPath);
            if (!mediaFile) continue;

            const extName = mediaPath.split('.').pop();
            const mime = extensionToMime(extName);
            if (mime === null) {
                // KNOWN SHORTCOMING: vector metafiles (emf/wmf) can't render
                // in the editor. Skip rather than emit a broken image; the
                // cells survive and the spreadsheet stays valid.
                console.warn(
                    `[Notesheet] M18: skipping non-renderable image (${mediaPath}) on sheet ${link.sheetIndex} — emf/wmf can't render in the editor`,
                );
                continue;
            }

            const bytes = await mediaFile.async('uint8array');

            const from = anchor.from;
            const ext =
                anchor.ext != null
                    ? { width: anchor.ext.cx / EMU_PER_PX, height: anchor.ext.cy / EMU_PER_PX }
                    : undefined;

            let toCol: number;
            let toRow: number;
            let toColOff: number;
            let toRowOff: number;
            if (anchor.to) {
                toCol = anchor.to.col;
                toRow = anchor.to.row;
                toColOff = anchor.to.colOff;
                toRowOff = anchor.to.rowOff;
            } else if (ext) {
                const spanCols = Math.max(1, Math.round(ext.width / APPROX_COL_W_PX));
                const spanRows = Math.max(1, Math.round(ext.height / APPROX_ROW_H_PX));
                toCol = from.col + spanCols;
                toRow = from.row + spanRows;
                const remColPx = ext.width - spanCols * APPROX_COL_W_PX;
                const remRowPx = ext.height - spanRows * APPROX_ROW_H_PX;
                toColOff = Math.max(0, Math.round(remColPx * EMU_PER_PX));
                toRowOff = Math.max(0, Math.round(remRowPx * EMU_PER_PX));
            } else {
                // No to-cell and no ext — fall back to a chart-sized span.
                toCol = from.col + 6;
                toRow = from.row + 14;
                toColOff = 0;
                toRowOff = 0;
            }

            // Derive a stable imageId from the media part number when present
            // (xl/media/imageN.ext), else a per-sheet running index.
            const mediaNumM = mediaPath.match(/image(\d+)\./);
            const imageId = mediaNumM ? parseInt(mediaNumM[1], 10) : perSheetSeq;
            perSheetSeq += 1;

            out.push({
                sheetIndex: link.sheetIndex,
                imageId,
                mime,
                buffer: bytes,
                anchor: {
                    fromCol: from.col,
                    fromColOff: from.colOff,
                    fromRow: from.row,
                    fromRowOff: from.rowOff,
                    toCol,
                    toColOff,
                    toRow,
                    toRowOff,
                },
                ...(ext ? { ext } : {}),
            });
        }
    }

    return out;
}
