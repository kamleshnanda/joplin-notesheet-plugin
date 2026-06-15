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

import { buildFilenameToSheetId } from './sheetIdResolver';

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
            // A sheet may reference more than one drawing part — emit every
            // one (no early break), so images in a second drawing aren't lost.
            out.push({ sheetIndex, drawingPath });
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

// Any namespace prefix (or none): `xdr:`, the default ns, or a third-party
// alias like `a1:`. We match an optional `<prefix>:` before the local name.
const NS = '(?:[A-Za-z_][\\w.-]*:)?';

function parseAnchorPoint(body: string): ParsedPoint | null {
    const col = body.match(new RegExp(`<${NS}col>(\\d+)</${NS}col>`));
    const colOff = body.match(new RegExp(`<${NS}colOff>(-?\\d+)</${NS}colOff>`));
    const row = body.match(new RegExp(`<${NS}row>(\\d+)</${NS}row>`));
    const rowOff = body.match(new RegExp(`<${NS}rowOff>(-?\\d+)</${NS}rowOff>`));
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

// Walk every anchor block in DOCUMENT order and emit one RawImageAnchor per
// <pic> it contains. Namespace-tolerant (any prefix or none). Handles all
// three anchor kinds:
//   - twoCellAnchor   <from>+<to>
//   - oneCellAnchor   <from>+<ext>  (to-cell synthesized downstream)
//   - absoluteAnchor  <pos>+<ext>   (no cells → from-cell synthesized at 0,0)
// A single anchor may hold multiple <pic> (e.g. a grouped <grpSp>); each is
// emitted as its own image so none are silently dropped.
function walkImageAnchors(drawingXml: string): RawImageAnchor[] {
    const out: RawImageAnchor[] = [];
    const re = new RegExp(
        `<${NS}(twoCellAnchor|oneCellAnchor|absoluteAnchor)\\b[^>]*>([\\s\\S]*?)</${NS}\\1>`,
        'g',
    );
    let match: RegExpExecArray | null;
    while ((match = re.exec(drawingXml)) !== null) {
        const kind = match[1];
        const body = match[2];
        // Only image anchors: must contain a <pic>.
        if (!new RegExp(`<${NS}pic\\b`).test(body)) continue;

        // Anchor-level geometry shared by every <pic> in this block.
        let from: ParsedPoint | null = null;
        const fromBody = body.match(new RegExp(`<${NS}from>([\\s\\S]*?)</${NS}from>`));
        if (fromBody) from = parseAnchorPoint(fromBody[1]);

        let to: ParsedPoint | null = null;
        const toBody = body.match(new RegExp(`<${NS}to>([\\s\\S]*?)</${NS}to>`));
        if (toBody) to = parseAnchorPoint(toBody[1]);

        // Extent (EMU). <ext cx="..." cy="..."/> — present on oneCellAnchor
        // and absoluteAnchor (and ignored on twoCellAnchor).
        let ext: { cx: number; cy: number } | null = null;
        const extM = body.match(new RegExp(`<${NS}ext\\b[^>]*\\bcx="(\\d+)"[^>]*\\bcy="(\\d+)"`));
        if (extM) ext = { cx: parseInt(extM[1], 10), cy: parseInt(extM[2], 10) };

        // absoluteAnchor has no <from> cell — anchor at A1 (0,0). The <pos> is
        // an absolute EMU offset from the sheet origin; approximate it as the
        // from-cell offset so the image lands near its real position. (Exact
        // EMU-position fidelity is the separate A3 item.)
        if (!from && kind === 'absoluteAnchor') {
            const pos = body.match(
                new RegExp(`<${NS}pos\\b[^>]*\\bx="(-?\\d+)"[^>]*\\by="(-?\\d+)"`),
            );
            from = {
                col: 0,
                colOff: pos ? Math.max(0, parseInt(pos[1], 10)) : 0,
                row: 0,
                rowOff: pos ? Math.max(0, parseInt(pos[2], 10)) : 0,
            };
        }
        if (!from) continue;

        // Emit one image per <pic> in the anchor (grouped images / multi-pic).
        const picRe = new RegExp(`<${NS}pic\\b[^>]*>([\\s\\S]*?)</${NS}pic>`, 'g');
        let picM: RegExpExecArray | null;
        while ((picM = picRe.exec(body)) !== null) {
            const picBody = picM[1];
            const embedM =
                picBody.match(/<(?:[A-Za-z_][\w.-]*:)?blip\b[^>]*\sr:embed="([^"]+)"/) ??
                picBody.match(/<(?:[A-Za-z_][\w.-]*:)?blip\b[^>]*\sembed="([^"]+)"/);
            if (!embedM) continue;
            out.push({ rEmbed: embedM[1], from, to, ext });
        }
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
    // Running image index PER sheetIndex (not per drawing link) so the imageId
    // stays unique even when a sheet references multiple drawing parts.
    const seqBySheet = new Map<number, number>();
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

            // imageId must be UNIQUE per (sheet, anchor occurrence) — it feeds
            // the snapshot drawingId `image-imported-${sheet}-${imageId}`, and
            // two anchors on one sheet can reference the SAME media part (e.g.
            // the same logo placed twice). Keying off the media number alone
            // would collide and the second image would overwrite the first in
            // the drawing map. Use a per-sheet anchor sequence (stable within
            // a single import; the media bytes still round-trip identically).
            const imageId = seqBySheet.get(link.sheetIndex) ?? 0;
            seqBySheet.set(link.sheetIndex, imageId + 1);

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
