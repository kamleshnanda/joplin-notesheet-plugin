// M18 A2: shape drawing import from .xlsx — PRESERVE-ONLY.
//
// Univer 0.23 cannot render shapes (DRAWING_SHAPE is an unbacked stub:
// sheets-drawing-ui mounts only DRAWING_IMAGE + DRAWING_DOM). So we do NOT
// build a Univer drawing entry for shapes and we do NOT render them in the
// editor or in Joplin's HTML/PDF export. Instead we PRESERVE them: each
// standalone <xdr:sp> anchor is captured VERBATIM and stashed (per sheet) in
// a passthrough resource (SHEET_NOTESHEET_SHAPES_PLUGIN). The editor's
// resource hook carries that string through save/reload untouched, and export
// injects the anchors back into the worksheet's drawing part so Excel renders
// them. Round-trips losslessly; invisible inside Joplin (documented gap).
//
// Reads from the ORIGINAL buffer (like the image/chart importers). Unlike
// charts, shapes do NOT crash exceljs's reconcile loop, so no strip step is
// needed — but reading zip-direct keeps the anchor XML byte-exact (exceljs
// does not surface shapes at all).

import JSZip from 'jszip';

import { buildFilenameToSheetId } from './sheetIdResolver';

// One preserved shape anchor. anchorXml is the FULL <xdr:*Anchor>...</...>
// element, verbatim, ready to splice back into a drawing part on export.
export interface ImportedShapeDrawing {
    // Resolved workbook sheetId (matches exceljs ws.id and the snapshot's
    // `sheet-${id}` subUnitId), NOT the worksheet filename number.
    sheetIndex: number;
    anchorXml: string;
}

function resolveRelTarget(baseFolder: string, target: string): string {
    if (target.startsWith('/')) return target.slice(1);
    const segs = baseFolder.split('/').filter(Boolean);
    for (const seg of target.split('/')) {
        if (seg === '..') segs.pop();
        else if (seg !== '.' && seg !== '') segs.push(seg);
    }
    return segs.join('/');
}

// Any namespace prefix (or none) before a drawing-ml local name.
const NS = '(?:[A-Za-z_][\\w.-]*:)?';

interface SheetDrawingLink {
    sheetIndex: number;
    drawingPath: string;
}

async function findSheetDrawingLinks(zip: JSZip): Promise<SheetDrawingLink[]> {
    const out: SheetDrawingLink[] = [];
    const filenameToSheetId = await buildFilenameToSheetId(zip);
    const sheetRelsRe = /^xl\/worksheets\/_rels\/sheet(\d+)\.xml\.rels$/;
    for (const path of Object.keys(zip.files)) {
        const fm = path.match(sheetRelsRe);
        if (!fm) continue;
        const filenameNum = parseInt(fm[1], 10);
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
            out.push({
                sheetIndex,
                drawingPath: resolveRelTarget('xl/worksheets', targetMatch[1]),
            });
        }
    }
    return out;
}

// Walk every anchor block in document order; keep only those that contain a
// standalone <xdr:sp> and NEITHER a <xdr:pic> (image) NOR a <*:graphicFrame>
// (chart). Returns the full anchor element XML, verbatim.
function extractShapeAnchors(drawingXml: string): string[] {
    const out: string[] = [];
    const re = new RegExp(
        `<${NS}(twoCellAnchor|oneCellAnchor|absoluteAnchor)\\b[^>]*>[\\s\\S]*?</${NS}\\1>`,
        'g',
    );
    let m: RegExpExecArray | null;
    while ((m = re.exec(drawingXml)) !== null) {
        const anchor = m[0];
        const hasShape = new RegExp(`<${NS}sp\\b`).test(anchor);
        if (!hasShape) continue;
        // Exclude image anchors and chart anchors — those are A1 / M17.
        if (new RegExp(`<${NS}pic\\b`).test(anchor)) continue;
        if (new RegExp(`<${NS}graphicFrame\\b`).test(anchor)) continue;
        out.push(anchor);
    }
    return out;
}

// Read every preserve-only shape anchor out of the original xlsx buffer.
export async function readShapesFromXlsxZip(
    buffer: ArrayBuffer | Uint8Array | Buffer,
): Promise<ImportedShapeDrawing[]> {
    let zip: JSZip;
    try {
        zip = await JSZip.loadAsync(buffer as ArrayBuffer);
    } catch {
        return [];
    }
    const links = await findSheetDrawingLinks(zip);
    if (links.length === 0) return [];

    const out: ImportedShapeDrawing[] = [];
    for (const link of links) {
        const file = zip.file(link.drawingPath);
        if (!file) continue;
        const xml = await file.async('string');
        for (const anchorXml of extractShapeAnchors(xml)) {
            out.push({ sheetIndex: link.sheetIndex, anchorXml });
        }
    }
    return out;
}
