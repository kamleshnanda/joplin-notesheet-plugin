// Shared drawing-import helper: resolve an xl/worksheets/sheet{N}.xml FILENAME
// number to the workbook's logical sheetId.
//
// THE PROBLEM THIS SOLVES: the worksheet filename number is NOT the sheet's
// identity. On an EDITED workbook (a sheet deleted/reordered/inserted) the two
// drift apart — e.g. a deleted leading sheet leaves sheet1.xml pointing at
// sheetId=2. xlsx.ts keys snapshot subUnitIds off exceljs's `ws.id`, which is
// the workbook.xml `<sheet sheetId="...">` value. A drawing importer that uses
// the raw filename number as the sheet index will therefore assign the drawing
// to the wrong `sheet-N` — and xlsx.ts's `if (!sheets[subUnitId]) continue`
// guard then SILENTLY DROPS it. Both the image importer (xlsxImageImport.ts)
// and the chart importer (xlsxChartImport.ts) MUST resolve through this.
//
// Resolution: workbook.xml.rels maps rId -> worksheet filename; workbook.xml
// maps sheetId -> rId. Compose them to get filename number -> sheetId. Returns
// an empty map when the workbook parts are absent — callers then fall back to
// the filename number, which is correct for a clean single-sheet save.

import JSZip from 'jszip';

// EMU per pixel at 96 DPI.
const EMU_PER_PX = 9525;

// The eight-field EMU anchor (from/to cell + offset), as stashed on a drawing
// at import under `_srcAnchorEmu`.
export interface SrcAnchorEmu {
    fromCol: number;
    fromColOff: number;
    fromRow: number;
    fromRowOff: number;
    toCol: number;
    toColOff: number;
    toRow: number;
    toRowOff: number;
}

interface PxPoint {
    column?: number;
    columnOffset?: number;
    row?: number;
    rowOffset?: number;
}

// M18 A3: resolve a drawing's OOXML EMU anchor for export.
//
// A drawing imported from .xlsx carries `_srcAnchorEmu` (the EXACT source EMU
// offsets). Univer's px `sheetTransform` is a lossy projection of that (EMU →
// round(px) → EMU drifts ~⅓px, and the editor only ever updates the px form).
// So:
//   - If `_srcAnchorEmu` is present AND the px transform still matches it
//     (drawing not moved in the editor), reproduce the source EMU EXACTLY.
//   - Otherwise (no stash, or the user moved the drawing so px diverged),
//     fall back to px × 9525.
// `matchesStash` compares each px field to round(emu / 9525); if every field
// agrees the drawing is unmoved and the exact EMU is safe to use.
export function resolveAnchorEmu(
    srcAnchorEmu: SrcAnchorEmu | undefined,
    from: PxPoint,
    to: PxPoint,
): SrcAnchorEmu {
    const px = {
        fromCol: from.column ?? 0,
        fromColOff: from.columnOffset ?? 0,
        fromRow: from.row ?? 0,
        fromRowOff: from.rowOffset ?? 0,
        toCol: to.column ?? 0,
        toColOff: to.columnOffset ?? 0,
        toRow: to.row ?? 0,
        toRowOff: to.rowOffset ?? 0,
    };
    if (srcAnchorEmu && anchorUnmoved(srcAnchorEmu, px)) {
        return srcAnchorEmu;
    }
    return {
        fromCol: px.fromCol,
        fromColOff: Math.round(px.fromColOff * EMU_PER_PX),
        fromRow: px.fromRow,
        fromRowOff: Math.round(px.fromRowOff * EMU_PER_PX),
        toCol: px.toCol,
        toColOff: Math.round(px.toColOff * EMU_PER_PX),
        toRow: px.toRow,
        toRowOff: Math.round(px.toRowOff * EMU_PER_PX),
    };
}

// True when every px field equals the stash's EMU field projected to px —
// i.e. the editor hasn't moved the drawing since import.
function anchorUnmoved(
    emu: SrcAnchorEmu,
    px: SrcAnchorEmu, // here the *Off fields hold PX, cells hold indices
): boolean {
    return (
        px.fromCol === emu.fromCol &&
        px.fromRow === emu.fromRow &&
        px.toCol === emu.toCol &&
        px.toRow === emu.toRow &&
        px.fromColOff === Math.round(emu.fromColOff / EMU_PER_PX) &&
        px.fromRowOff === Math.round(emu.fromRowOff / EMU_PER_PX) &&
        px.toColOff === Math.round(emu.toColOff / EMU_PER_PX) &&
        px.toRowOff === Math.round(emu.toRowOff / EMU_PER_PX)
    );
}

// M18 A2: Notesheet-private resource holding preserve-only shape anchors,
// keyed by subUnitId -> string[] of verbatim <xdr:*Anchor> XML. Lives here
// (a leaf module both xlsx.ts and the shape modules already depend on) to
// avoid a circular import between xlsx.ts and drawings/xlsxShape.ts.
export const NOTESHEET_SHAPES_RESOURCE = 'SHEET_NOTESHEET_SHAPES_PLUGIN';

export async function buildFilenameToSheetId(zip: JSZip): Promise<Map<number, number>> {
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
