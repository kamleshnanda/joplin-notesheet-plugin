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
