// M17 fidelity (manual-test issue 11): a worksheet's default column
// width must survive the import → export round-trip.
//
// The bug: 11-stacked-bar-chart.xlsx ships `<sheetFormatPr
// baseColWidth="10"/>` and a single explicit `<col>` for column A.
// Column B (and every other column) inherits the wider ~11.5-char
// default. exceljs doesn't read baseColWidth, so on export the default
// collapsed to exceljs's 8.43 — column B rendered visibly narrower than
// the source. We now read sheetFormatPr's defaultColWidth/baseColWidth
// zip-direct at import and re-emit it on export.
//
// Anchored to the source workbook (Excel-canonical): we first read what
// the source's default width is, then assert the round-tripped export
// preserves it — NOT a hardcoded literal.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');

async function sheetFormatPr(buffer: ArrayBuffer | Buffer): Promise<string | null> {
    const zip = await JSZip.loadAsync(buffer);
    const sheetPath = Object.keys(zip.files).find((k) => /^xl\/worksheets\/sheet1\.xml$/i.test(k));
    if (!sheetPath) return null;
    const xml = await zip.file(sheetPath)!.async('string');
    const m = xml.match(/<sheetFormatPr\b[^>]*>/i);
    return m ? m[0] : null;
}

// The effective default width Excel applies to a column with no explicit
// <col>: defaultColWidth if present, else baseColWidth + 1.
function effectiveDefaultWidth(fmt: string | null): number | null {
    if (!fmt) return null;
    const def = fmt.match(/\bdefaultColWidth="([\d.]+)"/i);
    if (def) return parseFloat(def[1]);
    const base = fmt.match(/\bbaseColWidth="([\d.]+)"/i);
    if (base) return parseFloat(base[1]) + 1;
    return null;
}

describe('M17 default column width round-trip (issue 11)', () => {
    test('11-stacked-bar-chart.xlsx: source has a wide default column width', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx'));
        const srcDefault = effectiveDefaultWidth(await sheetFormatPr(buf as unknown as Buffer));
        expect(srcDefault).not.toBeNull();
        // Source default is wider than exceljs's 8.43 fallback — that gap
        // is exactly what produced the "column B too narrow" report.
        expect(srcDefault!).toBeGreaterThan(8.43);
    });

    test('import captures the source default width onto the sheet record', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx'));
        const srcDefault = effectiveDefaultWidth(await sheetFormatPr(buf as unknown as Buffer))!;
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const sheet = Object.values(
            (snap as { sheets: Record<string, { defaultColWidthChars?: number }> }).sheets,
        )[0];
        expect(sheet.defaultColWidthChars).toBeCloseTo(srcDefault, 5);
    });

    test('export preserves the default width (within Excel rounding)', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx'));
        const srcDefault = effectiveDefaultWidth(await sheetFormatPr(buf as unknown as Buffer))!;
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const outDefault = effectiveDefaultWidth(await sheetFormatPr(out));
        expect(outDefault).not.toBeNull();
        // The round-tripped default stays at the source's wide value, NOT
        // exceljs's 8.43 — column B no longer shrinks.
        expect(outDefault!).toBeGreaterThan(8.43);
        expect(outDefault!).toBeCloseTo(srcDefault, 1);
    });

    test('a workbook with no sheetFormatPr default does not gain a spurious one', async () => {
        // 01-bar-simple has explicit per-column widths but (likely) no
        // baseColWidth/defaultColWidth override; export must not invent a
        // default that would resize its columns.
        const buf = readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const sheet = Object.values(
            (snap as { sheets: Record<string, { defaultColWidthChars?: number }> }).sheets,
        )[0];
        // Either the source genuinely had a default (captured) or it
        // didn't (undefined). When undefined, export must not write one.
        if (sheet.defaultColWidthChars === undefined) {
            const out = await snapshotToXlsxBuffer(snap);
            const fmt = await sheetFormatPr(out);
            if (fmt) {
                expect(/\bdefaultColWidth=/i.test(fmt)).toBe(false);
            }
        } else {
            expect(sheet.defaultColWidthChars).toBeGreaterThan(0);
        }
    });
});
