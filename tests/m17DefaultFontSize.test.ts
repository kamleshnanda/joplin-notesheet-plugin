// M17 fidelity: the workbook's DEFAULT font size must survive the
// import → export round-trip.
//
// The bug: source workbooks use Aptos Narrow (Body) 12, stored as the
// first <font><sz val="12"/> in xl/styles.xml (size-less cells inherit
// it). exceljs doesn't surface this and writes its own 11pt fallback on
// export, so a 12pt workbook came out 11pt. We now read the default size
// at import (defaultStyle.fs) and patch the exported styles.xml default
// font back.
//
// Anchored to the Excel-canonical source: read the source's default
// size, then assert the export preserves it — NOT a hardcoded literal.

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

// The default font size = the <sz> on the FIRST <font> in <fonts>
// (the workbook default that size-less cells inherit).
async function defaultFontSize(buffer: ArrayBuffer | Buffer): Promise<number | null> {
    const zip = await JSZip.loadAsync(buffer);
    const stylesPath = Object.keys(zip.files).find((p) => /^xl\/styles\.xml$/i.test(p));
    if (!stylesPath) return null;
    const xml = await zip.file(stylesPath)!.async('string');
    const fonts = xml.match(/<fonts\b[^>]*>([\s\S]*?)<\/fonts>/);
    if (!fonts) return null;
    const firstFont = fonts[1].match(/<font>([\s\S]*?)<\/font>/);
    if (!firstFont) return null;
    const sz = firstFont[1].match(/<sz\s+val="([\d.]+)"/);
    return sz ? parseFloat(sz[1]) : null;
}

describe('M17 default font size round-trip', () => {
    test('01-bar-simple.xlsx: source default font is 12pt (Excel-canonical anchor)', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const srcSize = await defaultFontSize(buf as unknown as Buffer);
        expect(srcSize).toBe(12);
    });

    test('import captures the default font size onto defaultStyle.fs', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const srcSize = await defaultFontSize(buf as unknown as Buffer)!;
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const ds = (snap as { defaultStyle?: { fs?: number } }).defaultStyle;
        expect(ds?.fs).toBe(srcSize);
    });

    test('export preserves the default font size (no 12 → 11 shrink)', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const srcSize = await defaultFontSize(buf as unknown as Buffer)!;
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const outSize = await defaultFontSize(out);
        expect(outSize).toBe(srcSize);
        // Specifically: it did NOT collapse to exceljs's 11pt fallback.
        expect(outSize).not.toBe(11);
    });

    test('round-trip is idempotent: re-importing the export keeps 12pt', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap1);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(out) as unknown as Buffer);
        const ds = (snap2 as { defaultStyle?: { fs?: number } }).defaultStyle;
        expect(ds?.fs).toBe(12);
    });
});
