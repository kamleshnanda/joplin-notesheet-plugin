// M13 workstream C: rotated text round-trip.
//
// The fixture-anchored happy-path tests live in m12FixtureRoundTrip.test.ts
// (CCW 45°, 90°, CW 45° encoded as -45). This file covers the edge cases
// that the fixture doesn't exercise:
//
//   1. Excel's "vertical stacked" rotation mode (textRotation=255 in
//      OOXML, surfaced as the string 'vertical' in exceljs's API). Maps
//      to Univer's ITextRotation.v=1.
//   2. Round-trip stability: import → export → re-import preserves the
//      angle exactly.
//   3. Cells with no rotation must NOT emit a tr field.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import ExcelJS from 'exceljs';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, { v?: unknown; s?: string }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
}

function styleAt(snap: SnapshotShape, sheetIdx: number, row: number, col: number): Record<string, unknown> | null {
    const sheet = snap.sheets[snap.sheetOrder[sheetIdx]];
    const cell = sheet.cellData[row]?.[col];
    return cell?.s ? snap.styles[cell.s] : null;
}

async function importBuf(buf: Buffer): Promise<SnapshotShape> {
    return await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
}

describe('M13 — rotated text edge cases', () => {
    test('import: textRotation=\'vertical\' (Excel mode 255) maps to {a:0, v:1}', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'stacked';
        ws.getCell('A1').alignment = { textRotation: 'vertical' };
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf);
        const tr = styleAt(snap, 0, 0, 0)?.tr as { a?: number; v?: number } | undefined;
        expect(tr).toBeDefined();
        expect(tr?.v).toBe(1);
        // Univer encodes vertical-stacked with a=0, v=1.
        expect(tr?.a).toBe(0);
    });

    test('import: cells without rotation do NOT emit tr', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'plain';
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf);
        const style = styleAt(snap, 0, 0, 0);
        // Either the cell has no style at all (everything default) or it
        // has a style but no tr field. Both are acceptable; the
        // invariant we care about is that we don't emit tr={a:0} as a
        // no-op, which would bloat the snapshot and confuse Univer.
        expect(style?.tr).toBeUndefined();
    });

    test('import: textRotation=0 explicitly set is treated as no-rotation', async () => {
        // Some workbooks (defensively saved) write textRotation=0
        // explicitly. exceljs surfaces this as alignment.textRotation = 0.
        // We treat it the same as "no rotation": don't bloat style.tr.
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'zero';
        ws.getCell('A1').alignment = { textRotation: 0 };
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf);
        expect(styleAt(snap, 0, 0, 0)?.tr).toBeUndefined();
    });

    test('round-trip: each rotation angle survives export → re-import unchanged', async () => {
        // Build a workbook with one cell per angle, round-trip, re-import,
        // confirm the angles match exactly. Catches any encoding drift in
        // the export-side mapping (e.g. signed-int vs OOXML's 91..180 CW
        // encoding for negative angles).
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        const angles = [15, 30, 45, 60, 75, 90, -15, -30, -45, -60, -75, -90];
        angles.forEach((a, i) => {
            const c = ws.getCell(i + 1, 1);
            c.value = `rot_${a}`;
            c.alignment = { textRotation: a };
        });
        const buf0 = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf0);
        const buf1 = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const wb1 = new ExcelJS.Workbook();
        await wb1.xlsx.load(buf1 as unknown as Parameters<typeof wb1.xlsx.load>[0]);
        const ws1 = wb1.getWorksheet('Sheet1')!;
        angles.forEach((a, i) => {
            expect(ws1.getCell(i + 1, 1).alignment?.textRotation).toBe(a);
        });
    });

    test('round-trip: vertical stacked mode survives export → re-import as \'vertical\'', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.getCell('A1').value = 'stacked';
        ws.getCell('A1').alignment = { textRotation: 'vertical' };
        const buf0 = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await importBuf(buf0);
        const buf1 = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const wb1 = new ExcelJS.Workbook();
        await wb1.xlsx.load(buf1 as unknown as Parameters<typeof wb1.xlsx.load>[0]);
        const ws1 = wb1.getWorksheet('Sheet1')!;
        expect(ws1.getCell('A1').alignment?.textRotation).toBe('vertical');
    });
});
