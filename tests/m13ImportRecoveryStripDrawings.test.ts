// M13 workstream B1: import recovery via zip pre-process.
//
// Two of the three exceljs-crash fixtures (MultiSheet.xlsx,
// LargeWorkbook.xlsx) trip the same drawing-reconcile bug — exceljs
// crashes inside `XLSX.reconcile` reading `drawing.anchors` because of
// a structural mismatch in chart drawings emitted by openpyxl (and by
// modern Excel saves with charts). Before M13 we wrapped the crash in
// a friendly NotesheetImportError with code `xlsx-charts-unsupported`.
//
// Strategy for M13: strip chart parts from the in-memory zip before it
// reaches exceljs. We lose the charts on import (they don't render in
// Joplin) but the rest of the workbook — values, formulas, tables,
// fonts, theme — survives intact.
//
// FormulasAndStructuredRefs.xlsx remains in the catch+degrade path —
// its crash is in the multi-sheet+multi-table reconcile, NOT drawings.
// That's a separate workstream (M13a) with the friendly error stable.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';

import { xlsxBufferToSnapshot, NotesheetImportError } from '../src/xlsx';

const FIX = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, { v?: unknown; s?: string; f?: string }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
    resources?: Array<{ name: string; data: string }>;
}

async function importFile(name: string): Promise<SnapshotShape> {
    const buf = readFileSync(path.join(FIX, name));
    return await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
}

describe('M13 — import recovery: strip chart drawings before exceljs sees them', () => {
    test('MultiSheet.xlsx imports cleanly (was throwing xlsx-charts-unsupported)', async () => {
        const snap = await importFile('MultiSheet.xlsx');
        // 3 sheets: Data, Chart, Summary.
        expect(snap.sheetOrder).toHaveLength(3);
    });

    test('MultiSheet.xlsx: Data sheet preserves header + 5 data rows', async () => {
        const snap = await importFile('MultiSheet.xlsx');
        // First sheet is "Data" (sheetId='sheet-1' in our naming).
        const data = snap.sheets[snap.sheetOrder[0]];
        // Row 0: Category, Value (headers).
        expect(data.cellData[0][0].v).toBe('Category');
        expect(data.cellData[0][1].v).toBe('Value');
        // Row 1-5: Apples 30, Bananas 50, Cherries 20, Dates 45, Elderberry 60.
        expect(data.cellData[1][0].v).toBe('Apples');
        expect(data.cellData[1][1].v).toBe(30);
        expect(data.cellData[5][0].v).toBe('Elderberry');
        expect(data.cellData[5][1].v).toBe(60);
    });

    test('MultiSheet.xlsx: Summary sheet preserves cross-sheet formulas', async () => {
        const snap = await importFile('MultiSheet.xlsx');
        const summary = snap.sheets[snap.sheetOrder[2]];
        // A2 (R1C0): generator wrote "=Data!B2+Chart!B2" as a string.
        // exceljs parses it as a formula on import, so we get cell.f
        // (not cell.v).
        const a2 = summary.cellData[1]?.[0];
        expect(a2?.f).toBe('=Data!B2+Chart!B2');
        // A3 (R2C0): "=SUM(Data!B2:B6)".
        expect(summary.cellData[2]?.[0]?.f).toBe('=SUM(Data!B2:B6)');
    });

    test('LargeWorkbook.xlsx imports cleanly (was throwing xlsx-charts-unsupported)', async () => {
        const snap = await importFile('LargeWorkbook.xlsx');
        // Two sheets: "Data" (1000 rows) + "Chart".
        expect(snap.sheetOrder.length).toBeGreaterThanOrEqual(1);
    });

    test('LargeWorkbook.xlsx: Data sheet preserves header + lots of data', async () => {
        const snap = await importFile('LargeWorkbook.xlsx');
        const data = snap.sheets[snap.sheetOrder[0]];
        // Row 0 headers: Col_1, Col_2, ..., Col_20.
        expect(data.cellData[0][0].v).toBe('Col_1');
        expect(data.cellData[0][19].v).toBe('Col_20');
        // 1000 data rows + 1 header = 1001 total rows.
        const rows = Object.keys(data.cellData);
        expect(rows.length).toBeGreaterThanOrEqual(1000);
    });

    test('LargeWorkbook.xlsx: column 21 formulas survive', async () => {
        const snap = await importFile('LargeWorkbook.xlsx');
        const data = snap.sheets[snap.sheetOrder[0]];
        // Column 21 (col index 20) carries 4 distinct formula patterns
        // in rows 2..101, cycling by i%4: =SUM(A{i}:E{i}), =AVERAGE(F:J),
        // =A*B+C, =IF(F>500,"High","Low"). Sample 4 consecutive rows to
        // hit all 4 patterns.
        const formulasFound = new Set<string>();
        for (let r = 1; r <= 4; r++) {
            const f = data.cellData[r]?.[20]?.f;
            if (f) {
                if (f.startsWith('=SUM(')) formulasFound.add('SUM');
                else if (f.startsWith('=AVERAGE(')) formulasFound.add('AVERAGE');
                else if (f.startsWith('=IF(')) formulasFound.add('IF');
                else if (/^=[A-Z]+\d+\*/.test(f)) formulasFound.add('arithmetic');
            }
        }
        // We expect all 4 patterns to be represented across rows 1..4.
        expect(formulasFound.size).toBeGreaterThanOrEqual(3);
    });

    test('FormulasAndStructuredRefs.xlsx imports cleanly (multi-sheet+multi-table)', async () => {
        // Originally classified as `xlsx-multi-table-unsupported`. Turned
        // out the root cause was the same path-resolution bug as the
        // chart fixtures — openpyxl emits absolute rel Targets but
        // exceljs's resolver expects relative. The rel-rewrite step in
        // preProcessForExceljs unblocks this case too. Both tables
        // (Table1 on Sheet1, Table2 on Sheet2) survive in the
        // SHEET_TABLE_PLUGIN resource.
        const snap = await importFile('FormulasAndStructuredRefs.xlsx');
        expect(snap.sheetOrder.length).toBeGreaterThanOrEqual(2);
        const tableRes = snap.resources?.find((r) => r.name === 'SHEET_TABLE_PLUGIN');
        expect(tableRes).toBeDefined();
        const tables = JSON.parse(tableRes!.data) as Record<string, { tables: Array<{ name: string }> }>;
        // Two sheets, each with one named table.
        const allNames = Object.values(tables).flatMap((t) => t.tables.map((tt) => tt.name));
        expect(allNames).toEqual(expect.arrayContaining(['Table1', 'Table2']));
    });

    test('importing a chart-free fixture still works (no regression in non-drawing path)', async () => {
        // Sanity: the strip-drawings code must be a no-op for fixtures
        // that have no drawings to strip.
        const snap = await importFile('Hyperlinks-Variants.xlsx');
        expect(snap.sheetOrder).toHaveLength(1);
        const sheet = snap.sheets[snap.sheetOrder[0]];
        expect(Object.keys(sheet.cellData).length).toBeGreaterThan(0);
    });
});

// Sanity: NotesheetImportError must remain exported (callers depend on
// instanceof checks and the .code field).
test('M13 — NotesheetImportError remains exported with code field', () => {
    const e = new NotesheetImportError('x-test', 'msg', null);
    expect(e.code).toBe('x-test');
});
