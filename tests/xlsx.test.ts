import ExcelJS from 'exceljs';
import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

// Build a fresh xlsx from exceljs and feed it through xlsxBufferToSnapshot.
async function buildXlsx(
    populate: (ws: ExcelJS.Worksheet, wb: ExcelJS.Workbook) => void,
): Promise<Buffer> {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    populate(ws, wb);
    const buf = await wb.xlsx.writeBuffer();
    return Buffer.from(buf as ArrayBuffer);
}

interface CellRecord {
    v?: string | number | boolean;
    f?: string;
    t?: number;
    s?: string;
}

interface Snapshot {
    sheetOrder: string[];
    sheets: Record<
        string,
        {
            id: string;
            name: string;
            cellData: Record<number, Record<number, CellRecord>>;
            mergeData?: Array<{
                startRow: number;
                endRow: number;
                startColumn: number;
                endColumn: number;
            }>;
        }
    >;
    styles: Record<string, Record<string, unknown>>;
}

describe('xlsx → snapshot import', () => {
    test('values: string, number, boolean', async () => {
        const buf = await buildXlsx((ws) => {
            ws.getCell('A1').value = 'hello';
            ws.getCell('A2').value = 42;
            ws.getCell('A3').value = true;
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        expect(sheet.cellData[0][0]).toEqual({ v: 'hello', t: 1 });
        expect(sheet.cellData[1][0]).toEqual({ v: 42, t: 2 });
        expect(sheet.cellData[2][0]).toEqual({ v: true, t: 3 });
    });

    test('formula cell preserves formula and cached result', async () => {
        const buf = await buildXlsx((ws) => {
            ws.getCell('A1').value = 1;
            ws.getCell('A2').value = 2;
            ws.getCell('A3').value = {
                formula: 'SUM(A1:A2)',
                result: 3,
            } as ExcelJS.CellFormulaValue;
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        expect(sheet.cellData[2][0].f).toBe('=SUM(A1:A2)');
        expect(sheet.cellData[2][0].v).toBe(3);
    });

    test('font formatting (bold + italic + color) is captured', async () => {
        const buf = await buildXlsx((ws) => {
            const c = ws.getCell('A1');
            c.value = 'styled';
            c.font = {
                bold: true,
                italic: true,
                color: { argb: 'FFFF0000' },
                size: 14,
                name: 'Arial',
            };
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const styleId = sheet.cellData[0][0].s as string;
        expect(styleId).toBeDefined();
        const style = snap.styles[styleId];
        expect(style.bl).toBe(1);
        expect(style.it).toBe(1);
        expect(style.fs).toBe(14);
        expect(style.ff).toBe('Arial');
        expect(style.cl).toEqual({ rgb: '#FF0000' });
    });

    test('background fill is captured', async () => {
        const buf = await buildXlsx((ws) => {
            const c = ws.getCell('A1');
            c.value = 'bg';
            c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const style = snap.styles[sheet.cellData[0][0].s as string];
        expect(style.bg).toEqual({ rgb: '#FFFF00' });
    });

    test('alignment and wrap text', async () => {
        const buf = await buildXlsx((ws) => {
            const c = ws.getCell('A1');
            c.value = 'aligned';
            c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const style = snap.styles[sheet.cellData[0][0].s as string];
        expect(style.ht).toBe(2);
        expect(style.vt).toBe(2);
        expect(style.tb).toBe(3);
    });

    test('number format pattern is captured', async () => {
        const buf = await buildXlsx((ws) => {
            const c = ws.getCell('A1');
            c.value = 1234.5;
            c.numFmt = '#,##0.00';
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const style = snap.styles[sheet.cellData[0][0].s as string];
        expect(style.n).toEqual({ pattern: '#,##0.00' });
    });

    test('merged cells', async () => {
        const buf = await buildXlsx((ws) => {
            ws.getCell('A1').value = 'span';
            ws.mergeCells('A1:C2');
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        expect(sheet.mergeData).toEqual([{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 2 }]);
    });

    test('identical styles share a single style id', async () => {
        const buf = await buildXlsx((ws) => {
            const a = ws.getCell('A1');
            a.value = 'x';
            a.font = { bold: true };
            const b = ws.getCell('B1');
            b.value = 'y';
            b.font = { bold: true };
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        expect(sheet.cellData[0][0].s).toBe(sheet.cellData[0][1].s);
    });

    test('empty workbook yields one empty sheet', async () => {
        const wb = new ExcelJS.Workbook();
        wb.addWorksheet('Empty');
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        expect(snap.sheetOrder.length).toBe(1);
    });

    test('date cell stored as Excel serial number, not ISO string', async () => {
        // Excel renders dates by combining a numeric serial with the cell's
        // numFmt pattern. If we store the date as an ISO string, no formatter
        // fires and the user sees "2025-12-02T00:00:00.000Z" instead of
        // "12/2/2025". Confirm we serialize as numbers + carry the numFmt.
        const buf = await buildXlsx((ws) => {
            const c = ws.getCell('A1');
            c.value = new Date(Date.UTC(2025, 11, 2));
            c.numFmt = 'm/d/yy';
        });
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as Snapshot;
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[0][0];
        // 2025-12-02 UTC → 45,993 days since 1899-12-30.
        expect(cell.v).toBe(45993);
        expect(cell.t).toBe(2);
        const styleId = cell.s as string;
        expect(snap.styles[styleId]?.n).toEqual({ pattern: 'm/d/yy' });
    });
});

describe('snapshot → xlsx export', () => {
    test('round-trip preserves values, formulas, fonts, fills, alignment, numFmt, merges', async () => {
        const original: Snapshot = {
            sheetOrder: ['s1'],
            styles: {
                'style-1': {
                    bl: 1,
                    it: 1,
                    fs: 12,
                    ff: 'Calibri',
                    cl: { rgb: '#0000FF' },
                    bg: { rgb: '#CCCCCC' },
                    ht: 2,
                    vt: 2,
                    tb: 3,
                    n: { pattern: '#,##0.00' },
                },
            },
            sheets: {
                s1: {
                    id: 's1',
                    name: 'Round',
                    cellData: {
                        0: {
                            0: { v: 'header', t: 1, s: 'style-1' },
                            1: { v: 100, t: 2 },
                        },
                        1: {
                            0: { v: 50, t: 2 },
                        },
                        2: {
                            0: { f: '=SUM(A1:A2)', v: 150, t: 2 },
                        },
                    },
                    mergeData: [{ startRow: 0, endRow: 0, startColumn: 2, endColumn: 4 }],
                },
            },
        };

        const xlsx = await snapshotToXlsxBuffer(
            original as unknown as Parameters<typeof snapshotToXlsxBuffer>[0],
        );
        const reimported = (await xlsxBufferToSnapshot(Buffer.from(xlsx))) as unknown as Snapshot;
        const sheet = reimported.sheets[reimported.sheetOrder[0]];

        expect(sheet.cellData[0][0].v).toBe('header');
        expect(sheet.cellData[0][1].v).toBe(100);
        expect(sheet.cellData[1][0].v).toBe(50);
        expect(sheet.cellData[2][0].f).toBe('=SUM(A1:A2)');
        expect(sheet.cellData[2][0].v).toBe(150);

        const styleId = sheet.cellData[0][0].s as string;
        const style = reimported.styles[styleId];
        expect(style.bl).toBe(1);
        expect(style.it).toBe(1);
        expect(style.fs).toBe(12);
        expect(style.ff).toBe('Calibri');
        expect(style.cl).toEqual({ rgb: '#0000FF' });
        expect(style.bg).toEqual({ rgb: '#CCCCCC' });
        expect(style.ht).toBe(2);
        expect(style.vt).toBe(2);
        expect(style.tb).toBe(3);
        expect(style.n).toEqual({ pattern: '#,##0.00' });

        expect(sheet.mergeData).toEqual([{ startRow: 0, endRow: 0, startColumn: 2, endColumn: 4 }]);
    });

    test('empty snapshot exports a workbook with one sheet', async () => {
        const empty = {
            sheetOrder: [],
            sheets: {},
            styles: {},
        };
        const xlsx = await snapshotToXlsxBuffer(
            empty as unknown as Parameters<typeof snapshotToXlsxBuffer>[0],
        );
        expect(xlsx.byteLength).toBeGreaterThan(0);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(Buffer.from(xlsx) as unknown as Parameters<typeof wb.xlsx.load>[0]);
        expect(wb.worksheets.length).toBeGreaterThanOrEqual(1);
    });
});
