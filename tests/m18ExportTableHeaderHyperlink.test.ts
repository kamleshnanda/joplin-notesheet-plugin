// Regression: a hyperlink in a cell that is ALSO a table header must survive
// export to .xlsx.
//
// The bug (found in live manual testing): exceljs's ws.addTable().store()
// overwrites every header-row cell with a PLAIN STRING (the column name),
// clobbering the { text, hyperlink } value snapshotToXlsxBuffer had already
// written from cellData. Result: the exported workbook showed the header text
// but the hyperlink was gone. Data-row hyperlinks were fine (only headers are
// rewritten by addTable), and the existing hyperlink export test used a bare
// cell with NO table — so nothing covered the header-cell intersection.
//
// The fix re-applies the hyperlink (and rich text) to header cells AFTER
// addTable runs.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import ExcelJS from 'exceljs';

import { snapshotToXlsxBuffer } from '../src/xlsx';

// A cell.p carrying a HYPERLINK custom range (rangeType 0) — the shape the
// importer produces and the exporter reads.
function hyperlinkCellP(text: string, url: string) {
    return {
        id: '__INTERNAL_EDITOR__DOCS_NORMAL',
        body: {
            dataStream: text,
            customRanges: [
                {
                    startIndex: 0,
                    endIndex: text.length - 1,
                    rangeId: 'r1',
                    rangeType: 0,
                    properties: { url },
                },
            ],
        },
    };
}

describe('M18: hyperlink in a table-header cell survives export', () => {
    test('header cell A1 keeps its hyperlink after addTable', async () => {
        const snap = {
            id: 'wb-th',
            name: 'Spreadsheet',
            appVersion: '0.1.0',
            locale: 'enUS',
            sheetOrder: ['s1'],
            styles: {},
            sheets: {
                s1: {
                    id: 's1',
                    name: 'Sheet1',
                    cellData: {
                        // Row 0 = table header row. A1 is the header AND a link.
                        0: {
                            0: {
                                v: 'Vendor',
                                t: 1,
                                p: hyperlinkCellP('Vendor', 'https://vendor.example/'),
                            },
                            1: { v: 'Investment', t: 1 },
                        },
                        1: { 0: { v: 'Acme', t: 1 }, 1: { v: 100, t: 2 } },
                        2: { 0: { v: 'Globex', t: 1 }, 1: { v: 200, t: 2 } },
                    },
                    rowCount: 100,
                    columnCount: 26,
                    defaultColumnWidth: 73,
                    defaultRowHeight: 19,
                    mergeData: [],
                    rowData: {},
                    columnData: {},
                },
            },
            resources: [
                {
                    name: 'SHEET_TABLE_PLUGIN',
                    data: JSON.stringify({
                        s1: {
                            tables: [
                                {
                                    id: 'tbl-1',
                                    name: 'Table1',
                                    range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
                                    options: { showHeader: true },
                                    filters: {},
                                    columns: [
                                        {
                                            id: 'c0',
                                            displayName: 'Vendor',
                                            dataType: 'string',
                                            formula: '',
                                            meta: {},
                                            style: {},
                                        },
                                        {
                                            id: 'c1',
                                            displayName: 'Investment',
                                            dataType: 'number',
                                            formula: '',
                                            meta: {},
                                            style: {},
                                        },
                                    ],
                                    meta: {},
                                },
                            ],
                            tableFilteredOutRows: [],
                        },
                    }),
                },
            ],
        };

        const buf = await snapshotToXlsxBuffer(
            snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0],
        );

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(buf as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet('Sheet1')!;
        const a1 = ws.getCell('A1');

        // The header text is preserved AND the hyperlink survived addTable.
        expect(a1.text).toBe('Vendor');
        expect(a1.isHyperlink).toBe(true); // was false before the fix (clobbered)
        expect(a1.hyperlink).toBe('https://vendor.example/');

        // Sanity: the table itself still exported.
        expect(ws.getTable('Table1')).toBeDefined();
    });
});
