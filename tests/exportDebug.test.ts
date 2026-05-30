// Standalone debug test to write a real-world-shaped export to /tmp so we
// can unzip and inspect the generated table.xml. Skipped in CI, run by name
// when needed:  npm test -- --testPathPattern=exportDebug
import { writeFileSync } from 'fs';
import ExcelJS from 'exceljs';
import { snapshotToXlsxBuffer } from '../src/xlsx';

describe.skip('export debug — writes /tmp/test-export.xlsx for manual inspection', () => {
    test('round-trip yields a table.xml Excel can open', async () => {
        const snap = {
            id: 'wb-1',
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
                        0: { 0: { v: 'Vendor', t: 1 }, 1: { v: 'Investment', t: 1 }, 2: { v: 'Profit', t: 1 } },
                        1: { 0: { v: 'Acme', t: 1 }, 1: { v: 100, t: 2 }, 2: { f: '=Table1[[#This Row],[Investment]]*0.05', v: 5, t: 2 } },
                        2: { 0: { v: 'Beta', t: 1 }, 1: { v: 200, t: 2 }, 2: { f: '=Table1[[#This Row],[Investment]]*0.05', v: 10, t: 2 } },
                        3: { 0: { v: 'Gamma', t: 1 }, 1: { v: 300, t: 2 }, 2: { f: '=Table1[[#This Row],[Investment]]*0.05', v: 15, t: 2 } },
                    },
                    rowCount: 100, columnCount: 26, defaultColumnWidth: 73, defaultRowHeight: 19,
                    mergeData: [], rowData: {}, columnData: {},
                },
            },
            resources: [{
                name: 'SHEET_TABLE_PLUGIN',
                data: JSON.stringify({
                    s1: {
                        tables: [{
                            id: 'tbl-1', name: 'Table1',
                            range: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 2 },
                            options: { showHeader: true, showFooter: false },
                            filters: { tableColumnFilterList: [] },
                            columns: [
                                { id: 'c1', displayName: 'Vendor', dataType: 'string', formula: '', meta: {}, style: {} },
                                { id: 'c2', displayName: 'Investment', dataType: 'number', formula: '', meta: {}, style: {} },
                                { id: 'c3', displayName: 'Profit', dataType: 'number', formula: '', meta: {}, style: {} },
                            ],
                            meta: {},
                        }],
                        tableFilteredOutRows: [],
                    },
                }),
            }],
        };
        const buf = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        writeFileSync('/tmp/test-export.xlsx', Buffer.from(buf));
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(Buffer.from(buf) as unknown as Parameters<typeof wb.xlsx.load>[0]);
        const ws = wb.getWorksheet('Sheet1')!;
        const tables = (ws as unknown as { getTables: () => Array<{ name: string; table: { tableRef?: string; columns: Array<{ name: string }> } }> }).getTables();
        console.log('table count:', tables.length);
        console.log('table[0]:', JSON.stringify({
            name: tables[0]?.name,
            tableRef: tables[0]?.table.tableRef,
            columns: tables[0]?.table.columns.map((c) => c.name),
        }));
    });
});
