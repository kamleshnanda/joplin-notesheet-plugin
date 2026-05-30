// Round-trip test for M9: a Univer snapshot containing table resource data
// must produce an .xlsx whose tables can be read back by exceljs (and by
// Excel itself). This was the bug — pre-M9 the exporter dropped tables.

import ExcelJS from 'exceljs';
import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, { v?: unknown; f?: string; s?: string }>>;
    }>;
    styles: Record<string, Record<string, unknown>>;
    resources?: Array<{ name: string; data: string }>;
}

describe('M9 — export preserves tables from snapshot resources', () => {
    test('snapshot with table resource → exported xlsx contains the table', async () => {
        // Build a snapshot with one table on Sheet1 spanning A1:B3 (header + 2 data rows).
        const snapshot = {
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
                        0: {
                            0: { v: 'Vendor', t: 1 },
                            1: { v: 'Investment', t: 1 },
                        },
                        1: {
                            0: { v: 'Acme', t: 1 },
                            1: { v: 100, t: 2 },
                        },
                        2: {
                            0: { v: 'Beta', t: 1 },
                            1: { v: 250, t: 2 },
                        },
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
            // Univer tables persist here. Exact shape comes from
            // @univerjs/sheets-table; this is a minimal plausible value.
            resources: [{
                name: 'SHEET_TABLE_PLUGIN',
                data: JSON.stringify({
                    s1: {
                        tables: [{
                            id: 'tbl-1',
                            name: 'Table1',
                            range: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
                            options: { showHeader: true },
                            filters: {},
                            columns: [
                                { id: 'col-vendor', displayName: 'Vendor', dataType: 'string', formula: '', meta: {}, style: {} },
                                { id: 'col-inv', displayName: 'Investment', dataType: 'number', formula: '', meta: {}, style: {} },
                            ],
                            meta: {},
                        }],
                        tableFilteredOutRows: [],
                    },
                }),
            }],
        };

        const xlsxBuf = await snapshotToXlsxBuffer(snapshot as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const wb = new ExcelJS.Workbook();
        await wb.xlsx.load(Buffer.from(xlsxBuf) as unknown as Parameters<typeof wb.xlsx.load>[0]);

        const ws = wb.getWorksheet('Sheet1');
        expect(ws).toBeTruthy();

        // Cells should still be present.
        expect(ws!.getCell('A1').value).toBe('Vendor');
        expect(ws!.getCell('B2').value).toBe(100);

        // M9 makes export emit table definitions. Confirm the table came
        // through with the right name and column shape.
        const tables = (ws as unknown as { getTables: () => Array<{ name: string; table: { columns: Array<{ name: string }> } }> }).getTables();
        expect(tables.length).toBe(1);
        expect(tables[0].name).toBe('Table1');
        expect(tables[0].table.columns.map((c) => c.name)).toEqual(['Vendor', 'Investment']);
    });
});

describe('M9 — import builds SHEET_TABLE_PLUGIN resource from xlsx tables', () => {
    test('xlsx with one named table → snapshot resource carries it', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Sheet1');
        ws.addTable({
            name: 'Table1',
            ref: 'A1',
            headerRow: true,
            columns: [{ name: 'Vendor' }, { name: 'Investment' }],
            rows: [
                ['Acme', 100],
                ['Beta', 250],
            ],
        });
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as SnapshotShape;

        expect(Array.isArray(snap.resources)).toBe(true);
        const tableEntry = snap.resources!.find((r) => r.name === 'SHEET_TABLE_PLUGIN');
        expect(tableEntry).toBeDefined();
        // Critical contract: data is JSON-stringified, not an object.
        expect(typeof tableEntry!.data).toBe('string');
        const parsed = JSON.parse(tableEntry!.data) as Record<string, { tables: Array<{ name: string; range: { endRow: number }; columns: Array<{ displayName: string }> }> }>;
        const sheetIds = Object.keys(parsed);
        expect(sheetIds.length).toBe(1);
        const tables = parsed[sheetIds[0]].tables;
        expect(tables.length).toBe(1);
        expect(tables[0].name).toBe('Table1');
        expect(tables[0].columns.map((c) => c.displayName)).toEqual(['Vendor', 'Investment']);
        // Range covers header + 2 data rows = rows 0..2 (0-based).
        expect(tables[0].range.endRow).toBe(2);

        // The schema-version sibling resource is also present.
        const schemaEntry = snap.resources!.find((r) => r.name === 'NOTESHEET_TABLE_SCHEMA');
        expect(schemaEntry).toBeDefined();
        expect(JSON.parse(schemaEntry!.data)).toEqual({ version: '0.23' });
    });

    test('xlsx without tables → no SHEET_TABLE_PLUGIN resource emitted', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('Plain');
        ws.getCell('A1').value = 'no table here';
        const buf = Buffer.from((await wb.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as SnapshotShape;

        // Either resources is undefined OR it has no SHEET_TABLE_PLUGIN entry.
        const entry = snap.resources?.find((r) => r.name === 'SHEET_TABLE_PLUGIN');
        expect(entry).toBeUndefined();
    });

    test('full xlsx → snapshot → xlsx round-trip preserves the table', async () => {
        const wb1 = new ExcelJS.Workbook();
        const ws1 = wb1.addWorksheet('Sheet1');
        ws1.addTable({
            name: 'SalesTable',
            ref: 'A1',
            headerRow: true,
            columns: [{ name: 'Region' }, { name: 'Q1' }, { name: 'Q2' }],
            rows: [
                ['North', 100, 120],
                ['South', 80, 95],
            ],
        });
        const buf1 = Buffer.from((await wb1.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf1);
        const buf2 = await snapshotToXlsxBuffer(snap);

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(Buffer.from(buf2) as unknown as Parameters<typeof wb2.xlsx.load>[0]);
        const ws2 = wb2.getWorksheet('Sheet1')!;
        const tables = (ws2 as unknown as { getTables: () => Array<{ name: string; table: { columns: Array<{ name: string }> } }> }).getTables();
        expect(tables.length).toBe(1);
        expect(tables[0].name).toBe('SalesTable');
        expect(tables[0].table.columns.map((c) => c.name)).toEqual(['Region', 'Q1', 'Q2']);
        // Data cells survived.
        expect(ws2.getCell('A2').value).toBe('North');
        expect(ws2.getCell('B3').value).toBe(80);
    });
});

describe('M9 — table import bypasses exceljs read bugs', () => {
    test('table column with <calculatedColumnFormula> nested child does not truncate the column list', async () => {
        // Reproduce the exact bug: exceljs's <tableColumn> parser breaks when
        // a column has nested children, causing it to drop subsequent columns.
        // Build raw OOXML by hand because exceljs can't write nested column
        // children. We assemble a minimal xlsx with a broken-from-exceljs's-
        // perspective table.xml and verify our reader handles all 8 columns.
        const JSZipMod = (await import('jszip')).default;
        const zip = new JSZipMod();
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
        zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
        zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
        zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
        zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`);
        // Worksheet has 8 columns of headers in row 1 and one data row in row 2.
        zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:H2"/>
<sheetData>
<row r="1"><c r="A1" t="str"><v>Vendor</v></c><c r="B1" t="str"><v>Start Date</v></c><c r="C1" t="str"><v>Investment</v></c><c r="D1" t="str"><v>Balance</v></c><c r="E1" t="str"><v>Profit</v></c><c r="F1" t="str"><v>Quoted Return</v></c><c r="G1" t="str"><v>Type</v></c><c r="H1" t="str"><v>Producing</v></c></row>
<row r="2"><c r="A2" t="str"><v>Acme</v></c><c r="C2"><v>100</v></c></row>
</sheetData>
<tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`);
        zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`);
        // Table.xml: 8 columns; the 5th column has a <calculatedColumnFormula>
        // child (the bug-trigger). headerRowCount attribute is absent (per
        // OOXML spec the default is 1, but exceljs reads false).
        zip.file('xl/tables/table1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:H2" totalsRowShown="0">
<tableColumns count="8">
<tableColumn id="1" name="Vendor"/>
<tableColumn id="2" name="Start Date"/>
<tableColumn id="3" name="Investment"/>
<tableColumn id="4" name="Balance"/>
<tableColumn id="5" name="Profit"><calculatedColumnFormula>(Table1[[#This Row],[Balance]]-Table1[[#This Row],[Investment]])/Table1[[#This Row],[Investment]]</calculatedColumnFormula></tableColumn>
<tableColumn id="6" name="Quoted Return"/>
<tableColumn id="7" name="Type"/>
<tableColumn id="8" name="Producing"/>
</tableColumns>
<tableStyleInfo name="TableStyleMedium2"/>
</table>`);

        const buf = Buffer.from(await zip.generateAsync({ type: 'nodebuffer' }));
        const snap = (await xlsxBufferToSnapshot(buf)) as unknown as SnapshotShape;

        const entry = snap.resources!.find((r) => r.name === 'SHEET_TABLE_PLUGIN');
        expect(entry).toBeDefined();
        const parsed = JSON.parse(entry!.data) as Record<string, { tables: Array<{ name: string; columns: Array<{ displayName: string }>; options: { showHeader?: boolean } }> }>;
        const sheetId = Object.keys(parsed)[0];
        const tables = parsed[sheetId].tables;
        expect(tables.length).toBe(1);
        // The whole point: 8 columns, not 5.
        expect(tables[0].columns.map((c) => c.displayName)).toEqual([
            'Vendor', 'Start Date', 'Investment', 'Balance', 'Profit', 'Quoted Return', 'Type', 'Producing',
        ]);
        // OOXML default for missing headerRowCount is 1 (true).
        expect(tables[0].options.showHeader).toBe(true);
    });
});

describe('M9 — table style + stripes round-trip', () => {
    test('Excel table style name and showRowStripes survive xlsx → snapshot → xlsx', async () => {
        const JSZipMod = (await import('jszip')).default;
        const zip = new JSZipMod();
        // Same shape as the column-truncation test, but the tableStyleInfo
        // declares TableStyleMedium2 + showRowStripes="1" (matching the
        // user's real source file).
        zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
        zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);
        zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);
        zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
        zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>`);
        zip.file('xl/worksheets/sheet1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:B2"/>
<sheetData>
<row r="1"><c r="A1" t="str"><v>X</v></c><c r="B1" t="str"><v>Y</v></c></row>
<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>
</sheetData>
<tableParts count="1"><tablePart r:id="rId1"/></tableParts>
</worksheet>`);
        zip.file('xl/worksheets/_rels/sheet1.xml.rels', `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>
</Relationships>`);
        zip.file('xl/tables/table1.xml', `<?xml version="1.0" encoding="UTF-8"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B2">
<tableColumns count="2">
<tableColumn id="1" name="X"/>
<tableColumn id="2" name="Y"/>
</tableColumns>
<tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/>
</table>`);

        const buf1 = Buffer.from(await zip.generateAsync({ type: 'nodebuffer' }));
        const snap = await xlsxBufferToSnapshot(buf1);
        const buf2 = await snapshotToXlsxBuffer(snap);

        // Re-open and check the table.xml has the right style + stripes.
        const zip2 = await JSZipMod.loadAsync(buf2 as ArrayBuffer);
        const tableXml = await zip2.files['xl/tables/table1.xml'].async('string');
        expect(tableXml).toContain('name="TableStyleMedium2"');
        expect(tableXml).toContain('showRowStripes="1"');
    });
});

describe('M9 — borders round-trip', () => {
    test('cell borders survive xlsx → snapshot → xlsx', async () => {
        const wb1 = new ExcelJS.Workbook();
        const ws1 = wb1.addWorksheet('Sheet1');
        const c = ws1.getCell('A1');
        c.value = 'bordered';
        c.border = {
            top: { style: 'thin', color: { argb: 'FF0000FF' } },
            bottom: { style: 'thick', color: { argb: 'FFFF0000' } },
            left: { style: 'medium' },
            right: { style: 'dashed' },
        };
        const buf1 = Buffer.from((await wb1.xlsx.writeBuffer()) as ArrayBuffer);
        const snap = await xlsxBufferToSnapshot(buf1);
        const buf2 = await snapshotToXlsxBuffer(snap);

        const wb2 = new ExcelJS.Workbook();
        await wb2.xlsx.load(Buffer.from(buf2) as unknown as Parameters<typeof wb2.xlsx.load>[0]);
        const ws2 = wb2.getWorksheet('Sheet1')!;
        const cell = ws2.getCell('A1');
        expect(cell.border?.top?.style).toBe('thin');
        expect(cell.border?.top?.color?.argb?.toUpperCase()).toBe('FF0000FF');
        expect(cell.border?.bottom?.style).toBe('thick');
        expect(cell.border?.bottom?.color?.argb?.toUpperCase()).toBe('FFFF0000');
        expect(cell.border?.left?.style).toBe('medium');
        expect(cell.border?.right?.style).toBe('dashed');
    });
});
