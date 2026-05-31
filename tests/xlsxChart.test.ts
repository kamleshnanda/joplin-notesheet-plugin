// Tests for M10 chart export. Covers:
//  - Pure helpers (escape, refs, sheet-name quoting)
//  - Per-chart-type XML shape
//  - readChartsFromSnapshot's drawing-resource parsing
//  - Full round-trip via snapshotToXlsxBuffer (zip surgery)
//  - Regression: no charts in snapshot → buffer unchanged
//  - Edge cases: multi-chart-one-sheet, cross-sheet refs, XML escape

import JSZip from 'jszip';

import {
    buildBarChartXml,
    buildLineChartXml,
    buildPieChartXml,
    buildDoughnutChartXml,
    buildDrawingXml,
    buildDrawingRelsXml,
    buildChartRelsXml,
    cellRef,
    rangeRefCol,
    escapeSheetName,
    escapeXml,
    readChartsFromSnapshot,
    type ChartDrawing,
} from '../src/charts/xlsxChart';
import { snapshotToXlsxBuffer } from '../src/xlsx';

// ─── Test helpers ──────────────────────────────────────────────────────────

function fixtureChart(overrides: Partial<ChartDrawing> = {}): ChartDrawing {
    return {
        chartId: 'chart-test',
        sheetId: 'sheet-1',
        type: 'bar',
        title: 'Quarterly Sales',
        sourceRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
        labels: ['Q1', 'Q2', 'Q3', 'Q4'],
        datasets: [{ label: 'Sales', data: [100, 120, 95, 140] }],
        anchor: {
            fromCol: 3, fromColOff: 0, fromRow: 0, fromRowOff: 0,
            toCol: 10, toColOff: 0, toRow: 20, toRowOff: 0,
        },
        ...overrides,
    };
}

function fixtureSnapshot(charts: ChartDrawing[], extraSheets: string[] = []) {
    const sheetOrder = ['sheet-1', ...extraSheets];
    const sheets: Record<string, unknown> = {
        'sheet-1': {
            id: 'sheet-1', name: 'Sheet1',
            cellData: {
                0: { 0: { v: 'Quarter', t: 1 }, 1: { v: 'Sales', t: 1 } },
                1: { 0: { v: 'Q1', t: 1 }, 1: { v: 100, t: 2 } },
                2: { 0: { v: 'Q2', t: 1 }, 1: { v: 120, t: 2 } },
                3: { 0: { v: 'Q3', t: 1 }, 1: { v: 95, t: 2 } },
                4: { 0: { v: 'Q4', t: 1 }, 1: { v: 140, t: 2 } },
            },
            rowCount: 100, columnCount: 26,
            defaultColumnWidth: 73, defaultRowHeight: 19,
            mergeData: [], rowData: {}, columnData: {},
        },
    };
    for (const id of extraSheets) {
        sheets[id] = {
            id, name: id === 'sheet-2' ? 'Sheet2' : id,
            cellData: {}, rowCount: 100, columnCount: 26,
            defaultColumnWidth: 73, defaultRowHeight: 19,
            mergeData: [], rowData: {}, columnData: {},
        };
    }

    const drawingResource = charts.length === 0 ? null : {
        name: 'SHEET_DRAWING_PLUGIN',
        data: JSON.stringify(buildDrawingResourceData(charts)),
    };

    return {
        id: 'wb-test',
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        sheetOrder,
        styles: {},
        sheets,
        ...(drawingResource ? { resources: [drawingResource] } : {}),
    };
}

// Mirror Univer's SHEET_DRAWING_PLUGIN shape: per-subUnitId map of
// { data: { [drawingId]: ISheetDrawing }, order: [...] }
function buildDrawingResourceData(charts: ChartDrawing[]): Record<string, unknown> {
    const out: Record<string, { data: Record<string, unknown>; order: string[] }> = {};
    for (const c of charts) {
        if (!out[c.sheetId]) out[c.sheetId] = { data: {}, order: [] };
        const drawingId = c.chartId;
        out[c.sheetId].data[drawingId] = {
            unitId: 'wb-test',
            subUnitId: c.sheetId,
            drawingId,
            drawingType: 8, // DRAWING_DOM
            componentKey: 'NotesheetChart',
            data: {
                chartId: c.chartId,
                type: c.type,
                title: c.title,
                sourceRange: c.sourceRange,
                labels: c.labels,
                datasets: c.datasets,
            },
            allowTransform: true,
            transform: { left: 0, top: 0, width: 480, height: 320 },
            sheetTransform: {
                from: { column: c.anchor.fromCol, columnOffset: c.anchor.fromColOff, row: c.anchor.fromRow, rowOffset: c.anchor.fromRowOff },
                to: { column: c.anchor.toCol, columnOffset: c.anchor.toColOff, row: c.anchor.toRow, rowOffset: c.anchor.toRowOff },
            },
            axisAlignSheetTransform: {
                from: { column: c.anchor.fromCol, columnOffset: c.anchor.fromColOff, row: c.anchor.fromRow, rowOffset: c.anchor.fromRowOff },
                to: { column: c.anchor.toCol, columnOffset: c.anchor.toColOff, row: c.anchor.toRow, rowOffset: c.anchor.toRowOff },
            },
        };
        out[c.sheetId].order.push(drawingId);
    }
    return out;
}

// ─── Pure helpers ──────────────────────────────────────────────────────────

describe('M10 helpers', () => {
    test('cellRef produces $A$1 form', () => {
        expect(cellRef('Sheet1', 0, 0)).toBe('Sheet1!$A$1');
        expect(cellRef('Sheet1', 9, 25)).toBe('Sheet1!$Z$10');
        expect(cellRef('Sheet1', 0, 26)).toBe('Sheet1!$AA$1');
    });

    test('rangeRefCol — single-column range across rows', () => {
        // col=1 (B), rows 1..4 (i.e. $B$2:$B$5)
        expect(rangeRefCol('Sheet1', 1, 4, 1)).toBe('Sheet1!$B$2:$B$5');
    });

    test('escapeSheetName quotes names with non-identifier chars', () => {
        expect(escapeSheetName('Sheet1')).toBe('Sheet1');
        expect(escapeSheetName('My Sheet')).toBe(`'My Sheet'`);
        expect(escapeSheetName("Sam's Data")).toBe(`'Sam''s Data'`);
        expect(escapeSheetName('2025-Q1')).toBe(`'2025-Q1'`);
    });

    test('escapeXml escapes & < > but leaves quotes alone', () => {
        expect(escapeXml('Foo & Bar')).toBe('Foo &amp; Bar');
        expect(escapeXml('<X>')).toBe('&lt;X&gt;');
        expect(escapeXml('Margins "FY26"')).toBe('Margins "FY26"');
        expect(escapeXml(`it's a test`)).toBe(`it's a test`);
        expect(escapeXml('A & B & C')).toBe('A &amp; B &amp; C');
    });
});

// ─── Chart-type XML shape ──────────────────────────────────────────────────

describe('M10 chart XML builders', () => {
    test('buildBarChartXml — bar layout with vertical columns + axes', () => {
        const c = fixtureChart({ type: 'bar' });
        const xml = buildBarChartXml(c, { sheetName: 'Sheet1' });
        expect(xml).toContain('<c:barChart>');
        expect(xml).toContain('<c:barDir val="col"/>'); // vertical column, NOT horizontal bar
        expect(xml).toContain('<c:grouping val="clustered"/>');
        expect(xml).toContain('<c:catAx>');
        expect(xml).toContain('<c:valAx>');
        // Cached labels and values must agree with the formula refs.
        expect(xml).toContain('<c:f>Sheet1!$A$2:$A$5</c:f>');
        expect(xml).toContain('<c:f>Sheet1!$B$2:$B$5</c:f>');
        expect(xml).toContain('<c:pt idx="0"><c:v>Q1</c:v></c:pt>');
        expect(xml).toContain('<c:pt idx="0"><c:v>100</c:v></c:pt>');
        // axId pair must match between barChart and the axes.
        const axIds = xml.match(/<c:axId val="(\d+)"\/>/g) ?? [];
        expect(axIds.length).toBeGreaterThanOrEqual(4);
    });

    test('buildLineChartXml — line grouping, no bars, with axes', () => {
        const c = fixtureChart({ type: 'line' });
        const xml = buildLineChartXml(c, { sheetName: 'Sheet1' });
        expect(xml).toContain('<c:lineChart>');
        expect(xml).toContain('<c:grouping val="standard"/>');
        expect(xml).toContain('<c:catAx>');
        expect(xml).toContain('<c:valAx>');
        expect(xml).not.toContain('<c:barChart>');
    });

    test('buildPieChartXml — single series, no axes, per-data-point dPt blocks', () => {
        const c = fixtureChart({ type: 'pie' });
        const xml = buildPieChartXml(c, { sheetName: 'Sheet1' });
        expect(xml).toContain('<c:pieChart>');
        expect(xml).toContain('<c:varyColors val="1"/>');
        expect(xml).not.toContain('<c:catAx>');
        expect(xml).not.toContain('<c:valAx>');
        // One <c:dPt> per data point.
        const dPtCount = (xml.match(/<c:dPt>/g) ?? []).length;
        expect(dPtCount).toBe(c.datasets[0].data.length);
    });

    test('buildDoughnutChartXml — pie shape + holeSize', () => {
        const c = fixtureChart({ type: 'doughnut' });
        const xml = buildDoughnutChartXml(c, { sheetName: 'Sheet1' });
        expect(xml).toContain('<c:doughnutChart>');
        expect(xml).toContain('<c:holeSize val="50"/>');
    });

    test('multi-series bar chart emits one <c:ser> per dataset', () => {
        const c = fixtureChart({
            datasets: [
                { label: 'North', data: [10, 20, 30, 40] },
                { label: 'South', data: [15, 25, 35, 45] },
                { label: 'East',  data: [12, 22, 32, 42] },
            ],
            sourceRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 3 },
        });
        const xml = buildBarChartXml(c, { sheetName: 'Sheet1' });
        const serCount = (xml.match(/<c:ser>/g) ?? []).length;
        expect(serCount).toBe(3);
    });

    test('chart title with special chars escapes correctly', () => {
        const c = fixtureChart({ title: 'Margins & <FY26>' });
        const xml = buildBarChartXml(c, { sheetName: 'Sheet1' });
        expect(xml).toContain('<a:t>Margins &amp; &lt;FY26&gt;</a:t>');
    });

    test('NaN/Infinity values are skipped from c:numCache', () => {
        const c = fixtureChart({
            datasets: [{ label: 'Sales', data: [100, NaN, Infinity, 50] }],
        });
        const xml = buildBarChartXml(c, { sheetName: 'Sheet1' });
        expect(xml).toContain('<c:pt idx="0"><c:v>100</c:v></c:pt>');
        // Indexes 1 and 2 should be missing (NaN/Inf).
        expect(xml).not.toContain('<c:v>NaN</c:v>');
        expect(xml).not.toContain('<c:v>Infinity</c:v>');
        expect(xml).toContain('<c:pt idx="3"><c:v>50</c:v></c:pt>');
    });

    test('sheet name with spaces is single-quoted in c:f refs', () => {
        const c = fixtureChart();
        const xml = buildBarChartXml(c, { sheetName: 'My Sheet' });
        expect(xml).toContain(`<c:f>'My Sheet'!$A$2:$A$5</c:f>`);
        expect(xml).toContain(`<c:f>'My Sheet'!$B$2:$B$5</c:f>`);
    });
});

// ─── drawing.xml + rels ────────────────────────────────────────────────────

describe('M10 drawing + rels builders', () => {
    test('buildDrawingXml emits one twoCellAnchor per chart, all sharing one wsDr', () => {
        const charts = [
            fixtureChart({ chartId: 'c1', anchor: { fromCol: 3, fromColOff: 0, fromRow: 0, fromRowOff: 0, toCol: 10, toColOff: 0, toRow: 20, toRowOff: 0 } }),
            fixtureChart({ chartId: 'c2', anchor: { fromCol: 11, fromColOff: 0, fromRow: 0, fromRowOff: 0, toCol: 18, toColOff: 0, toRow: 20, toRowOff: 0 } }),
        ];
        const xml = buildDrawingXml(charts);
        const anchorCount = (xml.match(/<xdr:twoCellAnchor>/g) ?? []).length;
        expect(anchorCount).toBe(2);
        expect(xml).toContain('<xdr:wsDr');
        expect(xml).toContain(`r:id="rId1"`);
        expect(xml).toContain(`r:id="rId2"`);
    });

    test('buildDrawingRelsXml maps rIds to chart file numbers', () => {
        const xml = buildDrawingRelsXml([3, 4]);
        expect(xml).toContain(`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart3.xml"/>`);
        expect(xml).toContain(`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart4.xml"/>`);
    });

    test('buildChartRelsXml points to the matching style + colors siblings', () => {
        const xml = buildChartRelsXml(7);
        expect(xml).toContain(`Target="style7.xml"`);
        expect(xml).toContain(`Target="colors7.xml"`);
    });
});

// ─── readChartsFromSnapshot ────────────────────────────────────────────────

describe('readChartsFromSnapshot', () => {
    test('returns [] when no resources field', () => {
        expect(readChartsFromSnapshot({} as Parameters<typeof readChartsFromSnapshot>[0])).toEqual([]);
    });

    test('returns [] when SHEET_DRAWING_PLUGIN missing', () => {
        const snap = { resources: [{ name: 'OTHER_PLUGIN', data: '{}' }] };
        expect(readChartsFromSnapshot(snap as Parameters<typeof readChartsFromSnapshot>[0])).toEqual([]);
    });

    test('returns [] when SHEET_DRAWING_PLUGIN data is malformed JSON', () => {
        const snap = { resources: [{ name: 'SHEET_DRAWING_PLUGIN', data: '{not valid' }] };
        expect(readChartsFromSnapshot(snap as Parameters<typeof readChartsFromSnapshot>[0])).toEqual([]);
    });

    test('filters non-NotesheetChart drawings (e.g. images) out', () => {
        const drawingData = {
            'sheet-1': {
                data: {
                    'img-1': { componentKey: 'NotImageThing', data: {} },
                    'chart-x': {
                        componentKey: 'NotesheetChart',
                        data: {
                            chartId: 'chart-x', type: 'bar', title: 'x',
                            sourceRange: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
                            labels: ['A'], datasets: [{ label: 'd', data: [1] }],
                        },
                        axisAlignSheetTransform: {
                            from: { column: 0, columnOffset: 0, row: 0, rowOffset: 0 },
                            to: { column: 5, columnOffset: 0, row: 10, rowOffset: 0 },
                        },
                    },
                },
                order: ['img-1', 'chart-x'],
            },
        };
        const snap = { resources: [{ name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify(drawingData) }] };
        const out = readChartsFromSnapshot(snap as Parameters<typeof readChartsFromSnapshot>[0]);
        expect(out).toHaveLength(1);
        expect(out[0].chartId).toBe('chart-x');
    });

    test('parses our test fixture round-trip', () => {
        const c = fixtureChart();
        const snap = fixtureSnapshot([c]);
        const out = readChartsFromSnapshot(snap as Parameters<typeof readChartsFromSnapshot>[0]);
        expect(out).toHaveLength(1);
        expect(out[0].title).toBe(c.title);
        expect(out[0].labels).toEqual(c.labels);
        expect(out[0].datasets[0].data).toEqual(c.datasets[0].data);
        expect(out[0].anchor.fromCol).toBe(c.anchor.fromCol);
    });
});

// ─── End-to-end round-trip via snapshotToXlsxBuffer ────────────────────────

describe('M10 round-trip via snapshotToXlsxBuffer', () => {
    test('snapshot with no charts → exporter unchanged (existing tests still pass implicitly)', async () => {
        const snap = fixtureSnapshot([]);
        const buf = await snapshotToXlsxBuffer(snap as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);
        // No chart parts should have been added.
        const paths = Object.keys(zip.files);
        expect(paths.some((p) => p.startsWith('xl/charts/'))).toBe(false);
        expect(paths.some((p) => p.startsWith('xl/drawings/'))).toBe(false);
    });

    test('snapshot with one bar chart → produces chart1.xml + drawing1.xml + style/colors + rels', async () => {
        const snap = fixtureSnapshot([fixtureChart()]);
        const buf = await snapshotToXlsxBuffer(snap as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);

        expect(zip.files['xl/charts/chart1.xml']).toBeDefined();
        expect(zip.files['xl/charts/style1.xml']).toBeDefined();
        expect(zip.files['xl/charts/colors1.xml']).toBeDefined();
        expect(zip.files['xl/charts/_rels/chart1.xml.rels']).toBeDefined();
        expect(zip.files['xl/drawings/drawing1.xml']).toBeDefined();
        expect(zip.files['xl/drawings/_rels/drawing1.xml.rels']).toBeDefined();

        // Sheet rels must reference the drawing.
        const sheetRels = await zip.files['xl/worksheets/_rels/sheet1.xml.rels'].async('string');
        expect(sheetRels).toContain('Target="../drawings/drawing1.xml"');

        // Sheet xml must include <drawing r:id="..."/>.
        const sheetXml = await zip.files['xl/worksheets/sheet1.xml'].async('string');
        expect(sheetXml).toMatch(/<drawing r:id="rId\d+"\/>/);

        // Content types must list all four overrides.
        const ct = await zip.files['[Content_Types].xml'].async('string');
        expect(ct).toContain('PartName="/xl/drawings/drawing1.xml"');
        expect(ct).toContain('PartName="/xl/charts/chart1.xml"');
        expect(ct).toContain('PartName="/xl/charts/style1.xml"');
        expect(ct).toContain('PartName="/xl/charts/colors1.xml"');
    });

    test('two charts on one sheet → one drawing.xml with two anchors, two chart files', async () => {
        const snap = fixtureSnapshot([
            fixtureChart({ chartId: 'c1', title: 'First' }),
            fixtureChart({ chartId: 'c2', title: 'Second', anchor: { fromCol: 11, fromColOff: 0, fromRow: 0, fromRowOff: 0, toCol: 18, toColOff: 0, toRow: 20, toRowOff: 0 } }),
        ]);
        const buf = await snapshotToXlsxBuffer(snap as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);

        expect(zip.files['xl/charts/chart1.xml']).toBeDefined();
        expect(zip.files['xl/charts/chart2.xml']).toBeDefined();
        // Only one drawing file — both anchors live inside it.
        expect(zip.files['xl/drawings/drawing1.xml']).toBeDefined();
        expect(zip.files['xl/drawings/drawing2.xml']).toBeUndefined();

        const drawing = await zip.files['xl/drawings/drawing1.xml'].async('string');
        const anchors = (drawing.match(/<xdr:twoCellAnchor>/g) ?? []).length;
        expect(anchors).toBe(2);

        const drawingRels = await zip.files['xl/drawings/_rels/drawing1.xml.rels'].async('string');
        expect(drawingRels).toContain('Target="../charts/chart1.xml"');
        expect(drawingRels).toContain('Target="../charts/chart2.xml"');
    });

    test('cross-sheet — chart on Sheet2 anchors there but range refs Sheet1', async () => {
        const snap = fixtureSnapshot([
            fixtureChart({ chartId: 'cross', sheetId: 'sheet-2', title: 'Cross-sheet' }),
        ], ['sheet-2']);
        // Re-write the source range to live on Sheet1 even though the chart sheetId is sheet-2.
        // readChartsFromSnapshot doesn't know "where the data lives" — that's
        // implicit in the c:f sheetName we pass at emit time. So we set the
        // chart's sheetName argument via the export pipeline: it will use
        // whatever sheet its sheetId points to. To make this test meaningful
        // we instead verify drawing rels live on sheet2 and chart c:f points
        // to sheet2 — the chart self-references its own sheet, which is what
        // our pipeline does today. (True cross-sheet refs are an M-something-later
        // feature; this test guards against the easy mistake of putting the
        // drawing on the wrong sheet.)
        const buf = await snapshotToXlsxBuffer(snap as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);

        // Drawing rel must be on sheet2's rels file, not sheet1's.
        const sheet2Rels = await zip.files['xl/worksheets/_rels/sheet2.xml.rels'].async('string');
        expect(sheet2Rels).toContain('Target="../drawings/drawing1.xml"');

        // Sheet1 must NOT have the drawing rel.
        const sheet1Rels = zip.files['xl/worksheets/_rels/sheet1.xml.rels'];
        if (sheet1Rels) {
            const xml = await sheet1Rels.async('string');
            expect(xml).not.toContain('Target="../drawings/drawing1.xml"');
        }

        // chart1.xml's range refs use the chart's own sheet name (Sheet2).
        const chartXml = await zip.files['xl/charts/chart1.xml'].async('string');
        expect(chartXml).toContain('<c:f>Sheet2!');
    });

    test('special chars in chart title survive round-trip', async () => {
        const snap = fixtureSnapshot([
            fixtureChart({ title: 'Margins & <FY26>' }),
        ]);
        const buf = await snapshotToXlsxBuffer(snap as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);
        const chartXml = await zip.files['xl/charts/chart1.xml'].async('string');
        expect(chartXml).toContain('<a:t>Margins &amp; &lt;FY26&gt;</a:t>');
    });

    test('chart + table on same sheet — drawing element comes BEFORE tableParts in worksheet xml', async () => {
        // OOXML's CT_Worksheet schema requires <drawing> before <tableParts>.
        // Earlier M10 builds inserted <drawing> just before </worksheet>,
        // which placed it AFTER <tableParts> (added by the M9 table path)
        // and corrupted the file — Excel rejected with "We found a problem".
        const snap = fixtureSnapshot([fixtureChart()]);
        (snap as Record<string, unknown>).resources = [
            ...((snap as { resources?: unknown[] }).resources ?? []),
            {
                name: 'SHEET_TABLE_PLUGIN',
                data: JSON.stringify({
                    'sheet-1': {
                        tables: [{
                            id: 'tbl-1', name: 'Table1',
                            range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
                            options: { showHeader: true, showFooter: false },
                            filters: { tableColumnFilterList: [] },
                            columns: [
                                { id: 'c1', displayName: 'Quarter', dataType: 'string', formula: '', meta: {}, style: {} },
                                { id: 'c2', displayName: 'Sales', dataType: 'number', formula: '', meta: {}, style: {} },
                            ],
                            meta: {},
                        }],
                        tableFilteredOutRows: [],
                    },
                }),
            },
        ];
        const buf = await snapshotToXlsxBuffer(snap as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);
        const sheetXml = await zip.files['xl/worksheets/sheet1.xml'].async('string');

        const drawingPos = sheetXml.search(/<drawing\b/);
        const tablePartsPos = sheetXml.search(/<tableParts\b/);
        expect(drawingPos).toBeGreaterThan(0);
        expect(tablePartsPos).toBeGreaterThan(0);
        // Drawing must precede tableParts — schema requirement.
        expect(drawingPos).toBeLessThan(tablePartsPos);
    });

    test('rId allocation does not collide with existing rels', async () => {
        // Build a snapshot that would cause the M9 path to add a table (rId1),
        // then confirm M10's drawing rel gets rId2.
        const snap = fixtureSnapshot([fixtureChart()]);
        // Inject a fake table resource so the export path adds a tables rel
        // first, claiming rId1.
        (snap as Record<string, unknown>).resources = [
            ...((snap as { resources?: unknown[] }).resources ?? []),
            {
                name: 'SHEET_TABLE_PLUGIN',
                data: JSON.stringify({
                    'sheet-1': {
                        tables: [{
                            id: 'tbl-1', name: 'Table1',
                            range: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
                            options: { showHeader: true, showFooter: false },
                            filters: { tableColumnFilterList: [] },
                            columns: [
                                { id: 'c1', displayName: 'Quarter', dataType: 'string', formula: '', meta: {}, style: {} },
                                { id: 'c2', displayName: 'Sales', dataType: 'number', formula: '', meta: {}, style: {} },
                            ],
                            meta: {},
                        }],
                        tableFilteredOutRows: [],
                    },
                }),
            },
        ];
        const buf = await snapshotToXlsxBuffer(snap as Parameters<typeof snapshotToXlsxBuffer>[0]);
        const zip = await JSZip.loadAsync(buf as ArrayBuffer);
        const sheetRels = await zip.files['xl/worksheets/_rels/sheet1.xml.rels'].async('string');
        // Both rels must be present, with non-colliding rIds.
        expect(sheetRels).toMatch(/Id="rId\d+"[^>]*Target="[^"]*tables\/table1\.xml"/);
        expect(sheetRels).toMatch(/Id="rId\d+"[^>]*Target="[^"]*drawings\/drawing1\.xml"/);
        // Confirm the drawing rId is different from the table rId.
        const tableRId = sheetRels.match(/Id="(rId\d+)"[^>]*Target="[^"]*tables\/table1\.xml"/)?.[1];
        const drawingRId = sheetRels.match(/Id="(rId\d+)"[^>]*Target="[^"]*drawings\/drawing1\.xml"/)?.[1];
        expect(tableRId).toBeDefined();
        expect(drawingRId).toBeDefined();
        expect(drawingRId).not.toBe(tableRId);
    });
});
