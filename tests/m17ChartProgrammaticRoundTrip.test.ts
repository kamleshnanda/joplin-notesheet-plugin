// M17 feature-6: programmatic round-trip pack.
//
// A pack of in-memory snapshots authored DIRECTLY in test code — each
// carrying one (or two) NotesheetChart drawings in the
// SHEET_DRAWING_PLUGIN resource — exercises Notesheet's OWN
// emit-and-import round-trip on edge cases the Excel-authored fixture
// set (`tests/fixtures/charts/`) doesn't cover or covers only thinly:
// negative values, single-data-point series, very long labels, an
// empty doughnut, XML-special characters in title + categories,
// cross-sheet references, and two charts on one sheet.
//
// LINT SENTINEL (feature-6 criterion 4): this test MUST NOT invoke
// `new ExcelJS.Workbook()`. exceljs's chart write API is thin/uneven —
// M10 deliberately bypasses it and post-processes the zip
// (`injectChartsIntoZip`). Every snapshot here is built as a plain
// object in the SHAPE `readChartsFromSnapshot` consumes, then driven
// through the real export pipeline (`snapshotToXlsxBuffer`) and the
// real import pipeline (`xlsxBufferToSnapshot`). The sentinel is
// enforced at runtime by the final test in this file, which reads this
// source file and asserts the forbidden constructor never appears.
//
// Each case anchors to its ORIGINAL snapshot's chart-drawing fields,
// NOT to a hardcoded literal or to the re-imported snapshot: it reads
// the original drawing, round-trips, reads the re-imported drawing, and
// asserts originalChart.X === reimportedChart.X for each pinned field
// (type, sourceRange, anchor, labels, datasets). `chartId` is
// deliberately excluded — the import path regenerates a synthetic id.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

// ─── Snapshot authoring helpers ────────────────────────────────────────────

type ChartType = 'bar' | 'line' | 'pie' | 'doughnut';

interface ChartSpec {
    sheetId: string;
    drawingId: string;
    type: ChartType;
    title: string;
    // { startRow, endRow, startColumn, endColumn }. Convention (matches
    // buildSeriesXml/buildPieSeriesXml): header row at startRow, label
    // column at startColumn, data rows startRow+1..endRow, series columns
    // startColumn+1..endColumn. So labels.length === endRow - startRow and
    // datasets.length === endColumn - startColumn.
    sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    sourceSheetName?: string;
    labels: string[];
    datasets: Array<{ label?: string; data: number[] }>;
    // Cell anchor. Offsets are intentionally 0: the import resource
    // builder treats <xdr:colOff>/<xdr:rowOff> as EMU and divides by
    // 9525, so only a 0 offset survives a programmatic round-trip
    // cleanly. Column/row indices round-trip exactly.
    anchor?: { fromCol: number; fromRow: number; toCol: number; toRow: number };
}

interface SheetSpec {
    id: string;
    name: string;
}

// Build a UniverSnapshot in the same shape xlsxBufferToSnapshot emits for
// imported charts — `SHEET_DRAWING_PLUGIN` resource whose `data` is a
// JSON-stringified { [subUnitId]: { data: { [drawingId]: ISheetDrawing },
// order: string[] } } map, each chart drawing carrying
// componentKey === 'NotesheetChart'. cellData is intentionally omitted:
// the chart caches (strCache/numCache) carry the values, so the
// round-trip exercises the export-cache → import-cache path directly.
function buildChartSnapshot(sheetSpecs: SheetSpec[], charts: ChartSpec[]): Record<string, unknown> {
    const sheets: Record<string, unknown> = {};
    const sheetOrder: string[] = [];
    for (const s of sheetSpecs) {
        sheetOrder.push(s.id);
        sheets[s.id] = {
            id: s.id,
            name: s.name,
            cellData: {},
            rowCount: 100,
            columnCount: 26,
            defaultColumnWidth: 73,
            defaultRowHeight: 19,
        };
    }

    const drawingResource: Record<string, { data: Record<string, unknown>; order: string[] }> = {};
    for (const c of charts) {
        const anchor = c.anchor ?? { fromCol: 4, fromRow: 1, toCol: 11, toRow: 16 };
        if (!drawingResource[c.sheetId]) {
            drawingResource[c.sheetId] = { data: {}, order: [] };
        }
        drawingResource[c.sheetId].data[c.drawingId] = {
            unitId: 'workbook',
            subUnitId: c.sheetId,
            drawingId: c.drawingId,
            drawingType: 8, // DRAWING_DOM
            componentKey: 'NotesheetChart',
            allowTransform: true,
            data: {
                chartId: c.drawingId,
                type: c.type,
                title: c.title,
                sourceRange: c.sourceRange,
                ...(c.sourceSheetName ? { sourceSheetName: c.sourceSheetName } : {}),
                labels: c.labels,
                datasets: c.datasets,
            },
            axisAlignSheetTransform: {
                from: { column: anchor.fromCol, columnOffset: 0, row: anchor.fromRow, rowOffset: 0 },
                to: { column: anchor.toCol, columnOffset: 0, row: anchor.toRow, rowOffset: 0 },
            },
        };
        drawingResource[c.sheetId].order.push(c.drawingId);
    }

    return {
        id: 'workbook-feature6',
        sheetOrder,
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        styles: {},
        sheets,
        resources: [
            { name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify(drawingResource) },
        ],
    };
}

// ─── Round-trip + collection ───────────────────────────────────────────────

interface CollectedChart {
    subUnitId: string;
    chartId: string;
    type: string;
    title: string;
    sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    sourceSheetName?: string;
    labels: string[];
    datasets: Array<{ label?: string; data: number[] }>;
    anchor: {
        fromCol: number; fromColOff: number; fromRow: number; fromRowOff: number;
        toCol: number; toColOff: number; toRow: number; toRowOff: number;
    };
}

// Read every NotesheetChart drawing out of a snapshot's
// SHEET_DRAWING_PLUGIN resource, in `order`, flattening the fields the
// pack pins. Anchor is read from axisAlignSheetTransform (the
// xlsx-aligned transform both import and export agree on).
function collectCharts(snap: unknown): CollectedChart[] {
    const resources = (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data) as Record<string, { data: Record<string, any>; order?: string[] }>;
    const out: CollectedChart[] = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        const order = sub.order ?? Object.keys(sub.data);
        for (const id of order) {
            const d = sub.data[id];
            if (d?.componentKey !== 'NotesheetChart') continue;
            const tx = d.axisAlignSheetTransform ?? d.sheetTransform ?? { from: {}, to: {} };
            out.push({
                subUnitId,
                chartId: d.data.chartId,
                type: d.data.type,
                title: d.data.title,
                sourceRange: d.data.sourceRange,
                sourceSheetName: d.data.sourceSheetName,
                labels: d.data.labels ?? [],
                datasets: (d.data.datasets ?? []).map((ds: any) => ({
                    ...(ds.label !== undefined ? { label: ds.label } : {}),
                    data: ds.data ?? [],
                })),
                anchor: {
                    fromCol: tx.from?.column ?? 0, fromColOff: tx.from?.columnOffset ?? 0,
                    fromRow: tx.from?.row ?? 0, fromRowOff: tx.from?.rowOffset ?? 0,
                    toCol: tx.to?.column ?? 0, toColOff: tx.to?.columnOffset ?? 0,
                    toRow: tx.to?.row ?? 0, toRowOff: tx.to?.rowOffset ?? 0,
                },
            });
        }
    }
    return out;
}

async function roundTrip(snapshot: Record<string, unknown>): Promise<{
    first: CollectedChart[];
    second: CollectedChart[];
}> {
    const first = collectCharts(snapshot);
    const buf = await snapshotToXlsxBuffer(snapshot);
    const snap2 = await xlsxBufferToSnapshot(Buffer.from(buf) as unknown as Buffer);
    const second = collectCharts(snap2);
    return { first, second };
}

// Assert the pinned fields (type, sourceRange, labels, datasets, anchor
// col/row) survive a round-trip. chartId is excluded by design.
function expectPinnedFieldsEqual(a: CollectedChart, b: CollectedChart): void {
    expect(b.type).toBe(a.type);
    expect(b.sourceRange).toEqual(a.sourceRange);
    expect(b.labels).toEqual(a.labels);
    expect(b.datasets.length).toBe(a.datasets.length);
    for (let i = 0; i < a.datasets.length; i++) {
        expect(b.datasets[i].data).toEqual(a.datasets[i].data);
        if (a.datasets[i].label !== undefined) {
            expect(b.datasets[i].label).toBe(a.datasets[i].label);
        }
    }
    // Anchor column/row indices (offsets authored as 0; see ChartSpec).
    expect(b.anchor.fromCol).toBe(a.anchor.fromCol);
    expect(b.anchor.fromRow).toBe(a.anchor.fromRow);
    expect(b.anchor.toCol).toBe(a.anchor.toCol);
    expect(b.anchor.toRow).toBe(a.anchor.toRow);
}

const ONE_SHEET: SheetSpec[] = [{ id: 'sheet-1', name: 'Sheet1' }];

describe('feature-6: programmatic round-trip pack', () => {
    test('Case A — bar chart with mixed positive/negative values', async () => {
        const snap = buildChartSnapshot(ONE_SHEET, [{
            sheetId: 'sheet-1',
            drawingId: 'chart-A',
            type: 'bar',
            title: 'Net change',
            sourceRange: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
            labels: ['Q1', 'Q2', 'Q3', 'Q4'],
            datasets: [{ label: 'Delta', data: [3, -2, 5, -1] }],
            anchor: { fromCol: 3, fromRow: 0, toCol: 10, toRow: 15 },
        }]);
        const { first, second } = await roundTrip(snap);
        expect(first).toHaveLength(1);
        expect(second).toHaveLength(1);
        // The negatives specifically must survive.
        expect(first[0].datasets[0].data).toEqual([3, -2, 5, -1]);
        expect(second[0].datasets[0].data).toEqual([3, -2, 5, -1]);
        expectPinnedFieldsEqual(first[0], second[0]);
    });

    test('Case B — line chart with a single-data-point series', async () => {
        const snap = buildChartSnapshot(ONE_SHEET, [{
            sheetId: 'sheet-1',
            drawingId: 'chart-B',
            type: 'line',
            title: 'Single point',
            sourceRange: { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
            labels: ['OnlyPoint'],
            datasets: [{ label: 'Series 1', data: [42] }],
        }]);
        const { first, second } = await roundTrip(snap);
        expect(first).toHaveLength(1);
        expect(second).toHaveLength(1);
        expect(second[0].labels).toHaveLength(1);
        expect(second[0].datasets[0].data).toHaveLength(1);
        expectPinnedFieldsEqual(first[0], second[0]);
    });

    test('Case C — pie chart with a very long category label', async () => {
        const longLabel = 'Engineering, Research, Development & Long-Tail Operations Division';
        expect(longLabel.length).toBeGreaterThan(30);
        const snap = buildChartSnapshot(ONE_SHEET, [{
            sheetId: 'sheet-1',
            drawingId: 'chart-C',
            type: 'pie',
            title: 'Headcount',
            sourceRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
            labels: [longLabel, 'Sales', 'Support'],
            datasets: [{ label: 'People', data: [120, 45, 30] }],
        }]);
        const { first, second } = await roundTrip(snap);
        expect(first).toHaveLength(1);
        expect(second).toHaveLength(1);
        // The long label must survive verbatim.
        expect(second[0].labels[0]).toBe(longLabel);
        expectPinnedFieldsEqual(first[0], second[0]);
    });

    test('Case D — doughnut chart with an empty series (zero data rows)', async () => {
        const snap = buildChartSnapshot(ONE_SHEET, [{
            sheetId: 'sheet-1',
            drawingId: 'chart-D',
            type: 'doughnut',
            title: 'Empty',
            sourceRange: { startRow: 0, endRow: 0, startColumn: 0, endColumn: 1 },
            labels: [],
            datasets: [{ label: 'Series 1', data: [] }],
        }]);
        const { first, second } = await roundTrip(snap);
        expect(first).toHaveLength(1);
        // Spec accepts EITHER outcome: DROP (0 drawings) or PRESERVE.
        // The implementation lands on PRESERVE — but NOT with truly
        // empty arrays. M10 export emits `<c:ptCount val="0"/>` caches
        // over a degenerate range; on re-import the empty caches trigger
        // the resolveDataFromCells fallback, which reads the (empty)
        // cellData over the reconstructed range and fabricates a single
        // placeholder point: labels === [''] and datasets[0].data === [0].
        // The load-bearing correctness properties this case pins: the
        // round-trip does NOT crash, produces at most one drawing, keeps
        // type === 'doughnut', and NEVER invents meaningful data (every
        // surviving label is '' and every surviving value is 0). See
        // PROGRESS.md ## Notes for the PRESERVE-with-empty-placeholder
        // choice.
        expect(second.length === 0 || second.length === 1).toBe(true);
        if (second.length === 1) {
            expect(second[0].type).toBe('doughnut');
            expect(second[0].labels.length).toBeLessThanOrEqual(1);
            expect(second[0].labels.every((l) => l === '')).toBe(true);
            for (const ds of second[0].datasets) {
                expect(ds.data.length).toBeLessThanOrEqual(1);
                expect(ds.data.every((v) => v === 0)).toBe(true);
            }
        }
    });

    test('Case E — bar chart with XML-special chars in title and category names', async () => {
        const title = 'Sales & "Profit" <Q1> \'2024\'';
        const labels = ['R&D', 'x<y', 'a>b'];
        const snap = buildChartSnapshot(ONE_SHEET, [{
            sheetId: 'sheet-1',
            drawingId: 'chart-E',
            type: 'bar',
            title,
            sourceRange: { startRow: 0, endRow: 3, startColumn: 0, endColumn: 1 },
            labels,
            datasets: [{ label: 'Q&A <count>', data: [10, 20, 30] }],
        }]);
        const { first, second } = await roundTrip(snap);
        expect(first).toHaveLength(1);
        expect(second).toHaveLength(1);
        // escapeXml + decodeXmlEntities must be inverse on this set.
        expect(second[0].title).toBe(title);
        expect(second[0].labels).toEqual(labels);
        expect(second[0].datasets[0].label).toBe('Q&A <count>');
        expectPinnedFieldsEqual(first[0], second[0]);
    });

    test('Case F — cross-sheet chart (chart on Sheet2 references data on Sheet1)', async () => {
        const snap = buildChartSnapshot(
            [{ id: 'sheet-1', name: 'Sheet1' }, { id: 'sheet-2', name: 'Sheet2' }],
            [{
                sheetId: 'sheet-2',
                drawingId: 'chart-F',
                type: 'bar',
                title: 'Cross-sheet',
                sourceRange: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
                sourceSheetName: 'Sheet1',
                labels: ['Jan', 'Feb'],
                datasets: [{ label: 'Revenue', data: [100, 200] }],
            }],
        );
        const { first, second } = await roundTrip(snap);
        expect(first).toHaveLength(1);
        expect(second).toHaveLength(1);
        // Chart lives on Sheet2 in both snapshots.
        expect(first[0].subUnitId).toBe('sheet-2');
        expect(second[0].subUnitId).toBe('sheet-2');
        // The cross-sheet reference points at Sheet1 in both.
        expect(first[0].sourceSheetName).toBe('Sheet1');
        expect(second[0].sourceSheetName).toBe('Sheet1');
        expectPinnedFieldsEqual(first[0], second[0]);
    });

    test('Case G — two charts on one sheet round-trip with distinct chartIds', async () => {
        const snap = buildChartSnapshot(ONE_SHEET, [
            {
                sheetId: 'sheet-1',
                drawingId: 'chart-G1',
                type: 'bar',
                title: 'First',
                sourceRange: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
                labels: ['a', 'b'],
                datasets: [{ label: 'one', data: [1, 2] }],
                anchor: { fromCol: 0, fromRow: 0, toCol: 7, toRow: 14 },
            },
            {
                sheetId: 'sheet-1',
                drawingId: 'chart-G2',
                type: 'line',
                title: 'Second',
                sourceRange: { startRow: 0, endRow: 2, startColumn: 4, endColumn: 5 },
                labels: ['c', 'd'],
                datasets: [{ label: 'two', data: [3, 4] }],
                anchor: { fromCol: 8, fromRow: 0, toCol: 15, toRow: 14 },
            },
        ]);
        const { first, second } = await roundTrip(snap);
        expect(first).toHaveLength(2);
        expect(second).toHaveLength(2);
        // Both on sheet-1.
        expect(second.every((c) => c.subUnitId === 'sheet-1')).toBe(true);
        // Distinct chartIds after round-trip.
        expect(second[0].chartId).not.toBe(second[1].chartId);
        // Document order preserved: pinned fields match per index.
        expectPinnedFieldsEqual(first[0], second[0]);
        expectPinnedFieldsEqual(first[1], second[1]);
    });

    test('Lint sentinel — this test file uses no ExcelJS workbook constructor', () => {
        const src = readFileSync(__filename, 'utf8');
        // Strip comments (// ... and /* ... */) so the sentinel checks
        // only EXECUTABLE code, not the prose above that legitimately
        // names the forbidden API to explain why it's forbidden.
        const codeOnly = src
            .replace(/\/\*[\s\S]*?\*\//g, '')
            .replace(/\/\/[^\n]*/g, '');
        // The forbidden form is the constructor call (exceljs's chart
        // write API is thin; M10 post-processes the zip instead).
        expect(/new\s+ExcelJS\s*\.\s*Workbook\s*\(/.test(codeOnly)).toBe(false);
        // Belt-and-suspenders: the module is never even imported here.
        expect(/\bimport\b[^\n]*\bexceljs\b/i.test(codeOnly)).toBe(false);
    });
});
