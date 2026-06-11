// M17 feature-5: Excel-authored chart fixture survives Notesheet's
// import → export → re-import cycle.
//
// This pins import + export are inverse operations on the chart subset,
// parallel to how M9 (tables) and M15 (CF) already pin inverse round-trip
// on their respective subsets.
//
// Anchored to the FIRST snapshot, NOT to a hardcoded literal:
// snapshot-vs-snapshot inverseness, NOT Excel parity. The M10 emit may
// legitimately produce slightly different XML (palette colours, spPr
// details) that nonetheless re-imports to the same logical chart.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');

interface ChartDrawing {
    componentKey: string;
    data: {
        chartId: string;
        type: string;
        title: string;
        sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
        sourceSheetName?: string;
        labels: string[];
        datasets: Array<{ label?: string; data: number[] }>;
    };
}

function collectChartDrawings(snap: unknown): Array<{ subUnitId: string; drawing: ChartDrawing }> {
    const resources = (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data);
    const out: Array<{ subUnitId: string; drawing: ChartDrawing }> = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        const order: string[] = sub.order ?? Object.keys(sub.data);
        for (const id of order) {
            const d = sub.data[id];
            if (d?.componentKey === 'NotesheetChart') out.push({ subUnitId, drawing: d as ChartDrawing });
        }
    }
    return out;
}

async function roundTrip(fixturePath: string): Promise<{
    first: ReturnType<typeof collectChartDrawings>;
    second: ReturnType<typeof collectChartDrawings>;
    exportedBuffer: ArrayBuffer;
}> {
    const buf = readFileSync(fixturePath);
    const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
    const buf2 = await snapshotToXlsxBuffer(snap1);
    const snap2 = await xlsxBufferToSnapshot(Buffer.from(buf2) as unknown as Buffer);
    return {
        first: collectChartDrawings(snap1),
        second: collectChartDrawings(snap2),
        exportedBuffer: buf2,
    };
}

const TYPE_FIXTURES = [
    { fixture: '01-bar-simple.xlsx', type: 'bar' },
    { fixture: '02-line-multi-series.xlsx', type: 'line' },
    { fixture: '03-pie-single.xlsx', type: 'pie' },
    { fixture: '04-doughnut.xlsx', type: 'doughnut' },
];

describe('M17 feature-5: chart bidirectional round-trip across types', () => {
    for (const { fixture, type } of TYPE_FIXTURES) {
        test(`${fixture} round-trip preserves type=${type} + sourceRange + labels + datasets`, async () => {
            const fp = path.join(FIXTURES_DIR, fixture);
            const { first, second } = await roundTrip(fp);

            expect(first).toHaveLength(1);
            expect(second).toHaveLength(1);

            const a = first[0].drawing.data;
            const b = second[0].drawing.data;

            expect(b.type).toBe(a.type);
            expect(b.type).toBe(type);
            expect(b.sourceRange).toEqual(a.sourceRange);
            expect(b.labels).toEqual(a.labels);
            expect(b.datasets.length).toBe(a.datasets.length);
            for (let i = 0; i < a.datasets.length; i++) {
                expect(b.datasets[i].data).toEqual(a.datasets[i].data);
                if (a.datasets[i].label !== undefined) {
                    expect(b.datasets[i].label).toBe(a.datasets[i].label);
                }
            }
        });
    }

    test('07-chart-cross-sheet.xlsx: chart on Sheet2 keeps its source range pointing at Sheet1', async () => {
        const fp = path.join(FIXTURES_DIR, '07-chart-cross-sheet.xlsx');
        const { first, second, exportedBuffer } = await roundTrip(fp);

        expect(first).toHaveLength(1);
        expect(second).toHaveLength(1);

        // First snapshot — chart anchored on Sheet2, sourceSheetName = Sheet1.
        expect(first[0].subUnitId).toBe('sheet-2');
        expect(first[0].drawing.data.sourceSheetName).toBe('Sheet1');

        // Second snapshot — same shape after round-trip.
        expect(second[0].subUnitId).toBe('sheet-2');
        expect(second[0].drawing.data.sourceSheetName).toBe('Sheet1');
        expect(second[0].drawing.data.sourceRange).toEqual(first[0].drawing.data.sourceRange);

        // Catch the M10 trap: the exported xlsx's <c:f> formula prefix
        // must still be Sheet1!, not Sheet2! (the chart's containing
        // sheet). Without sourceSheetName plumbing M10 silently rebuilds
        // the formula prefix as the chart's host sheet, which lets the
        // cached labels/values pass equality but breaks Excel's
        // re-evaluation from cells.
        const exportedZip = await JSZip.loadAsync(exportedBuffer);
        const chartFiles = Object.keys(exportedZip.files).filter((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
        expect(chartFiles.length).toBeGreaterThan(0);
        const chartXml = await exportedZip.file(chartFiles[0])!.async('string');
        const formulas = [...chartXml.matchAll(/<c:f>([^<]+)<\/c:f>/g)].map((m) => m[1]);
        expect(formulas.length).toBeGreaterThan(0);
        for (const f of formulas) {
            expect(f.startsWith('Sheet1!')).toBe(true);
        }
    });
});
