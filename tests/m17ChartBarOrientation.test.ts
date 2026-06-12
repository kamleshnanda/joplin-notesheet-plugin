// M17: bar-orientation fidelity. Excel's <c:barDir val="bar"/> means
// horizontal bars, val="col" means vertical (Excel's default for "Column
// Chart"). The Notesheet ChartType union has only one 'bar' literal, so
// orientation flows through `meta.barDir` and the renderer routes it to
// Chart.js's `options.indexAxis`.
//
// Anchored UPSTREAM: each fixture's expected barDir is independently
// parsed from `xl/charts/chart1.xml` via JSZip+regex, NOT from what
// Notesheet emits.

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

interface ChartDrawing {
    componentKey: string;
    data: {
        type: string;
        meta?: { barDir?: 'bar' | 'col'; unsupportedSourceType?: string };
    };
}

function collectChartDrawings(snap: unknown): ChartDrawing[] {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data);
    const out: ChartDrawing[] = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        const order: string[] = sub.order ?? Object.keys(sub.data);
        for (const id of order) {
            const d = sub.data[id];
            if (d?.componentKey === 'NotesheetChart') out.push(d as ChartDrawing);
        }
    }
    return out;
}

// Independently parse <c:barDir val="..."/> from the source chart XML.
// Returns 'bar' or 'col' or null if absent.
async function readSourceBarDirs(fixturePath: string): Promise<Array<'bar' | 'col' | null>> {
    const buf = readFileSync(fixturePath);
    const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
    const out: Array<'bar' | 'col' | null> = [];
    const chartFiles = Object.keys(zip.files)
        .filter((p) => /^xl\/charts\/chart\d+\.xml$/.test(p))
        .sort();
    for (const p of chartFiles) {
        const xml = await zip.file(p)!.async('string');
        const m = xml.match(/<(?:c:)?barDir\s+val="(bar|col)"/);
        out.push(m ? (m[1] as 'bar' | 'col') : null);
    }
    return out;
}

describe('M17: bar orientation surfaces via meta.barDir', () => {
    test('01-bar-simple: source <c:barDir val="bar"/> survives import as meta.barDir="bar"', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const sourceDirs = await readSourceBarDirs(fp);
        // Sanity check the fixture itself.
        expect(sourceDirs[0]).toBe('bar');

        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.type).toBe('bar');
        expect(drawings[0].data.meta?.barDir).toBe('bar');
    });

    test('07-chart-cross-sheet: source <c:barDir val="col"/> survives import as meta.barDir="col"', async () => {
        const fp = path.join(FIXTURES_DIR, '07-chart-cross-sheet.xlsx');
        const sourceDirs = await readSourceBarDirs(fp);
        expect(sourceDirs[0]).toBe('col');

        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.type).toBe('bar');
        expect(drawings[0].data.meta?.barDir).toBe('col');
    });

    test('non-bar charts (line/pie/doughnut) carry no barDir', async () => {
        for (const fixture of ['02-line-multi-series', '03-pie-single', '04-doughnut']) {
            const fp = path.join(FIXTURES_DIR, `${fixture}.xlsx`);
            const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
            const drawings = collectChartDrawings(snap);
            expect(drawings).toHaveLength(1);
            expect(drawings[0].data.meta?.barDir).toBeUndefined();
        }
    });

    test('round-trip: horizontal bar fixture (01) re-exports with <c:barDir val="bar"/>', async () => {
        // Pre-fix: M10 hardcoded <c:barDir val="col"/>, so a re-imported
        // horizontal bar chart silently flipped to vertical on export.
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(out as ArrayBuffer);
        const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
        const chartXml = await zip.file(chartFile!)!.async('string');
        expect(chartXml).toMatch(/<c:barDir\s+val="bar"/);
    });

    test('round-trip: vertical bar fixture (07) re-exports with <c:barDir val="col"/>', async () => {
        const fp = path.join(FIXTURES_DIR, '07-chart-cross-sheet.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(out as ArrayBuffer);
        const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
        const chartXml = await zip.file(chartFile!)!.async('string');
        expect(chartXml).toMatch(/<c:barDir\s+val="col"/);
    });

    test('06-two-charts-one-sheet: each chart preserves its own barDir', async () => {
        const fp = path.join(FIXTURES_DIR, '06-two-charts-one-sheet.xlsx');
        const sourceDirs = await readSourceBarDirs(fp);
        // Sanity: 06 ships two charts with potentially different orientations.
        expect(sourceDirs).toHaveLength(2);

        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(2);
        // The walk order in xlsxChartImport matches document anchors
        // which in 06 line up with chart1/chart2.xml. Pin each.
        for (let i = 0; i < drawings.length; i++) {
            if (drawings[i].data.type === 'bar') {
                expect(drawings[i].data.meta?.barDir).toBe(sourceDirs[i]);
            }
        }
    });
});
