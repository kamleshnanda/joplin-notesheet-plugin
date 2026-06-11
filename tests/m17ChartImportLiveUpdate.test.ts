// M17 feature-4: imported charts wire into the live data bus.
//
// Gates against the M13-class trap: trackedCharts was populated only by the
// in-editor insertChart() flow. A chart imported from .xlsx subscribed to
// nothing — editing a source-range cell never re-pushed data, and the
// chart showed stale values forever. In a single screenshot a stale chart
// looks identical to a fresh one, so the runtime gap could only be caught
// by Jest. This test is that gate.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot } from '../src/xlsx';
import { decodeCellRef } from '../src/charts/xlsxChartImport';
import { extractDataFromSnapshot } from '../src/charts/extractData';
import {
    pushChartUpdate,
    subscribeChartUpdate,
    _resetChartBus,
} from '../src/charts/dataBus';
import {
    populateTrackedChartsFromSnapshot,
    trackedCharts,
} from '../src/charts/trackedCharts';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');

interface ChartDrawing {
    componentKey: string;
    data: {
        chartId: string;
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

// Independently parse source-XML labels and per-series data values.
async function readSourceLabelsAndData(fixturePath: string): Promise<{
    labels: string[];
    datasetsData: number[][];
}> {
    const buf = readFileSync(fixturePath);
    const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
    const chartXml = await zip.file('xl/charts/chart1.xml')!.async('string');
    let labels: string[] = [];
    const datasetsData: number[][] = [];
    const serRe = /<c:ser>([\s\S]*?)<\/c:ser>/g;
    let m: RegExpExecArray | null;
    while ((m = serRe.exec(chartXml)) !== null) {
        const ser = m[1];
        if (labels.length === 0) {
            const cat = ser.match(/<c:cat>([\s\S]*?)<\/c:cat>/);
            if (cat) labels = readPtCacheValues(cat[1]).map(String);
        }
        const val = ser.match(/<c:val>([\s\S]*?)<\/c:val>/);
        const arr = val ? readPtCacheValues(val[1]).map(Number) : [];
        datasetsData.push(arr);
    }
    return { labels, datasetsData };
}

function readPtCacheValues(cacheBody: string): string[] {
    const ptRe = /<c:pt\b[^>]*idx="(\d+)"[^>]*>[\s\S]*?<c:v>([\s\S]*?)<\/c:v>[\s\S]*?<\/c:pt>/g;
    const map = new Map<number, string>();
    let max = -1;
    let pm: RegExpExecArray | null;
    while ((pm = ptRe.exec(cacheBody)) !== null) {
        const idx = parseInt(pm[1], 10);
        map.set(idx, pm[2]);
        if (idx > max) max = idx;
    }
    const out: string[] = [];
    for (let i = 0; i <= max; i++) out.push(map.get(i) ?? '');
    return out;
}

beforeEach(() => {
    _resetChartBus();
    trackedCharts.clear();
});

describe('M17 feature-4: imported charts wire into the live data bus', () => {
    test('imported chart drawing has non-empty chartId and sourceRange matches source XML', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const buf = readFileSync(fp);
        const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
        const chartXml = await zip.file('xl/charts/chart1.xml')!.async('string');
        const catFormula = chartXml.match(/<c:cat>[\s\S]*?<c:f>([\s\S]*?)<\/c:f>/)![1];
        const valFormula = chartXml.match(/<c:val>[\s\S]*?<c:f>([\s\S]*?)<\/c:f>/)![1];
        const catRef = decodeCellRef(catFormula)!;
        const valRef = decodeCellRef(valFormula)!;
        const expected = {
            startRow: Math.min(catRef.startRow - 1, valRef.startRow - 1),
            endRow: Math.max(catRef.endRow, valRef.endRow),
            startColumn: Math.min(catRef.startColumn, valRef.startColumn),
            endColumn: Math.max(catRef.endColumn, valRef.endColumn),
        };

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].drawing.data.chartId).toMatch(/.+/);
        expect(drawings[0].drawing.data.sourceRange).toEqual(expected);
    });

    test('subscribeChartUpdate keys off the imported chartId verbatim', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const chartId = collectChartDrawings(snap)[0].drawing.data.chartId;

        const received: string[][] = [];
        const unsubscribe = subscribeChartUpdate(chartId, (data) => {
            received.push(data.labels);
        });

        const fakeChartData = {
            labels: ['x', 'y'],
            datasets: [{ data: [1, 2], backgroundColor: '#ff0000' }],
        };
        pushChartUpdate(chartId, fakeChartData);
        unsubscribe();

        expect(received).toHaveLength(1);
        expect(received[0]).toEqual(['x', 'y']);
    });

    test('populateTrackedChartsFromSnapshot lands every imported chart in trackedCharts', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        const chartId = drawings[0].drawing.data.chartId;
        const expectedRange = drawings[0].drawing.data.sourceRange;

        expect(trackedCharts.size).toBe(0);
        populateTrackedChartsFromSnapshot(snap as Record<string, unknown>);

        expect(trackedCharts.has(chartId)).toBe(true);
        const tracked = trackedCharts.get(chartId)!;
        expect(tracked.id).toBe(chartId);
        expect(tracked.sourceRange.startRow).toBe(expectedRange.startRow);
        expect(tracked.sourceRange.endRow).toBe(expectedRange.endRow);
        expect(tracked.sourceRange.startColumn).toBe(expectedRange.startColumn);
        expect(tracked.sourceRange.endColumn).toBe(expectedRange.endColumn);
        expect(tracked.sourceRange.subUnitId).toBe(drawings[0].subUnitId);
    });

    test('end-to-end: snapshot-load + extractDataFromSnapshot + bus delivers fresh data', async () => {
        // The chart's source range covers HEADER row + DATA rows. Notesheet's
        // extract-for-chart convention (mirroring extractRangeAsChartData) is:
        // first column → labels (including the header label),
        // remaining columns → series (with the header label coerced to NaN
        // since the header isn't numeric). The source XML's <c:cat>/<c:val>
        // caches hold DATA-only labels/values — those are what Excel displays
        // when re-rendering, not what extractDataFromSnapshot returns.
        //
        // The contract this test pins: extractDataFromSnapshot, run against
        // the imported snapshot's cell data, returns the same data the
        // editor-side extractRangeAsChartData would return when the user
        // edits a cell. We anchor to the source XML's data values
        // (truth.datasetsData[0]) — those values match the cell values
        // exceljs landed in the snapshot's cellData[row][col].v.
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const truth = await readSourceLabelsAndData(fp);

        populateTrackedChartsFromSnapshot(snap as Record<string, unknown>);
        const drawings = collectChartDrawings(snap);
        const chartId = drawings[0].drawing.data.chartId;
        const tracked = trackedCharts.get(chartId)!;

        const fresh = extractDataFromSnapshot(
            snap,
            tracked.sourceRange,
            tracked.sourceSheetName,
        );

        // labels = header row + cat-cache values (truth.labels are data-only).
        expect(fresh.labels.length).toBe(truth.labels.length + 1);
        expect(fresh.labels.slice(1)).toEqual(truth.labels);
        // series-0 data = NaN (header) + numCache values.
        expect(Number.isNaN(fresh.datasets[0].data[0])).toBe(true);
        expect(fresh.datasets[0].data.slice(1)).toEqual(truth.datasetsData[0]);

        const received: Array<typeof fresh> = [];
        subscribeChartUpdate(chartId, (data) => { received.push(data); });
        pushChartUpdate(chartId, fresh);
        expect(received).toHaveLength(1);
        expect(received[0].labels.slice(1)).toEqual(truth.labels);
        expect(received[0].datasets[0].data.slice(1)).toEqual(truth.datasetsData[0]);
    });

    test('no second renderer branch keyed off imported-vs-authored', () => {
        // Static sentinel keeps a future refactor honest. If a second
        // renderer branch is added for imports, this test fails and
        // the operator gets a chance to re-evaluate.
        const editorViewSrc = readFileSync(
            path.join(__dirname, '..', 'src', 'editorView.tsx'),
            'utf-8',
        );
        const pushCallSites = (editorViewSrc.match(/\bpushChartUpdate\s*\(/g) ?? []).length;
        expect(pushCallSites).toBe(1);
        expect(editorViewSrc).not.toMatch(/isImported/);
    });
});
