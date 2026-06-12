// M17 feature-2: per-chart-type fidelity.
//
// For each of the four supported types (bar / line / pie / doughnut), the
// imported snapshot's chart-drawing carries the matching ChartType literal.
// Unsupported types (radar, scatter, ...) fall back to 'bar' AND tag the
// chart with `meta.unsupportedSourceType` for evaluator visibility.
//
// All assertions anchor to source XML, NOT to our own emit:
//   * Type strings parsed independently from <c:barChart>/<c:lineChart>/
//     <c:pieChart>/<c:doughnutChart> element names in the chart XML.
//   * Series count parsed independently from the count of <c:ser> elements.
//
// Per feedback_pge_fidelity_test_gap.md.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot } from '../src/xlsx';
import { buildSyntheticChartXlsx, RADAR_CHART_BODY } from './util/m17BuildSyntheticXlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');

interface ChartDrawing {
    componentKey: string;
    data: {
        type: string;
        labels: string[];
        datasets: Array<{ label?: string; data: number[] }>;
        meta?: { unsupportedSourceType?: string };
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

// Parse the type string from the source chart XML — anchored upstream of
// our import code. Returns the chart-type element name without the c: prefix
// (e.g. 'bar' for <c:barChart>).
async function readSourceType(fixturePath: string): Promise<string> {
    const buf = readFileSync(fixturePath);
    const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
    const chartFile = zip.file('xl/charts/chart1.xml');
    if (!chartFile) throw new Error(`no chart1.xml in ${fixturePath}`);
    const xml = await chartFile.async('string');
    const m = xml.match(
        /<c:(bar|line|pie|doughnut|radar|scatter|area|bubble|surface|stock)Chart\b/,
    );
    if (!m) throw new Error(`no chart-type element found in ${fixturePath}`);
    return m[1];
}

async function countSourceSeries(fixturePath: string): Promise<number> {
    const buf = readFileSync(fixturePath);
    const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
    const chartFile = zip.file('xl/charts/chart1.xml');
    if (!chartFile) throw new Error(`no chart1.xml in ${fixturePath}`);
    const xml = await chartFile.async('string');
    return (xml.match(/<c:ser\b/g) ?? []).length;
}

describe('M17 feature-2: chart type fidelity', () => {
    test('01-bar-simple.xlsx imports as type=bar (matches source XML)', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        expect(await readSourceType(fp)).toBe('bar');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.type).toBe('bar');
    });

    test('02-line-multi-series.xlsx imports as type=line with 3 datasets (matches source <c:ser> count)', async () => {
        const fp = path.join(FIXTURES_DIR, '02-line-multi-series.xlsx');
        expect(await readSourceType(fp)).toBe('line');
        const expectedSeriesCount = await countSourceSeries(fp);
        expect(expectedSeriesCount).toBe(3); // sanity: the fixture has 3 series

        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.type).toBe('line');
        expect(drawings[0].data.datasets.length).toBe(expectedSeriesCount);
    });

    test('03-pie-single.xlsx imports as type=pie (matches source XML)', async () => {
        const fp = path.join(FIXTURES_DIR, '03-pie-single.xlsx');
        expect(await readSourceType(fp)).toBe('pie');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.type).toBe('pie');
    });

    test('04-doughnut.xlsx imports as type=doughnut (matches source XML)', async () => {
        const fp = path.join(FIXTURES_DIR, '04-doughnut.xlsx');
        expect(await readSourceType(fp)).toBe('doughnut');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.type).toBe('doughnut');
    });

    test('radar chart falls back to bar with meta.unsupportedSourceType=radar + console.warn', async () => {
        const buf = await buildSyntheticChartXlsx({ chartXml: RADAR_CHART_BODY });
        const warn = jest.spyOn(console, 'warn').mockImplementation(() => {
            /* swallow */
        });
        try {
            const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
            const drawings = collectChartDrawings(snap);
            expect(drawings).toHaveLength(1);
            expect(drawings[0].data.type).toBe('bar');
            expect(drawings[0].data.meta?.unsupportedSourceType).toBe('radar');
            expect(warn).toHaveBeenCalled();
        } finally {
            warn.mockRestore();
        }
    });
});
