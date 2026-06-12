// M17: chart title and bar gap-width round-trip.
//
// Title: Excel allows multi-run titles
// (<a:r><a:t>...</a:t></a:r><a:r><a:t>...</a:t></a:r>) — each run is a
// formatting variation but the displayed title is all runs concatenated.
// Earlier the import resolver matched only the first <a:t>, truncating
// "Investment vs Balance so far" to "Investment".
//
// Gap width: Excel <c:gapWidth val="N"/> is the gap between bar groups
// as a percent of bar width. Source values vary (150, 182, 219). M10
// previously hardcoded 182, making exported bars narrower than source
// for some fixtures and wider for others. Now plumbed via meta.barGapWidth.

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
        title: string;
        meta?: { barGapWidth?: number };
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

async function exportedChartXml(fp: string): Promise<string> {
    const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
    const out = await snapshotToXlsxBuffer(snap);
    const zip = await JSZip.loadAsync(out as ArrayBuffer);
    const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
    return zip.file(chartFile!)!.async('string');
}

describe('M17 chart title (multi-run concatenation)', () => {
    test('11-stacked-bar imports full title "Investment vs Balance so far"', async () => {
        const fp = path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx');
        // Sanity: source XML has two <a:t> runs.
        const sourceXml = await JSZip.loadAsync(readFileSync(fp) as unknown as ArrayBuffer).then(
            (z) => z.file('xl/charts/chart1.xml')!.async('string'),
        );
        const runs = [...sourceXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]);
        // First two runs are inside <c:title>; subsequent runs are
        // axis labels / etc. Just verify the first two concatenate.
        expect(runs.slice(0, 2).join('')).toBe('Investment vs Balance so far');

        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.title).toBe('Investment vs Balance so far');
    });

    test('single-run titles still import correctly (01)', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings[0].data.title).toBeTruthy();
        expect(drawings[0].data.title.length).toBeGreaterThan(0);
    });

    test('11 round-trip preserves full title', async () => {
        const fp = path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx');
        const xml = await exportedChartXml(fp);
        // Exported title — single <a:t> with the full concatenated
        // string. We don't try to preserve the multi-run structure
        // (formatting per-run is out of scope) but the displayed text
        // must round-trip.
        expect(xml).toMatch(/<a:t>Investment vs Balance so far<\/a:t>/);
    });
});

describe('M17 bar gap-width (round-trip)', () => {
    test('11 imports meta.barGapWidth=150 (matches source XML)', async () => {
        const fp = path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings[0].data.meta?.barGapWidth).toBe(150);
    });

    test('11 round-trip emits <c:gapWidth val="150"/>', async () => {
        const fp = path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx');
        const xml = await exportedChartXml(fp);
        expect(xml).toMatch(/<c:gapWidth\s+val="150"/);
    });

    test('07 (gapWidth=219) round-trips with same value', async () => {
        const fp = path.join(FIXTURES_DIR, '07-chart-cross-sheet.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings[0].data.meta?.barGapWidth).toBe(219);
        const xml = await exportedChartXml(fp);
        expect(xml).toMatch(/<c:gapWidth\s+val="219"/);
    });

    test('non-bar charts (line/pie/doughnut) carry no barGapWidth', async () => {
        for (const fixture of ['02-line-multi-series', '03-pie-single', '04-doughnut']) {
            const fp = path.join(FIXTURES_DIR, `${fixture}.xlsx`);
            const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
            const drawings = collectChartDrawings(snap);
            expect(drawings[0].data.meta?.barGapWidth).toBeUndefined();
        }
    });
});
