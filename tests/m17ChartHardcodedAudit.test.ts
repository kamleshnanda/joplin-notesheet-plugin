// M17 Option-B audit follow-through: 6 chart-XML attributes that used
// to be hardcoded on export now round-trip through `meta`. This test
// pins each value's import/export and the renderer routing where one
// exists.
//
// Coverage:
//   1. holeSize (doughnut)             — 04-doughnut.xlsx ships 75
//   2. lineSmooth (per-series)         — 02-line-multi-series ships 0
//   3. lineMarkerOn (chart-level)      — 02 has no chart-level <c:marker>
//   4. dispBlanksAs (any chart)        — every fixture ships 'gap'
//   5. crossBetween (bar/line)         — every bar fixture ships 'between'
//   6. (axId is intentionally NOT plumbed — within-chart placeholder
//      identifiers, not data-affecting; documented in xlsxChart.ts.)

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
        type: string;
        meta?: {
            holeSize?: number;
            lineSmooth?: boolean;
            lineMarkerOn?: boolean;
            dispBlanksAs?: 'gap' | 'zero' | 'span';
            crossBetween?: 'between' | 'midCat';
        };
    };
}

function collectChartDrawings(snap: unknown): ChartDrawing[] {
    const resources = (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
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

describe('M17 hardcoded-audit follow-through: holeSize round-trip', () => {
    test('04-doughnut imports as meta.holeSize=75 (matches source XML)', async () => {
        const fp = path.join(FIXTURES_DIR, '04-doughnut.xlsx');
        // Sanity: source has <c:holeSize val="75"/>, not the prior
        // hardcoded 50.
        const sourceXml = await JSZip.loadAsync(readFileSync(fp) as unknown as ArrayBuffer)
            .then((z) => z.file('xl/charts/chart1.xml')!.async('string'));
        expect(sourceXml).toMatch(/<c:holeSize\s+val="75"/);

        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.meta?.holeSize).toBe(75);
    });

    test('04-doughnut round-trip emits <c:holeSize val="75"/>', async () => {
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '04-doughnut.xlsx'));
        expect(xml).toMatch(/<c:holeSize\s+val="75"/);
    });

    test('non-doughnut charts (03-pie) carry no holeSize', async () => {
        const fp = path.join(FIXTURES_DIR, '03-pie-single.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings[0].data.meta?.holeSize).toBeUndefined();
    });
});

describe('M17 hardcoded-audit follow-through: line smooth + markers', () => {
    test('02-line-multi-series imports lineSmooth=false (source <c:smooth val="0"/>)', async () => {
        const fp = path.join(FIXTURES_DIR, '02-line-multi-series.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings[0].data.meta?.lineSmooth).toBe(false);
    });

    test('02 round-trip emits <c:smooth val="0"/> on every series', async () => {
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '02-line-multi-series.xlsx'));
        // 02 has 3 series — expect 3 <c:smooth> elements with val="0".
        const matches = [...xml.matchAll(/<c:smooth\s+val="0"\/>/g)];
        expect(matches.length).toBeGreaterThanOrEqual(3);
        // None smoothed.
        expect(xml).not.toMatch(/<c:smooth\s+val="1"/);
    });

    test('non-line charts carry no lineSmooth/lineMarkerOn meta', async () => {
        for (const fixture of ['01-bar-simple', '03-pie-single', '04-doughnut']) {
            const fp = path.join(FIXTURES_DIR, `${fixture}.xlsx`);
            const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
            const drawings = collectChartDrawings(snap);
            expect(drawings[0].data.meta?.lineSmooth).toBeUndefined();
            expect(drawings[0].data.meta?.lineMarkerOn).toBeUndefined();
        }
    });
});

describe('M17 hardcoded-audit follow-through: dispBlanksAs + crossBetween', () => {
    test('all fixtures import dispBlanksAs="gap" (source default)', async () => {
        const fixtures = [
            '01-bar-simple', '02-line-multi-series', '03-pie-single',
            '04-doughnut', '11-stacked-bar-chart',
        ];
        for (const fixture of fixtures) {
            const fp = path.join(FIXTURES_DIR, `${fixture}.xlsx`);
            const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
            const drawings = collectChartDrawings(snap);
            expect(drawings[0].data.meta?.dispBlanksAs).toBe('gap');
        }
    });

    test('bar/line fixtures import crossBetween="between"', async () => {
        for (const fixture of ['01-bar-simple', '02-line-multi-series', '11-stacked-bar-chart']) {
            const fp = path.join(FIXTURES_DIR, `${fixture}.xlsx`);
            const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
            const drawings = collectChartDrawings(snap);
            expect(drawings[0].data.meta?.crossBetween).toBe('between');
        }
    });

    test('pie/doughnut have no crossBetween (no value axis)', async () => {
        for (const fixture of ['03-pie-single', '04-doughnut']) {
            const fp = path.join(FIXTURES_DIR, `${fixture}.xlsx`);
            const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
            const drawings = collectChartDrawings(snap);
            expect(drawings[0].data.meta?.crossBetween).toBeUndefined();
        }
    });

    test('11 round-trip preserves dispBlanksAs and crossBetween in exported XML', async () => {
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx'));
        expect(xml).toMatch(/<c:dispBlanksAs\s+val="gap"/);
        expect(xml).toMatch(/<c:crossBetween\s+val="between"/);
    });
});
