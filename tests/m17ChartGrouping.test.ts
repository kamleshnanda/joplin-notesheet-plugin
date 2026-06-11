// M17: bar/line grouping (clustered / stacked / percentStacked /
// standard). 11-stacked-bar-chart's source XML has
// <c:grouping val="stacked"/>. Without a meta.barGrouping plumb,
// Notesheet rendered AND re-exported as clustered, silently changing
// the chart's intent.
//
// Anchored UPSTREAM: each fixture's expected grouping is independently
// parsed from <c:grouping val="..."/> via JSZip+regex, NOT from what
// Notesheet emits.

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
        meta?: { barGrouping?: 'clustered' | 'stacked' | 'percentStacked' | 'standard' };
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

async function readSourceGrouping(fp: string): Promise<string | null> {
    const zip = await JSZip.loadAsync(readFileSync(fp) as unknown as ArrayBuffer);
    const xml = await zip.file('xl/charts/chart1.xml')!.async('string');
    const m = xml.match(/<(?:c:)?grouping\s+val="(clustered|stacked|percentStacked|standard)"/);
    return m ? m[1] : null;
}

async function exportedChartXml(fp: string): Promise<string> {
    const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
    const out = await snapshotToXlsxBuffer(snap);
    const zip = await JSZip.loadAsync(out as ArrayBuffer);
    const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
    return zip.file(chartFile!)!.async('string');
}

describe('M17 chart grouping (stacked / clustered) round-trip', () => {
    test('11-stacked-bar imports as meta.barGrouping=stacked (matches source XML)', async () => {
        const fp = path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx');
        expect(await readSourceGrouping(fp)).toBe('stacked');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].data.meta?.barGrouping).toBe('stacked');
    });

    test('11 round-trip: exported chart XML carries <c:grouping val="stacked"/> + <c:overlap val="100"/>', async () => {
        const fp = path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx');
        const xml = await exportedChartXml(fp);
        expect(xml).toMatch(/<c:grouping\s+val="stacked"/);
        // Stacked bars must have overlap=100 so segments stack flush.
        expect(xml).toMatch(/<c:overlap\s+val="100"/);
    });

    test('clustered fixtures (01) keep clustered + no overlap=100 on round-trip', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const xml = await exportedChartXml(fp);
        expect(xml).toMatch(/<c:grouping\s+val="clustered"/);
        // Clustered emits no overlap element; assert by absence.
        expect(xml).not.toMatch(/<c:overlap\s+val="100"/);
    });

    test('line chart (02) preserves its grouping (standard)', async () => {
        const fp = path.join(FIXTURES_DIR, '02-line-multi-series.xlsx');
        const sourceGrouping = await readSourceGrouping(fp);
        // 02's source uses 'standard' (line default — non-stacked).
        expect(sourceGrouping).toBe('standard');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings[0].data.meta?.barGrouping).toBe('standard');
        const xml = await exportedChartXml(fp);
        expect(xml).toMatch(/<c:grouping\s+val="standard"/);
    });
});
