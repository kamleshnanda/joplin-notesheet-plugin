// M18-pulled-into-M17: chart trendline import → meta → export round-trip.
//
// Excel stores a per-series fitted line as <c:ser><c:trendline> with a
// <c:trendlineType> and optional <c:dispRSqr>/<c:dispEq> label flags.
// 10-bar-with-trendline.xlsx carries a linear trendline with both labels
// on. We now parse it into meta.trendline, render it (computed least-
// squares overlay — not assertable here), and re-emit <c:trendline> on
// export so the round-trip preserves it.
//
// Anchored to the Excel-canonical source: read what the source's
// trendline is, then assert import captured it and export reproduces it.

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
const FIXTURE = '10-bar-with-trendline.xlsx';

function firstChartMeta(snap: unknown): Record<string, any> | undefined {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return undefined;
    const parsed = JSON.parse(entry.data);
    const sub = Object.values(parsed)[0] as { data: Record<string, any> };
    const drawing = Object.values(sub.data)[0] as { data: { meta?: Record<string, any> } };
    return drawing.data.meta;
}

async function firstChartXml(buffer: ArrayBuffer | Buffer): Promise<string> {
    const zip = await JSZip.loadAsync(buffer);
    const k = Object.keys(zip.files).find((p) => /^xl\/charts\/chart\d+\.xml$/.test(p));
    if (!k) throw new Error('no chart xml');
    return zip.file(k)!.async('string');
}

describe('M17 chart trendline round-trip (fixture 10)', () => {
    test('source has a linear trendline with dispEq + dispRSqr (Excel-canonical anchor)', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const xml = await firstChartXml(buf as unknown as Buffer);
        expect(/<c:trendline>/.test(xml)).toBe(true);
        expect(/<c:trendlineType\s+val="linear"/.test(xml)).toBe(true);
        expect(/<c:dispRSqr\s+val="1"/.test(xml)).toBe(true);
        expect(/<c:dispEq\s+val="1"/.test(xml)).toBe(true);
    });

    test('import captures the trendline onto meta.trendline', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const meta = firstChartMeta(snap);
        expect(meta?.trendline).toBeDefined();
        expect(meta!.trendline.type).toBe('linear');
        expect(meta!.trendline.dispRSqr).toBe(true);
        expect(meta!.trendline.dispEq).toBe(true);
    });

    test('export re-emits the trendline (round-trip preserves it)', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const xml = await firstChartXml(out);
        expect(/<c:trendline>/.test(xml)).toBe(true);
        expect(/<c:trendlineType\s+val="linear"/.test(xml)).toBe(true);
        expect(/<c:dispRSqr\s+val="1"/.test(xml)).toBe(true);
        expect(/<c:dispEq\s+val="1"/.test(xml)).toBe(true);
    });

    test('round-trip is idempotent: re-import keeps the trendline meta', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap1);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(out) as unknown as Buffer);
        const meta = firstChartMeta(snap2);
        expect(meta?.trendline?.type).toBe('linear');
    });

    test('charts WITHOUT a trendline gain no spurious <c:trendline> on export', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        expect(firstChartMeta(snap)?.trendline).toBeUndefined();
        const out = await snapshotToXlsxBuffer(snap);
        const xml = await firstChartXml(out);
        expect(/<c:trendline>/.test(xml)).toBe(false);
    });
});
