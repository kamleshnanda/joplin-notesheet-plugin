// M17 fidelity (manual-test issues 3, 6): pie / doughnut slice labels
// must survive the import → export round-trip.
//
// The bug: Excel stores a pie's slice-label flags (showCatName=1,
// showPercent=1) on the SERIES-level <c:ser><c:dLbls>, while writing a
// chart-level <c:dLbls> with everything OFF. Our importer read only the
// chart-level block, so it imported "no labels" — and the re-exported
// pie had no slice labels. We now fall back to the series-level dLbls
// when the chart-level one shows nothing.
//
// Anchored to the Excel-canonical source: we first read what the source
// turns on (at whichever level), then assert the export reproduces it.

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

// Read every <c:show*> flag from ANY <c:dLbls> in a chart xml (chart- or
// series-level), OR'd together — "does this chart show this label kind
// anywhere". Mirrors how Excel actually renders: a flag set at either
// level lights the labels.
function anyShowFlags(chartXml: string): {
    showCatName: boolean;
    showPercent: boolean;
    showVal: boolean;
} {
    const has = (name: string) => new RegExp(`<(?:c:)?${name}\\s+val="1"`).test(chartXml);
    return {
        showCatName: has('showCatName'),
        showPercent: has('showPercent'),
        showVal: has('showVal'),
    };
}

async function chartXmls(
    buffer: ArrayBuffer | Buffer,
    predicate: (xml: string) => boolean,
): Promise<string[]> {
    const zip = await JSZip.loadAsync(buffer);
    const out: string[] = [];
    for (const k of Object.keys(zip.files)) {
        if (!/^xl\/charts\/chart\d+\.xml$/.test(k)) continue;
        const xml = await zip.file(k)!.async('string');
        if (predicate(xml)) out.push(xml);
    }
    return out;
}

const isPieish = (xml: string) => /<(?:c:)?(pie|doughnut)Chart\b/.test(xml);

describe('M17 pie/doughnut data-label round-trip (issues 3/6)', () => {
    test('03-pie-single.xlsx: source shows slice labels (cat name + percent)', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '03-pie-single.xlsx'));
        const pies = await chartXmls(buf as unknown as Buffer, isPieish);
        expect(pies.length).toBe(1);
        const flags = anyShowFlags(pies[0]);
        // The reference: the source DOES show slice labels somewhere.
        expect(flags.showCatName).toBe(true);
        expect(flags.showPercent).toBe(true);
    });

    test('03-pie-single.xlsx: import captures the slice-label flags onto meta.dLbls', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '03-pie-single.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const resources =
            (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!;
        const parsed = JSON.parse(entry.data);
        const sub = Object.values(parsed)[0] as { data: Record<string, any> };
        const drawing = Object.values(sub.data)[0] as {
            data: { type: string; meta?: { dLbls?: Record<string, boolean> } };
        };
        expect(drawing.data.type).toBe('pie');
        // The fix: the series-level flags reached meta.dLbls.
        expect(drawing.data.meta?.dLbls?.showCatName).toBe(true);
        expect(drawing.data.meta?.dLbls?.showPercent).toBe(true);
    });

    test('03-pie-single.xlsx: export re-emits the slice labels', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '03-pie-single.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const pies = await chartXmls(out, isPieish);
        expect(pies.length).toBe(1);
        const flags = anyShowFlags(pies[0]);
        expect(flags.showCatName).toBe(true);
        expect(flags.showPercent).toBe(true);
    });

    test('06-two-charts-one-sheet.xlsx: the pie keeps its labels; the bar does not gain spurious ones', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '06-two-charts-one-sheet.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);

        const pies = await chartXmls(out, isPieish);
        expect(pies.length).toBeGreaterThanOrEqual(1);
        const pieFlags = anyShowFlags(pies[0]);
        expect(pieFlags.showCatName || pieFlags.showPercent).toBe(true);

        // The bar chart on the same sheet must NOT show category/percent
        // labels (its source series-level dLbls turned nothing on).
        const bars = await chartXmls(out, (xml) => /<(?:c:)?barChart\b/.test(xml));
        expect(bars.length).toBeGreaterThanOrEqual(1);
        const barFlags = anyShowFlags(bars[0]);
        expect(barFlags.showCatName).toBe(false);
        expect(barFlags.showPercent).toBe(false);
    });

    test('round-trip is idempotent: re-importing the export keeps the pie label flags', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '03-pie-single.xlsx'));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap1);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(out) as unknown as Buffer);
        const resources =
            (snap2 as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!;
        const parsed = JSON.parse(entry.data);
        const sub = Object.values(parsed)[0] as { data: Record<string, any> };
        const drawing = Object.values(sub.data)[0] as {
            data: { meta?: { dLbls?: Record<string, boolean> } };
        };
        expect(drawing.data.meta?.dLbls?.showCatName).toBe(true);
        expect(drawing.data.meta?.dLbls?.showPercent).toBe(true);
    });
});
