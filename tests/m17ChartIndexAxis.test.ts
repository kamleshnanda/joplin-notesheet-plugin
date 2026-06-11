// M17: charts whose source XML has no <c:cat> element (Excel infers
// row index 1..N as the X-axis) must round-trip without losing series
// or treating column 0 as a label column.
//
// 11-stacked-bar-chart.xlsx is the canonical fixture: two value series
// (Investment, Balance) referencing $A$2:$A$16 and $B$2:$B$16, with NO
// <c:cat> in the chart XML. Before this fix:
//   - Joplin's import resolver synthesized labels from column 0
//     (Investment values became X-axis labels — wrong)
//   - Joplin's M10 export assumed first column = labels, dropped the
//     second series (Investment values became X-axis, Balance series
//     was missing)
// After: meta.categoryAxisType === 'index' flows through both paths.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');
const FIXTURE_11 = path.join(FIXTURES_DIR, '11-stacked-bar-chart.xlsx');

interface ChartDrawing {
    componentKey: string;
    data: {
        type: string;
        labels: string[];
        datasets: Array<{ label?: string; data: number[] }>;
        meta?: { categoryAxisType?: 'index' | 'category'; legendPos?: string };
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

describe('M17 index-axis charts (no <c:cat>)', () => {
    test('11 imports as index-axis with both series intact and labels NOT synthesized from col 0', async () => {
        const snap = await xlsxBufferToSnapshot(readFileSync(FIXTURE_11) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        const d = drawings[0];

        // Two value series — Investment and Balance.
        expect(d.data.datasets).toHaveLength(2);
        const seriesLabels = d.data.datasets.map((s) => s.label).filter(Boolean);
        expect(seriesLabels).toEqual(expect.arrayContaining(['Investment', 'Balance']));

        // First series's data is the Investment column values.
        const investment = d.data.datasets.find((s) => s.label === 'Investment');
        expect(investment?.data[0]).toBe(5500);
        expect(investment?.data.length).toBe(15);

        // Second series's data is the Balance column values.
        const balance = d.data.datasets.find((s) => s.label === 'Balance');
        expect(balance?.data[0]).toBe(6798.22);
        expect(balance?.data.length).toBe(15);

        // categoryAxisType marker present.
        expect(d.data.meta?.categoryAxisType).toBe('index');

        // Labels are NOT contaminated with column-0 values. Either empty
        // (synthesized 1..N at render time) or matching some plausible
        // category set; either way they must NOT include the Investment
        // values.
        for (const lbl of d.data.labels) {
            expect(lbl).not.toBe('5500');
            expect(lbl).not.toBe('2084.51');
        }
    });

    test('11 round-trip preserves both series; export emits <c:ser> per column with no <c:cat>', async () => {
        const snap = await xlsxBufferToSnapshot(readFileSync(FIXTURE_11) as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(out as ArrayBuffer);
        const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
        expect(chartFile).toBeDefined();
        const chartXml = await zip.file(chartFile!)!.async('string');

        // Two <c:ser> blocks emitted (was 1 before the fix — Balance series
        // was being dropped because export treated column 0 as labels).
        const serMatches = chartXml.match(/<c:ser>/g) ?? [];
        expect(serMatches.length).toBe(2);

        // No <c:cat> on either series — index-axis charts use row index
        // 1..N (Excel's behaviour when <c:cat> is absent).
        expect(chartXml).not.toMatch(/<c:cat>/);

        // Each series's values formula points at its own column. The
        // exported sheet name is what xlsxBufferToSnapshot landed it as
        // (Excel's "Sheet1" → snapshot sheet whose name is "Sheet1").
        // Investment column → $A$1:$A$15 (data starts at row 1 in
        // 0-indexed snapshot, which Excel writes as $A$2:$A$16). Match
        // a relaxed pattern.
        const valFormulas = [...chartXml.matchAll(/<c:val>[\s\S]*?<c:f>([^<]+)<\/c:f>/g)].map((m) => m[1]);
        expect(valFormulas).toHaveLength(2);
        // Series 0 should hit column A, series 1 column B.
        expect(valFormulas[0]).toMatch(/\$A\$\d+:\$A\$\d+/);
        expect(valFormulas[1]).toMatch(/\$B\$\d+:\$B\$\d+/);
    });

    test('11 round-trip preserves the legendPos="b" (bottom)', async () => {
        const snap = await xlsxBufferToSnapshot(readFileSync(FIXTURE_11) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        // Source XML has <c:legendPos val="b"/>. Confirm it landed in meta.
        expect(drawings[0].data.meta?.legendPos).toBe('b');

        // And the export emits it back.
        const out = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(out as ArrayBuffer);
        const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
        const chartXml = await zip.file(chartFile!)!.async('string');
        expect(chartXml).toMatch(/<c:legendPos\s+val="b"/);
    });

    test('category-axis charts (e.g. 01) still emit <c:cat> on export (no regression)', async () => {
        const snap = await xlsxBufferToSnapshot(readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx')) as unknown as Buffer);
        const out = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(out as ArrayBuffer);
        const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
        const chartXml = await zip.file(chartFile!)!.async('string');
        // Category-axis charts MUST still ship <c:cat>.
        expect(chartXml).toMatch(/<c:cat>/);
    });
});
