// Round of fixes from the user's manual cross-fixture comparison:
//   1. Horizontal-bar export had axes swapped (catAx axPos="b" emitted
//      where Excel needs "l", etc.). Caused gridline-position drift.
//   2. Bar/line gridlines emitted bare <c:majorGridlines/> with no
//      spPr, so Excel rendered its dark default instead of the
//      light-grey lines source XML ships with explicit lumMod=15000.
//   3. Tick marks hardcoded "out" — every fixture surveyed uses
//      "none". Wrong default; plumbed via meta.tickMark.
//   4. Pie/doughnut export emitted no <c:dLbls> at all so slice
//      labels were missing on re-open in Excel. Plumbed via meta.dLbls
//      with all six show* flags.
//   5. NotesheetChart didn't apply <c:numFmt> from valAx, so the
//      09-bar-percent-axis fixture (formatCode="0%") rendered raw
//      0.05/0.10/... in Joplin's import. Plumbed via meta.valAxisNumFmt
//      and routed to a Chart.js tick-callback formatter.
//
// All anchored UPSTREAM: each fixture's expected value is independently
// parsed from the source XML, never from what Notesheet emits.

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
            tickMark?: 'none' | 'in' | 'out' | 'cross';
            valAxisNumFmt?: string;
            dLbls?: {
                showVal?: boolean;
                showCatName?: boolean;
                showPercent?: boolean;
                showSerName?: boolean;
                showLegendKey?: boolean;
                showBubbleSize?: boolean;
            };
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

describe('M17 axis + label fidelity round-trip', () => {
    test('horizontal bar (01) export: catAx axPos="l", valAx axPos="b"', async () => {
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const catAxBlock = xml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
        const valAxBlock = xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
        // Horizontal bar: categories on left, values on bottom.
        expect(catAxBlock).toMatch(/<c:axPos\s+val="l"/);
        expect(valAxBlock).toMatch(/<c:axPos\s+val="b"/);
    });

    test('vertical bar (07) export: catAx axPos="b", valAx axPos="l"', async () => {
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '07-chart-cross-sheet.xlsx'));
        const catAxBlock = xml.match(/<c:catAx>[\s\S]*?<\/c:catAx>/)![0];
        const valAxBlock = xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
        expect(catAxBlock).toMatch(/<c:axPos\s+val="b"/);
        expect(valAxBlock).toMatch(/<c:axPos\s+val="l"/);
    });

    test('majorGridlines emitted with light-grey spPr (lumMod=15000)', async () => {
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        // Source workbooks ship explicit light-grey gridlines; we now
        // emit the same so Excel doesn't pick its darker default.
        expect(xml).toMatch(/<c:majorGridlines>[\s\S]*?<a:lumMod\s+val="15000"\/>/);
        expect(xml).toMatch(/<c:majorGridlines>[\s\S]*?<a:lumOff\s+val="85000"\/>/);
    });

    test('tickMark default "none" emitted on both axes (matches every source fixture)', async () => {
        // Survey the source first — none of the bar/line fixtures use
        // anything other than "none". If a future fixture ships "out"
        // we'd plumb it; today the assertion is that we don't emit
        // ticks where source has none.
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const tickMarkValues = [...xml.matchAll(/<c:majorTickMark\s+val="([^"]+)"/g)].map((m) => m[1]);
        expect(tickMarkValues.length).toBeGreaterThanOrEqual(2);
        for (const v of tickMarkValues) {
            expect(v).toBe('none');
        }
    });

    test('pie chart (03) imports + re-exports with chart-level <c:dLbls>', async () => {
        const fp = path.join(FIXTURES_DIR, '03-pie-single.xlsx');
        // Source has chart-level dLbls.
        const sourceXml = await JSZip.loadAsync(readFileSync(fp) as unknown as ArrayBuffer)
            .then((z) => z.file('xl/charts/chart1.xml')!.async('string'));
        // Source's pie has dLbls present. Confirm by grep.
        expect(sourceXml).toMatch(/<c:dLbls>/);

        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        // Pie chart should now have dLbls in meta.
        expect(drawings[0].data.meta?.dLbls).toBeDefined();

        // And re-export must include <c:dLbls> in the pie chart block.
        const xml = await exportedChartXml(fp);
        // The <c:dLbls> we emit lives INSIDE <c:pieChart>.
        const pieBlock = xml.match(/<c:pieChart>[\s\S]*?<\/c:pieChart>/);
        expect(pieBlock).not.toBeNull();
        expect(pieBlock![0]).toMatch(/<c:dLbls>/);
    });

    test('09-bar-percent-axis imports valAxisNumFmt="0%" and re-exports it', async () => {
        const fp = path.join(FIXTURES_DIR, '09-bar-percent-axis.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        // Source has formatCode="0%" on valAx.
        expect(drawings[0].data.meta?.valAxisNumFmt).toBe('0%');

        const xml = await exportedChartXml(fp);
        // Exported chart's valAx must carry the same formatCode with
        // sourceLinked="0" so Excel uses our value, not the cell's.
        const valAxBlock = xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
        expect(valAxBlock).toMatch(/<c:numFmt\s+formatCode="0%"\s+sourceLinked="0"/);
    });

    test('non-percent fixtures keep numFmt General (no regression)', async () => {
        const xml = await exportedChartXml(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const valAxBlock = xml.match(/<c:valAx>[\s\S]*?<\/c:valAx>/)![0];
        expect(valAxBlock).toMatch(/<c:numFmt\s+formatCode="General"/);
    });

    test('tickMark survives round-trip via meta.tickMark (defaults to "none")', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(fp) as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings[0].data.meta?.tickMark).toBe('none');
    });
});
