// M17 fidelity (manual-test issues 1, 2, 7): exported bar/line charts
// must carry the SAME axis-line styling the source Excel chart uses —
// a thin light-grey CATEGORY axis line and a NO-LINE value axis — NOT
// Excel's dark legacy default that appears when <c:spPr> is omitted.
//
// The bug: our M10 export emitted neither axis's <c:spPr>, so Excel
// painted a dark solid line on BOTH axes. Against the source (which
// hides the value axis and lightens the category axis) this read as
// "Joplin introduced axes / a vertical line that wasn't there".
//
// These tests anchor to the EXCEL-CANONICAL reference (the source
// fixture's own chart XML), not to our own emit: we first assert what
// the source does, then assert our export reproduces it. They also pin
// the import→export round-trip of the catAxisLine/valAxisLine meta.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');

async function firstChartXml(buffer: ArrayBuffer | Buffer): Promise<string> {
    const zip = await JSZip.loadAsync(buffer);
    const chartPath = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
    if (!chartPath) throw new Error('no chart xml');
    return zip.file(chartPath)!.async('string');
}

// Classify an axis's OWN line (ignoring the gridlines block) as the
// source/export wrote it: 'none' (<a:ln><a:noFill/>), 'line' (any
// visible <a:ln ...>), or 'absent' (no <c:spPr> at all → Excel paints
// its dark default).
function axisLineKind(chartXml: string, axTag: 'catAx' | 'valAx'): 'none' | 'line' | 'absent' {
    const block = chartXml.match(new RegExp(`<(?:c:)?${axTag}>([\\s\\S]*?)</(?:c:)?${axTag}>`));
    if (!block) return 'absent';
    const body = block[1].replace(/<(?:c:)?majorGridlines>[\s\S]*?<\/(?:c:)?majorGridlines>/g, '');
    const spPr = body.match(/<(?:c:)?spPr>([\s\S]*?)<\/(?:c:)?spPr>/);
    if (!spPr) return 'absent';
    if (/<a:ln>\s*<a:noFill\/>\s*<\/a:ln>/.test(spPr[1])) return 'none';
    if (/<a:ln\b/.test(spPr[1])) return 'line';
    return 'absent';
}

const BAR_LINE_FIXTURES = [
    '01-bar-simple.xlsx',
    '02-line-multi-series.xlsx',
    '07-chart-cross-sheet.xlsx',
    '09-bar-percent-axis.xlsx',
    '11-stacked-bar-chart.xlsx',
];

describe('M17 axis-line fidelity (issues 1/2/7)', () => {
    for (const fixture of BAR_LINE_FIXTURES) {
        test(`${fixture}: source styles cat=grey-line, val=no-line (Excel-canonical anchor)`, async () => {
            const buf = readFileSync(path.join(FIXTURES_DIR, fixture));
            const srcXml = await firstChartXml(buf as unknown as Buffer);
            // The reference: every modern-template chart hides the value
            // axis line and draws a light category axis line.
            expect(axisLineKind(srcXml, 'catAx')).toBe('line');
            expect(axisLineKind(srcXml, 'valAx')).toBe('none');
        });

        test(`${fixture}: our export reproduces the source axis-line styling`, async () => {
            const buf = readFileSync(path.join(FIXTURES_DIR, fixture));
            const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
            const out = await snapshotToXlsxBuffer(snap);
            const outXml = await firstChartXml(out);
            // The fix: BOTH axes carry an explicit <c:spPr> (never
            // 'absent', which is what made Excel paint dark defaults).
            expect(axisLineKind(outXml, 'catAx')).not.toBe('absent');
            expect(axisLineKind(outXml, 'valAx')).not.toBe('absent');
            // And they match the source's intent: cat has a line, val
            // has none.
            expect(axisLineKind(outXml, 'catAx')).toBe('line');
            expect(axisLineKind(outXml, 'valAx')).toBe('none');
        });
    }

    test('catAxisLine/valAxisLine meta round-trips import → export', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const resources = (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
        expect(entry).toBeDefined();
        const parsed = JSON.parse(entry!.data);
        const sub = Object.values(parsed)[0] as { data: Record<string, any> };
        const drawing = Object.values(sub.data)[0] as { data: { meta?: { catAxisLine?: string; valAxisLine?: string } } };
        // Import captured the source's modern-template axis lines.
        expect(drawing.data.meta?.catAxisLine).toBe('grey');
        expect(drawing.data.meta?.valAxisLine).toBe('none');
    });

    test('a chart whose source has a visible value-axis line keeps it on export', async () => {
        // Synthesize a snapshot whose chart meta explicitly asks for a
        // grey value-axis line (the non-default case). Export must honor
        // it rather than forcing 'none'. Anchors the meta override path.
        const drawingResource = {
            'sheet-1': {
                data: {
                    'chart-x': {
                        unitId: 'workbook', subUnitId: 'sheet-1', drawingId: 'chart-x',
                        drawingType: 8, componentKey: 'NotesheetChart', allowTransform: true,
                        data: {
                            chartId: 'chart-x', type: 'bar', title: 'T',
                            sourceRange: { startRow: 0, endRow: 2, startColumn: 0, endColumn: 1 },
                            labels: ['a', 'b'], datasets: [{ label: 's', data: [1, 2] }],
                            meta: { catAxisLine: 'grey', valAxisLine: 'grey' },
                        },
                        axisAlignSheetTransform: {
                            from: { column: 3, columnOffset: 0, row: 0, rowOffset: 0 },
                            to: { column: 9, columnOffset: 0, row: 14, rowOffset: 0 },
                        },
                    },
                },
                order: ['chart-x'],
            },
        };
        const snap = {
            id: 'wb', sheetOrder: ['sheet-1'], name: 'S', appVersion: '0.1.0', locale: 'enUS',
            styles: {}, sheets: { 'sheet-1': { id: 'sheet-1', name: 'Sheet1', cellData: {}, rowCount: 100, columnCount: 26 } },
            resources: [{ name: 'SHEET_DRAWING_PLUGIN', data: JSON.stringify(drawingResource) }],
        };
        const out = await snapshotToXlsxBuffer(snap as unknown as Record<string, unknown>);
        const outXml = await firstChartXml(out);
        expect(axisLineKind(outXml, 'catAx')).toBe('line');
        expect(axisLineKind(outXml, 'valAx')).toBe('line');
    });
});
