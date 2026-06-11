// M17: chart legend survives export. Excel's <c:legend> element belongs
// inside <c:chart> after <c:plotArea>. Notesheet's M10 emitter previously
// omitted it entirely — multi-series line/bar charts and pie/doughnut
// charts opened in Excel without a legend, even though the runtime
// NotesheetChart component shows one. This test pins the export side.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');

async function exportedChartXml(buf: Buffer): Promise<string> {
    const out = await snapshotToXlsxBuffer(await xlsxBufferToSnapshot(buf as unknown as Buffer));
    const zip = await JSZip.loadAsync(out as ArrayBuffer);
    const chartFile = Object.keys(zip.files).find((k) => /^xl\/charts\/chart\d+\.xml$/.test(k));
    if (!chartFile) throw new Error('no chart in export');
    return zip.file(chartFile)!.async('string');
}

describe('M17 chart legend on export', () => {
    test('multi-series line chart (02) exports with <c:legend> at source position', async () => {
        const xml = await exportedChartXml(readFileSync(path.join(FIXTURES_DIR, '02-line-multi-series.xlsx')));
        expect(xml).toMatch(/<c:legend>/);
        // 02's source XML has <c:legendPos val="b"/> (bottom) — the
        // round-trip should preserve it via meta.legendPos. If the
        // import side ever drops the field, this falls back to 'r'
        // and the test fails — which is what we want.
        expect(xml).toMatch(/<c:legendPos\s+val="b"/);
    });

    test('pie chart (03) exports with <c:legend>', async () => {
        const xml = await exportedChartXml(readFileSync(path.join(FIXTURES_DIR, '03-pie-single.xlsx')));
        expect(xml).toMatch(/<c:legend>/);
    });

    test('doughnut chart (04) exports with <c:legend>', async () => {
        const xml = await exportedChartXml(readFileSync(path.join(FIXTURES_DIR, '04-doughnut.xlsx')));
        expect(xml).toMatch(/<c:legend>/);
    });

    test('single-series bar chart (01) exports without <c:legend> (Excel default)', async () => {
        const xml = await exportedChartXml(readFileSync(path.join(FIXTURES_DIR, '01-bar-simple.xlsx')));
        expect(xml).not.toMatch(/<c:legend>/);
    });

    test('legend element sits between </c:plotArea> and <c:plotVisOnly> per ECMA-376', async () => {
        // Order inside <c:chart>: title, autoTitleDeleted, plotArea,
        // legend, plotVisOnly, dispBlanksAs. Excel rejects out-of-order.
        const xml = await exportedChartXml(readFileSync(path.join(FIXTURES_DIR, '02-line-multi-series.xlsx')));
        const plotAreaEnd = xml.indexOf('</c:plotArea>');
        const legendStart = xml.indexOf('<c:legend>');
        const plotVisOnly = xml.indexOf('<c:plotVisOnly');
        expect(plotAreaEnd).toBeGreaterThan(0);
        expect(legendStart).toBeGreaterThan(plotAreaEnd);
        expect(plotVisOnly).toBeGreaterThan(legendStart);
    });
});
