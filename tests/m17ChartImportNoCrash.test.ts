// M17 feature-1: chart import doesn't crash + snapshot carries
// SHEET_DRAWING_PLUGIN with at least one chart drawing per fixture.
//
// All assertions anchor UPSTREAM of Notesheet's emit:
//   * chart-type element name parsed independently from the source XML
//     (NOT from `xlsxBufferToSnapshot`'s output's `type` field).
//   * sourceRange parsed independently from the <c:cat>/<c:val> formula
//     refs in the chart XML, NOT from the snapshot's emitted sourceRange.
//   * anchor coordinates parsed independently from <xdr:from>/<xdr:to>
//     children in the drawing XML, NOT from the snapshot's emitted anchor.
//
// Per feedback_pge_fidelity_test_gap.md: anchoring tests to your own
// behaviour gives you confidence your code is consistent, NOT that it's
// correct.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync, readdirSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { NotesheetImportError, xlsxBufferToSnapshot } from '../src/xlsx';
import { decodeCellRef, readChartsFromXlsxZip } from '../src/charts/xlsxChartImport';
import {
    buildSyntheticChartXlsx,
    NO_CAT_CHART_BODY,
    RADAR_CHART_BODY,
    STANDARD_BAR_CHART_BODY,
} from './util/m17BuildSyntheticXlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');

interface ChartDrawing {
    componentKey: string;
    data: {
        chartId: string;
        type: string;
        title: string;
        sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
        sourceSheetName?: string;
        labels: string[];
        datasets: Array<{ label?: string; data: number[] }>;
        meta?: { unsupportedSourceType?: string };
    };
    sheetTransform: {
        from: { column: number; row: number };
        to: { column: number; row: number };
    };
}

// Walk the snapshot's SHEET_DRAWING_PLUGIN resource and return a flat
// list of ALL chart drawings across all subUnits, preserving order.
function collectChartDrawings(snap: unknown): Array<{ subUnitId: string; drawing: ChartDrawing }> {
    const resources = (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data);
    const out: Array<{ subUnitId: string; drawing: ChartDrawing }> = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        const order: string[] = sub.order ?? Object.keys(sub.data);
        for (const id of order) {
            const d = sub.data[id];
            if (d?.componentKey === 'NotesheetChart') out.push({ subUnitId, drawing: d as ChartDrawing });
        }
    }
    return out;
}

// Parse the fixture's chart XML to extract the source-of-truth values
// the snapshot must match. Returns the raw values from the source XML.
async function readSourceTruth(fixturePath: string): Promise<Array<{
    chartFile: string;
    drawingFile: string;
    anchorIndex: number;
    typeFromXml: string;
    sourceSheetName: string;
    sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
    anchorFrom: { col: number; row: number };
    anchorTo: { col: number; row: number };
}>> {
    const buf = readFileSync(fixturePath);
    const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
    const out: Array<{
        chartFile: string;
        drawingFile: string;
        anchorIndex: number;
        typeFromXml: string;
        sourceSheetName: string;
        sourceRange: { startRow: number; endRow: number; startColumn: number; endColumn: number };
        anchorFrom: { col: number; row: number };
        anchorTo: { col: number; row: number };
    }> = [];

    // Walk every drawing in the workbook in sheet order.
    const sheetRelPaths = Object.keys(zip.files)
        .filter((p) => /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(p))
        .sort();
    for (const relPath of sheetRelPaths) {
        const xml = await zip.file(relPath)!.async('string');
        const dr = xml.match(/Type="[^"]*\/drawing"[^/]*Target="([^"]+)"/);
        if (!dr) continue;
        const target = dr[1].replace(/^\.\.\//, '');
        const drawingPath = `xl/${target}`;
        const drawingFile = zip.file(drawingPath);
        if (!drawingFile) continue;
        const drawingXml = await drawingFile.async('string');
        const drawingRelsXml = await zip.file(drawingPath.replace(/^(.*\/)([^/]+)$/, '$1_rels/$2.rels'))!.async('string');

        // Build rId -> chart path map.
        const relMap = new Map<string, string>();
        const relWalkRe = /<Relationship\s+([^>]*?)\/>/g;
        let relMatchIter: RegExpExecArray | null;
        while ((relMatchIter = relWalkRe.exec(drawingRelsXml)) !== null) {
            const idMatch = relMatchIter[1].match(/Id="([^"]+)"/);
            const tgtMatch = relMatchIter[1].match(/Target="([^"]+)"/);
            if (idMatch && tgtMatch) relMap.set(idMatch[1], tgtMatch[1]);
        }

        // Walk anchors in document order.
        const anchorRe = /<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b[^>]*>([\s\S]*?)<\/xdr:\1>/g;
        let anchorIndex = 0;
        let anchorMatchIter: RegExpExecArray | null;
        while ((anchorMatchIter = anchorRe.exec(drawingXml)) !== null) {
            anchorIndex++;
            const body = anchorMatchIter[2];
            const idMatch = body.match(/<c:chart\b[^>]*r:id="([^"]+)"/);
            if (!idMatch) continue;
            const target = relMap.get(idMatch[1]);
            if (!target) continue;
            const chartPath = `xl/${target.replace(/^\.\.\//, '')}`;
            const chartXml = await zip.file(chartPath)!.async('string');

            const fromBody = body.match(/<xdr:from>([\s\S]*?)<\/xdr:from>/)![1];
            const toBody = body.match(/<xdr:to>([\s\S]*?)<\/xdr:to>/);
            const anchorFrom = {
                col: parseInt(fromBody.match(/<xdr:col>(\d+)<\/xdr:col>/)![1], 10),
                row: parseInt(fromBody.match(/<xdr:row>(\d+)<\/xdr:row>/)![1], 10),
            };
            const anchorTo = toBody ? {
                col: parseInt(toBody[1].match(/<xdr:col>(\d+)<\/xdr:col>/)![1], 10),
                row: parseInt(toBody[1].match(/<xdr:row>(\d+)<\/xdr:row>/)![1], 10),
            } : { col: anchorFrom.col + 6, row: anchorFrom.row + 14 };

            // Type from XML: regex over chart-type element name.
            const typeMatch = chartXml.match(/<c:(bar|line|pie|doughnut|radar|scatter|area)Chart\b/);
            const typeFromXml = typeMatch ? typeMatch[1] : 'unknown';

            // Source range — bounding box of the cat-ref + each val-ref. Note
            // the FIRST <c:f> in a series is the series-name ref nested in
            // <c:tx><c:strRef>; categories come SECOND inside <c:cat>, values
            // THIRD inside <c:val>. We must explicitly scope to the cat+val
            // sub-elements rather than grabbing the first <c:f>.
            const refs: { startRow: number; endRow: number; startColumn: number; endColumn: number; sheet: string }[] = [];
            const seriesRe = /<c:ser>([\s\S]*?)<\/c:ser>/g;
            let seriesMatchIter: RegExpExecArray | null;
            let labelsRangeRow: number | null = null;
            while ((seriesMatchIter = seriesRe.exec(chartXml)) !== null) {
                const ser = seriesMatchIter[1];
                const catMatch = ser.match(/<c:cat>([\s\S]*?)<\/c:cat>/);
                if (catMatch) {
                    const catFormula = catMatch[1].match(/<c:f>([\s\S]*?)<\/c:f>/);
                    if (catFormula) {
                        const ref = decodeCellRef(catFormula[1]);
                        if (ref) {
                            refs.push({ ...ref, sheet: ref.sheetName });
                            if (labelsRangeRow == null) labelsRangeRow = ref.startRow;
                        }
                    }
                }
                const valMatch = ser.match(/<c:val>([\s\S]*?)<\/c:val>/);
                if (valMatch) {
                    const valFormula = valMatch[1].match(/<c:f>([\s\S]*?)<\/c:f>/);
                    if (valFormula) {
                        const ref = decodeCellRef(valFormula[1]);
                        if (ref) refs.push({ ...ref, sheet: ref.sheetName });
                    }
                }
            }
            if (refs.length === 0) continue;
            let startRow = refs[0].startRow;
            let endRow = refs[0].endRow;
            let startColumn = refs[0].startColumn;
            let endColumn = refs[0].endColumn;
            // Series-header (label) row sits one row above the data row when
            // a <c:cat> ref exists; include it so the source range covers
            // the same area Notesheet emits.
            if (labelsRangeRow != null) startRow = Math.min(startRow, labelsRangeRow - 1);
            for (const r of refs) {
                if (r.startRow < startRow) startRow = r.startRow;
                if (r.endRow > endRow) endRow = r.endRow;
                if (r.startColumn < startColumn) startColumn = r.startColumn;
                if (r.endColumn > endColumn) endColumn = r.endColumn;
            }
            if (startRow < 0) startRow = 0;
            const sourceSheetName = refs[0].sheet;
            out.push({
                chartFile: chartPath,
                drawingFile: drawingPath,
                anchorIndex,
                typeFromXml,
                sourceSheetName,
                sourceRange: { startRow, endRow, startColumn, endColumn },
                anchorFrom,
                anchorTo,
            });
        }
    }
    return out;
}

describe('M17 feature-1: chart import does not crash', () => {
    test('01-bar-simple.xlsx → 1 bar chart, anchored to source XML', async () => {
        const fp = path.join(FIXTURES_DIR, '01-bar-simple.xlsx');
        const buf = readFileSync(fp);
        const truth = await readSourceTruth(fp);
        expect(truth).toHaveLength(1);
        expect(truth[0].typeFromXml).toBe('bar');

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].drawing.data.type).toBe('bar');
        expect(drawings[0].drawing.data.sourceRange).toEqual(truth[0].sourceRange);
        expect(drawings[0].drawing.data.chartId).toMatch(/.+/);
        expect(drawings[0].drawing.sheetTransform.from.column).toBe(truth[0].anchorFrom.col);
        expect(drawings[0].drawing.sheetTransform.from.row).toBe(truth[0].anchorFrom.row);
        expect(drawings[0].drawing.sheetTransform.to.column).toBe(truth[0].anchorTo.col);
        expect(drawings[0].drawing.sheetTransform.to.row).toBe(truth[0].anchorTo.row);
    });

    test('02-line-multi-series.xlsx → line chart with N series matching source', async () => {
        const fp = path.join(FIXTURES_DIR, '02-line-multi-series.xlsx');
        const buf = readFileSync(fp);
        const truth = await readSourceTruth(fp);
        expect(truth[0].typeFromXml).toBe('line');

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].drawing.data.type).toBe('line');
        expect(drawings[0].drawing.data.sourceRange).toEqual(truth[0].sourceRange);
        // Source has 3 <c:ser> elements — verify by re-parsing the source.
        const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
        const chartXml = await zip.file(truth[0].chartFile)!.async('string');
        const seriesMatches = chartXml.match(/<c:ser>/g) ?? [];
        expect(drawings[0].drawing.data.datasets.length).toBe(seriesMatches.length);
    });

    test('03-pie-single.xlsx → pie chart', async () => {
        const fp = path.join(FIXTURES_DIR, '03-pie-single.xlsx');
        const buf = readFileSync(fp);
        const truth = await readSourceTruth(fp);
        expect(truth[0].typeFromXml).toBe('pie');

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].drawing.data.type).toBe('pie');
        expect(drawings[0].drawing.data.sourceRange).toEqual(truth[0].sourceRange);
    });

    test('04-doughnut.xlsx → doughnut chart', async () => {
        const fp = path.join(FIXTURES_DIR, '04-doughnut.xlsx');
        const buf = readFileSync(fp);
        const truth = await readSourceTruth(fp);
        expect(truth[0].typeFromXml).toBe('doughnut');

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        expect(drawings[0].drawing.data.type).toBe('doughnut');
        expect(drawings[0].drawing.data.sourceRange).toEqual(truth[0].sourceRange);
    });

    test('06-two-charts-one-sheet.xlsx → 2 charts in DOCUMENT ORDER (multi-anchor walk)', async () => {
        const fp = path.join(FIXTURES_DIR, '06-two-charts-one-sheet.xlsx');
        const buf = readFileSync(fp);
        const truth = await readSourceTruth(fp);
        expect(truth).toHaveLength(2);
        // The fixture packs both anchors into ONE drawing1.xml. Walking
        // by zip-key order would still pair (anchor[0] -> chart1,
        // anchor[1] -> chart2), but the load-bearing case is
        // multi-anchor-in-one-drawing. Verify the snapshot has exactly
        // 2 drawings on the same subUnit.
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(2);
        expect(drawings[0].subUnitId).toBe(drawings[1].subUnitId);
        // Both drawings must have distinct chartIds.
        expect(drawings[0].drawing.data.chartId).not.toBe(drawings[1].drawing.data.chartId);
        // Snapshot anchors line up with source XML anchors in DOCUMENT order.
        expect(drawings[0].drawing.sheetTransform.from.column).toBe(truth[0].anchorFrom.col);
        expect(drawings[1].drawing.sheetTransform.from.column).toBe(truth[1].anchorFrom.col);
    });

    test('07-chart-cross-sheet.xlsx → chart on Sheet2 references Sheet1 data', async () => {
        const fp = path.join(FIXTURES_DIR, '07-chart-cross-sheet.xlsx');
        const buf = readFileSync(fp);
        const truth = await readSourceTruth(fp);
        expect(truth[0].sourceSheetName).toBe('Sheet1');

        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const drawings = collectChartDrawings(snap);
        expect(drawings).toHaveLength(1);
        // The chart drawing lives on Sheet2 (subUnitId = sheet-2).
        expect(drawings[0].subUnitId).toBe('sheet-2');
        // The data ref still resolves to Sheet1 — sourceSheetName is plumbed.
        expect(drawings[0].drawing.data.sourceSheetName).toBe('Sheet1');
    });

    test('all chart fixtures import without throwing', async () => {
        // Filter out Excel temp-lock files (~$*.xlsx) that appear when a
        // user has the fixture open in Excel locally — they're not
        // valid xlsx archives.
        const fixtures = readdirSync(FIXTURES_DIR)
            .filter((f) => f.endsWith('.xlsx') && !f.startsWith('~$'))
            .sort();
        // Sanity floor: at least the 10 hand-crafted operator fixtures.
        // Additional regression fixtures (StackedBarChart, etc.) push
        // this higher and are also expected to import cleanly.
        expect(fixtures.length).toBeGreaterThanOrEqual(10);
        for (const fixture of fixtures) {
            const buf = readFileSync(path.join(FIXTURES_DIR, fixture));
            const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
            const sheetOrder = (snap as { sheetOrder: string[] }).sheetOrder;
            expect(Array.isArray(sheetOrder)).toBe(true);
            expect(sheetOrder.length).toBeGreaterThanOrEqual(1);
            const drawings = collectChartDrawings(snap);
            const expectedCharts = fixture === '06-two-charts-one-sheet.xlsx' ? 2 : 1;
            expect(drawings.length).toBeGreaterThanOrEqual(expectedCharts);
        }
    });

    test('synthetic <xdr:oneCellAnchor> drawing imports with synthesized to-anchor', async () => {
        const buf = await buildSyntheticChartXlsx({
            chartXml: STANDARD_BAR_CHART_BODY,
            anchorKind: 'oneCellAnchor',
        });
        const charts = await readChartsFromXlsxZip(buf as unknown as Buffer);
        expect(charts).toHaveLength(1);
        expect(charts[0].type).toBe('bar');
        // oneCellAnchor synthesizes to = from + (6 cols, 14 rows). Pin one
        // shape; document-anchor approximation per ## Notes.
        expect(charts[0].anchor.fromCol).toBe(0);
        expect(charts[0].anchor.fromRow).toBe(0);
        expect(charts[0].anchor.toCol).toBe(6);
        expect(charts[0].anchor.toRow).toBe(14);
    });

    test('synthetic missing-chart-target drawing → drops chart with warn (no throw)', async () => {
        const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
        try {
            const buf = await buildSyntheticChartXlsx({
                chartXml: STANDARD_BAR_CHART_BODY,
                omitChartRel: false,
                anchorKind: 'twoCellAnchor',
            });
            // Manually corrupt: re-zip the buffer with a dangling chart rel.
            const zip = await JSZip.loadAsync(buf as unknown as ArrayBuffer);
            zip.remove('xl/charts/chart1.xml');
            const corrupted = await zip.generateAsync({ type: 'nodebuffer' });
            const charts = await readChartsFromXlsxZip(corrupted as unknown as Buffer);
            expect(charts).toHaveLength(0);
            expect(warnSpy).toHaveBeenCalled();
        } finally {
            warnSpy.mockRestore();
        }
    });

    test('synthetic chart with no <c:cat> → recoverable (line chart with empty labels)', async () => {
        const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
        try {
            const buf = await buildSyntheticChartXlsx({ chartXml: NO_CAT_CHART_BODY });
            const charts = await readChartsFromXlsxZip(buf as unknown as Buffer);
            // The chart still has <c:val> with a usable formula ref, so it
            // imports as a 'line' chart with empty labels[]. The shape is
            // recoverable; we don't drop it.
            expect(charts).toHaveLength(1);
            expect(charts[0].type).toBe('line');
            expect(charts[0].labels).toEqual([]);
            expect(charts[0].datasets[0].data).toEqual([1, 2, 3]);
        } finally {
            warnSpy.mockRestore();
        }
    });

    test('synthetic radar-chart fixture → falls back to bar with meta.unsupportedSourceType', async () => {
        const warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => undefined);
        try {
            const buf = await buildSyntheticChartXlsx({ chartXml: RADAR_CHART_BODY });
            const charts = await readChartsFromXlsxZip(buf as unknown as Buffer);
            expect(charts).toHaveLength(1);
            expect(charts[0].type).toBe('bar');
            expect(charts[0].meta?.unsupportedSourceType).toBe('radar');
            expect(warnSpy).toHaveBeenCalled();
        } finally {
            warnSpy.mockRestore();
        }
    });

    test('xlsx-charts-unsupported error class still defined for non-strip-covered crash classes', async () => {
        // Cover the safety net: if readChartsFromXlsxZip returns []
        // (e.g. because the strip path can't parse a future shape),
        // exceljs's load encounters the chart-bearing buffer and throws
        // the legacy `anchors` reference error. The wrap still
        // produces NotesheetImportError — the error class is intact.
        expect(NotesheetImportError).toBeDefined();
        const dummy = new NotesheetImportError('xlsx-charts-unsupported', 'test', new Error('original'));
        expect(dummy.code).toBe('xlsx-charts-unsupported');
    });
});
