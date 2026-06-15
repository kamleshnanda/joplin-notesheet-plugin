// M18 A1 follow-up — REPRO: the chart importer mis-maps charts on an edited
// workbook where the worksheet FILENAME number (xl/worksheets/sheetN.xml)
// differs from the logical sheetId (workbook.xml <sheet sheetId=...>).
//
// xlsx.ts keys snapshot subUnitIds off the workbook sheetId (exceljs ws.id).
// The IMAGE importer resolves filename->sheetId via workbook.xml(.rels); the
// CHART importer (xlsxChartImport.ts findSheetDrawingLinks) uses the raw
// filename number as sheetIndex, so on a drifted workbook the chart is
// assigned to the wrong `sheet-N` — and the `if (!sheets[subUnitId]) continue`
// guard in xlsx.ts then SILENTLY DROPS it.
//
// We synthesize the drift from the real 07-chart-cross-sheet fixture by
// rewriting workbook.xml so the chart-bearing worksheet (sheet2.xml) maps to
// a HIGHER sheetId than its filename number. A correct importer puts the chart
// on the sheet whose name matches; the buggy importer drops it.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURE = path.join(__dirname, 'fixtures', 'charts', '07-chart-cross-sheet.xlsx');

interface ChartDrawing {
    componentKey?: string;
}
function collectChartDrawings(snap: unknown): Array<{ subUnitId: string; drawingId: string }> {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data);
    const out: Array<{ subUnitId: string; drawingId: string }> = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        for (const id of Object.keys(sub.data ?? {})) {
            const d = sub.data[id] as ChartDrawing;
            if (d?.componentKey === 'NotesheetChart') out.push({ subUnitId, drawingId: id });
        }
    }
    return out;
}

function sheetName(snap: unknown, subUnitId: string): string | undefined {
    const sheets = (snap as { sheets?: Record<string, { name?: string }> }).sheets ?? {};
    return sheets[subUnitId]?.name;
}

// Rewrite workbook.xml so the two sheets carry NON-CONTIGUOUS sheetIds that
// don't match their worksheet filename numbers — simulating an edited
// workbook (e.g. a deleted leading sheet). The r:id->filename mapping in
// workbook.xml.rels is left untouched; only the sheetId attributes change.
async function buildDriftedFixture(): Promise<Buffer> {
    const zip = await JSZip.loadAsync(readFileSync(FIXTURE));
    const wbPath = 'xl/workbook.xml';
    let wbXml = await zip.file(wbPath)!.async('string');
    // Force sheetIds to 5 and 6 (filenames remain sheet1.xml / sheet2.xml).
    // Map by r:id so the chart-bearing sheet (rId pointing at sheet2.xml)
    // gets sheetId 6 while its filename number stays 2.
    let n = 5;
    wbXml = wbXml.replace(
        /(sheetId=")(\d+)(")/g,
        (_full: string, pre: string, _id: string, post: string) => `${pre}${n++}${post}`,
    );
    zip.file(wbPath, wbXml);
    return zip.generateAsync({ type: 'nodebuffer' });
}

describe('M18 — chart importer sheetId/filename drift (REPRO + fix)', () => {
    test('baseline: undrifted fixture imports its chart', async () => {
        const snap = await xlsxBufferToSnapshot(readFileSync(FIXTURE) as unknown as Buffer);
        const charts = collectChartDrawings(snap);
        expect(charts.length).toBeGreaterThanOrEqual(1);
    });

    test('drifted workbook: chart still imported AND on the correct named sheet', async () => {
        const drifted = await buildDriftedFixture();
        const snap = await xlsxBufferToSnapshot(drifted as unknown as Buffer);

        // 1. The chart must not be silently dropped.
        const charts = collectChartDrawings(snap);
        expect(charts.length).toBeGreaterThanOrEqual(1);

        // 2. It must land on a subUnitId that actually exists in the snapshot
        //    (the bug assigns it to a `sheet-N` that doesn't exist → dropped,
        //    or to the wrong existing sheet).
        const sheets = (snap as { sheets?: Record<string, unknown> }).sheets ?? {};
        for (const c of charts) {
            expect(sheets[c.subUnitId]).toBeDefined();
        }

        // 3. The chart in 07-chart-cross-sheet is anchored on the SECOND sheet
        //    ("Sheet2"). After drift it must still be on the sheet named
        //    "Sheet2", not relocated by the filename-number assumption.
        const hostNames = charts.map((c) => sheetName(snap, c.subUnitId));
        expect(hostNames).toContain('Sheet2');
    });
});
