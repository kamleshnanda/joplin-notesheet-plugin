// M18 A1 (Phase 4c): chart + image coexistence on ONE sheet (HIGHEST RISK).
//
// Excel honors exactly ONE <drawing r:id> per worksheet. If image export
// created a SECOND drawing part for a sheet that already has a chart drawing,
// one of them would silently vanish. injectImagesIntoZip MUST merge the
// <xdr:pic> into the sheet's existing (chart) drawing part instead.
//
// Strategy: import a real chart fixture, programmatically add a native image
// drawing to the SAME sheet's SHEET_DRAWING_PLUGIN entry, export, then assert
// the exported zip has BOTH a chart-referencing drawing AND an <xdr:pic> in
// the SAME drawing part, with non-colliding drawing/rId/media numbers — and a
// re-import yields exactly one chart + one image.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';
import type { UniverSnapshot } from '../src/snapshot';

const CHART_FIXTURES_DIR = path.join(__dirname, 'fixtures', 'charts');
const IMAGE_FIXTURES_DIR = path.join(__dirname, 'fixtures', 'images');

// A tiny valid 1x1 png (transparent), base64. Keeps the test self-contained
// without depending on a fixture's exact byte count.
const TINY_PNG_BASE64 =
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

function getDrawingResource(snap: UniverSnapshot): {
    subUnitId: string;
    parsed: Record<string, { data: Record<string, unknown>; order: string[] }>;
} {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) throw new Error('no SHEET_DRAWING_PLUGIN');
    const parsed = JSON.parse(entry.data) as Record<
        string,
        { data: Record<string, unknown>; order: string[] }
    >;
    const subUnitId = Object.keys(parsed)[0];
    return { subUnitId, parsed };
}

// Inject a native image drawing onto the SAME sheet as the chart.
function addImageDrawingToSnapshot(snap: UniverSnapshot, subUnitId: string, source: string): void {
    const resources = (snap as { resources: Array<{ name: string; data: string }> }).resources;
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!;
    const parsed = JSON.parse(entry.data) as Record<
        string,
        { data: Record<string, unknown>; order: string[] }
    >;
    const drawingId = 'image-coexist-1';
    parsed[subUnitId].data[drawingId] = {
        unitId: 'workbook',
        subUnitId,
        drawingId,
        drawingType: 0,
        imageSourceType: 'BASE64',
        source,
        allowTransform: true,
        transform: {
            flipY: false,
            flipX: false,
            angle: 0,
            skewX: 0,
            skewY: 0,
            left: 400,
            top: 200,
            width: 120,
            height: 90,
        },
        sheetTransform: {
            from: { column: 10, columnOffset: 0, row: 5, rowOffset: 0 },
            to: { column: 13, columnOffset: 0, row: 12, rowOffset: 0 },
        },
        axisAlignSheetTransform: {
            from: { column: 10, columnOffset: 0, row: 5, rowOffset: 0 },
            to: { column: 13, columnOffset: 0, row: 12, rowOffset: 0 },
        },
    };
    parsed[subUnitId].order.push(drawingId);
    entry.data = JSON.stringify(parsed);
}

function countDrawings(snap: UniverSnapshot): { charts: number; images: number } {
    const { parsed } = getDrawingResource(snap);
    let charts = 0;
    let images = 0;
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        for (const id of Object.keys(sub.data)) {
            const d = sub.data[id] as { componentKey?: string; drawingType?: number };
            if (d.componentKey === 'NotesheetChart') charts += 1;
            else if (d.drawingType === 0) images += 1;
        }
    }
    return { charts, images };
}

describe('M18 A1: chart + image coexist on one sheet', () => {
    test('exported zip merges <xdr:pic> into the chart drawing (no second drawing part)', async () => {
        const buf = readFileSync(path.join(CHART_FIXTURES_DIR, '01-bar-simple.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);

        // Chart-only baseline.
        expect(countDrawings(snap)).toEqual({ charts: 1, images: 0 });

        const { subUnitId } = getDrawingResource(snap);
        const source = `data:image/png;base64,${TINY_PNG_BASE64}`;
        addImageDrawingToSnapshot(snap, subUnitId, source);
        expect(countDrawings(snap)).toEqual({ charts: 1, images: 1 });

        const exported = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(exported);

        // The sheet must reference exactly ONE drawing part.
        const sheetRelKeys = Object.keys(zip.files).filter((p) =>
            /^xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(p),
        );
        let drawingRefCount = 0;
        const drawingTargets: string[] = [];
        for (const k of sheetRelKeys) {
            const xml = await zip.files[k].async('string');
            const matches = xml.match(/Type="[^"]*\/drawing"[^>]*Target="([^"]+)"/g) ?? [];
            drawingRefCount += matches.length;
            for (const mm of matches) {
                const t = mm.match(/Target="([^"]+)"/);
                if (t) drawingTargets.push(t[1]);
            }
        }
        expect(drawingRefCount).toBe(1);

        // There must be exactly ONE drawing part on disk, and it must hold
        // BOTH the chart graphicFrame AND the <xdr:pic>.
        const drawingKeys = Object.keys(zip.files).filter((p) =>
            /^xl\/drawings\/drawing\d+\.xml$/.test(p),
        );
        expect(drawingKeys).toHaveLength(1);
        const drawingXml = await zip.files[drawingKeys[0]].async('string');
        // Chart graphicFrame ref.
        expect(/<a:graphicData[^>]*chart/.test(drawingXml) || /<c:chart\b/.test(drawingXml)).toBe(
            true,
        );
        // Image pic.
        expect(/<xdr:pic>/.test(drawingXml)).toBe(true);

        // The drawing rels must reference BOTH a chart target and an image
        // target with DISTINCT rIds.
        const drawingNum = drawingKeys[0].match(/drawing(\d+)\.xml$/)![1];
        const drawingRelsXml =
            await zip.files[`xl/drawings/_rels/drawing${drawingNum}.xml.rels`].async('string');
        const chartRel = drawingRelsXml.match(
            /<Relationship\s+Id="(rId\d+)"[^>]*Type="[^"]*\/chart"[^>]*Target="([^"]+)"/,
        );
        const imageRel = drawingRelsXml.match(
            /<Relationship\s+Id="(rId\d+)"[^>]*Type="[^"]*\/image"[^>]*Target="([^"]+)"/,
        );
        expect(chartRel).not.toBeNull();
        expect(imageRel).not.toBeNull();
        expect(chartRel![1]).not.toBe(imageRel![1]); // distinct rIds

        // The pic's blip embed rId matches the image rel's rId.
        const blipEmbed = drawingXml.match(/<a:blip[^>]*r:embed="(rId\d+)"/);
        expect(blipEmbed).not.toBeNull();
        expect(blipEmbed![1]).toBe(imageRel![1]);

        // Media part exists (non-colliding name) and is a png.
        const mediaKeys = Object.keys(zip.files).filter((p) =>
            /^xl\/media\/image\d+\.png$/.test(p),
        );
        expect(mediaKeys.length).toBeGreaterThanOrEqual(1);

        // No collision between drawing targets referenced and what exists.
        expect(drawingTargets).toHaveLength(1);
    });

    test('re-import of the coexist export yields one chart + one image', async () => {
        const buf = readFileSync(path.join(CHART_FIXTURES_DIR, '01-bar-simple.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const { subUnitId } = getDrawingResource(snap);
        // Use a real, larger png so its byte payload is unambiguous.
        const humanBuf = readFileSync(path.join(IMAGE_FIXTURES_DIR, 'HumanImage-SingleSheet.xlsx'));
        const humanSnap = await xlsxBufferToSnapshot(humanBuf as unknown as Buffer);
        const humanImg = getDrawingResource(humanSnap);
        const humanSource = Object.values(humanImg.parsed[humanImg.subUnitId].data)
            .map((d) => (d as { source?: string }).source)
            .find((s) => typeof s === 'string')!;

        addImageDrawingToSnapshot(snap, subUnitId, humanSource);
        const exported = await snapshotToXlsxBuffer(snap);
        const reimported = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);

        expect(countDrawings(reimported)).toEqual({ charts: 1, images: 1 });
    });
});
