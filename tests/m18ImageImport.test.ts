// M18 A1 (Phase 4a): image drawings import from .xlsx into the snapshot's
// SHEET_DRAWING_PLUGIN resource as NATIVE Univer image drawings.
//
// Anchored to EXTERNAL ground truth — the fixture's real anchor values were
// read directly out of exceljs's getImages() API (see PROGRESS.md / the task
// brief). We assert the snapshot carries a native image drawing
// (drawingType === 0, imageSourceType === 'BASE64', a data: URI source) at
// the exact from-cell the source file declares.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';

import { xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'images');

interface ImageDrawing {
    drawingType: number;
    imageSourceType?: string;
    source?: string;
    sheetTransform?: {
        from?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number };
        to?: { column?: number; columnOffset?: number; row?: number; rowOffset?: number };
    };
}

function collectImageDrawings(
    snap: unknown,
): Array<{ subUnitId: string; drawingId: string; drawing: ImageDrawing }> {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data);
    const out: Array<{ subUnitId: string; drawingId: string; drawing: ImageDrawing }> = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        const order: string[] = sub.order ?? Object.keys(sub.data);
        for (const id of order) {
            const d = sub.data[id];
            // Native image drawings are drawingType 0 (DRAWING_IMAGE) and
            // carry NO componentKey (charts use 8 + 'NotesheetChart').
            if (d?.drawingType === 0 && d?.componentKey === undefined) {
                out.push({ subUnitId, drawingId: id, drawing: d as ImageDrawing });
            }
        }
    }
    return out;
}

describe('M18 A1: image import → native Univer image drawing', () => {
    test('HumanImage-SingleSheet.xlsx → one base64 png image drawing at the source anchor', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, 'HumanImage-SingleSheet.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const images = collectImageDrawings(snap);

        expect(images).toHaveLength(1);
        const { drawing } = images[0];

        // Native Univer image drawing shape.
        expect(drawing.drawingType).toBe(0);
        expect(drawing.imageSourceType).toBe('BASE64');
        expect(typeof drawing.source).toBe('string');
        expect(drawing.source!.startsWith('data:image/png;base64,')).toBe(true);
        // The base64 payload should be non-trivial (the source png is ~628 KB).
        expect(drawing.source!.length).toBeGreaterThan(100000);

        // External ground truth: source tl = {col:1, row:3}.
        expect(drawing.sheetTransform?.from?.column).toBe(1);
        expect(drawing.sheetTransform?.from?.row).toBe(3);
        // br = {col:8, row:30}.
        expect(drawing.sheetTransform?.to?.column).toBe(8);
        expect(drawing.sheetTransform?.to?.row).toBe(30);
    });

    test('Multi-sheet-StlyedImages.xlsx → jpeg/jpg images with ext-synthesized to-cells', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, 'Multi-sheet-StlyedImages.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const images = collectImageDrawings(snap);

        // Two images: one jpeg, one jpg, on two different sheets. Both anchors
        // ship `ext` (no br), so the importer must synthesize a to-cell.
        expect(images).toHaveLength(2);
        for (const { drawing } of images) {
            expect(drawing.drawingType).toBe(0);
            expect(drawing.imageSourceType).toBe('BASE64');
            expect(drawing.source!.startsWith('data:image/jpeg;base64,')).toBe(true);
            // to-cell must be strictly past the from-cell (synthesized span).
            expect(drawing.sheetTransform!.to!.column!).toBeGreaterThan(
                drawing.sheetTransform!.from!.column!,
            );
            expect(drawing.sheetTransform!.to!.row!).toBeGreaterThan(
                drawing.sheetTransform!.from!.row!,
            );
        }
    });
});
