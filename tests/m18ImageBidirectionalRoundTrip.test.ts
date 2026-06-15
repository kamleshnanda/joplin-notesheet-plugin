// M18 A1 (Phase 4b): image drawing import -> export -> re-import round-trip.
//
// Pins that import + export are inverse operations on the image subset
// (parallel to the M17 chart round-trip). Anchored to the FIRST snapshot,
// NOT a hardcoded literal: snapshot-vs-snapshot inverseness. Sub-cell offset
// precision is a separate backlog item (A3), so we tolerate offset drift but
// assert the cell col/row indices are exact and the base64 image bytes are
// byte-equal.
//
// Also JSZip-asserts the EXPORTED buffer carries xl/media/image*.png + a
// <xdr:pic> in a drawing part (the on-disk OOXML proof, not just re-import).

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'images');

interface ImageDrawing {
    drawingType: number;
    imageSourceType?: string;
    source?: string;
    sheetTransform?: {
        from?: { column?: number; row?: number };
        to?: { column?: number; row?: number };
    };
}

function collectImageDrawings(snap: unknown): Array<{ subUnitId: string; drawing: ImageDrawing }> {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data);
    const out: Array<{ subUnitId: string; drawing: ImageDrawing }> = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        const order: string[] = sub.order ?? Object.keys(sub.data);
        for (const id of order) {
            const d = sub.data[id];
            if (d?.drawingType === 0 && d?.componentKey === undefined) {
                out.push({ subUnitId, drawing: d as ImageDrawing });
            }
        }
    }
    return out;
}

describe('M18 A1: image bidirectional round-trip', () => {
    test('HumanImage-SingleSheet.xlsx: image survives import -> export -> re-import', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, 'HumanImage-SingleSheet.xlsx'));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap1);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);

        const first = collectImageDrawings(snap1);
        const second = collectImageDrawings(snap2);

        // Count preserved.
        expect(first).toHaveLength(1);
        expect(second).toHaveLength(1);

        // Source base64 byte-equality (the image payload survives the trip).
        expect(second[0].drawing.source).toBe(first[0].drawing.source);

        // Anchor cell col/row exactness (offsets may drift — A3 scope).
        expect(second[0].drawing.sheetTransform?.from?.column).toBe(
            first[0].drawing.sheetTransform?.from?.column,
        );
        expect(second[0].drawing.sheetTransform?.from?.row).toBe(
            first[0].drawing.sheetTransform?.from?.row,
        );
        expect(second[0].drawing.sheetTransform?.to?.column).toBe(
            first[0].drawing.sheetTransform?.to?.column,
        );
        expect(second[0].drawing.sheetTransform?.to?.row).toBe(
            first[0].drawing.sheetTransform?.to?.row,
        );
    });

    test('exported buffer carries xl/media/image*.png + a <xdr:pic> in a drawing part', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, 'HumanImage-SingleSheet.xlsx'));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap1);

        const zip = await JSZip.loadAsync(exported);
        const mediaKeys = Object.keys(zip.files).filter((p) =>
            /^xl\/media\/image\d+\.png$/.test(p),
        );
        expect(mediaKeys.length).toBeGreaterThanOrEqual(1);

        // The media part should be a real, sizeable png.
        const mediaBytes = await zip.files[mediaKeys[0]].async('uint8array');
        expect(mediaBytes.length).toBeGreaterThan(100000);
        // PNG magic number.
        expect(Array.from(mediaBytes.slice(0, 4))).toEqual([0x89, 0x50, 0x4e, 0x47]);

        // Some drawing part contains an <xdr:pic> with a blip embed ref.
        const drawingKeys = Object.keys(zip.files).filter((p) =>
            /^xl\/drawings\/drawing\d+\.xml$/.test(p),
        );
        expect(drawingKeys.length).toBeGreaterThanOrEqual(1);
        let foundPic = false;
        for (const dk of drawingKeys) {
            const xml = await zip.files[dk].async('string');
            if (/<xdr:pic>/.test(xml) && /<a:blip[^>]*r:embed="rId\d+"/.test(xml)) {
                foundPic = true;
            }
        }
        expect(foundPic).toBe(true);

        // [Content_Types].xml carries a Default-by-extension for png.
        const ct = await zip.files['[Content_Types].xml'].async('string');
        expect(/<Default\s+Extension="png"\s+ContentType="image\/png"/.test(ct)).toBe(true);
    });

    test('Multi-sheet-StlyedImages.xlsx: two jpeg images survive round-trip', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, 'Multi-sheet-StlyedImages.xlsx'));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap1);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);

        const first = collectImageDrawings(snap1);
        const second = collectImageDrawings(snap2);
        expect(first).toHaveLength(2);
        expect(second).toHaveLength(2);

        // Byte-equality on each image source (sort by source to pair them).
        const firstSources = first.map((f) => f.drawing.source).sort();
        const secondSources = second.map((f) => f.drawing.source).sort();
        expect(secondSources).toEqual(firstSources);

        // Each image must stay bound to the SAME NAMED sheet across the
        // round-trip. The subUnitId NUMBER legitimately changes (export
        // renumbers sheetIds to contiguous 1..N), so asserting subUnitId
        // equality would be a false alarm — assert by sheet *name*, which is
        // the real fidelity contract. (See M18 A1 stabilization review.)
        const nameOf = (snap: unknown, subUnitId: string): string | undefined => {
            const sheets = (snap as { sheets?: Record<string, { name?: string }> }).sheets ?? {};
            return sheets[subUnitId]?.name;
        };
        const byName = (
            snap: unknown,
            items: Array<{ subUnitId: string; drawing: ImageDrawing }>,
        ) => items.map((i) => `${nameOf(snap, i.subUnitId)}::${i.drawing.source}`).sort();
        expect(byName(snap2, second)).toEqual(byName(snap1, first));

        // Exported zip carries jpeg media + a jpeg Default content-type.
        const zip = await JSZip.loadAsync(exported);
        const jpegKeys = Object.keys(zip.files).filter((p) =>
            /^xl\/media\/image\d+\.jpeg$/.test(p),
        );
        expect(jpegKeys.length).toBe(2);
        const ct = await zip.files['[Content_Types].xml'].async('string');
        expect(/<Default\s+Extension="jpeg"\s+ContentType="image\/jpeg"/.test(ct)).toBe(true);
    });
});
