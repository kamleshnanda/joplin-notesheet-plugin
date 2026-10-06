// M18: rotated-image round-trip. An Excel <a:xfrm rot="..."> on a picture was
// previously dropped on import (the snapshot hardcoded transform.angle = 0) and
// never re-emitted on export, so a rotated image imported and exported back
// axis-aligned. This pins the fix across all three sites: import parses rot,
// the snapshot carries transform.angle (DEGREES), and export re-emits
// <a:xfrm rot> in 60000ths of a degree.
//
// Anchored to a real Excel-authored fixture (RotatedImage-MultiSheet.xlsx):
// sheet 1 holds an axis-aligned image, sheet 2 an image rotated rot=20172693
// (= 336.21155°). Ground-truth rot value read straight out of the source
// xl/drawings/drawing2.xml.

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
const FIXTURE = 'RotatedImage-MultiSheet.xlsx';

// The source rot on sheet 2's picture, in 60000ths of a degree, and its
// degree equivalent. Read directly from the fixture's xl/drawings/drawing2.xml.
const SRC_ROT_EMU_DEG = 20172693;
const SRC_ANGLE_DEG = SRC_ROT_EMU_DEG / 60000; // 336.21155

interface ImageDrawing {
    drawingType: number;
    componentKey?: string;
    transform?: { angle?: number };
    sheetTransform?: { angle?: number };
    axisAlignSheetTransform?: { angle?: number };
}

function collectAngles(snap: unknown): number[] {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return [];
    const parsed = JSON.parse(entry.data);
    const out: number[] = [];
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        const order: string[] = sub.order ?? Object.keys(sub.data);
        for (const id of order) {
            const d = sub.data[id] as ImageDrawing;
            if (d?.drawingType === 0 && d?.componentKey === undefined) {
                out.push(d.transform?.angle ?? 0);
            }
        }
    }
    return out.sort((a, b) => a - b);
}

// The rotated image drawing (angle > 0), for asserting the angle lands on the
// fields Univer actually reads at load time (sheetTransform), not just the
// live `transform` — Univer recomputes `transform` from sheetTransform on load,
// so an angle only on `transform` renders axis-aligned (the shipped bug).
function findRotatedDrawing(snap: unknown): ImageDrawing | null {
    const resources =
        (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
    const entry = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
    if (!entry) return null;
    const parsed = JSON.parse(entry.data);
    for (const subUnitId of Object.keys(parsed)) {
        const sub = parsed[subUnitId];
        for (const id of Object.keys(sub.data)) {
            const d = sub.data[id] as ImageDrawing;
            if (
                d?.drawingType === 0 &&
                d?.componentKey === undefined &&
                (d.transform?.angle ?? 0)
            ) {
                return d;
            }
        }
    }
    return null;
}

// Pull every rot value out of every drawing part in an exported buffer.
async function collectExportedRots(buffer: ArrayBuffer): Promise<number[]> {
    const zip = await JSZip.loadAsync(buffer);
    const rots: number[] = [];
    for (const p of Object.keys(zip.files)) {
        if (!/^xl\/drawings\/drawing\d+\.xml$/.test(p)) continue;
        const xml = await zip.files[p].async('string');
        for (const m of xml.matchAll(/<a:xfrm\b[^>]*\brot="(-?\d+)"/g)) {
            rots.push(parseInt(m[1], 10));
        }
    }
    return rots.sort((a, b) => a - b);
}

describe('M18: rotated-image round-trip', () => {
    test('import captures the source rotation as transform.angle (degrees)', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);

        const angles = collectAngles(snap);
        // Two images: one axis-aligned (0°), one rotated (~336.21°).
        expect(angles).toHaveLength(2);
        expect(angles[0]).toBe(0);
        expect(angles[1]).toBeCloseTo(SRC_ANGLE_DEG, 3);
    });

    test('rotation lands on sheetTransform (the field Univer reads on load)', async () => {
        // Regression guard for the shipped bug: Univer recomputes the live
        // `transform` from `sheetTransform` on load, so the angle MUST be on
        // sheetTransform + axisAlignSheetTransform, not only on `transform`.
        // A snapshot with the angle only on `transform` renders axis-aligned.
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);

        const rotated = findRotatedDrawing(snap);
        expect(rotated).not.toBeNull();
        expect(rotated!.transform?.angle).toBeCloseTo(SRC_ANGLE_DEG, 3);
        expect(rotated!.sheetTransform?.angle).toBeCloseTo(SRC_ANGLE_DEG, 3);
        expect(rotated!.axisAlignSheetTransform?.angle).toBeCloseTo(SRC_ANGLE_DEG, 3);
    });

    test('export re-emits <a:xfrm rot> for the rotated image only', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap);

        const rots = await collectExportedRots(exported);
        // Exactly one rot emitted (the axis-aligned image emits no <a:xfrm>).
        expect(rots).toHaveLength(1);
        // Round-trips to the same 60000ths-of-a-degree value (±1 unit of
        // rounding — the degree value is fractional).
        expect(Math.abs(rots[0] - SRC_ROT_EMU_DEG)).toBeLessThanOrEqual(1);
    });

    test('rotation survives a full import → export → re-import', async () => {
        const buf = readFileSync(path.join(FIXTURES_DIR, FIXTURE));
        const snap1 = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap1);
        const snap2 = await xlsxBufferToSnapshot(Buffer.from(exported) as unknown as Buffer);

        const a1 = collectAngles(snap1);
        const a2 = collectAngles(snap2);
        expect(a2).toHaveLength(a1.length);
        // Angle preserved within rounding on both images.
        for (let i = 0; i < a1.length; i++) {
            expect(a2[i]).toBeCloseTo(a1[i], 2);
        }
    });
});
