// M18 A3: preserve the EMU sub-cell anchor offset losslessly across the
// .xlsx round-trip for IMAGE and CHART drawings.
//
// Before A3, both paths stored the anchor offset as ROUNDED PIXELS in the
// snapshot (Math.round(emu / 9525)), so a round-trip drifted by up to ~half a
// pixel (3175 EMU). Worse, the CHART export read that pixel value as if it
// were EMU (a latent unit bug) — so any chart with a sub-cell offset exported
// at ~1/9525 of its intended offset (effectively 0).
//
// Fix: import stashes the ORIGINAL EMU anchor on the drawing under a private
// `_srcAnchorEmu` key (Univer ignores it). Export prefers `_srcAnchorEmu`
// (exact) and falls back to px×9525 only for editor-authored/moved drawings
// that never had a source anchor.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';
import { xlsxBufferToSnapshot, snapshotToXlsxBuffer } from '../src/xlsx';

const IMAGES = path.join(__dirname, 'fixtures', 'images');
const CHARTS = path.join(__dirname, 'fixtures', 'charts');

// Pull every <xdr:from>/<xdr:to> colOff/rowOff from a drawing part.
async function anchorOffsets(buf: ArrayBuffer): Promise<number[]> {
    const zip = await JSZip.loadAsync(buf);
    const dpath = Object.keys(zip.files).find((p) => /xl\/drawings\/drawing\d+\.xml$/.test(p));
    if (!dpath) return [];
    const xml = await zip.files[dpath].async('string');
    return [...xml.matchAll(/<xdr:(?:colOff|rowOff)>(-?\d+)<\/xdr:(?:colOff|rowOff)>/g)].map((m) =>
        parseInt(m[1], 10),
    );
}

async function sourceOffsets(file: string): Promise<number[]> {
    const zip = await JSZip.loadAsync(readFileSync(file));
    const dpath = Object.keys(zip.files).find((p) => /xl\/drawings\/drawing\d+\.xml$/.test(p));
    const xml = await zip.files[dpath!].async('string');
    return [...xml.matchAll(/<xdr:(?:colOff|rowOff)>(-?\d+)<\/xdr:(?:colOff|rowOff)>/g)].map((m) =>
        parseInt(m[1], 10),
    );
}

describe('M18 A3 — EMU sub-cell anchor offset fidelity', () => {
    test('image anchor offsets survive round-trip with ZERO EMU drift', async () => {
        const file = path.join(IMAGES, 'HumanImage-SingleSheet.xlsx');
        const original = await sourceOffsets(file);
        const snap = await xlsxBufferToSnapshot(readFileSync(file) as unknown as Buffer);
        const exported = await snapshotToXlsxBuffer(snap);
        const roundTripped = await anchorOffsets(exported);
        // Exact equality — not "within a pixel". The source EMU is preserved.
        expect(roundTripped).toEqual(original);
    });

    test('the snapshot image drawing carries the original EMU anchor (_srcAnchorEmu)', async () => {
        const file = path.join(IMAGES, 'HumanImage-SingleSheet.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(file) as unknown as Buffer);
        const resources =
            (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const res = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
        const parsed = JSON.parse((res as { data: string }).data);
        const sub = parsed['sheet-1'];
        const drawing = sub.data[Object.keys(sub.data)[0]];
        expect(drawing._srcAnchorEmu).toBeDefined();
        // Carries the four from/to offsets in EMU.
        expect(typeof drawing._srcAnchorEmu.fromColOff).toBe('number');
    });

    test('chart anchor sub-cell offset exports in EMU, not mis-scaled to ~0', async () => {
        // Synthesize a chart snapshot whose anchor has a real sub-cell offset,
        // then export and assert the OOXML colOff is the EMU value, not the
        // pixel value (the latent unit bug exported px-as-EMU → ~0).
        const file = path.join(CHARTS, '01-bar-simple.xlsx');
        const snap: {
            resources?: Array<{ name: string; data: string }>;
        } = await xlsxBufferToSnapshot(readFileSync(file) as unknown as Buffer);
        const res = (snap.resources ?? []).find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!;
        const parsed = JSON.parse(res.data);
        const sub = parsed[Object.keys(parsed)[0]];
        const id = Object.keys(sub.data)[0];
        const drawing = sub.data[id];
        // Give it a known px offset AND the matching source EMU (as the import
        // would now stash). 50 px ≈ 476250 EMU.
        drawing.sheetTransform.from.columnOffset = 50;
        drawing.axisAlignSheetTransform.from.columnOffset = 50;
        drawing._srcAnchorEmu = {
            fromColOff: 476250,
            fromRowOff: drawing._srcAnchorEmu?.fromRowOff ?? 0,
            toColOff: drawing._srcAnchorEmu?.toColOff ?? 0,
            toRowOff: drawing._srcAnchorEmu?.toRowOff ?? 0,
        };
        res.data = JSON.stringify(parsed);

        const exported = await snapshotToXlsxBuffer(snap as never);
        const offs = await anchorOffsets(exported);
        // The from.colOff must be the EMU value (476250), not 50 (px-as-EMU).
        expect(offs).toContain(476250);
        expect(offs).not.toContain(50);
    });

    test('editor-MOVED drawing ignores the stale _srcAnchorEmu (px×9525 fallback)', async () => {
        // Simulate a user dragging an imported image in the editor: the px
        // sheetTransform changes but _srcAnchorEmu still holds the OLD source
        // EMU. Export must use the NEW px position (×9525), not the stale EMU.
        const file = path.join(IMAGES, 'HumanImage-SingleSheet.xlsx');
        const snap = await xlsxBufferToSnapshot(readFileSync(file) as unknown as Buffer);
        const resources =
            (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const res = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN')!;
        const parsed = JSON.parse(res.data);
        const sub = parsed['sheet-1'];
        const drawing = sub.data[Object.keys(sub.data)[0]];
        // Move it: change from-cell column + offset in BOTH transforms; leave
        // the stale _srcAnchorEmu untouched (as the editor would).
        drawing.sheetTransform.from.column = 0;
        drawing.sheetTransform.from.columnOffset = 10;
        drawing.axisAlignSheetTransform.from.column = 0;
        drawing.axisAlignSheetTransform.from.columnOffset = 10;
        res.data = JSON.stringify(parsed);

        const exported = await snapshotToXlsxBuffer(snap);
        const zip = await JSZip.loadAsync(exported);
        const dpath = Object.keys(zip.files).find((p) => /xl\/drawings\/drawing\d+\.xml$/.test(p))!;
        const xml = await zip.files[dpath].async('string');
        const fromColOff = parseInt(xml.match(/<xdr:from>[\s\S]*?<xdr:colOff>(-?\d+)/)![1], 10);
        // Moved → 10 px × 9525 = 95250 EMU. NOT the stale source (787400).
        expect(fromColOff).toBe(95250);
    });
});
