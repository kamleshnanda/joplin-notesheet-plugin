// Mock @univerjs/sheets-table so importing src/xlsx (transitively pulled by
// the SUT) doesn't drag the whole Univer ESM graph through jest's CJS
// transform. Mirrors the pattern in m12FixturePinDowns.test.ts.
jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

// Pin-down regression tests for M12+ import recovery.
//
// Three project-owned fixtures — MultiSheet.xlsx, LargeWorkbook.xlsx,
// FormulasAndStructuredRefs.xlsx — once crashed inside exceljs's reconcile
// pipeline on a bare `await wb.xlsx.load(buf)`. All three now import cleanly
// after Notesheet pre-processes the buffer:
//   * MultiSheet — chart drawings stripped pre-load (M17).
//   * LargeWorkbook + FormulasAndStructuredRefs — their ABSOLUTE relationship
//     Targets (`Target="/xl/tables/table1.xml"`) are normalized to relative
//     pre-load (normalizeAbsoluteRelTargets). exceljs couldn't resolve an
//     absolute target, so its worksheet table-reduce crashed on
//     `tables[table.name]`. These were originally MIS-diagnosed as a
//     "multiple sheets with named tables" limitation and gated behind a
//     friendly xlsx-multi-table-unsupported error; the real cause was the
//     target format, and a same-shaped file with relative targets always
//     imported fine. Fixed 2026-06.
//
// For any crash class we DON'T pre-empt, xlsxBufferToSnapshot() still
// catches and rethrows a typed NotesheetImportError with a user-actionable
// message (consumed by the dialog in src/index.ts + the editor status bar).
//
// REGRESSION HISTORY: if a future exceljs/preprocessing change reintroduces
// a load crash for any of these, the "imports cleanly" assertions below flip
// to rejection — that's the signal the pre-processing path regressed.

import { readFileSync } from 'fs';
import path from 'path';

import { xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'fixtures', 'formatting-testdata');

async function loadAndCatch(file: string): Promise<unknown> {
    const buf = readFileSync(path.join(FIXTURES_DIR, file));
    try {
        await xlsxBufferToSnapshot(buf as unknown as Buffer);
        return null;
    } catch (e) {
        return e;
    }
}

// ─── Crashing fixtures: must reject with NotesheetImportError ─────────

describe('M12 import recovery — crashing fixtures throw typed errors', () => {
    test('MultiSheet.xlsx → snapshot with N charts (M17 chart import)', async () => {
        // M17 (feature-1): MultiSheet.xlsx now imports cleanly via the
        // pre-load chart strip. The fixture's chart parts are read
        // independently and re-emitted as a SHEET_DRAWING_PLUGIN
        // resource on the snapshot; exceljs loads the stripped buffer
        // without touching the offending drawings. The `xlsx-charts-
        // unsupported` error class stays defined for future drawing-
        // related crash classes outside the strip path's coverage.
        const err = await loadAndCatch('MultiSheet.xlsx');
        expect(err).toBeNull();
        const buf = readFileSync(path.join(FIXTURES_DIR, 'MultiSheet.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const resources =
            (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const drawing = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
        expect(drawing).toBeDefined();
        const parsed = JSON.parse(drawing!.data);
        let chartCount = 0;
        for (const su of Object.keys(parsed)) {
            chartCount += Object.keys(parsed[su].data).length;
        }
        expect(chartCount).toBeGreaterThanOrEqual(1);
    });

    // FORMERLY xlsx-multi-table-unsupported. Both LargeWorkbook.xlsx and
    // FormulasAndStructuredRefs.xlsx were misdiagnosed as a "multiple sheets
    // with named tables" limitation. The TRUE cause was ABSOLUTE relationship
    // Targets (`Target="/xl/tables/table1.xml"`), which exceljs can't resolve
    // — leaving the table part unlinked so its worksheet table-reduce crashed
    // on `tables[table.name]`. A workbook with the SAME multi-sheet/multi-table
    // layout but RELATIVE targets (spreadsheet1.xlsx) always imported fine.
    // normalizeAbsoluteRelTargets() rewrites absolute → relative before load,
    // so both now import cleanly. (If a future regression reintroduces the
    // crash, these flip back to rejection assertions.)
    test('LargeWorkbook.xlsx imports cleanly (absolute rel-targets normalized)', async () => {
        const err = await loadAndCatch('LargeWorkbook.xlsx');
        expect(err).toBeNull();
        const buf = readFileSync(path.join(FIXTURES_DIR, 'LargeWorkbook.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        expect(
            Object.keys((snap as { sheets?: Record<string, unknown> }).sheets ?? {}).length,
        ).toBe(2);
    });

    test('FormulasAndStructuredRefs.xlsx imports cleanly (absolute rel-targets normalized)', async () => {
        const err = await loadAndCatch('FormulasAndStructuredRefs.xlsx');
        expect(err).toBeNull();
        const buf = readFileSync(path.join(FIXTURES_DIR, 'FormulasAndStructuredRefs.xlsx'));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        // Both sheets survive AND both named tables round-trip.
        expect(
            Object.keys((snap as { sheets?: Record<string, unknown> }).sheets ?? {}).length,
        ).toBe(2);
        const resources =
            (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const tableRes = resources.find((r) => r.name === 'SHEET_TABLE_PLUGIN');
        expect(tableRes).toBeDefined();
        const parsed = JSON.parse(tableRes!.data);
        const tableCount = Object.keys(parsed)
            .filter((k) => !k.startsWith('_'))
            .reduce((n, sheetId) => n + Object.keys(parsed[sheetId] ?? {}).length, 0);
        expect(tableCount).toBeGreaterThanOrEqual(2);
    });
});

// ─── Importable fixtures: must NOT trip the catch (no false positives) ─
//
// REGRESSION HISTORY: an over-broad catch (e.g. one that wrapped every
// import error in NotesheetImportError without a stack-frame check) would
// hide real bugs in non-crashing fixtures. This block asserts each of
// the 8 importable fixtures completes without throwing AND yields a
// snapshot with the expected sheetOrder shape, so a future regression
// in the wrap that accidentally short-circuits a clean load is caught.

describe('M12 import recovery — importable fixtures still import cleanly', () => {
    const importable = [
        'BordersAndCellColors.xlsx',
        'ConditionalFormatting-Variants.xlsx',
        'EmptyAndDegenerate.xlsx',
        'FormattingSmorgasboard.xlsx',
        'FormattingSmorgasboard-NonAptosClassicThemeWithConditionalFormatting.xlsx',
        'Hyperlinks-Variants.xlsx',
        'MaliciousValues.xlsx',
        'MergedCellsAndAlignment.xlsx',
        'NumberFormats.xlsx',
        'RichTextInOneCell.xlsx',
    ];

    test.each(importable)('%s imports without throwing', async (file) => {
        const buf = readFileSync(path.join(FIXTURES_DIR, file));
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        // Light shape assertion — the wrap must not corrupt the
        // snapshot returned for working fixtures. We don't pin specific
        // sheet counts here (those are covered elsewhere); we just check
        // that we got a real snapshot back.
        expect(snap).toBeDefined();
        const sheetOrder = (snap as unknown as { sheetOrder: string[] }).sheetOrder;
        expect(Array.isArray(sheetOrder)).toBe(true);
        expect(sheetOrder.length).toBeGreaterThanOrEqual(1);
    });
});
