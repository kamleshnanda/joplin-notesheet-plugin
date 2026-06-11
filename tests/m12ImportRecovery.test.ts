// Mock @univerjs/sheets-table so importing src/xlsx (transitively pulled by
// the SUT) doesn't drag the whole Univer ESM graph through jest's CJS
// transform. Mirrors the pattern in m12FixturePinDowns.test.ts.
jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

// Pin-down regression tests for M12 import-error recovery.
//
// Three project-owned fixtures — MultiSheet.xlsx, LargeWorkbook.xlsx,
// FormulasAndStructuredRefs.xlsx — crash inside exceljs's reconcile
// pipeline when loaded. The crashes reproduce on a bare
// `await wb.xlsx.load(buf)` with zero Notesheet code involved, so they
// are pre-existing exceljs bugs we can't fix without forking the package.
//
// Until those fixtures import cleanly (would require either an exceljs
// patch or a switch to a different reader), xlsxBufferToSnapshot()
// catches the crash and rethrows a typed NotesheetImportError with a
// user-actionable message. The user-facing dialog in src/index.ts and
// the editor status bar in src/editorView.tsx both consume `error.message`
// already, so the wrapped error's friendly text reaches the user without
// any caller-side change.
//
// WOULD HAVE CAUGHT (2026-05-31): the original behavior leaked a raw
// `TypeError: Cannot read properties of undefined (reading 'anchors')`
// stack into the Joplin error dialog. Users had no way to tell whether
// the file was corrupt, whether they'd hit a Notesheet bug, or whether
// to file an issue. After this wrap landed, the dialog reads "This .xlsx
// contains chart drawings that Notesheet can't import yet" — actionable.
//
// REGRESSION HISTORY: future agents may be tempted to "fix" the underlying
// exceljs crash by patching node_modules or upgrading. Two of these
// crashes (anchors, name-in-tables-reduce) reproduce on the latest
// 4.4.0 of exceljs as of 2026-06. If a future exceljs version DOES
// resolve them, these tests will fail at the rejection assertion — that's
// the signal to flip the fixture into the importable list and delete
// the corresponding case here.

import { readFileSync } from 'fs';
import path from 'path';

import { NotesheetImportError, xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');

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
        const resources = (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const drawing = resources.find((r) => r.name === 'SHEET_DRAWING_PLUGIN');
        expect(drawing).toBeDefined();
        const parsed = JSON.parse(drawing!.data);
        let chartCount = 0;
        for (const su of Object.keys(parsed)) {
            chartCount += Object.keys(parsed[su].data).length;
        }
        expect(chartCount).toBeGreaterThanOrEqual(1);
    });

    test('LargeWorkbook.xlsx → xlsx-multi-table-unsupported (post-M17, second crash class)', async () => {
        // Pre-M17: this fixture surfaced the `anchors` crash class
        // (chart-bearing drawing). M17's pre-load chart strip removed
        // the chart parts; what remains underneath is the multi-table
        // reduce crash in exceljs's worksheet.js:920 — the same
        // structural issue FormulasAndStructuredRefs.xlsx exhibits.
        // The error class flipped from xlsx-charts-unsupported to
        // xlsx-multi-table-unsupported as a consequence.
        const err = await loadAndCatch('LargeWorkbook.xlsx');
        expect(err).toBeInstanceOf(NotesheetImportError);
        const e = err as NotesheetImportError;
        expect(e.code).toBe('xlsx-multi-table-unsupported');
        expect(e.message).not.toMatch(/anchors/i);
        expect(e.message).not.toMatch(/undefined/i);
        expect(e.message).not.toMatch(/TypeError/i);
        expect(e.cause).toBeDefined();
    });

    test('FormulasAndStructuredRefs.xlsx → xlsx-multi-table-unsupported', async () => {
        const err = await loadAndCatch('FormulasAndStructuredRefs.xlsx');
        expect(err).toBeInstanceOf(NotesheetImportError);
        const e = err as NotesheetImportError;
        expect(e.code).toBe('xlsx-multi-table-unsupported');

        // Same negative checks as above. The raw exceljs error here is
        // "Cannot read properties of undefined (reading 'name')" — the
        // bare word "name" is too generic to forbid in the message
        // (our friendly text legitimately uses "named tables"), so we
        // only check that the cryptic shape is gone.
        expect(e.message).not.toMatch(/undefined/i);
        expect(e.message).not.toMatch(/TypeError/i);
        expect(e.message).not.toMatch(/Cannot read properties/);

        // Friendly explanation must mention what's unsupported.
        expect(e.message).toMatch(/table/i);
        expect(e.message).toMatch(/Notesheet/);

        expect(e.cause).toBeDefined();
        expect((e.cause as Error).message).toMatch(/name/);
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
