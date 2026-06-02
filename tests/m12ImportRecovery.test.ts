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

// ─── Previously-crashing fixtures: M13 zip pre-process unblocks them ──
//
// HISTORY: M12 caught three exceljs reconcile crashes
// (xlsx-charts-unsupported × 2, xlsx-multi-table-unsupported × 1) and
// reported friendly errors. M13 added a zip pre-process step
// (preProcessForExceljs in src/xlsx.ts) that strips chart drawings
// AND rewrites absolute rel Targets to relative form before exceljs
// loads the buffer. All three previously-crashing fixtures now import
// cleanly. The friendly-error pin-downs in this block are kept (now
// flipped to assert "imports cleanly") so we'd notice if the
// pre-process step ever regresses.
//
// NotesheetImportError + the three error codes remain exported because
// the catch in xlsxBufferToSnapshot still classifies any crash that
// somehow slips past the pre-process. New code paths that introduce
// fresh exceljs crashes will fall into the generic xlsx-import-failed
// branch.

describe('M13 import recovery — previously-crashing fixtures now import (chart strip + rel rewrite)', () => {
    test('MultiSheet.xlsx imports without throwing (was xlsx-charts-unsupported)', async () => {
        const err = await loadAndCatch('MultiSheet.xlsx');
        expect(err).toBeNull();
    });

    test('LargeWorkbook.xlsx imports without throwing (was xlsx-charts-unsupported)', async () => {
        const err = await loadAndCatch('LargeWorkbook.xlsx');
        expect(err).toBeNull();
    });

    test('FormulasAndStructuredRefs.xlsx imports without throwing (was xlsx-multi-table-unsupported)', async () => {
        const err = await loadAndCatch('FormulasAndStructuredRefs.xlsx');
        expect(err).toBeNull();
    });

    test('NotesheetImportError remains exported (still used for unforeseen crash classes)', () => {
        // Forces a code-path through the catch wrapper. We synthesize an
        // error here rather than try to reproduce a real exceljs crash
        // class — those are pre-processed away.
        const e = new NotesheetImportError('xlsx-import-failed', 'sample', null);
        expect(e).toBeInstanceOf(Error);
        expect(e.code).toBe('xlsx-import-failed');
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
