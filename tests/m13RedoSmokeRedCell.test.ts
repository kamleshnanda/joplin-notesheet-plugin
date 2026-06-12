// Pin-down for the PR #17 smoke-seed leak.
//
// HISTORY
//
// PR #17 (the very first PGE harness cycle) added a "smoke seed" inline
// to `emptySnapshot()` in `src/snapshot.ts` so the harness could prove
// end-to-end wiring by rendering a single red "harness-smoke-OK" cell
// in a fresh Notesheet. That seed never got reverted, so every "New
// Spreadsheet" command in the deployed plugin shipped with cell A1
// pre-filled — operator caught it 2026-06-03 mid-M13/E review.
//
// CURRENT CONTRACT
//
// `emptySnapshot()` returns a workbook whose first sheet has empty
// `cellData`. The PGE harness's smoke fixture is built independently
// in `scripts/pge/create-seeded-notesheet.js` and is no longer wired
// into production code.
//
// WOULD HAVE CAUGHT (had it shipped earlier)
//
// - Any code path that re-introduces a debug seed into `emptySnapshot()`
//   would fail this pin-down before the .jpl ships.
// - Any future "Import .xlsx as empty" path that re-uses
//   `emptySnapshot()` and silently inherits a seed.

import { emptySnapshot, extractSnapshot, wrapSnapshot } from '../src/snapshot';

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<
        string,
        {
            id: string;
            name: string;
            cellData: Record<number, Record<number, { v?: unknown; s?: string }>>;
        }
    >;
    styles: Record<string, unknown>;
}

describe('emptySnapshot — PR #17 smoke-seed leak (reverted)', () => {
    test('emptySnapshot has no A1 cell entry', () => {
        const snap = emptySnapshot() as unknown as SnapshotShape;
        const sheetId = snap.sheetOrder[0];
        const sheet = snap.sheets[sheetId];

        // The fix: cellData must have no row 0 (or row 0 with no col 0).
        // Both shapes are equivalent for Univer ("no entry" === "default
        // empty"); we accept either.
        const row0 = sheet?.cellData?.[0];
        if (row0 !== undefined) {
            expect(row0[0]).toBeUndefined();
        }
    });

    test('emptySnapshot has no smoke-style entry in styles', () => {
        const snap = emptySnapshot() as unknown as SnapshotShape;
        // The seed shipped a `pge-smoke-red` style entry. Production code
        // must not ship harness-only style ids.
        expect(snap.styles).toBeDefined();
        expect(snap.styles['pge-smoke-red']).toBeUndefined();
    });

    test('emptySnapshot is a valid Univer IWorkbookData shape (sheetOrder + sheets + styles)', () => {
        // Even with no seed, the snapshot must still satisfy the
        // IWorkbookData shape Univer expects at load time.
        const snap = emptySnapshot() as unknown as SnapshotShape;

        expect(Array.isArray(snap.sheetOrder)).toBe(true);
        expect(snap.sheetOrder.length).toBeGreaterThanOrEqual(1);

        const firstSheetId = snap.sheetOrder[0];
        expect(snap.sheets[firstSheetId]).toBeDefined();
        expect(snap.sheets[firstSheetId].id).toBe(firstSheetId);
        expect(typeof snap.sheets[firstSheetId].name).toBe('string');
        expect(typeof snap.sheets[firstSheetId].cellData).toBe('object');

        expect(snap.styles).toBeDefined();
        expect(typeof snap.styles).toBe('object');
    });

    test('emptySnapshot survives wrap → extract round-trip without mutation', () => {
        const snap = emptySnapshot();
        const body = wrapSnapshot(snap);
        const result = extractSnapshot(body);

        expect(result.ok).toBe(true);
        if (!result.ok) return;

        const round = result.snapshot as unknown as SnapshotShape;
        expect(round.sheetOrder).toEqual((snap as unknown as SnapshotShape).sheetOrder);
        const sheetId = round.sheetOrder[0];
        // Empty cellData survives the JSON round-trip as `{}`.
        expect(round.sheets[sheetId].cellData).toEqual({});
    });
});
