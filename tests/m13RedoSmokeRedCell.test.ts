// Pre-implementation pin-down for the M13-redo PGE harness smoke feature
// (`feature-1-smoke-red-cell` in BUILD_PLAN.md, sourced from OPERATOR_ASK.md).
//
// PURPOSE OF THIS FILE
//
// The smoke feature is intentionally NOT YET IMPLEMENTED. These tests
// exist to lay the test bed BEFORE the generator session runs, so:
//
//   1. The generator has a clear, mechanical specification of the
//      snapshot shape it must produce.
//   2. When the generator's first build lands, jest will go from
//      red → green and the harness will see deterministic progress.
//   3. If a future contributor accidentally regresses the seed (e.g.
//      reverts `emptySnapshot()` back to "no cells"), this file fails
//      LOUD before the user even opens Joplin.
//
// SEAM IDENTIFIED
//
// The single function that produces the initial empty workbook for a
// brand-new Notesheet note is `emptySnapshot()` in `src/snapshot.ts`
// (lines 65-85). It is the only call-site of the "no prior snapshot"
// path — `src/index.ts:103` does `wrapSnapshot(emptySnapshot())` inside
// the `newSpreadsheet` command's `execute`. Putting the seed there
// makes it pure-testable and centralises the shape in one place.
//
// WOULD HAVE CAUGHT
//
// - A generator that edits `src/index.ts` to inline the seeded snapshot
//   but leaves `emptySnapshot()` untouched, then a future "Import
//   .xlsx as empty" path silently re-uses `emptySnapshot()` and ships
//   without the seed — the generator's smoke screenshot would still
//   pass, but the snapshot helper would no longer match its declared
//   contract.
// - A generator that puts the value `'harness-smoke-OK'` on the cell
//   but forgets the style/color — Univer would render it black-on-
//   white, the visual gate would fail, but the failing test message
//   would point straight at the missing style entry.
// - A generator that synthesizes a style with `cl: { rgb: 'FF0000' }`
//   (no leading `#`) — Univer's style resolver requires the `#` prefix
//   on srgbClr values; without it the cell renders default-coloured.
//
// REGRESSION HISTORY
//
// (None yet — this is feature-1, the very first cycle of the harness.)
//
// SPEC TRACEABILITY
//
// - OPERATOR_ASK.md lines 18-30: cell A1 = "harness-smoke-OK" in red.
// - BUILD_PLAN.md feature-1-smoke-red-cell, lines 22-44: identical
//   acceptance criteria + the explicit reminder that the cell-level
//   style must reference a `style-N` entry whose `cl.rgb === '#FF0000'`.
//
// The acceptance criteria in BUILD_PLAN.md are visual (the evaluator
// reads the screenshot with the human eye). These tests assert the
// SNAPSHOT-LEVEL invariants that necessarily precede the visual
// outcome. The visual gate is the evaluator's job; jest is the fast
// pre-check.

import { emptySnapshot, extractSnapshot, wrapSnapshot } from '../src/snapshot';

const EXPECTED_CELL_TEXT = 'harness-smoke-OK';
const EXPECTED_RGB = '#FF0000';

interface SeededSnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        id: string;
        name: string;
        cellData: Record<number, Record<number, { v?: unknown; s?: string }>>;
    }>;
    styles: Record<string, { cl?: { rgb?: string } }>;
}

describe('M13-redo smoke (feature-1-smoke-red-cell): empty snapshot scaffold', () => {
    test('import: emptySnapshot() preserves cell A1 = "harness-smoke-OK" literal value', () => {
        const snap = emptySnapshot() as unknown as SeededSnapshotShape;
        const sheetId = snap.sheetOrder[0];
        const sheet = snap.sheets[sheetId];
        const a1 = sheet?.cellData?.[0]?.[0];

        expect(a1).toBeDefined();
        expect(a1!.v).toBe(EXPECTED_CELL_TEXT);
    });

    test('import: emptySnapshot() preserves cell A1 font color = #FF0000 via styles entry', () => {
        const snap = emptySnapshot() as unknown as SeededSnapshotShape;
        const sheetId = snap.sheetOrder[0];
        const sheet = snap.sheets[sheetId];
        const a1 = sheet?.cellData?.[0]?.[0];

        // The cell MUST reference a style by id, not inline. Univer's
        // style resolver only honours the `cl` field when it lives in
        // `snapshot.styles[<id>]` and the cell carries `s: <id>`.
        expect(a1).toBeDefined();
        expect(typeof a1!.s).toBe('string');
        expect(a1!.s!.length).toBeGreaterThan(0);

        const styleEntry = snap.styles[a1!.s as string];
        expect(styleEntry).toBeDefined();
        expect(styleEntry!.cl).toBeDefined();
        // Hash-prefixed uppercase hex; lowercase or missing-# variants
        // do not render as red in Univer.
        expect(styleEntry!.cl!.rgb).toBe(EXPECTED_RGB);
    });

    test('import: emptySnapshot() preserves "all other cells empty" — only A1 has a value', () => {
        const snap = emptySnapshot() as unknown as SeededSnapshotShape;
        const sheetId = snap.sheetOrder[0];
        const sheet = snap.sheets[sheetId];
        const row0 = sheet?.cellData?.[0] ?? {};

        // Every other column in row 0 must be absent or value-less.
        for (const colKey of Object.keys(row0)) {
            if (colKey === '0') continue;
            const cell = row0[Number(colKey)];
            // The harness contract: no `v` (or undefined v) on any
            // non-A1 cell in row 0. Style-only cells are tolerated
            // (none should exist, but a future formatting-default
            // could legitimately add one — the spec only forbids
            // visible TEXT in non-A1 cells).
            expect(cell?.v === undefined || cell?.v === '').toBe(true);
        }

        // Rows 1+ must have no cellData entries at all (this is a
        // freshly-created sheet — only A1 was seeded).
        for (const rowKey of Object.keys(sheet.cellData)) {
            if (rowKey === '0') continue;
            const row = sheet.cellData[Number(rowKey)];
            const filled = Object.values(row).filter((c) => c?.v !== undefined && c?.v !== '');
            expect(filled).toEqual([]);
        }
    });

    test('import: emptySnapshot() preserves Univer\'s required workbook shape (sheetOrder + sheets + styles)', () => {
        // Positive structural sanity: even with the seed, the snapshot
        // must still satisfy the IWorkbookData shape Univer expects at
        // load time. If the generator regresses this (e.g. drops
        // `sheetOrder` or stuffs the seed under a key Univer ignores),
        // Univer fails to bootUniver and the editor view is blank —
        // no amount of style-tweaking will recover.
        const snap = emptySnapshot() as unknown as SeededSnapshotShape;

        expect(Array.isArray(snap.sheetOrder)).toBe(true);
        expect(snap.sheetOrder.length).toBeGreaterThanOrEqual(1);

        const firstSheetId = snap.sheetOrder[0];
        expect(snap.sheets[firstSheetId]).toBeDefined();
        expect(snap.sheets[firstSheetId].id).toBe(firstSheetId);
        expect(typeof snap.sheets[firstSheetId].name).toBe('string');
        expect(typeof snap.sheets[firstSheetId].cellData).toBe('object');

        // styles must be an object — even an empty seed without colors
        // requires the empty-object placeholder, because Univer's
        // resolver does `snapshot.styles[id]` and bombs if styles is
        // missing entirely.
        expect(snap.styles).toBeDefined();
        expect(typeof snap.styles).toBe('object');
    });

    test('import: emptySnapshot() seed survives wrap → extract round-trip without mutation', () => {
        // The on-disk format is a markdown fence wrapping JSON.
        // `JSON.stringify` followed by `JSON.parse` should preserve
        // every field. If the generator accidentally puts non-
        // serialisable content in the seed (e.g. a `Date` instance),
        // this test catches it before Joplin sees a mangled note.
        const snap = emptySnapshot();
        const body = wrapSnapshot(snap);
        const result = extractSnapshot(body);

        expect(result.ok).toBe(true);
        if (!result.ok) return;

        const round = result.snapshot as unknown as SeededSnapshotShape;
        const sheetId = round.sheetOrder[0];
        const a1 = round.sheets[sheetId].cellData[0][0];
        expect(a1.v).toBe(EXPECTED_CELL_TEXT);

        const styleEntry = round.styles[a1.s as string];
        expect(styleEntry?.cl?.rgb).toBe(EXPECTED_RGB);
    });
});
