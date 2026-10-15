// M19 feature-1 (BACKLOG E3): arrow keys clamp at the sheet edge.
//
// Ground truth: Excel leaves the active cell where it is when an arrow key
// would move it past the sheet edge. Univer 0.23 wraps instead, because
// `findNextRange` in `@univerjs/sheets-ui` (lib/es/index.js, "selection-utils.ts"
// region) defaults `isGoBack = true`: Left in column A jumps to the last
// column of the previous row, and Up / Left at A1 jumps to the bottom-right
// cell. Upstream closed univer#6988 as by-design and advises overriding.
//
// The fix vetoes `sheet.command.move-selection` from a BeforeCommandExecute
// listener. These cases drive the pure edge-decision helper against a fake
// worksheet model (same method names as the core `Worksheet` the listener
// hands it) so the decision is testable without booting Univer.
//
// `Direction` is read from the REAL `@univerjs/core` export (the same enum
// the arrow shortcut items pass as `staticParameters.direction`; editorView
// injects it via `@univerjs/presets`, which re-exports core verbatim). Jest
// can't load Univer's CJS build in-process (it pulls ESM-only deps such as
// lodash-es and nanoid 5 — which is why other suites `jest.mock` Univer), so
// the enum is read in a child Node process, which loads it natively. Nothing
// here hard-codes 0..3.

import { execFileSync } from 'child_process';
import { existsSync, readFileSync } from 'fs';
import { dirname, join } from 'path';
import {
    MOVE_SELECTION_COMMAND_ID,
    makeArrowEdgeClamp,
    type ArrowEdgeRange,
    type ArrowEdgeSheet,
    type DirectionEnum,
} from '../src/editor/arrowEdgeClamp';

const Direction: DirectionEnum = JSON.parse(
    execFileSync(
        process.execPath,
        [
            '--input-type=module',
            '-e',
            "const m = await import('@univerjs/presets'); process.stdout.write(JSON.stringify(m.Direction));",
        ],
        { cwd: __dirname, encoding: 'utf8' },
    ),
);
type Direction = number;

// `JumpOver.moveGap` lives in @univerjs/sheets-ui, which can't be imported in
// plain Node (it pulls in CSS). Read the value from the installed build instead,
// so an upgrade that renumbers the enum fails the contract test below rather
// than silently changing what the clamp exempts. editorView injects the real
// enum value from @univerjs/preset-sheets-core.
// Package root of an installed dependency, resolved from `fromDir` the way Node
// does (main entry, then walk up to the package.json whose name matches). These
// packages' `exports` maps don't expose ./package.json. Resolving sheets-ui
// from @univerjs/preset-sheets-core (a direct dependency) avoids relying on npm
// hoisting a transitive package to the top level.
function packageRoot(name: string, fromDir: string): string {
    let dir = dirname(require.resolve(name, { paths: [fromDir] }));
    for (;;) {
        const pj = join(dir, 'package.json');
        if (existsSync(pj) && JSON.parse(readFileSync(pj, 'utf8')).name === name) return dir;
        const up = dirname(dir);
        if (up === dir) throw new Error(`no package root for ${name}`);
        dir = up;
    }
}
const SHEETS_UI_ES = readFileSync(
    join(
        packageRoot('@univerjs/sheets-ui', packageRoot('@univerjs/preset-sheets-core', __dirname)),
        'lib',
        'es',
        'index.js',
    ),
    'utf8',
);
const jumpOverMatch = /JumpOver\[JumpOver\["moveGap"\] = (\d+)\]/.exec(SHEETS_UI_ES);
const JUMP_OVER_GAP = jumpOverMatch ? Number(jumpOverMatch[1]) : NaN;
const shouldClampArrowMove = makeArrowEdgeClamp(Direction, JUMP_OVER_GAP);

interface FakeSheetOpts {
    rows?: number;
    cols?: number;
    hiddenRows?: number[];
    hiddenCols?: number[];
    merges?: ArrowEdgeRange[];
}

function fakeSheet(opts: FakeSheetOpts = {}): ArrowEdgeSheet {
    const rows = opts.rows ?? 1000;
    const cols = opts.cols ?? 20;
    const hiddenRows = new Set(opts.hiddenRows ?? []);
    const hiddenCols = new Set(opts.hiddenCols ?? []);
    const merges = opts.merges ?? [];
    return {
        getRowCount: () => rows,
        getColumnCount: () => cols,
        getRowVisible: (r: number) => !hiddenRows.has(r),
        getColVisible: (c: number) => !hiddenCols.has(c),
        getMergeData: () => merges,
    };
}

/** A1-style single cell → 0-based range. Only handles A–Z, enough here. */
function cell(a1: string): ArrowEdgeRange {
    const m = a1.match(/^([A-Z])(\d+)$/);
    if (!m) throw new Error(`bad ref ${a1}`);
    const c = m[1].charCodeAt(0) - 65;
    const r = Number(m[2]) - 1;
    return { startRow: r, endRow: r, startColumn: c, endColumn: c };
}

function range(a1: string): ArrowEdgeRange {
    const [a, b] = a1.split(':');
    const s = cell(a);
    const e = cell(b);
    return {
        startRow: s.startRow,
        endRow: e.endRow,
        startColumn: s.startColumn,
        endColumn: e.endColumn,
    };
}

const ROWS = 1000;
const COLS = 20; // last column is T (index 19)
const sheet = fakeSheet({ rows: ROWS, cols: COLS });
const lastCol = (r: number): ArrowEdgeRange => ({
    startRow: r,
    endRow: r,
    startColumn: COLS - 1,
    endColumn: COLS - 1,
});
const lastRow = (c: number): ArrowEdgeRange => ({
    startRow: ROWS - 1,
    endRow: ROWS - 1,
    startColumn: c,
    endColumn: c,
});

describe('M19 arrow edge clamp — Univer contract', () => {
    it('reads the four Direction members from @univerjs/presets (what editorView uses)', () => {
        for (const k of ['UP', 'RIGHT', 'DOWN', 'LEFT'] as const) {
            expect(typeof Direction[k]).toBe('number');
        }
        expect(new Set([Direction.UP, Direction.RIGHT, Direction.DOWN, Direction.LEFT]).size).toBe(
            4,
        );
    });
    it('reads JumpOver.moveGap from the installed @univerjs/sheets-ui (Ctrl/Cmd+arrow)', () => {
        expect(JUMP_OVER_GAP).toBe(1);
        // MoveSelectionCommand only takes the non-wrapping gap path for moveGap.
        expect(SHEETS_UI_ES).toContain('jumpOver === 1 ? findNextGapRange');
    });
    it('targets the Univer 0.23 MoveSelectionCommand id', () => {
        expect(MOVE_SELECTION_COMMAND_ID).toBe('sheet.command.move-selection');
    });
});

describe('M19 arrow edge clamp — cancel at the sheet edge (Excel behaviour)', () => {
    it.each<[string, ArrowEdgeRange, Direction]>([
        ['A5 Left', cell('A5'), Direction.LEFT],
        ['C1 Up', cell('C1'), Direction.UP],
        ['A1 Left', cell('A1'), Direction.LEFT],
        ['A1 Up', cell('A1'), Direction.UP],
        ['last-column Right', lastCol(4), Direction.RIGHT],
        ['last-row Down', lastRow(2), Direction.DOWN],
    ])('%s → cancel', (_label, r, direction) => {
        expect(shouldClampArrowMove({ direction }, r, sheet)).toBe(true);
    });
});

describe('M19 arrow edge clamp — normal movement is untouched', () => {
    it.each<[string, ArrowEdgeRange, Direction]>([
        ['B5 Left', cell('B5'), Direction.LEFT],
        ['A5 Right', cell('A5'), Direction.RIGHT],
        ['A5 Up', cell('A5'), Direction.UP],
        ['A1 Down', cell('A1'), Direction.DOWN],
        ['A1 Right', cell('A1'), Direction.RIGHT],
        ['last-column Left', lastCol(4), Direction.LEFT],
        ['last-row Up', lastRow(2), Direction.UP],
    ])('%s → allow', (_label, r, direction) => {
        expect(shouldClampArrowMove({ direction }, r, sheet)).toBe(false);
    });
});

describe('M19 arrow edge clamp — hidden rows / columns are not visible cells', () => {
    it('column A hidden, B5 active, Left → cancel', () => {
        const s = fakeSheet({ hiddenCols: [0] });
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, cell('B5'), s)).toBe(true);
    });
    it('columns A and B hidden, C5 active, Left → cancel', () => {
        const s = fakeSheet({ hiddenCols: [0, 1] });
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, cell('C5'), s)).toBe(true);
    });
    it('column B hidden only, C5 active, Left → allow (A is still visible)', () => {
        const s = fakeSheet({ hiddenCols: [1] });
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, cell('C5'), s)).toBe(false);
    });
    it('row 1 hidden, C2 active, Up → cancel', () => {
        const s = fakeSheet({ hiddenRows: [0] });
        expect(shouldClampArrowMove({ direction: Direction.UP }, cell('C2'), s)).toBe(true);
    });
    it('last column hidden, second-to-last active, Right → cancel', () => {
        const s = fakeSheet({ cols: COLS, hiddenCols: [COLS - 1] });
        const r: ArrowEdgeRange = {
            startRow: 4,
            endRow: 4,
            startColumn: COLS - 2,
            endColumn: COLS - 2,
        };
        expect(shouldClampArrowMove({ direction: Direction.RIGHT }, r, s)).toBe(true);
    });
    it('last row hidden, second-to-last active, Down → cancel', () => {
        const s = fakeSheet({ rows: ROWS, hiddenRows: [ROWS - 1] });
        const r: ArrowEdgeRange = {
            startRow: ROWS - 2,
            endRow: ROWS - 2,
            startColumn: 2,
            endColumn: 2,
        };
        expect(shouldClampArrowMove({ direction: Direction.DOWN }, r, s)).toBe(true);
    });
});

describe('M19 arrow edge clamp — merge-aware (outer edge on the movement side)', () => {
    const mergeAB = fakeSheet({ merges: [range('A5:B6')] });
    const mergeBC = fakeSheet({ merges: [range('B5:C6')] });

    it('active cell inside merge A5:B6 (selection = merge), Left → cancel', () => {
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, range('A5:B6'), mergeAB)).toBe(
            true,
        );
    });
    it('active cell inside merge A5:B6 (selection = interior cell B6), Left → cancel', () => {
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, cell('B6'), mergeAB)).toBe(true);
    });
    it('active cell inside merge B5:C6 (selection = merge), Left → allow', () => {
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, range('B5:C6'), mergeBC)).toBe(
            false,
        );
    });
    it('active cell inside merge B5:C6 (selection = interior cell C6), Left → allow', () => {
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, cell('C6'), mergeBC)).toBe(
            false,
        );
    });
    it('merge A1:B2, Up → cancel; Down → allow', () => {
        const s = fakeSheet({ merges: [range('A1:B2')] });
        expect(shouldClampArrowMove({ direction: Direction.UP }, cell('B2'), s)).toBe(true);
        expect(shouldClampArrowMove({ direction: Direction.DOWN }, cell('B2'), s)).toBe(false);
    });
    it('merge reaching the last column, Right from its interior → cancel', () => {
        const m: ArrowEdgeRange = {
            startRow: 4,
            endRow: 5,
            startColumn: COLS - 3,
            endColumn: COLS - 1,
        };
        const s = fakeSheet({ merges: [m] });
        const interior: ArrowEdgeRange = {
            startRow: 5,
            endRow: 5,
            startColumn: COLS - 3,
            endColumn: COLS - 3,
        };
        expect(shouldClampArrowMove({ direction: Direction.RIGHT }, interior, s)).toBe(true);
    });
    it('a merge elsewhere does not affect B5 Left', () => {
        const s = fakeSheet({ merges: [range('D1:E2')] });
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, cell('B5'), s)).toBe(false);
    });
});

describe('M19 arrow edge clamp — jumpOver, formula-editor and malformed calls', () => {
    it.each<[string, Record<string, unknown>]>([
        [
            'jumpOver = moveGap (Ctrl/Cmd+arrow)',
            { direction: Direction.LEFT, jumpOver: JUMP_OVER_GAP },
        ],
        ['extra = formula-editor', { direction: Direction.LEFT, extra: 'formula-editor' }],
        ['extra = any string', { direction: Direction.LEFT, extra: 'something' }],
    ])('A1 with %s → allow', (_label, params) => {
        expect(shouldClampArrowMove(params, cell('A1'), sheet)).toBe(false);
    });

    // Univer only skips the wrapping findNextRange for jumpOver === moveGap;
    // moveStopeOne (0), moveStepPage (2) and moveStepEnd (3) still wrap, so
    // they must be vetoed at the edge like a plain arrow.
    it.each<[string, number]>([
        ['moveStopeOne', 0],
        ['moveStepPage', 2],
        ['moveStepEnd', 3],
    ])('A1 Left with jumpOver = %s (%i) → cancel (still wraps in Univer)', (_n, jumpOver) => {
        expect(
            shouldClampArrowMove({ direction: Direction.LEFT, jumpOver }, cell('A1'), sheet),
        ).toBe(true);
    });
    it('B5 Left with jumpOver = 0 → allow (not at the edge)', () => {
        expect(
            shouldClampArrowMove({ direction: Direction.LEFT, jumpOver: 0 }, cell('B5'), sheet),
        ).toBe(false);
    });

    it('no params → allow (the command itself returns false)', () => {
        expect(shouldClampArrowMove(undefined, cell('A1'), sheet)).toBe(false);
    });
    it('missing direction → allow', () => {
        expect(shouldClampArrowMove({}, cell('A1'), sheet)).toBe(false);
    });
    it('no active selection → allow', () => {
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, null, sheet)).toBe(false);
    });
});

// KNOWN SHORTCOMING (README "Known gaps → Editor", BACKLOG E5, milestone TBD):
// Univer moves a multi-cell selection from the range's EDGE, not from the
// active cell. Excel collapses A5:C5 (active C5) + Left to B5; Notesheet
// cancels because the range's left edge is already in column A. Pinned so a
// change in either direction is a deliberate decision, not an accident.
describe('M19 arrow edge clamp — KNOWN SHORTCOMING: multi-cell selection moves from its edge', () => {
    it('A5:C5 + Left → cancel (Excel would move the active cell C5 to B5)', () => {
        expect(shouldClampArrowMove({ direction: Direction.LEFT }, range('A5:C5'), sheet)).toBe(
            true,
        );
    });
    it('A5:C5 + Right → allow (the range edge is not at the sheet edge)', () => {
        expect(shouldClampArrowMove({ direction: Direction.RIGHT }, range('A5:C5'), sheet)).toBe(
            false,
        );
    });
});
