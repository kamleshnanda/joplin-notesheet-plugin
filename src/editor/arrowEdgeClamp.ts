// M19 feature-1 (BACKLOG E3) — arrow keys clamp at the sheet edge, as in Excel.
//
// Univer 0.23 wraps instead: `findNextRange` in `@univerjs/sheets-ui`
// (lib/es/index.js, "selection-utils.ts" region) defaults `isGoBack = true`,
// so Left in column A jumps to the last column of the previous row and Up /
// Left at A1 jumps to the bottom-right cell. Upstream closed univer#6988 as
// by-design and advises overriding.
//
// editorView.tsx vetoes `sheet.command.move-selection` from a
// `BeforeCommandExecute` listener when this helper says the move would wrap.
// A command-level veto (rather than a shortcut override) is needed because
// committing a cell edit with an arrow key calls MoveSelectionCommand directly
// from the editor-bridge `_moveSelection` path and never goes through
// ShortcutService.
//
// The decision mirrors MoveSelectionCommand + findNextRange exactly:
//   - `getStartRange` keeps the selection range's extent along the movement
//     axis (it only collapses the orthogonal axis to the primary cell), so the
//     edge that matters is the RANGE's outer edge on the movement side. For a
//     merged cell the selection range is the merge, so this is the merge's
//     outer edge. We also expand by any intersecting merge, so a caller that
//     passes an interior cell of a merge gets the same answer.
//   - findNextRange steps one row/column past that edge and skips hidden
//     rows/columns. If it runs off the sheet (no visible cell beyond), the
//     `isGoBack` branch wraps — that is exactly the case we veto.
//   - `jumpOver` (Ctrl/Cmd+arrow → findNextGapRange, already clamps) and
//     `extra` (the formula editor's ref-selection moves) are never vetoed.
//
// Pure: no Univer import. `Direction` is injected (editorView passes the real
// enum from `@univerjs/presets`, which re-exports `@univerjs/core`), so no enum
// value is hard-coded here and Jest can drive it without booting Univer.
//
// Univer 1.x upgrade (BACKLOG F5) checklist: this depends on the facade's
// `BeforeCommandExecute` event + `event.cancel`, on the command id below and
// on the `{ direction, jumpOver, extra }` param shape. Re-verify all three.
// Also re-check the cancel leak: in Univer 0.23 core `executeCommand`, a
// `CanceledError` thrown by a before-listener skips the command-stack
// `dispose()`, so each vetoed keystroke leaves one stale
// `_commandExecutionStack` entry (reset on every note open; harmless today —
// only `syncExecuteCommand` reads it, to set a mutation's `trigger`, and every
// `trigger` reader compares other ids). Drop this note if upstream wraps the
// before-listeners in try/finally.

export const MOVE_SELECTION_COMMAND_ID = 'sheet.command.move-selection';

/** Structural copy of `@univerjs/core`'s `Direction` enum (values injected). */
export interface DirectionEnum {
    readonly UP: number;
    readonly RIGHT: number;
    readonly DOWN: number;
    readonly LEFT: number;
}

/** Subset of `@univerjs/core` `IRange` this helper reads. */
export interface ArrowEdgeRange {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}

/** Subset of `@univerjs/core` `Worksheet` (method names match exactly). */
export interface ArrowEdgeSheet {
    getRowCount(): number;
    getColumnCount(): number;
    getRowVisible(row: number): boolean;
    getColVisible(col: number): boolean;
    getMergeData(): ArrowEdgeRange[];
}

/** Returns `true` when the move-selection call must be cancelled. */
export type ArrowEdgeClamp = (
    params: unknown,
    range: ArrowEdgeRange | null,
    sheet: ArrowEdgeSheet,
) => boolean;

interface MoveSelectionParamsLike {
    direction?: unknown;
    jumpOver?: unknown;
    extra?: unknown;
}

function intersects(a: ArrowEdgeRange, b: ArrowEdgeRange): boolean {
    return (
        a.startRow <= b.endRow &&
        b.startRow <= a.endRow &&
        a.startColumn <= b.endColumn &&
        b.startColumn <= a.endColumn
    );
}

/** Grow `range` until no merge partially overlaps it. */
function expandByMerges(range: ArrowEdgeRange, merges: ArrowEdgeRange[]): ArrowEdgeRange {
    const out = { ...range };
    let grew = true;
    while (grew) {
        grew = false;
        for (const m of merges) {
            if (!intersects(out, m)) continue;
            if (m.startRow < out.startRow) {
                out.startRow = m.startRow;
                grew = true;
            }
            if (m.endRow > out.endRow) {
                out.endRow = m.endRow;
                grew = true;
            }
            if (m.startColumn < out.startColumn) {
                out.startColumn = m.startColumn;
                grew = true;
            }
            if (m.endColumn > out.endColumn) {
                out.endColumn = m.endColumn;
                grew = true;
            }
        }
    }
    return out;
}

export function makeArrowEdgeClamp(direction: DirectionEnum): ArrowEdgeClamp {
    return (params, range, sheet) => {
        if (!params || typeof params !== 'object' || !range) return false;
        const p = params as MoveSelectionParamsLike;
        // Ctrl/Cmd+arrow (findNextGapRange) and formula-editor moves keep
        // Univer's own behaviour. Any presence of the key counts.
        if (p.jumpOver !== undefined && p.jumpOver !== null) return false;
        if (p.extra !== undefined && p.extra !== null) return false;

        const r = expandByMerges(range, sheet.getMergeData() ?? []);
        switch (p.direction) {
            case direction.LEFT: {
                for (let c = r.startColumn - 1; c >= 0; c--) {
                    if (sheet.getColVisible(c)) return false;
                }
                return true;
            }
            case direction.RIGHT: {
                const n = sheet.getColumnCount();
                for (let c = r.endColumn + 1; c < n; c++) {
                    if (sheet.getColVisible(c)) return false;
                }
                return true;
            }
            case direction.UP: {
                for (let row = r.startRow - 1; row >= 0; row--) {
                    if (sheet.getRowVisible(row)) return false;
                }
                return true;
            }
            case direction.DOWN: {
                const n = sheet.getRowCount();
                for (let row = r.endRow + 1; row < n; row++) {
                    if (sheet.getRowVisible(row)) return false;
                }
                return true;
            }
            default:
                return false;
        }
    };
}
