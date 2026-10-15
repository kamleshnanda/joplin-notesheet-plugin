// M19 feature-1 (BACKLOG E3) — arrow-key edge clamp. STUB: models today's
// behaviour (never veto), so tests/m19ArrowEdgeClamp.test.ts fails first.

export const MOVE_SELECTION_COMMAND_ID = 'sheet.command.move-selection';

export interface DirectionEnum {
    readonly UP: number;
    readonly RIGHT: number;
    readonly DOWN: number;
    readonly LEFT: number;
}

export interface ArrowEdgeRange {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}

export interface ArrowEdgeSheet {
    getRowCount(): number;
    getColumnCount(): number;
    getRowVisible(row: number): boolean;
    getColVisible(col: number): boolean;
    getMergeData(): ArrowEdgeRange[];
}

export type ArrowEdgeClamp = (
    params: unknown,
    range: ArrowEdgeRange | null,
    sheet: ArrowEdgeSheet,
) => boolean;

export function makeArrowEdgeClamp(_direction: DirectionEnum): ArrowEdgeClamp {
    return () => false;
}
