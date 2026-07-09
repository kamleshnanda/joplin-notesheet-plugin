// M18 finding #3: register a generic "Text to Number" command + Data-ribbon
// menu item that converts numeric-looking text (percent, grouped, currency,
// plain float) into a real number with the matching format — mirroring Excel's
// "Convert to Number", which Univer's built-in `sheet.command.text-to-number`
// refuses to do (its isRealNum gate rejects "0%", "1,234", "$5", etc.).
//
// Registration recipe verified against @univerjs 0.23.0 bundled source:
//   - Register AFTER univer.createUnit(): the UNIVER_SHEET plugins (sheets,
//     sheets-ui, sheets-numfmt) are lazily instantiated on createUnit, so
//     SheetsSelectionsService / SetNumfmtCommand / the ribbon menu schema only
//     exist afterwards. The ribbon rebuilds reactively on menuChanged$, so a
//     late mergeMenu still appears.
//   - Command mirrors the built-in TextToNumberCommand structure
//     (@univerjs/sheets: getSheetCommandTarget + SheetsSelectionsService +
//     SetRangeValuesMutation), but relaxes the numeric gate via
//     parseTextToNumber and applies the derived numfmt via the high-level
//     SetNumfmtCommand (@univerjs/sheets-numfmt), which also fixes up the cell
//     value-type and handles undo.
//   - The built-in Data → Text to Number item is hidden via the preset `menu`
//     config in editorView (commandId can't be repointed on an existing item),
//     and we add our own item under the same Data-ribbon group.

// Core DI + data types come from the presets barrel (re-exports @univerjs/core).
import {
    ICommandService,
    IUniverInstanceService,
    IUndoRedoService,
    CommandType,
    ObjectMatrix,
    sequenceExecute,
} from '@univerjs/presets';
// Sheet services, mutations, numfmt command, and the menu/ribbon symbols are
// all re-exported by @univerjs/preset-sheets-core (which the editor already
// imports UniverSheetsCorePreset from). Importing them here — rather than from
// the deep transitive @univerjs/sheets / @univerjs/ui / @univerjs/sheets-numfmt
// packages — keeps us on the same barrel the rest of the editor uses and avoids
// a version-skew risk if the presets bump their internal pins.
import {
    SheetsSelectionsService,
    getSheetCommandTarget,
    SetRangeValuesMutation,
    SetRangeValuesUndoMutationFactory,
    IMenuManagerService,
    RibbonPosition,
    RibbonDataGroup,
    MenuItemType,
} from '@univerjs/preset-sheets-core';

import { parseTextToNumber } from './textToNumber';

// Univer CellValueType.NUMBER. Hard-coded (2) to avoid importing the enum where
// the bundled export shape is awkward; matches CellValueType.NUMBER in 0.23.
const CELL_TYPE_NUMBER = 2;

export const NOTESHEET_TEXT_TO_NUMBER_COMMAND_ID = 'notesheet.command.text-to-number-generic';

// Univer's built-in Text-to-Number ribbon button id (sheets-ui). We repoint
// this existing item's command at ours rather than adding a second button.
const BUILTIN_TEXT_TO_NUMBER_MENU_ID = 'sheet.toolbar.text-to-number';

// A minimal structural view of the Univer accessor + services we touch. The
// bundled packages are loosely typed here (the presets barrel doesn't re-export
// everything), so we describe just what we call.
interface CellRaw {
    v?: unknown;
    t?: number;
    s?: unknown;
}
interface WorksheetLike {
    getCellRaw: (row: number, col: number) => CellRaw | null | undefined;
}
interface RangeLike {
    startRow: number;
    endRow: number;
    startColumn: number;
    endColumn: number;
}
interface Accessor {
    get: (id: unknown) => unknown;
}

// Register the command + menu on a live Univer injector. Call AFTER createUnit.
// Fail-soft: any wiring error is logged, never thrown (a missing service must
// not break the editor boot).
export function registerTextToNumberCommand(injector: { get: (id: unknown) => unknown }): void {
    try {
        const commandService = injector.get(ICommandService) as {
            registerCommand: (cmd: unknown) => unknown;
            hasCommand?: (id: string) => boolean;
            syncExecuteCommand: (id: string, params?: unknown) => unknown;
        };
        if (!commandService?.registerCommand) {
            console.warn('[Notesheet] Text-to-Number: ICommandService unavailable; skipping');
            return;
        }
        // Guard against double registration (the editor can re-boot on note
        // switch within one webview); registerCommand throws on a dup id.
        if (commandService.hasCommand?.(NOTESHEET_TEXT_TO_NUMBER_COMMAND_ID)) return;

        commandService.registerCommand({
            id: NOTESHEET_TEXT_TO_NUMBER_COMMAND_ID,
            type: CommandType.COMMAND,
            handler: (accessor: Accessor, params?: { ranges?: RangeLike[] }) => {
                const univerIS = accessor.get(IUniverInstanceService);
                // getSheetCommandTarget is loosely typed via the barrel; cast
                // the instance service through unknown to satisfy its signature.
                const target = (getSheetCommandTarget as (is: unknown) => unknown)(univerIS) as {
                    worksheet?: WorksheetLike;
                    unitId?: string;
                    subUnitId?: string;
                } | null;
                if (!target?.worksheet || !target.unitId || !target.subUnitId) return false;
                const { worksheet, unitId, subUnitId } = target;

                const selections = accessor.get(SheetsSelectionsService) as {
                    getCurrentSelections?: () => Array<{ range: RangeLike }>;
                };
                const ranges: RangeLike[] =
                    params?.ranges ??
                    (selections.getCurrentSelections?.() ?? []).map((s) => s.range);
                if (!ranges.length) return false;

                const newValues = new ObjectMatrix();
                let converted = 0;

                for (const rng of ranges) {
                    for (let r = rng.startRow; r <= rng.endRow; r++) {
                        for (let c = rng.startColumn; c <= rng.endColumn; c++) {
                            const cell = worksheet.getCellRaw(r, c);
                            if (!cell || cell.v == null) continue;
                            if (cell.t === CELL_TYPE_NUMBER) continue; // already numeric
                            const parsed = parseTextToNumber(String(cell.v));
                            if (!parsed) continue; // not a number → leave as text
                            // Write value + type + (optional) number format in ONE
                            // cell payload. Folding the numfmt into an inline style
                            // (s.n.pattern) on the SAME SetRangeValuesMutation is
                            // deliberate: a separate follow-up SetNumfmtCommand
                            // reverts the just-converted percent cells when the
                            // command is dispatched from the ribbon button (the
                            // button's layoutService.focus() re-enters the numfmt
                            // interceptor and re-derives the cell type from the
                            // pre-conversion text). One atomic mutation is
                            // dispatch-path-independent — it converts identically
                            // whether invoked from the menu or programmatically.
                            const payload: {
                                v: number;
                                t: number;
                                s?: { n: { pattern: string } };
                            } = { v: parsed.value, t: CELL_TYPE_NUMBER };
                            if (parsed.pattern) payload.s = { n: { pattern: parsed.pattern } };
                            newValues.setValue(r, c, payload);
                            converted++;
                        }
                    }
                }
                if (converted === 0) return true; // nothing to do — silent no-op

                // Run as a redo/undo pair so Ctrl+Z reverts the conversion.
                // A bare syncExecuteCommand(mutation) applies the change but
                // never registers it on the undo stack (mutations are the
                // low-level layer; undo lives at the command layer). Mirror the
                // built-in TextToNumberCommand: build the inverse mutation from
                // the CURRENT state via SetRangeValuesUndoMutationFactory,
                // sequenceExecute the redo, then pushUndoRedo.
                const undoRedoService = accessor.get(IUndoRedoService) as {
                    pushUndoRedo: (item: {
                        unitID: string;
                        undoMutations: unknown[];
                        redoMutations: unknown[];
                    }) => void;
                };
                const setParams = { unitId, subUnitId, cellValue: newValues.getMatrix() };
                const redos = [{ id: SetRangeValuesMutation.id, params: setParams }];
                const undos = [
                    {
                        id: SetRangeValuesMutation.id,
                        params: (
                            SetRangeValuesUndoMutationFactory as (a: unknown, p: unknown) => unknown
                        )(accessor, setParams),
                    },
                ];
                const seq = sequenceExecute as (m: unknown[], cs: unknown) => { result: boolean };
                if (seq(redos, commandService).result) {
                    undoRedoService.pushUndoRedo({
                        unitID: unitId,
                        undoMutations: undos,
                        redoMutations: redos,
                    });
                    return true;
                }
                return false;
            },
        });

        // Repoint the EXISTING built-in "Text to Number" ribbon button at our
        // command. The built-in item id is `sheet.toolbar.text-to-number`,
        // placed at [DATA][OTHERS] in the ribbon schema. mergeMenu deep-merges
        // by key, replacing the `menuItemFactory` function reference for a key
        // that already exists — so returning a factory with the SAME id + label
        // but our `commandId` makes the visible button dispatch our generic
        // command instead of Univer's numeral-only one. (Adding a SEPARATE item
        // didn't render, and hiding the built-in + adding our own is more
        // fragile than reusing the button the ribbon already draws.)
        const menuManager = injector.get(IMenuManagerService) as {
            mergeMenu?: (schema: unknown) => void;
        };
        if (menuManager?.mergeMenu) {
            menuManager.mergeMenu({
                [RibbonPosition.DATA]: {
                    [RibbonDataGroup.OTHERS]: {
                        [BUILTIN_TEXT_TO_NUMBER_MENU_ID]: {
                            menuItemFactory: () => ({
                                id: BUILTIN_TEXT_TO_NUMBER_MENU_ID,
                                commandId: NOTESHEET_TEXT_TO_NUMBER_COMMAND_ID,
                                type: MenuItemType.BUTTON,
                                title: 'notesheet.textToNumber.title',
                                tooltip: 'notesheet.textToNumber.tooltip',
                            }),
                        },
                    },
                },
            });
        }
    } catch (e) {
        console.warn('[Notesheet] Text-to-Number command registration failed', e);
    }
}
