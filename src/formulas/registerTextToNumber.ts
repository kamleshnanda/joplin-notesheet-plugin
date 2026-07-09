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
    CommandType,
    ObjectMatrix,
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
    SetNumfmtCommand,
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
                const numfmtValues: Array<{
                    row: number;
                    col: number;
                    pattern: string;
                    type: string;
                }> = [];
                let converted = 0;

                for (const rng of ranges) {
                    for (let r = rng.startRow; r <= rng.endRow; r++) {
                        for (let c = rng.startColumn; c <= rng.endColumn; c++) {
                            const cell = worksheet.getCellRaw(r, c);
                            if (!cell || cell.v == null) continue;
                            if (cell.t === CELL_TYPE_NUMBER) continue; // already numeric
                            const parsed = parseTextToNumber(String(cell.v));
                            if (!parsed) continue; // not a number → leave as text
                            newValues.setValue(r, c, {
                                v: parsed.value,
                                t: CELL_TYPE_NUMBER,
                            });
                            if (parsed.pattern) {
                                numfmtValues.push({
                                    row: r,
                                    col: c,
                                    pattern: parsed.pattern,
                                    type: 'unknown',
                                });
                            }
                            converted++;
                        }
                    }
                }
                if (converted === 0) return true; // nothing to do — silent no-op

                commandService.syncExecuteCommand(SetRangeValuesMutation.id, {
                    unitId,
                    subUnitId,
                    cellValue: newValues.getMatrix(),
                });
                if (numfmtValues.length) {
                    // High-level command: applies numfmt + fixes cell value-type
                    // + registers undo. `values` is keyed to the active sheet.
                    commandService.syncExecuteCommand((SetNumfmtCommand as { id: string }).id, {
                        values: numfmtValues,
                    });
                }
                return true;
            },
        });

        // Add our menu item under Data → Others (the built-in text-to-number
        // lives here too; it's hidden via the preset menu config). The ribbon
        // rebuilds reactively, so this late merge appears.
        const menuManager = injector.get(IMenuManagerService) as {
            mergeMenu?: (schema: unknown) => void;
        };
        if (menuManager?.mergeMenu) {
            menuManager.mergeMenu({
                [RibbonPosition.DATA]: {
                    [RibbonDataGroup.OTHERS]: {
                        [NOTESHEET_TEXT_TO_NUMBER_COMMAND_ID]: {
                            order: 1,
                            menuItemFactory: () => ({
                                id: NOTESHEET_TEXT_TO_NUMBER_COMMAND_ID,
                                commandId: NOTESHEET_TEXT_TO_NUMBER_COMMAND_ID,
                                type: MenuItemType.BUTTON,
                                title: 'notesheet.textToNumber.title',
                                tooltip: 'notesheet.textToNumber.tooltip',
                                icon: 'AutoNumberSingle',
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
