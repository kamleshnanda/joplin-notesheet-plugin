import joplin from 'api';
import { MenuItemLocation, ToolbarButtonLocation } from 'api/types';
import { emptySnapshot, wrapSnapshot } from './snapshot';

const LOG = '[Notesheet]';

// Notesheet — a spreadsheet plugin for Joplin.
//
// M1 (re-scoped): registers a "New Spreadsheet" command that creates a
// regular Joplin note pre-populated with an empty Univer-shaped snapshot
// inside a fenced ```notesheet code block. The note opens in Joplin's
// normal markdown editor — the Univer-backed Custom Editor view that
// renders the snapshot as an interactive spreadsheet lands in M2.
joplin.plugins.register({
    onStart: async function () {
        try {
            await joplin.commands.register({
                name: 'newSpreadsheet',
                label: 'New Spreadsheet',
                // Joplin bundles Font Awesome 5; 'fa-table-cells' (FA 6) does
                // not render. Sticking with 'fa-table' for now. A custom inline
                // SVG (the original Notesheet brand mark) is a future polish task.
                iconName: 'fas fa-table',
                execute: async () => {
                    try {
                        const folder = await joplin.workspace.selectedFolder();
                        if (!folder) {
                            await joplin.views.dialogs.showMessageBox(
                                'Please select a notebook first.',
                            );
                            return;
                        }
                        const body = wrapSnapshot(emptySnapshot());
                        const note = await joplin.data.post(['notes'], null, {
                            parent_id: folder.id,
                            title: 'New Spreadsheet',
                            body,
                        });
                        await joplin.commands.execute('openNote', note.id);
                    } catch (error) {
                        console.error(LOG, 'newSpreadsheet failed:', error);
                        await joplin.views.dialogs.showMessageBox(
                            'Failed to create spreadsheet: ' + (error as Error).message,
                        );
                    }
                },
            });

            // EditorToolbar (above the editor body) is the more visible spot
            // for "create" actions; NoteToolbar (top-right) is also added for
            // discoverability since users may look in either location.
            await joplin.views.toolbarButtons.create(
                'notesheetNewEditorButton',
                'newSpreadsheet',
                ToolbarButtonLocation.EditorToolbar,
            );
            await joplin.views.toolbarButtons.create(
                'notesheetNewNoteButton',
                'newSpreadsheet',
                ToolbarButtonLocation.NoteToolbar,
            );

            await joplin.views.menuItems.create(
                'notesheetNewMenu',
                'newSpreadsheet',
                MenuItemLocation.Tools,
                { accelerator: 'CmdOrCtrl+Shift+S' },
            );

            console.info(LOG, 'plugin loaded');
        } catch (error) {
            console.error(LOG, 'onStart failed:', error);
        }
    },
});
