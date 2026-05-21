import joplin from 'api';
import { MenuItemLocation, ToolbarButtonLocation } from 'api/types';
import { emptySnapshot, extractSnapshot, isNotesheetBody, wrapSnapshot } from './snapshot';

const LOG = '[Notesheet]';

joplin.plugins.register({
    onStart: async function () {
        try {
            // ── Custom Editor: Univer spreadsheet view ──
            // When the active note's body matches the notesheet fence sentinel,
            // this editor takes over and shows the Univer spreadsheet UI.
            await joplin.views.editors.register('notesheetEditor', {
                onActivationCheck: async (event) => {
                    const note = await joplin.data.get(['notes', event.noteId], { fields: ['body'] });
                    return isNotesheetBody(note?.body);
                },
                onSetup: async (handle) => {
                    let viewReady = false;

                    await joplin.views.editors.setHtml(handle, `
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <meta charset="utf-8">
                            <style>
                                html, body { margin: 0; padding: 0; width: 100%; height: 100%; overflow: hidden; }
                                #notesheet-univer-root { position: absolute; inset: 0; }
                            </style>
                        </head>
                        <body>
                            <div id="notesheet-univer-root"></div>
                        </body>
                        </html>
                    `);
                    await joplin.views.editors.addScript(handle, './editorView.js');

                    // Message bridge: editor view ←→ plugin process.
                    await joplin.views.editors.onMessage(handle, async (msg: any) => {
                        if (!msg || typeof msg !== 'object') return;

                        if (msg.type === 'ready') {
                            viewReady = true;
                            const note = await joplin.workspace.selectedNote();
                            if (!note) return;
                            const fullNote = await joplin.data.get(['notes', note.id], { fields: ['id', 'body'] });
                            const result = extractSnapshot(fullNote?.body);
                            if (result.ok) {
                                await joplin.views.editors.postMessage(handle, {
                                    type: 'load',
                                    snapshot: result.snapshot,
                                });
                            }
                            return;
                        }

                        if (msg.type === 'save' && msg.snapshot) {
                            const note = await joplin.workspace.selectedNote();
                            if (!note) return;
                            const body = wrapSnapshot(msg.snapshot);
                            await joplin.views.editors.saveNote(handle, {
                                noteId: note.id,
                                body,
                            });
                            return;
                        }
                    });

                    // External note changes (e.g. sync) — reload into editor.
                    // Guard: only send if the webview has signalled 'ready';
                    // otherwise postMessage fires before the handler is wired
                    // and Joplin logs a harmless but noisy warning.
                    await joplin.views.editors.onUpdate(handle, async (event) => {
                        if (!viewReady) return;
                        const result = extractSnapshot(event.newBody);
                        if (result.ok) {
                            await joplin.views.editors.postMessage(handle, {
                                type: 'load',
                                snapshot: result.snapshot,
                            });
                        }
                    });
                },
            });

            // ── "New Spreadsheet" command ──
            await joplin.commands.register({
                name: 'newSpreadsheet',
                label: 'New Spreadsheet',
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
                        // The Custom Editor activation check runs async after
                        // the note opens. Wait for it to complete before asking
                        // Joplin to show our editor. Retry a few times since
                        // the timing varies by machine load.
                        for (let attempt = 0; attempt < 10; attempt++) {
                            await new Promise(r => setTimeout(r, 300));
                            try {
                                await joplin.commands.execute('showEditorPlugin');
                                break;
                            } catch { /* retry */ }
                        }
                    } catch (error) {
                        console.error(LOG, 'newSpreadsheet failed:', error);
                        await joplin.views.dialogs.showMessageBox(
                            'Failed to create spreadsheet: ' + (error as Error).message,
                        );
                    }
                },
            });

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
