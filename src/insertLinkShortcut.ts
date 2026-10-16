// M19 feature-2 (BACKLOG E4) — host side of the insert-link shortcut.
// Why a separate command and not Cmd/Ctrl+K: see src/editor/insertLink.ts.
//
// Accelerator choice: `CmdOrCtrl+Alt+K` is unbound in Joplin's default keymap
// at the installed v3.7.21 (laurent22/joplin e41516e6,
// packages/lib/services/KeymapService.ts: neither `Option+Cmd+K` in the darwin
// table nor `Ctrl+Alt+K` in the default table; packages/app-desktop/gui/
// MenuBar.tsx has no literal K accelerator) and in every installed Univer
// shortcut (pinned in tests/m19InsertLinkShortcut.test.ts). Plugin menu items
// are listed in Tools → Options → Keyboard Shortcuts, so users can rebind it.

import { INSERT_LINK_MESSAGE } from './editor/insertLink';

export const INSERT_LINK_COMMAND = 'notesheetInsertLink';
export const INSERT_LINK_LABEL = 'Insert link in spreadsheet';
export const INSERT_LINK_ACCELERATOR = 'CmdOrCtrl+Alt+K';
export const INSERT_LINK_MENU_ID = 'notesheetInsertLinkMenu';

/** The subset of `joplin.views.editors` this module calls. */
export interface EditorViewsApi {
    isActive(handle: string): Promise<boolean>;
    isVisible(handle: string): Promise<boolean>;
    // Typed `void` in api/JoplinViewsEditor.d.ts, but plugin API calls are
    // proxied across the sandbox and return a Promise at runtime.
    postMessage(handle: string, message: unknown): void | Promise<void>;
}

/**
 * Post the insert-link message to the editor view the user is looking at.
 *
 * `enabledCondition: 'hasActivePluginEditor'` only says SOME plugin editor
 * view is active (Joplin's getActivePluginEditorView: `view.active`, not
 * `opened`), and Joplin creates one view per window. So check each known
 * handle and require it to be both active and visible: an active view the
 * user has toggled back to the markdown editor gets nothing. A handle whose
 * view is gone (isActive/isVisible throws) is dropped from the set. Like
 * Joplin's own getActivePluginEditorView, the first match wins.
 *
 * Returns the handle posted to, or null. Never throws.
 */
export async function postInsertLinkToActiveEditor(
    editors: EditorViewsApi,
    handles: Set<string>,
): Promise<string | null> {
    for (const handle of Array.from(handles)) {
        let target = false;
        try {
            target = (await editors.isActive(handle)) && (await editors.isVisible(handle));
        } catch {
            handles.delete(handle);
            continue;
        }
        if (!target) continue;
        try {
            await editors.postMessage(handle, INSERT_LINK_MESSAGE);
            return handle;
        } catch (e) {
            console.warn('[Notesheet] insert-link postMessage failed', e);
            return null;
        }
    }
    return null;
}
