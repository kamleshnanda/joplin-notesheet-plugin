// M19 feature-2 (BACKLOG E4) — open Univer's insert-link panel from a
// plugin-owned shortcut.
//
// Cmd/Ctrl+K can't reach the Univer iframe: Joplin's default keymap binds it
// to `textLink` (laurent22/joplin @ v3.7.21, packages/lib/services/
// KeymapService.ts L38 darwin / L94 default), registered as an Electron
// app-menu accelerator (packages/app-desktop/gui/MenuBar.tsx L732) that
// `editorCommandDeclarations.ts` L25-28 enables for plugin editors via
// `hasActivePluginEditor`. The app menu consumes the keystroke, and Joplin
// then logs "Received command, but editor is gone textLink". So src/index.ts
// registers `notesheetInsertLink` with its own accelerator; when it runs it
// posts INSERT_LINK_MESSAGE to the active editor view, and editorView.tsx
// hands it to `openUniverInsertLink` here.
//
// Pure: no Univer import. editorView injects the command id from
// `InsertLinkShortcut.id` (@univerjs/preset-sheets-hyper-link re-exports
// @univerjs/sheets-hyper-link-ui), i.e. exactly the operation Univer's own
// Cmd/Ctrl+K shortcut fires (`sheet.operation.insert-hyper-link-toolbar` in
// 0.23). tests/m19InsertLinkShortcut.test.ts pins that against the installed
// build. Univer 1.x upgrade (BACKLOG F5): re-verify the export and the id.

export const INSERT_LINK_MESSAGE_TYPE = 'insertLink';

export const INSERT_LINK_MESSAGE: { readonly type: typeof INSERT_LINK_MESSAGE_TYPE } = {
    type: INSERT_LINK_MESSAGE_TYPE,
};

export function isInsertLinkMessage(m: unknown): boolean {
    return (
        !!m && typeof m === 'object' && (m as { type?: unknown }).type === INSERT_LINK_MESSAGE_TYPE
    );
}

interface CommandApi {
    executeCommand?: (id: string, params?: unknown) => unknown;
}

/**
 * Dispatch Univer's insert-link operation. Returns true when the command was
 * handed to Univer. Fails closed (returns false, never throws) when there is
 * no Univer instance, no `executeCommand`, or no command id.
 */
export function openUniverInsertLink(api: unknown, commandId: string): boolean {
    const exec = (api as CommandApi | null | undefined)?.executeCommand;
    if (typeof exec !== 'function' || !commandId) return false;
    try {
        const result = exec.call(api, commandId);
        if (result && typeof (result as Promise<unknown>).catch === 'function') {
            (result as Promise<unknown>).catch((e) => {
                console.warn('[Notesheet] insert-link command failed', e);
            });
        }
        return true;
    } catch (e) {
        console.warn('[Notesheet] insert-link command failed', e);
        return false;
    }
}
