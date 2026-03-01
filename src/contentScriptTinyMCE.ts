/**
 * CodeMirror content script for Univer spreadsheet plugin.
 * 
 * In the Rich Text Editor, the markdown source is edited via CodeMirror.
 * This content script doesn't need to handle click events because:
 * - The Markdown viewer uses univerWebviewHandler.js for clicks
 * - The RTE preserves the original markdown via joplin-editable pattern
 * - CodeMirror shows raw markdown, not rendered HTML
 * 
 * This script is kept as a minimal stub to avoid registration errors.
 * VERSION: 2.2.0
 */

export default function(_context: any) {
    return {
        plugin: function(_CodeMirror: any) {
            // No-op: click handling is done in the Markdown viewer webview
            // via univerWebviewHandler.js, not in the CodeMirror editor.
        },
        codeMirrorResources: [],
        codeMirrorOptions: {},
        assets: function() {
            return [];
        },
    };
}
