// Notesheet content script.
//
// M0: stripped to a no-op. The v0 markdown-it plugin that rendered
// `univer-sheet://` links has been removed. M2 will reintroduce a
// content script that renders read-only HTML table previews of
// linked spreadsheet notes inside regular markdown notes.

export default function (_context: unknown) {
    return {
        plugin: function (_markdownIt: unknown, _options: unknown) {
            // No-op for M0.
        },
        assets: function () {
            return [];
        },
    };
}
