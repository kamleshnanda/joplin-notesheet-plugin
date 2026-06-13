// Data-loss regression (M18): editing a Notesheet note in Joplin's Rich
// Text (TinyMCE) editor and saving destroyed the `notesheet v=1` fence —
// TinyMCE serialized our rendered <table> back to a plain GFM table,
// obliterating the snapshot JSON (styles, charts, formulas, everything).
//
// Fix: wrap the rendered output in Joplin's documented `joplin-editable`
// container with a hidden `joplin-source` element carrying the ORIGINAL
// fence so Joplin's HTML→Markdown converter (turndown) reconstructs the
// fence verbatim instead of re-serializing the table. This is the same
// pattern every built-in Joplin renderer (mermaid/katex/fountain) uses.
//
// References (laurent22/joplin):
//   - ContentScriptType API docs, "Supporting the Rich Text Editor"
//   - packages/renderer/MdToHtml/rules/fountain.ts
//   - packages/turndown/src/commonmark-rules.js  (joplinSourceBlock rule)

import { renderFenceToken, renderNotesheetSnapshot } from '../src/contentScripts/notesheetRenderer';

// Minimal one-sheet snapshot body (the JSON that lives inside the fence).
const SNAPSHOT_JSON = JSON.stringify({
    sheetOrder: ['sheet-1'],
    sheets: {
        'sheet-1': {
            id: 'sheet-1',
            name: 'Sheet1',
            cellData: { 0: { 0: { v: 'Quarter' }, 1: { v: 'Sales' } } },
            rowCount: 5,
            columnCount: 3,
        },
    },
    styles: {},
});

// The exact fence Joplin stores (mirror of src/snapshot.ts wrapSnapshot):
// ```notesheet v=1\n<json>\n```
const FENCE_OPEN = '```notesheet v=1\n';
const FENCE_CLOSE = '\n```';

describe('M18 — Rich Text editor fence round-trip (data-loss fix)', () => {
    test('rendered output is wrapped in a joplin-editable container', () => {
        const html = renderFenceToken({ info: 'notesheet v=1', content: SNAPSHOT_JSON })!;
        expect(html).not.toBeNull();
        // The whole block must carry joplin-editable so TinyMCE treats it
        // as atomic (noneditable_class) and turndown uses the source block.
        expect(html).toMatch(/class="[^"]*\bjoplin-editable\b[^"]*"/);
    });

    test('contains a hidden joplin-source block carrying the ORIGINAL fence body', () => {
        const html = renderFenceToken({ info: 'notesheet v=1', content: SNAPSHOT_JSON })!;
        // A hidden joplin-source element exists.
        expect(html).toMatch(/<pre[^>]*\bclass="[^"]*\bjoplin-source\b[^"]*"[^>]*>/);
        expect(html).toMatch(/<pre[^>]*\bhidden\b/);
        // It declares the language so turndown round-trips correctly.
        expect(html).toContain('data-joplin-language="notesheet"');
    });

    test('source-open / source-close reconstruct the EXACT fence delimiters', () => {
        const html = renderFenceToken({ info: 'notesheet v=1', content: SNAPSHOT_JSON })!;
        // Newlines are encoded as &#10; inside the attribute values.
        const openAttr = (html.match(/data-joplin-source-open="([^"]*)"/) ?? [])[1];
        const closeAttr = (html.match(/data-joplin-source-close="([^"]*)"/) ?? [])[1];
        expect(openAttr).toBeDefined();
        expect(closeAttr).toBeDefined();
        // Decode &#10; → \n and assert they equal the canonical delimiters.
        const decode = (s: string) => s.replace(/&#10;/g, '\n').replace(/&amp;/g, '&');
        expect(decode(openAttr!)).toBe(FENCE_OPEN);
        expect(decode(closeAttr!)).toBe(FENCE_CLOSE + '\n');
    });

    test('the original JSON body is preserved verbatim in the source block', () => {
        const html = renderFenceToken({ info: 'notesheet v=1', content: SNAPSHOT_JSON })!;
        // The hidden source <pre> body, HTML-unescaped, must equal the
        // original snapshot JSON exactly — lossless reconstruction.
        const m = html.match(/<pre[^>]*class="[^"]*joplin-source[^"]*"[^>]*>([\s\S]*?)<\/pre>/);
        expect(m).not.toBeNull();
        const unescape = (s: string) =>
            s
                .replace(/&lt;/g, '<')
                .replace(/&gt;/g, '>')
                .replace(/&quot;/g, '"')
                .replace(/&#39;/g, "'")
                .replace(/&amp;/g, '&');
        expect(unescape(m![1])).toBe(SNAPSHOT_JSON);
    });

    test('the rendered table is still present (preview/export unaffected)', () => {
        const html = renderFenceToken({ info: 'notesheet v=1', content: SNAPSHOT_JSON })!;
        expect(html).toContain('notesheet-table');
        expect(html).toContain('Quarter');
        // The joplin-source block precedes the visible render so the source
        // is found first by turndown.
        expect(html.indexOf('joplin-source')).toBeLessThan(html.indexOf('notesheet-table'));
    });

    test('renderNotesheetSnapshot WITHOUT a fence body still renders (export path)', () => {
        // The HTML/PDF export path calls renderNotesheetSnapshot directly
        // with just the JSON; it should still produce the table. The
        // source-fence wrapper is only added when the original fence is
        // known (renderFenceToken passes it).
        const html = renderNotesheetSnapshot(SNAPSHOT_JSON)!;
        expect(html).toContain('notesheet-table');
    });
});
