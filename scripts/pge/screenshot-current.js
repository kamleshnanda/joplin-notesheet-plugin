// Open a note and screenshot the editor pane right where it is.
const playwright = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');

async function main() {
    const noteId = process.argv[2];
    const outPath = process.argv[3] || `/tmp/screenshot-${Date.now()}.png`;
    if (!noteId) { console.error('usage: screenshot-current.js <noteId> [outPath]'); process.exit(2); }

    execFileSync('open', [`joplin://x-callback-url/openNote?id=${noteId}`]);
    await new Promise((r) => setTimeout(r, 2000));

    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    try {
        const ctx = browser.contexts()[0];
        const pages = ctx.pages();
        let editorPage = null;
        for (const p of pages) {
            try {
                const t = await p.title();
                if (t === 'Joplin') { editorPage = p; break; }
            } catch {}
        }
        if (!editorPage) { console.error('no editor page'); process.exit(1); }
        // Wait a beat for chart to mount
        await editorPage.waitForTimeout(2500);
        await editorPage.screenshot({ path: outPath, fullPage: false });
        console.log(outPath);
    } finally { await browser.close(); }
}
main().catch((e) => { console.error(e); process.exit(1); });
