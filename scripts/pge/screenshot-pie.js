// One-shot: open a Notesheet note, scroll the Univer sheet right so a
// chart anchored off the initial viewport (e.g. 06's pie at ~col N) comes
// into view, then screenshot the editor webview frame. Used to visually
// verify the custom pie leader-line label plugin against Excel.
//
// usage: node screenshot-pie.js <noteId> [outPath] [scrollX]
const playwright = require('playwright');
const { execFileSync } = require('child_process');

async function main() {
    const noteId = process.argv[2];
    const outPath = process.argv[3] || `/tmp/pie-${Date.now()}.png`;
    const scrollX = parseInt(process.argv[4] || '700', 10);
    if (!noteId) { console.error('usage: screenshot-pie.js <noteId> [outPath] [scrollX]'); process.exit(2); }

    execFileSync('open', [`joplin://x-callback-url/openNote?id=${noteId}`]);
    await new Promise((r) => setTimeout(r, 2500));

    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    try {
        const ctx = browser.contexts()[0];
        let editorPage = null;
        for (const p of ctx.pages()) {
            try { if ((await p.title()) === 'Joplin') { editorPage = p; break; } } catch { /* ignore */ }
        }
        if (!editorPage) { console.error('no editor page'); process.exit(1); }

        // Find the UserWebviewIndex frame (where Univer mounts).
        await editorPage.waitForTimeout(1500);
        const frame = editorPage.frames().find((f) => /UserWebviewIndex\.html/.test(f.url()));
        if (!frame) { console.error('no UserWebviewIndex frame — note may not be a Notesheet'); process.exit(1); }

        // Scroll the Univer sheet so the off-screen chart shows. Use the
        // FUniver facade exposed on the frame (window.__notesheetUniverAPI):
        // select a far-right cell, which scrolls it into view. scrollX is
        // reinterpreted as the target column index (0-based).
        const targetCol = Math.max(1, Math.round(scrollX));
        const scrollResult = await frame.evaluate((col) => {
            const api = window.__notesheetUniverAPI;
            if (!api) return 'no api';
            const wb = api.getActiveWorkbook && api.getActiveWorkbook();
            const sheet = wb && (wb.getActiveSheet ? wb.getActiveSheet() : null);
            if (!sheet) return 'no sheet';
            try {
                // Selecting a far cell scrolls the viewport to it.
                const range = sheet.getRange(0, col);
                if (sheet.setActiveRange) sheet.setActiveRange(range);
                else if (sheet.setActiveSelection) sheet.setActiveSelection(range);
                if (sheet.scrollToCell) sheet.scrollToCell(0, col);
                return 'ok col=' + col;
            } catch (e) { return 'scroll threw: ' + e.message; }
        }, targetCol);
        console.error('scroll:', scrollResult);
        await editorPage.waitForTimeout(1800);

        await editorPage.screenshot({ path: outPath, fullPage: false });
        console.log(outPath);
    } finally { await browser.close(); }
}
main().catch((e) => { console.error(e); process.exit(1); });
