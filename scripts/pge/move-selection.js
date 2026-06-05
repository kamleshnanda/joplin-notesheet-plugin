// Click off A1 so the active-cell selection moves elsewhere — keeps A1
// header text rendering cleanly without an edit cursor overlay.
const playwright = require('playwright');
async function main() {
    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    try {
        const ctx = browser.contexts()[0];
        const pages = ctx.pages();
        let frame = null;
        for (let i = 0; i < 20 && !frame; i++) {
            for (const p of pages) {
                const c = p.frames().find(f => /UserWebviewIndex\.html/.test(f.url()));
                if (c) { frame = c; break; }
            }
            if (!frame) await pages[0].waitForTimeout(200);
        }
        if (!frame) throw new Error('no frame');
        const result = await frame.evaluate(() => {
            const api = window.__notesheetUniverAPI;
            const wb = api.getActiveWorkbook();
            const ws = wb.getActiveSheet();
            // Move active cell to N1 (off-screen-ish, blank, far from CF)
            const r = ws.getRange('N1');
            r.activate();
            return { ok: true };
        });
        console.log(JSON.stringify(result, null, 2));
    } finally { await browser.close(); }
}
main().catch(e => { console.error(e); process.exit(1); });
