// Activate a named sheet inside the currently-loaded workbook so the
// canvas + float-DOM screenshot shows that sheet's content.
//
// Used by feature-3 of M17 — `MultiSheet.xlsx` opens onto its first
// sheet ("Data") which contains only the source range cells; the chart
// float-DOM lives on the second sheet ("Chart"). The PGE harness has
// no other way to switch sheets without running through Univer's
// command bus.
const playwright = require('playwright');

async function main() {
    const targetSheetName = process.argv[2];
    if (!targetSheetName) {
        console.error('usage: activate-sheet.js <sheetName>');
        process.exit(2);
    }
    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    try {
        const ctx = browser.contexts()[0];
        const pages = ctx.pages();
        let frame = null;
        for (let i = 0; i < 30 && !frame; i++) {
            for (const p of pages) {
                const c = p.frames().find(f => /UserWebviewIndex\.html/.test(f.url()));
                if (c) { frame = c; break; }
            }
            if (!frame) await pages[0].waitForTimeout(200);
        }
        if (!frame) throw new Error('no UserWebviewIndex frame found within 6s');
        const result = await frame.evaluate((name) => {
            const api = window.__notesheetUniverAPI;
            if (!api) return { error: 'no __notesheetUniverAPI on window' };
            const wb = api.getActiveWorkbook ? api.getActiveWorkbook() : null;
            if (!wb) return { error: 'no active workbook' };
            // FWorkbook on Univer 0.23 typically exposes getSheetBySheetName,
            // getSheets, and a setActiveSheet on the underlying unit. Pick
            // whichever exists at runtime.
            let sheet = null;
            if (wb.getSheets) {
                const sheets = wb.getSheets();
                for (const s of sheets) {
                    if (typeof s.getSheetName === 'function' && s.getSheetName() === name) { sheet = s; break; }
                    if (typeof s.getName === 'function' && s.getName() === name) { sheet = s; break; }
                }
            }
            if (!sheet && wb.getSheetByName) sheet = wb.getSheetByName(name);
            if (!sheet) return { error: `no sheet named "${name}"` };
            try {
                if (sheet.activate) sheet.activate();
                else if (wb.setActiveSheet) wb.setActiveSheet(sheet);
            } catch (e) {
                return { error: `sheet.activate threw: ${e.message}` };
            }
            // Confirm by reading what's now active.
            const active = wb.getActiveSheet ? wb.getActiveSheet() : null;
            const activeName = active && (active.getSheetName ? active.getSheetName() : (active.getName ? active.getName() : null));
            return { ok: true, activeName };
        }, targetSheetName);
        console.log(JSON.stringify(result, null, 2));
        if (result?.error) process.exit(1);
    } finally {
        await browser.close();
    }
}
main().catch((e) => { console.error(e); process.exit(1); });
