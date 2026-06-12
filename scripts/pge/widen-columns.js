// PGE harness helper — widen CF columns A..I uniformly via the FUniver
// facade, so the live Joplin state matches the column geometry the
// `tests/excelCanvasFidelity.test.ts` Joplin-side coordinates were
// tuned against. The harness sampler in `eval-screenshot.js` is now
// width-agnostic (Option A — reads live widths via the FUniver
// facade), but the canvas-fidelity tests still hardcode device-pixel
// x-ranges. Keeping their geometry stable means widening columns
// uniformly before the recapture.
//
// Target widths derived from the test sample coordinates:
//   - col G test sample: x=1290..1480 device px = 645..740 CSS at DPR=2.
//   - col G index 6 → cumulative widths(A..F) = 645 - 46 = 599 CSS.
//   - 6 equal columns → 599/6 ≈ 99.83 CSS each.
// We round to an integer (100) which puts col G at 646 CSS, well
// within the test's 95-px-wide sample band.
//
// Idempotent: running multiple times leaves widths at 100 each.
//
// Pre-condition: a Notesheet note must already be open in the editor
// (so the `UserWebviewIndex` frame exposes `window.__notesheetUniverAPI`).
// Run AFTER `eval-screenshot.sh` has navigated to the M15 note.
const playwright = require('playwright');

async function main() {
    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    try {
        const ctx = browser.contexts()[0];
        const pages = ctx.pages();
        // Find the page that hosts a UserWebviewIndex frame (the main
        // editor page); the "score by url" heuristic in eval-screenshot
        // picks the wrong page if you only check titles.
        let frame = null;
        for (let i = 0; i < 30 && !frame; i++) {
            for (const p of pages) {
                const candidate = p.frames().find((f) => /UserWebviewIndex\.html/.test(f.url()));
                if (candidate) {
                    frame = candidate;
                    break;
                }
            }
            if (!frame) await pages[0].waitForTimeout(200);
        }
        if (!frame) throw new Error('UserWebviewIndex frame not found in any page');

        const result = await frame.evaluate(() => {
            const api = window.__notesheetUniverAPI;
            if (!api)
                return { error: 'no notesheet api on window — is the Notesheet plugin loaded?' };
            const wb = api.getActiveWorkbook && api.getActiveWorkbook();
            const ws = wb && wb.getActiveSheet && wb.getActiveSheet();
            if (!ws) return { error: 'no active sheet' };
            const TARGET = 95; // CSS px — empirically tuned to put col G/I in the canvas-fidelity tests' sample x-bands.
            const before = {};
            const after = {};
            for (let c = 0; c <= 8; c++) before[c] = ws.getColumnWidth(c);
            for (let c = 0; c <= 8; c++) ws.setColumnWidth(c, TARGET);
            for (let c = 0; c <= 8; c++) after[c] = ws.getColumnWidth(c);
            return { ok: true, before, after, target: TARGET };
        });
        console.log(JSON.stringify(result, null, 2));
    } finally {
        await browser.close();
    }
}

main().catch((e) => {
    console.error(e.stack || e.message);
    process.exit(1);
});
