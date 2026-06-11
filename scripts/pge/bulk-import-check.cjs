// Bulk runtime smoke: import every fixture into Joplin and confirm the
// Notesheet editor (UserWebviewIndex frame + Univer canvas) mounts —
// i.e. the import didn't crash the live editor. Prints PASS/FAIL per
// fixture. Does NOT screenshot (that's for spot-checks); this is the
// broad no-crash gate across all fixtures.
const playwright = require('playwright');
const { execFileSync } = require('child_process');
const path = require('path');

const FIXTURES = process.argv.slice(2);
if (FIXTURES.length === 0) { console.error('usage: bulk-import-check.cjs <fixture.xlsx>...'); process.exit(2); }

async function importFixture(name) {
    const out = execFileSync('bash', [path.join(__dirname, 'import-fixture.sh'), name], { encoding: 'utf8' });
    return out.trim().split('\n').pop().trim();
}

async function main() {
    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    const results = [];
    try {
        const ctx = browser.contexts()[0];
        let page = null;
        for (const p of ctx.pages()) { try { if ((await p.title()) === 'Joplin') { page = p; break; } } catch {} }
        if (!page) { console.error('no editor page'); process.exit(1); }

        for (const fx of FIXTURES) {
            let noteId, ok = false, detail = '';
            try {
                noteId = await importFixture(fx);
                execFileSync('open', [`joplin://x-callback-url/openNote?id=${noteId}`]);
                await page.waitForTimeout(2500);
                // Did the Notesheet editor mount?
                const frame = page.frames().find((f) => /UserWebviewIndex\.html/.test(f.url()));
                if (frame) {
                    // Univer canvas present = rendered without throwing.
                    const hasCanvas = await frame.evaluate(() => !!document.querySelector('canvas[id^="univer-sheet-main-canvas"]')).catch(() => false);
                    ok = hasCanvas;
                    detail = hasCanvas ? 'canvas mounted' : 'frame present, no canvas';
                } else {
                    detail = 'no UserWebviewIndex frame';
                }
            } catch (e) {
                detail = 'import threw: ' + e.message.split('\n')[0];
            }
            results.push({ fx, ok, detail });
            console.log(`${ok ? 'PASS' : 'FAIL'}  ${fx}  — ${detail}`);
        }
    } finally { await browser.close(); }
    const failed = results.filter((r) => !r.ok);
    console.log(`\n${results.length - failed.length}/${results.length} fixtures mounted OK`);
    process.exit(failed.length === 0 ? 0 : 1);
}
main().catch((e) => { console.error(e); process.exit(1); });
