// One-shot: open a Notesheet note, find the Univer host frame, and screenshot
// the outer Univer root (#notesheet-univer-root) so native image drawings —
// which render as overlays/sprites on the canvas — are captured. Also probes
// the frame DOM for <img>/canvas drawing evidence and writes a sidecar.
const playwright = require('playwright');
const { execFileSync } = require('child_process');
const fs = require('fs');

async function main() {
    const noteId = process.argv[2];
    const outPath = process.argv[3] || `/tmp/img-${Date.now()}.png`;
    if (!noteId) {
        console.error('usage: capture-image-render.js <noteId> [outPath]');
        process.exit(2);
    }

    execFileSync('open', [`joplin://x-callback-url/openNote?id=${noteId}`]);
    await new Promise((r) => setTimeout(r, 2500));

    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    try {
        const ctx = browser.contexts()[0];
        const pages = ctx.pages();
        let editorPage = null;
        for (const p of pages) {
            try {
                if ((await p.title()) === 'Joplin') {
                    editorPage = p;
                    break;
                }
            } catch {}
        }
        if (!editorPage) {
            console.error('no editor page');
            process.exit(1);
        }
        await editorPage.waitForTimeout(2500);

        // Find the UserWebviewIndex frame that hosts Univer.
        let univerFrame = null;
        for (const f of editorPage.frames()) {
            try {
                const has = await f.evaluate(
                    () => !!document.querySelector('canvas[id^="univer-sheet-main-canvas"]'),
                );
                if (has) {
                    univerFrame = f;
                    break;
                }
            } catch {}
        }

        let probe = { univerFrameFound: !!univerFrame };
        if (univerFrame) {
            probe = await univerFrame.evaluate(() => {
                const canvases = Array.from(document.querySelectorAll('canvas')).map((c) => ({
                    id: c.id,
                    w: c.width,
                    h: c.height,
                }));
                const imgs = Array.from(document.querySelectorAll('img')).map((i) => ({
                    src: (i.src || '').slice(0, 40),
                    w: i.naturalWidth,
                    h: i.naturalHeight,
                }));
                const root = document.querySelector('#notesheet-univer-root');
                return {
                    univerFrameFound: true,
                    canvasCount: canvases.length,
                    canvases,
                    imgCount: imgs.length,
                    imgs,
                    rootBox: root ? { w: root.clientWidth, h: root.clientHeight } : null,
                };
            });
            // Screenshot the outer root if present, else the main canvas.
            const rootLoc = univerFrame.locator('#notesheet-univer-root');
            const canvasLoc = univerFrame.locator('canvas[id^="univer-sheet-main-canvas"]').first();
            try {
                if (await rootLoc.count()) await rootLoc.screenshot({ path: outPath });
                else await canvasLoc.screenshot({ path: outPath });
            } catch (e) {
                await editorPage.screenshot({ path: outPath, fullPage: false });
                probe.screenshotFallback = String(e).slice(0, 120);
            }
        } else {
            await editorPage.screenshot({ path: outPath, fullPage: false });
        }
        fs.writeFileSync(outPath.replace(/\.png$/, '.probe.json'), JSON.stringify(probe, null, 2));
        console.log(outPath);
        console.log(JSON.stringify(probe, null, 2));
    } finally {
        await browser.close();
    }
}
main().catch((e) => {
    console.error(e);
    process.exit(1);
});
