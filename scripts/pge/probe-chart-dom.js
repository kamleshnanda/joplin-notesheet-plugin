// Open a note via joplin:// and probe the live float-DOM chart's
// rendered dimensions. Used to debug why imported charts appear
// oversized — compares the snapshot's stored transform with what
// the DOM actually paints.
const playwright = require('playwright');
const { execFileSync } = require('child_process');

async function main() {
    const noteId = process.argv[2];
    if (!noteId) {
        console.error('usage: probe-chart-dom.js <noteId>');
        process.exit(2);
    }
    execFileSync('open', [`joplin://x-callback-url/openNote?id=${noteId}`]);
    await new Promise((r) => setTimeout(r, 2000));

    const browser = await playwright.chromium.connectOverCDP('http://localhost:8315');
    try {
        const ctx = browser.contexts()[0];
        const pages = ctx.pages();
        let frame = null;
        for (let i = 0; i < 30 && !frame; i++) {
            for (const p of pages) {
                const c = p.frames().find((f) => /UserWebviewIndex\.html/.test(f.url()));
                if (c) {
                    frame = c;
                    break;
                }
            }
            if (!frame) await pages[0].waitForTimeout(200);
        }
        if (!frame) throw new Error('no frame');
        await frame.waitForTimeout(2000); // let chart-float-DOM mount

        const r = await frame.evaluate(() => {
            const out = { domChartContainers: [], snapshotTransforms: [] };
            // Look for float-DOM containers — Univer typically wraps
            // each in a div whose computed transform / left+top tells
            // the story.
            const candidates = Array.from(document.querySelectorAll('canvas')).filter(
                (c) =>
                    !c.id.includes('univer-sheet-main-canvas') &&
                    !c.id.includes('univer-sheet-engine-render'),
            );
            for (const c of candidates) {
                let el = c;
                let depth = 0;
                while (el && depth < 8) {
                    const rect = el.getBoundingClientRect();
                    out.domChartContainers.push({
                        depth,
                        tag: el.tagName,
                        cls: typeof el.className === 'string' ? el.className.slice(0, 80) : '',
                        id: el.id,
                        rect: {
                            left: Math.round(rect.left),
                            top: Math.round(rect.top),
                            width: Math.round(rect.width),
                            height: Math.round(rect.height),
                        },
                        canvasW: el.tagName === 'CANVAS' ? el.width : null,
                        canvasH: el.tagName === 'CANVAS' ? el.height : null,
                    });
                    el = el.parentElement;
                    depth++;
                }
                out.domChartContainers.push({ separator: '---' });
            }
            // Read snapshot from FUniver
            try {
                const api = window.__notesheetUniverAPI;
                const wb = api?.getActiveWorkbook?.();
                const snap = wb?.save?.() || wb?.getSnapshot?.();
                const r = (snap?.resources ?? []).find((rr) => rr?.name === 'SHEET_DRAWING_PLUGIN');
                if (r) {
                    const parsed = JSON.parse(r.data);
                    for (const sId of Object.keys(parsed)) {
                        for (const dId of Object.keys(parsed[sId].data)) {
                            const d = parsed[sId].data[dId];
                            out.snapshotTransforms.push({
                                drawingId: dId.slice(0, 30),
                                transform: d.transform,
                                sheetTransform: d.sheetTransform,
                            });
                        }
                    }
                }
            } catch (e) {
                out.snapErr = e.message;
            }
            return out;
        });
        console.log(JSON.stringify(r, null, 2));
    } finally {
        await browser.close();
    }
}
main().catch((e) => {
    console.error(e);
    process.exit(1);
});
