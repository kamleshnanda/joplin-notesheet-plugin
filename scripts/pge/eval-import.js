#!/usr/bin/env node
// PGE import smoke test — drives the plugin's REAL "Import .xlsx" toolbar
// path INSIDE the live Joplin webview (over CDP), the mirror of
// eval-export.js.
//
// WHY this exists: scripts/pge/import-fixture.sh runs xlsxBufferToSnapshot
// under NODE (main-process exceljs), so it can never catch a break in the
// WEBVIEW bundle's exceljs. M20 switched that bundle from exceljs's prebuilt
// dist/exceljs.min.js to bundling lib/ from source (webpack.config.js
// exceljsFromSourceResolve) — a change jest is blind to, because jest never
// goes through the webpack graph. This script feeds a file to the hidden
// <input type=file> the Import button clicks, accepts the "replace?"
// confirm(), and asserts handleImport() reaches "Imported." with no
// "[Notesheet] xlsx import failed" console error, then screenshots the
// webview.
//
// Usage:
//   PGE_NOTE_ID=<id> node scripts/pge/eval-import.js <feature-id> <path/to/file.xlsx>
// Writes screenshots/<feature-id>/import-<basename>-<utc>.png.
// Exit codes: 0 ok, 2 usage, 3 no playwright, 4 harness problem,
// 5 import failed in the webview.
//
// Lifecycle: same contract as eval-export.js — attach to an already-running
// Joplin (scripts/pge/launch-joplin.sh), never quit it, browser.close() in a
// finally only DETACHES.

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const EXIT_USAGE = 2;
const EXIT_NO_PLAYWRIGHT = 3;
const EXIT_HARNESS = 4;
const EXIT_IMPORT_FAILED = 5;

const FEATURE_ID = process.argv[2];
const XLSX_PATH = process.argv[3];
const NOTE_ID = process.env.PGE_NOTE_ID;
if (!FEATURE_ID || !XLSX_PATH || !NOTE_ID) {
    console.error('usage: PGE_NOTE_ID=<id> eval-import.js <feature-id> <file.xlsx>');
    process.exit(EXIT_USAGE);
}

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const OUT_DIR = path.join(REPO_ROOT, 'screenshots', FEATURE_ID);
const CDP_URL = `http://localhost:${process.env.PGE_CDP_PORT || '8315'}`;
// Must match src/editorView.tsx.
const FILE_INPUT_SELECTOR = '#notesheet-xlsx-input';
const STATUS_SELECTOR = '#notesheet-action-status';
const CANVAS_SELECTOR = 'canvas[id^="univer-sheet-main-canvas"]';

async function findWebviewFrame(page) {
    for (let attempt = 0; attempt < 50; attempt++) {
        const frame = page.frames().find((f) => /UserWebviewIndex\.html/.test(f.url()));
        if (frame) return frame;
        await page.waitForTimeout(200);
    }
    throw new Error('UserWebviewIndex frame did not appear (is the note a Notesheet?)');
}

async function main() {
    const absXlsx = path.resolve(XLSX_PATH);
    if (!fs.existsSync(absXlsx)) throw new Error(`no such file: ${absXlsx}`);
    fs.mkdirSync(OUT_DIR, { recursive: true });

    let chromium;
    try {
        ({ chromium } = require('playwright'));
    } catch {
        console.error('Playwright not installed (npm install --save-dev playwright).');
        process.exit(EXIT_NO_PLAYWRIGHT);
    }

    const browser = await chromium.connectOverCDP(CDP_URL);
    try {
        const pages = browser.contexts()[0]?.pages() ?? [];
        const page = pages.find((p) => /app\.asar\/index\.html$/.test(p.url()));
        if (!page) throw new Error('Joplin main window page not found over CDP');

        execFileSync('open', [`joplin://x-callback-url/openNote?id=${NOTE_ID}`], {
            stdio: ['ignore', 'ignore', 'pipe'],
            timeout: 5000,
        });
        const frame = await findWebviewFrame(page);
        await frame.waitForSelector(CANVAS_SELECTOR, { timeout: 15_000, state: 'attached' });

        const consoleErrors = [];
        page.on('console', (msg) => {
            if (msg.type() === 'error') consoleErrors.push(msg.text());
        });
        // handleImport() asks window.confirm() before replacing the workbook.
        page.on('dialog', (d) => void d.accept());

        console.error(`eval-import: importing ${path.basename(absXlsx)} via the webview`);
        await frame.setInputFiles(FILE_INPUT_SELECTOR, absXlsx);

        // setStatus() clears itself 3 s after the final message, so poll.
        const outcome = await frame.waitForFunction(
            (sel) => {
                const t = document.querySelector(sel)?.textContent || '';
                return t === 'Imported.' || t.startsWith('Import failed') ? t : null;
            },
            STATUS_SELECTOR,
            { timeout: 30_000, polling: 50 },
        );
        const status = await outcome.jsonValue();
        const importErrors = consoleErrors.filter((t) => t.includes('xlsx import failed'));

        // Give Univer a beat to paint the new workbook before the screenshot.
        await frame.waitForTimeout(1500);
        const utc = new Date().toISOString().replace(/[:.]/g, '-');
        const shot = path.join(
            OUT_DIR,
            `import-${path.basename(absXlsx, path.extname(absXlsx))}-${utc}.png`,
        );
        await page.screenshot({ path: shot });

        console.error(`eval-import: status="${status}" screenshot=${shot}`);
        if (status !== 'Imported.' || importErrors.length > 0) {
            for (const e of importErrors) console.error(`eval-import: console: ${e}`);
            process.exitCode = EXIT_IMPORT_FAILED;
        }
        console.log(shot);
    } finally {
        await browser.close();
    }
}

main().catch((e) => {
    console.error(`eval-import: ${e.message}`);
    process.exit(EXIT_HARNESS);
});
