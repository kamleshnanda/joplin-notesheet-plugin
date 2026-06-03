// Idempotent: hide Joplin's sidebar and note list, and close any
// DevTools page exposed by Joplin's CDP endpoint. Called by
// prep-joplin-window.sh before evaluator screenshots.
//
// What this script does:
//   - Closes any CDP-exposed DevTools page. Joplin's main process
//     opens DevTools as a separate Electron window; closing the page
//     tells the main process to dispose it.
//   - Probes the renderer DOM to learn whether sidebar / note list
//     are currently visible.
//   - For each visible pane, asks AppleScript (via `osascript`) to
//     activate Joplin and send Cmd+Alt+S / Cmd+Alt+L. Playwright's
//     `keyboard.press()` over CDP does NOT reach Joplin's
//     Electron-accelerator-routed shortcuts — empirically the renderer
//     receives the keydown but the command system never fires. OS-
//     level keystrokes via System Events do work.
//
// Stable selectors (Joplin desktop, current build):
//   - `.rli-sideBar`  — resizable layout item wrapping the notebooks
//                       sidebar.
//   - `.note-list`    — note list pane wrapper.
//
// Joplin hides panes by setting `display: none` somewhere up the
// tree; `element.offsetParent === null` is the canonical visibility
// check (it returns null whenever any ancestor has `display: none`).
//
// Shortcuts (Joplin View menu):
//   - Cmd+Option+S → Toggle sidebar
//   - Cmd+Option+L → Toggle note list

const { execFileSync } = require('child_process');

const CDP_PORT = process.env.PGE_CDP_PORT || '8315';
const CDP_URL = `http://localhost:${CDP_PORT}`;

function sendKeystroke(letter) {
    // Activate Joplin so the keystroke routes to it, then send
    // Cmd+Alt+<letter>. The 100ms delay is empirical — without it,
    // System Events sometimes loses the modifier press.
    const script = `
tell application "Joplin" to activate
delay 0.1
tell application "System Events"
    keystroke "${letter}" using {command down, option down}
end tell
    `.trim();
    execFileSync('osascript', ['-e', script], { stdio: ['ignore', 'pipe', 'inherit'] });
}

(async () => {
    let chromium;
    try {
        chromium = require('playwright').chromium;
    } catch {
        console.error('prep-joplin-panes: playwright not installed (npm i --save-dev playwright)');
        process.exit(3);
    }

    const browser = await chromium.connectOverCDP(CDP_URL);
    try {
        const context = browser.contexts()[0];
        if (!context) throw new Error('CDP attach succeeded but no BrowserContext');

        // 1. Close any DevTools page. The CDP page list exposes them
        // as `devtools://...` URLs. Closing the page tells Joplin's
        // main process to dispose the DevTools window.
        for (const p of context.pages()) {
            let url = '';
            try { url = p.url(); } catch {}
            if (/^devtools:\/\//.test(url)) {
                console.error(`prep-joplin-panes: closing DevTools page (${url.slice(0, 80)}...)`);
                try { await p.close(); }
                catch (e) { console.error(`prep-joplin-panes: close failed: ${e.message}`); }
            }
        }

        // 2. Find the Joplin editor page (renderer).
        const editor = context.pages().find((p) => /Resources\/app\.asar\/index\.html/.test(p.url()));
        if (!editor) {
            console.error('prep-joplin-panes: no Joplin editor page found; skipping pane prep.');
            return;
        }

        // 3. Probe pane visibility. `offsetParent === null` is true
        // whenever any ancestor has `display: none` (canonical check).
        async function paneVisible(selector) {
            return await editor.evaluate((sel) => {
                const el = document.querySelector(sel);
                return !!el && el.offsetParent !== null;
            }, selector);
        }

        // 4. Toggle if needed. AppleScript keystroke reaches Joplin's
        // Electron-accelerator-routed shortcut where Playwright's
        // CDP-routed keyboard.press does not.
        async function toggleIfVisible(name, selector, key) {
            if (!(await paneVisible(selector))) {
                console.error(`prep-joplin-panes: ${name} already hidden`);
                return;
            }
            console.error(`prep-joplin-panes: hiding ${name} (${selector}) via Cmd+Option+${key}`);
            sendKeystroke(key);
            // Wait for the layout to settle. Joplin's Redux dispatch
            // + React rerender + paint typically lands within 200ms.
            const deadline = Date.now() + 3_000;
            while (Date.now() < deadline) {
                if (!(await paneVisible(selector))) return;
                await editor.waitForTimeout(50);
            }
            console.error(`prep-joplin-panes: ${name} did not collapse within 3s after shortcut.`);
        }

        await toggleIfVisible('sidebar', '.rli-sideBar', 's');
        await toggleIfVisible('note list', '.note-list', 'l');

        const after = await editor.evaluate(() => ({
            sidebarVisible: !!document.querySelector('.rli-sideBar')?.offsetParent,
            noteListVisible: !!document.querySelector('.note-list')?.offsetParent,
            inner: [window.innerWidth, window.innerHeight],
        }));
        console.error(`prep-joplin-panes: post-state ${JSON.stringify(after)}`);
    } finally {
        await browser.close();
    }
})().catch((e) => {
    console.error('prep-joplin-panes failed:', e.stack || e.message);
    process.exit(1);
});
