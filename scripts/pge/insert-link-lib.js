// M19 feature-2 (BACKLOG E4) — shared live-harness helpers for the
// "Insert link in spreadsheet" shortcut. Used by insert-link-probe.js and by
// eval-screenshot.js's `insertLinkPanel` region. See insert-link-probe.js for
// the signals and why each action is a real OS-level input.
//
// Child processes: execFileSync with argv arrays only (no shell). The only
// values interpolated into AppleScript are a numeric PID and constants.

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const CANVAS = 'canvas[id^="univer-sheet-main-canvas"]';
const MENU_LABEL = 'Insert link in spreadsheet';
const JOPLIN_LINK_PROMPT = 'Insert Hyperlink';
const KEY_CODE_K = 40; // macOS virtual key code for K (kVK_ANSI_K)
const CMD_OPTION = ['command down', 'option down'];
const CMD = ['command down'];

// ── Installed-source constants (resolved from the direct dependency) ──────
function packageRoot(name, fromDir) {
    let dir = path.dirname(require.resolve(name, { paths: [fromDir] }));
    for (;;) {
        const pj = path.join(dir, 'package.json');
        if (fs.existsSync(pj) && JSON.parse(fs.readFileSync(pj, 'utf8')).name === name) return dir;
        const up = path.dirname(dir);
        if (up === dir) throw new Error(`no package root for ${name}`);
        dir = up;
    }
}
function univerConstants() {
    const preset = packageRoot('@univerjs/preset-sheets-hyper-link', REPO_ROOT);
    const ui = packageRoot('@univerjs/sheets-hyper-link-ui', preset);
    const es = fs.readFileSync(path.join(ui, 'lib', 'es', 'index.js'), 'utf8');
    const locale = fs.readFileSync(path.join(ui, 'lib', 'es', 'locale', 'en-US.js'), 'utf8');
    const pick = (re, src, what) => {
        const m = re.exec(src);
        if (!m) throw new Error(`could not read ${what} from installed sheets-hyper-link-ui`);
        return m[1];
    };
    return {
        version: JSON.parse(fs.readFileSync(path.join(ui, 'package.json'), 'utf8')).version,
        openPanelId: pick(
            /const OpenHyperLinkEditPanelOperation = \{\s*type: CommandType\.OPERATION,\s*id: "([^"]+)"/,
            es,
            'OpenHyperLinkEditPanelOperation.id',
        ),
        closePopupId: pick(
            /const CloseHyperLinkPopupOperation = \{\s*type: CommandType\.OPERATION,\s*id: "([^"]+)"/,
            es,
            'CloseHyperLinkPopupOperation.id',
        ),
        toolbarOpId: pick(
            /const InsertHyperLinkToolbarOperation = \{\s*type: CommandType\.OPERATION,\s*id: "([^"]+)"/,
            es,
            'InsertHyperLinkToolbarOperation.id',
        ),
        linkPlaceholder: pick(
            /linkPlaceholder:\s*"([^"]+)"/,
            locale,
            'hyperLink.form.linkPlaceholder',
        ),
    };
}

// ── Process guard: only ever drive the DEV-profile Joplin ─────────────────
function joplinMainProcesses() {
    let out = '';
    try {
        out = execFileSync('pgrep', ['-fl', 'Joplin.app/Contents/MacOS/Joplin'], {
            encoding: 'utf8',
        });
    } catch {
        return [];
    }
    return out
        .split('\n')
        .filter(Boolean)
        .map((l) => ({ pid: Number(l.split(' ')[0]), args: l }));
}
// Fails closed: throws (err.code 'NON_DEV_JOPLIN') if any non-dev Joplin is
// running, so no keystroke can ever land in the operator's real profile.
function devJoplinPid() {
    const procs = joplinMainProcesses();
    const nonDev = procs.filter((p) => !/--env dev\b/.test(p.args));
    if (nonDev.length) {
        const e = new Error('a non-dev Joplin is running; refusing to send keystrokes');
        e.code = 'NON_DEV_JOPLIN';
        throw e;
    }
    const dev = procs.filter((p) => /--env dev\b/.test(p.args));
    if (dev.length !== 1) throw new Error(`expected 1 dev Joplin, found ${dev.length}`);
    return dev[0].pid;
}

// ── System Events helpers (scoped to the dev PID) ─────────────────────────
function osa(script) {
    return execFileSync('osascript', ['-e', script], {
        encoding: 'utf8',
        timeout: 10_000,
        stdio: ['ignore', 'pipe', 'pipe'], // errors surface via the thrown Error
    }).trim();
}
function bringToFront(pid) {
    osa(
        `tell application "System Events" to set frontmost of (first process whose unix id is ${pid}) to true`,
    );
}
function isFrontmost(pid) {
    return (
        osa(
            `tell application "System Events" to get unix id of (first process whose frontmost is true)`,
        ) === String(pid)
    );
}
function pressK(pid, modifiers) {
    bringToFront(pid);
    osa(
        `tell application "System Events" to key code ${KEY_CODE_K} using {${modifiers.join(', ')}}`,
    );
}
function menuItemRef(pid) {
    return `menu item "${MENU_LABEL}" of menu "Tools" of menu bar 1 of (first process whose unix id is ${pid})`;
}
function menuItemInfo(pid) {
    let raw;
    try {
        raw = osa(
            `tell application "System Events" to get {enabled, value of attribute "AXMenuItemCmdChar", value of attribute "AXMenuItemCmdModifiers"} of ${menuItemRef(pid)}`,
        );
    } catch (e) {
        // Item absent (e.g. a `main` baseline build without the command).
        return { present: false, enabled: false, error: String(e.message).slice(0, 200) };
    }
    const [enabled, cmdChar, mods] = raw.split(',').map((s) => s.trim());
    // AXMenuItemCmdModifiers: bit0 Shift, bit1 Option, bit2 Control, bit3 "no Command".
    const m = Number(mods);
    return {
        present: true,
        enabled: enabled === 'true',
        cmdChar,
        modifiers: m,
        accelerator: Number.isFinite(m)
            ? [
                  m & 8 ? null : 'Cmd',
                  m & 4 ? 'Ctrl' : null,
                  m & 2 ? 'Option' : null,
                  m & 1 ? 'Shift' : null,
                  cmdChar,
              ]
                  .filter(Boolean)
                  .join('+')
            : null,
    };
}
async function menuEnabledSettled(page, pid, want) {
    // Joplin re-evaluates when-clauses on state change; poll briefly.
    let info = menuItemInfo(pid);
    for (let i = 0; i < 15 && info.enabled !== want; i++) {
        await page.waitForTimeout(200);
        info = menuItemInfo(pid);
    }
    return info;
}
function clickMenu(pid) {
    bringToFront(pid);
    osa(`tell application "System Events" to click ${menuItemRef(pid)}`);
}

// ── CDP helpers ───────────────────────────────────────────────────────────
function openNoteUrl(id) {
    execFileSync('open', [`joplin://x-callback-url/openNote?id=${id}`], {
        stdio: ['ignore', 'ignore', 'pipe'],
        timeout: 5000,
    });
}
async function openNotesheet(page, id) {
    openNoteUrl(id);
    await page.waitForTimeout(1500);
    let frame = null;
    for (let i = 0; i < 50 && !frame; i++) {
        frame = page.frames().find((f) => /UserWebviewIndex\.html/.test(f.url())) || null;
        if (!frame) await page.waitForTimeout(200);
    }
    if (!frame) throw new Error('UserWebviewIndex frame did not appear');
    await frame.waitForSelector(CANVAS, { timeout: 15_000, state: 'attached' });
    await frame.waitForFunction(
        (s) => {
            const c = document.querySelector(s);
            return c && c.width > 0 && c.height > 0 && !!window.__notesheetUniverAPI;
        },
        CANVAS,
        { timeout: 10_000, polling: 100 },
    );
    await frame.waitForTimeout(800);
    return frame;
}

async function installCommandLog(frame, ids) {
    await frame.evaluate((watch) => {
        const api = window.__notesheetUniverAPI;
        if (window.__linkProbe && window.__linkProbe.api === api) {
            window.__linkProbe.log.length = 0;
            return;
        }
        const log = [];
        const d = api.addEvent(api.Event.CommandExecuted, (e) => {
            if (watch.includes(e.id)) log.push(e.id);
        });
        window.__linkProbe = { api, log, dispose: () => d.dispose() };
    }, ids);
}

async function webviewState(frame, placeholder) {
    return frame.evaluate((ph) => {
        const api = window.__notesheetUniverAPI;
        const ws = api.getActiveWorkbook().getActiveSheet();
        const r = ws.getActiveRange();
        const input = Array.from(document.querySelectorAll('input')).find(
            (i) => i.getAttribute('placeholder') === ph,
        );
        const rect = input ? input.getBoundingClientRect() : null;
        const root = input ? input.closest('.univer-rounded-xl') : null;
        const rootRect = root ? root.getBoundingClientRect() : null;
        return {
            activeRange: r ? r.getA1Notation() : null,
            panelOpen: !!(rect && rect.width > 0 && rect.height > 0),
            panelInputFocused: !!input && document.activeElement === input,
            panelRoot: rootRect
                ? {
                      tag: root.tagName,
                      className: String(root.className).slice(0, 200),
                      rect: {
                          x: Math.round(rootRect.x),
                          y: Math.round(rootRect.y),
                          w: Math.round(rootRect.width),
                          h: Math.round(rootRect.height),
                      },
                  }
                : null,
            commandLog: window.__linkProbe ? window.__linkProbe.log.slice() : null,
        };
    }, placeholder);
}

async function joplinPromptVisible(page) {
    // Joplin's dialogs.prompt renders a `smalltalk` dialog in the main window:
    // <div class="smalltalk">…<div class="content-area">Insert Hyperlink<input …>
    return page.evaluate(
        (text) =>
            Array.from(document.querySelectorAll('.smalltalk .content-area')).some(
                (el) =>
                    (el.textContent || '').includes(text) && el.getBoundingClientRect().width > 0,
            ),
        JOPLIN_LINK_PROMPT,
    );
}

async function cancelJoplinPrompt(page) {
    if (!(await joplinPromptVisible(page))) return true;
    await page.click('.smalltalk [data-name="js-cancel"]');
    await page.waitForTimeout(300);
    return !(await joplinPromptVisible(page));
}

// Poll for an outcome instead of a fixed sleep (the accelerator → plugin →
// webview round trip measured ~300 ms; allow up to 2.5 s).
async function waitFor(page, fn, ms = 2500) {
    const t0 = Date.now();
    for (;;) {
        if (await fn()) return Date.now() - t0;
        if (Date.now() - t0 > ms) return null;
        await page.waitForTimeout(100);
    }
}

async function closePanel(frame, closeId) {
    // Harness cleanup only (not the action under test): reset the start state.
    await frame.evaluate((id) => window.__notesheetUniverAPI.executeCommand(id), closeId);
    // Let Univer finish closing before the next case (back-to-back presses
    // within ~400 ms of a programmatic close were seen to miss).
    await frame.waitForTimeout(500);
}

// Focus modes for the start state:
//   canvas   — the Univer grid has keyboard focus (B2 active);
//   editing  — B2 is in cell-edit mode (typed "ab");
//   blurred  — keyboard focus is outside the editor iframe (main window body).
async function prepareB2(page, frame, pid, k, mode = 'canvas') {
    await cancelJoplinPrompt(page);
    bringToFront(pid);
    await frame
        .locator(CANVAS)
        .first()
        .click({ position: { x: 300, y: 200 } });
    await frame.waitForTimeout(250);
    await frame.evaluate(() =>
        window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet().getRange('B2').activate(),
    );
    await frame.waitForTimeout(250);
    if (mode === 'editing') {
        await page.keyboard.type('ab');
        await frame.waitForTimeout(300);
    } else if (mode === 'blurred') {
        await page.evaluate(() => document.activeElement && document.activeElement.blur());
        await frame.waitForTimeout(250);
    }
    await installCommandLog(frame, [k.openPanelId, k.toolbarOpId]);
    const before = await webviewState(frame, k.linkPlaceholder);
    const promptBefore = await joplinPromptVisible(page);
    const front = isFrontmost(pid);
    const mainActive = await page.evaluate(() =>
        document.activeElement ? document.activeElement.tagName : null,
    );
    const focusOk = mode === 'blurred' ? mainActive !== 'IFRAME' : mainActive === 'IFRAME';
    return {
        focusMode: mode,
        before,
        promptBefore,
        frontmost: front,
        mainActiveElement: mainActive,
        startOk:
            before.activeRange === 'B2' && !before.panelOpen && !promptBefore && front && focusOk,
    };
}

module.exports = {
    CANVAS,
    CMD,
    CMD_OPTION,
    JOPLIN_LINK_PROMPT,
    MENU_LABEL,
    bringToFront,
    cancelJoplinPrompt,
    clickMenu,
    closePanel,
    devJoplinPid,
    installCommandLog,
    isFrontmost,
    joplinPromptVisible,
    menuEnabledSettled,
    menuItemInfo,
    openNoteUrl,
    openNotesheet,
    prepareB2,
    pressK,
    univerConstants,
    waitFor,
    webviewState,
};
