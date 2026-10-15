#!/usr/bin/env node
// M19 feature-2 (BACKLOG E4) live probe: "Insert link in spreadsheet".
//
// Every action under test is a REAL OS-level input, never executeCommand:
//   - keystroke path: System Events `key code 40` (K) with command+option,
//     sent to the frontmost Joplin (the dev-profile process, by PID);
//   - menu path: System Events clicks Tools → "Insert link in spreadsheet";
//   - Cmd/Ctrl+K regression: System Events `key code 40` with command only.
// CDP (Playwright connectOverCDP) is used only to position the start state
// (focus the Univer canvas, activate B2) and to READ the outcome.
//
// Signals (all read live, none hard-coded where the installed build has them):
//   - Univer panel: an <input placeholder=<hyperLink.form.linkPlaceholder>>
//     that is laid out (non-zero rect) inside the Notesheet webview. The
//     placeholder text is read from the installed sheets-hyper-link-ui en-US
//     locale. The panel root is that input's closest `.univer-rounded-xl`
//     ancestor (CellLinkEdit's root <div>; it has no id or data attribute).
//   - Univer command log: CommandExecuted for the open-edit-panel operation
//     (id read from the installed build).
//   - Joplin markdown link dialog: the `dialogs.prompt(_('Insert Hyperlink'))`
//     prompt from packages/app-desktop/gui/NoteEditor/NoteBody/CodeMirror/v6/
//     useEditorCommands.ts L83-84 (v3.7.21), found by its text in the main page.
//   - Menu item enabled flag + accelerator, via System Events AX attributes.
//
// Every case checks its START state first (startOk). A case whose start state
// is wrong is reported INVALID and counts as a failure, never as a pass.
//
// Shared helpers (process guard, System Events, DOM signals): insert-link-lib.js.
//
// Usage (Joplin dev profile running with CDP, see launch-joplin.sh):
//   PGE_NOTE_ID=<empty notesheet id> PGE_MD_NOTE_ID=<plain markdown note id> \
//     node scripts/pge/insert-link-probe.js <label>
// Output: screenshots/feature-2-m19-insert-link-alt-shortcut/probe-<label>-<utc>.json
// Exit: 0 all scored cases pass, 2 usage, 3 refused (non-dev Joplin running),
//       4 harness problem, 5 at least one case failed or was invalid.
// macOS only (System Events); on other platforms it exits 2.

const fs = require('fs');
const path = require('path');
const {
    CMD,
    CMD_OPTION,
    JOPLIN_LINK_PROMPT,
    cancelJoplinPrompt,
    clickMenu,
    closePanel,
    devJoplinPid,
    isFrontmost,
    joplinPromptVisible,
    menuEnabledSettled,
    openNoteUrl,
    openNotesheet,
    prepareB2,
    pressK,
    bringToFront,
    univerConstants,
    waitFor,
    webviewState,
} = require('./insert-link-lib');

const LABEL = process.argv[2];
const NOTE_ID = process.env.PGE_NOTE_ID;
const MD_NOTE_ID = process.env.PGE_MD_NOTE_ID;
if (!LABEL || !NOTE_ID || !MD_NOTE_ID || process.platform !== 'darwin') {
    console.error(
        'usage (macOS): PGE_NOTE_ID=<notesheet id> PGE_MD_NOTE_ID=<markdown note id> insert-link-probe.js <label>',
    );
    process.exit(2);
}

const FEATURE_ID = 'feature-2-m19-insert-link-alt-shortcut';
const REPO_ROOT = path.resolve(__dirname, '..', '..');
const OUT_DIR = path.join(REPO_ROOT, 'screenshots', FEATURE_ID);
const CDP_URL = `http://localhost:${process.env.PGE_CDP_PORT || '8315'}`;

async function cleanupEditing(page, frame, mode) {
    if (mode !== 'editing') return;
    await page.keyboard.press('Escape');
    await frame.waitForTimeout(200);
    await frame.evaluate(() =>
        window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet().getRange('B2').clear(),
    );
}

// One notesheet case: real keystroke (or menu click) from a given focus mode.
async function notesheetCase(page, frame, pid, k, consoleMsgs, c) {
    const prep = await prepareB2(page, frame, pid, k, c.mode);
    const mark = consoleMsgs.length;
    const r = { id: c.id, action: c.action, ...prep };
    if (prep.startOk) {
        try {
            if (c.action === 'menu') clickMenu(pid);
            else pressK(pid, c.mods);
        } catch (e) {
            r.actionError = String(e.message).split('\n')[0].slice(0, 200);
        }
        const panelMs = await waitFor(
            page,
            async () => (await webviewState(frame, k.linkPlaceholder)).panelOpen,
            c.expectPanel ? 2500 : 1500,
        );
        r.panelOpenedAfterMs = panelMs;
        r.after = await webviewState(frame, k.linkPlaceholder);
        r.joplinPromptAfter = await joplinPromptVisible(page);
        r.editorGoneLogs = consoleMsgs.slice(mark).filter((m) => /editor is gone/i.test(m.text));
        r.outcome = r.actionError
            ? 'action-failed (e.g. menu item absent)'
            : r.after.panelOpen
              ? 'univer-panel'
              : r.joplinPromptAfter
                ? 'joplin-prompt'
                : r.editorGoneLogs.length
                  ? 'joplin-textLink-editor-gone'
                  : 'nothing';
        if (r.after.panelOpen) await closePanel(frame, k.closePopupId);
        if (r.joplinPromptAfter) r.promptCancelled = await cancelJoplinPrompt(page);
        await cleanupEditing(page, frame, c.mode);
        r.pass = c.observe
            ? true
            : r.after.panelOpen &&
              !r.joplinPromptAfter &&
              r.after.commandLog.includes(k.openPanelId);
    } else {
        await cleanupEditing(page, frame, c.mode);
        r.pass = false;
    }
    r.expected = c.observe
        ? `OBSERVATION (not scored): ${c.note}`
        : `Univer panel open (input "${k.linkPlaceholder}"), ${k.openPanelId} executed, no Joplin "${JOPLIN_LINK_PROMPT}" prompt`;
    r.summary = `mode=${c.mode} startOk=${prep.startOk} outcome=${r.outcome} panelMs=${r.panelOpenedAfterMs} log=${JSON.stringify(r.after?.commandLog)}`;
    return r;
}

async function main() {
    let pid;
    try {
        pid = devJoplinPid();
    } catch (e) {
        console.error(`insert-link-probe: ${e.message}`);
        process.exit(e.code === 'NON_DEV_JOPLIN' ? 3 : 4);
    }
    const k = univerConstants();
    fs.mkdirSync(OUT_DIR, { recursive: true });
    const { chromium } = require('playwright');
    const browser = await chromium.connectOverCDP(CDP_URL);
    const consoleMsgs = [];
    const results = [];
    const log = (r) =>
        console.error(
            `${r.startOk === false ? 'INVALID' : r.pass ? 'PASS' : 'FAIL'} ${r.id}: ${r.summary}`,
        );
    try {
        const page = (browser.contexts()[0]?.pages() ?? []).find((p) =>
            /app\.asar\/index\.html$/.test(p.url()),
        );
        if (!page) throw new Error('Joplin main window page not found over CDP');
        page.on('console', (m) =>
            consoleMsgs.push({ type: m.type(), text: m.text().slice(0, 400) }),
        );
        page.on('pageerror', (e) =>
            consoleMsgs.push({ type: 'pageerror', text: String(e).slice(0, 400) }),
        );

        // ── Notesheet note ──
        const frame = await openNotesheet(page, NOTE_ID);
        const menuOnSheet = await menuEnabledSettled(page, pid, true);
        {
            const r = {
                id: 'menu-enabled-on-notesheet',
                menu: menuOnSheet,
                pass:
                    menuOnSheet.present === true &&
                    menuOnSheet.enabled === true &&
                    menuOnSheet.accelerator === 'Cmd+Option+K',
                expected: 'menu item enabled; AX accelerator Cmd+Option+K',
            };
            r.summary = `enabled=${menuOnSheet.enabled} accelerator=${menuOnSheet.accelerator}`;
            results.push(r);
            log(r);
        }

        const ALT = CMD_OPTION;
        const cases = [
            { id: 'keystroke-cmd-option-k-canvas', mode: 'canvas', mods: ALT, expectPanel: true },
            { id: 'keystroke-cmd-option-k-editing', mode: 'editing', mods: ALT, expectPanel: true },
            { id: 'keystroke-cmd-option-k-blurred', mode: 'blurred', mods: ALT, expectPanel: true },
            { id: 'menu-click-opens-panel', mode: 'canvas', action: 'menu', expectPanel: true },
            ...['canvas', 'editing', 'blurred'].map((mode) => ({
                id: `observe-cmd-k-${mode}`,
                mode,
                mods: CMD,
                observe: true,
                note: "plain Cmd+K is not ours (Joplin's textLink accelerator vs Univer's own Cmd+K); record who handles it",
            })),
        ];
        for (const c of cases) {
            const r = await notesheetCase(page, frame, pid, k, consoleMsgs, c);
            results.push(r);
            log(r);
        }

        // ── Plain markdown note ──
        openNoteUrl(MD_NOTE_ID);
        await page.waitForTimeout(2000);
        const mdState = await page.evaluate(() => {
            const cm = document.querySelector('.cm-editor .cm-content');
            const r = cm ? cm.getBoundingClientRect() : null;
            return { codeMirrorVisible: !!(r && r.width > 0 && r.height > 0) };
        });
        const menuOnMd = await menuEnabledSettled(page, pid, false);
        {
            const r = {
                id: 'menu-disabled-on-markdown-note',
                mdState,
                menu: menuOnMd,
                startOk: mdState.codeMirrorVisible,
                pass:
                    mdState.codeMirrorVisible &&
                    menuOnMd.present === true &&
                    menuOnMd.enabled === false,
                expected:
                    'markdown editor visible; menu item present but disabled (enabledCondition false)',
            };
            r.summary = `startOk=${r.startOk} enabled=${menuOnMd.enabled}`;
            results.push(r);
            log(r);
        }
        {
            const r = { id: 'markdown-cmd-k-opens-joplin-link-dialog', key: 'Cmd+K' };
            bringToFront(pid);
            await page.locator('.cm-editor .cm-content').first().click();
            await page.waitForTimeout(250);
            r.promptBefore = await joplinPromptVisible(page);
            r.frontmost = isFrontmost(pid);
            r.startOk = mdState.codeMirrorVisible && !r.promptBefore && r.frontmost;
            if (r.startOk) {
                pressK(pid, CMD);
                r.promptOpenedAfterMs = await waitFor(page, () => joplinPromptVisible(page));
                r.joplinPromptAfter = await joplinPromptVisible(page);
                r.pass = r.joplinPromptAfter;
                if (r.joplinPromptAfter) r.promptCancelled = await cancelJoplinPrompt(page);
            } else r.pass = false;
            r.expected = `Joplin's "${JOPLIN_LINK_PROMPT}" prompt opens`;
            r.summary = `startOk=${r.startOk} joplinPrompt=${r.joplinPromptAfter}`;
            results.push(r);
            log(r);
        }
        {
            const r = { id: 'markdown-cmd-option-k-is-noop', key: 'Cmd+Option+K' };
            const bodyBefore = await page.evaluate(
                () => document.querySelector('.cm-editor .cm-content')?.textContent ?? null,
            );
            r.promptBefore = await joplinPromptVisible(page);
            r.frontmost = isFrontmost(pid);
            r.startOk = mdState.codeMirrorVisible && !r.promptBefore && r.frontmost;
            if (r.startOk) {
                pressK(pid, CMD_OPTION);
                await page.waitForTimeout(1500);
                r.joplinPromptAfter = await joplinPromptVisible(page);
                const bodyAfter = await page.evaluate(
                    () => document.querySelector('.cm-editor .cm-content')?.textContent ?? null,
                );
                r.bodyUnchanged = bodyAfter === bodyBefore;
                r.pass = !r.joplinPromptAfter && r.bodyUnchanged;
            } else r.pass = false;
            r.expected = 'nothing happens: no prompt, markdown body unchanged';
            r.summary = `startOk=${r.startOk} joplinPrompt=${r.joplinPromptAfter} bodyUnchanged=${r.bodyUnchanged}`;
            results.push(r);
            log(r);
        }

        // Leave the notesheet note open for a following screenshot.
        await openNotesheet(page, NOTE_ID);

        const scored = results.filter((r) => !r.id.startsWith('observe-'));
        const relevantConsole = consoleMsgs.filter(
            (m) =>
                m.type === 'pageerror' ||
                ((m.type === 'error' || m.type === 'warning') && /\[Notesheet\]/.test(m.text)),
        );
        const out = {
            label: LABEL,
            generatedAt: new Date().toISOString(),
            platform: process.platform,
            joplinPid: pid,
            noteId: NOTE_ID,
            markdownNoteId: MD_NOTE_ID,
            univer: k,
            allPass: scored.every((r) => r.pass && r.startOk !== false) && !relevantConsole.length,
            passCount: scored.filter((r) => r.pass && r.startOk !== false).length,
            scoredCount: scored.length,
            consoleRelevant: relevantConsole,
            consoleEditorGone: consoleMsgs.filter((m) => /editor is gone/i.test(m.text)),
            results,
        };
        const utc = out.generatedAt.replace(/[:.]/g, '-');
        const file = path.join(OUT_DIR, `probe-${LABEL}-${utc}.json`);
        fs.writeFileSync(file, JSON.stringify(out, null, 2) + '\n');
        console.error(
            `insert-link-probe: ${out.passCount}/${out.scoredCount} pass; console relevant=${relevantConsole.length}`,
        );
        console.log(file);
        if (!out.allPass) process.exitCode = 5;
    } finally {
        await browser.close();
    }
}

main().catch((e) => {
    console.error(`insert-link-probe: ${e.stack || e.message}`);
    process.exit(4);
});
