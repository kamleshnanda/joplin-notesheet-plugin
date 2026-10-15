#!/usr/bin/env node
// M19 feature-1 (BACKLOG E3) live probe: arrow keys at the sheet edge.
//
// Drives REAL keystrokes (Playwright `page.keyboard.press`, i.e. CDP
// Input.dispatchKeyEvent) into the Notesheet webview after clicking the Univer
// canvas so the sheet has keyboard focus. Selections are SET via the facade
// (`FRange.activate()`) only to position the start cell, then READ back via
// `univerAPI.getActiveWorkbook().getActiveSheet().getActiveRange()` and the
// selection's current cell. `executeCommand` is never used to move: the bug
// lives in the shortcut / editor-bridge paths.
//
// It also records, per case, every `sheet.command.move-selection` the probe saw
// on BeforeCommandExecute (with the `cancel` flag our veto set, because the
// plugin's listener is registered first) and whether it reached
// CommandExecuted. On `main` the cancel flag is always false.
//
// Usage (Joplin dev profile running with CDP, see launch-joplin.sh):
//   PGE_NOTE_ID=<empty notesheet id> [PGE_DRAWING_NOTE_ID=<id>] \
//     node scripts/pge/arrow-edge-probe.js <label>
// <label> is e.g. `main` or `branch`. Output:
//   screenshots/feature-1-m19-arrow-keys-clamp-at-sheet-edge/probe-<label>-<utc>.json
// Exit codes: 0 ran (inspect `allPass` / per-case `pass`), 2 usage,
// 4 harness problem, 5 at least one case failed its expectation.
//
// Expectations are the M19 spec (Excel behaviour). On `main` many cases are
// EXPECTED to fail; that is the reproduced bug.

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

const LABEL = process.argv[2];
const NOTE_ID = process.env.PGE_NOTE_ID;
const DRAWING_NOTE_ID = process.env.PGE_DRAWING_NOTE_ID || '';
if (!LABEL || !NOTE_ID) {
    console.error('usage: PGE_NOTE_ID=<id> [PGE_DRAWING_NOTE_ID=<id>] arrow-edge-probe.js <label>');
    process.exit(2);
}

const FEATURE_ID = 'feature-1-m19-arrow-keys-clamp-at-sheet-edge';
const REPO_ROOT = path.resolve(__dirname, '..', '..');
const OUT_DIR = path.join(REPO_ROOT, 'screenshots', FEATURE_ID);
const CDP_URL = `http://localhost:${process.env.PGE_CDP_PORT || '8315'}`;
const CANVAS = 'canvas[id^="univer-sheet-main-canvas"]';
const MOD = process.platform === 'darwin' ? 'Meta' : 'Control';
const SETTLE_MS = 250;
// Width used by the "different column count" cases (default sheets have 20).
const WIDE_COLS = 50;

// 0-based column index → A1 letters (0 → A, 25 → Z, 49 → AX).
function colA1(col) {
    let s = '';
    let c = col + 1;
    while (c > 0) {
        const m = (c - 1) % 26;
        s = String.fromCharCode(65 + m) + s;
        c = Math.floor((c - 1) / 26);
    }
    return s;
}

async function openNote(page, id) {
    execFileSync('open', [`joplin://x-callback-url/openNote?id=${id}`], {
        stdio: ['ignore', 'ignore', 'pipe'],
        timeout: 5000,
    });
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

// Install a probe-side command log in the webview (idempotent per frame).
async function installCommandLog(frame) {
    await frame.evaluate(() => {
        const api = window.__notesheetUniverAPI;
        // The webview frame survives note switches but Univer is re-booted
        // (new facade), so re-install whenever the facade instance changed.
        if (window.__arrowProbe && window.__arrowProbe.api === api) {
            window.__arrowProbe.log.length = 0;
            return;
        }
        if (window.__arrowProbe) {
            try {
                window.__arrowProbe.dispose();
            } catch {
                /* old instance already disposed */
            }
        }
        const log = [];
        const d1 = api.addEvent(api.Event.BeforeCommandExecute, (e) => {
            if (e.id === 'sheet.command.move-selection' || e.id === 'sheet.command.move-drawing')
                log.push({
                    phase: 'before',
                    id: e.id,
                    params: JSON.parse(JSON.stringify(e.params ?? null)),
                    cancel: !!e.cancel,
                });
        });
        const d2 = api.addEvent(api.Event.CommandExecuted, (e) => {
            if (e.id === 'sheet.command.move-selection' || e.id === 'sheet.command.move-drawing')
                log.push({ phase: 'executed', id: e.id });
        });
        window.__arrowProbe = { api, log, dispose: () => (d1.dispose(), d2.dispose()) };
    });
}

async function state(frame) {
    return frame.evaluate(() => {
        const api = window.__notesheetUniverAPI;
        const ws = api.getActiveWorkbook().getActiveSheet();
        const r = ws.getActiveRange();
        const sel = ws.getSelection();
        const cur = sel && sel.getCurrentCell ? sel.getCurrentCell() : null;
        const toA1 = (row, col) => {
            let s = '';
            let c = col + 1;
            while (c > 0) {
                const m = (c - 1) % 26;
                s = String.fromCharCode(65 + m) + s;
                c = Math.floor((c - 1) / 26);
            }
            return `${s}${row + 1}`;
        };
        let stack = null;
        try {
            stack = api._commandService?._commandExecutionStack?.length ?? null;
        } catch {
            /* private; best effort */
        }
        const sheet = ws.getSheet();
        return {
            range: r ? r.getA1Notation() : null,
            current: cur ? toA1(cur.actualRow, cur.actualColumn) : null,
            selections: sel ? sel.getActiveRangeList().map((x) => x.getA1Notation()) : [],
            colAVisible: sheet.getColVisible(0),
            commandStackDepth: stack,
        };
    });
}

async function takeLog(frame) {
    return frame.evaluate(() => {
        const out = window.__arrowProbe.log.slice();
        window.__arrowProbe.log.length = 0;
        return out;
    });
}

async function activate(frame, a1) {
    await frame.evaluate((ref) => {
        const ws = window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet();
        ws.getRange(ref).activate();
    }, a1);
    await frame.waitForTimeout(150);
}

async function focusCanvas(frame) {
    // Click a cell well inside the grid (around B3) so the sheet owns focus.
    const box = await frame.locator(CANVAS).first().boundingBox();
    if (!box) throw new Error('canvas has no bounding box');
    await frame
        .locator(CANVAS)
        .first()
        .click({ position: { x: 120, y: 60 } });
    await frame.waitForTimeout(200);
}

async function lastRefs(frame) {
    return frame.evaluate(() => {
        const ws = window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet();
        const sheet = ws.getSheet();
        const rows = sheet.getRowCount();
        const cols = sheet.getColumnCount();
        const colA1 = (col) => {
            let s = '';
            let c = col + 1;
            while (c > 0) {
                const m = (c - 1) % 26;
                s = String.fromCharCode(65 + m) + s;
                c = Math.floor((c - 1) / 26);
            }
            return s;
        };
        return {
            rows,
            cols,
            lastColCell: `${colA1(cols - 1)}5`,
            lastRowCell: `C${rows}`,
            bottomRight: `${colA1(cols - 1)}${rows}`,
        };
    });
}

async function runKeyCase(page, frame, c) {
    await page.keyboard.press('Escape');
    await frame.waitForTimeout(100);
    if (c.setup) await c.setup();
    await activate(frame, c.start);
    if (c.preKeys)
        for (const k of c.preKeys) {
            await page.keyboard.press(k);
            await frame.waitForTimeout(150);
        }
    await frame.waitForTimeout(100);
    await takeLog(frame);
    const before = await state(frame);
    if (c.type) {
        await page.keyboard.type(c.type);
        await frame.waitForTimeout(SETTLE_MS);
    }
    await page.keyboard.press(c.key);
    await frame.waitForTimeout(SETTLE_MS);
    const after = await state(frame);
    const commands = await takeLog(frame);
    let extra = {};
    if (c.inspect) extra = await c.inspect();
    if (c.teardown) await c.teardown();
    const res = { id: c.id, desc: c.desc, key: c.key, before, after, commands, ...extra };
    // Harness guard: the case only counts if it really started where intended.
    const wantStart = c.expectStart || c.start;
    res.startOk = before.range === wantStart;
    res.pass = res.startOk && c.expect(res);
    res.expected = c.expectText;
    return res;
}

async function drawingCase(page) {
    const frame = await openNote(page, DRAWING_NOTE_ID);
    await installCommandLog(frame);
    const info = await frame.evaluate(() => {
        const api = window.__notesheetUniverAPI;
        const ws = api.getActiveWorkbook().getActiveSheet();
        const imgs = ws.getImages ? ws.getImages() : [];
        if (!imgs.length) return { error: 'no images on active sheet' };
        const img = imgs[0];
        const t = img._image?.transform || {};
        const cfg = ws.getSheet().getConfig();
        return {
            id: img.getId(),
            left: t.left,
            top: t.top,
            width: t.width,
            height: t.height,
            rowHeaderWidth: cfg.rowHeader?.width ?? 46,
            columnHeaderHeight: cfg.columnHeader?.height ?? 20,
        };
    });
    if (info.error) return { id: 'drawing-nudge', pass: false, error: info.error };
    const readLeft = () =>
        frame.evaluate((id) => {
            const ws = window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet();
            const img = ws.getImageById(id);
            return img?._image?.transform?.left ?? null;
        }, info.id);
    // Click the image centre (sheet coords + header offsets; note opens unscrolled).
    const x = info.rowHeaderWidth + info.left + Math.min(info.width / 2, 40);
    const y = info.columnHeaderHeight + info.top + Math.min(info.height / 2, 40);
    await frame.locator(CANVAS).first().click({ position: { x, y } });
    await frame.waitForTimeout(400);
    await takeLog(frame);
    const cellBefore = await state(frame);
    const leftBefore = await readLeft();
    await page.keyboard.press('ArrowRight');
    await frame.waitForTimeout(SETTLE_MS);
    await page.keyboard.press('ArrowRight');
    await frame.waitForTimeout(SETTLE_MS);
    const leftAfterRight = await readLeft();
    await page.keyboard.press('ArrowLeft');
    await frame.waitForTimeout(SETTLE_MS);
    await page.keyboard.press('ArrowLeft');
    await frame.waitForTimeout(SETTLE_MS);
    const leftAfterLeft = await readLeft();
    const cellAfter = await state(frame);
    const commands = await takeLog(frame);
    await page.keyboard.press('Escape');
    const pass =
        typeof leftBefore === 'number' &&
        leftAfterRight > leftBefore &&
        Math.abs(leftAfterLeft - leftBefore) < 0.5 &&
        !commands.some((e) => e.id === 'sheet.command.move-selection' && e.cancel);
    return {
        id: 'drawing-nudge',
        desc: 'drawing selected (TwoImages-MultiSheet.xlsx) → ArrowRight x2 then ArrowLeft x2 nudges it',
        clickAt: { x, y },
        image: info,
        transformLeft: {
            before: leftBefore,
            afterRightX2: leftAfterRight,
            afterLeftX2: leftAfterLeft,
        },
        activeCell: { before: cellBefore, after: cellAfter },
        commands,
        expected: 'transform.left increases on Right and returns on Left; no vetoed move-selection',
        pass,
    };
}

async function main() {
    fs.mkdirSync(OUT_DIR, { recursive: true });
    const { chromium } = require('playwright');
    const browser = await chromium.connectOverCDP(CDP_URL);
    const consoleMsgs = [];
    try {
        const pages = browser.contexts()[0]?.pages() ?? [];
        const page = pages.find((p) => /app\.asar\/index\.html$/.test(p.url()));
        if (!page) throw new Error('Joplin main window page not found over CDP');
        page.on('console', (m) => {
            if (m.type() === 'error' || m.type() === 'warning')
                consoleMsgs.push({ type: m.type(), text: m.text().slice(0, 400) });
        });
        page.on('pageerror', (e) =>
            consoleMsgs.push({ type: 'pageerror', text: String(e).slice(0, 400) }),
        );

        const frame = await openNote(page, NOTE_ID);
        await installCommandLog(frame);
        // Seed: D5 has a value (data edge for Ctrl/Cmd+Left); A5 empty.
        await frame.evaluate(() => {
            const ws = window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet();
            ws.getRange('D5').setValue('d5');
            ws.getRange('A5').clear();
            ws.showColumns(0, 1);
        });
        await focusCanvas(frame);
        // Warm-up: settle any pointer-driven selection from the focusing click.
        await activate(frame, 'H10');
        await page.keyboard.press('ArrowDown');
        await frame.waitForTimeout(300);
        const refs = await lastRefs(frame);
        const stays = (r) =>
            r.after.range === r.before.range && r.after.current === r.before.current;
        const at = (ref) => (r) => r.after.range === ref && r.after.current === ref;

        const cases = [
            { id: 'A5-left', start: 'A5', key: 'ArrowLeft', expect: stays, expectText: 'stays A5' },
            { id: 'C1-up', start: 'C1', key: 'ArrowUp', expect: stays, expectText: 'stays C1' },
            { id: 'A1-left', start: 'A1', key: 'ArrowLeft', expect: stays, expectText: 'stays A1' },
            { id: 'A1-up', start: 'A1', key: 'ArrowUp', expect: stays, expectText: 'stays A1' },
            {
                id: 'lastcol-right',
                start: refs.lastColCell,
                key: 'ArrowRight',
                expect: stays,
                expectText: `stays ${refs.lastColCell}`,
            },
            {
                id: 'lastrow-down',
                start: refs.lastRowCell,
                key: 'ArrowDown',
                expect: stays,
                expectText: `stays ${refs.lastRowCell}`,
            },
            {
                id: 'B5-left',
                start: 'B5',
                key: 'ArrowLeft',
                expect: at('A5'),
                expectText: 'moves to A5',
            },
            {
                id: 'A5-right',
                start: 'A5',
                key: 'ArrowRight',
                expect: at('B5'),
                expectText: 'moves to B5',
            },
            {
                id: 'A5-shift-left',
                start: 'A5',
                key: 'Shift+ArrowLeft',
                expect: stays,
                expectText: 'selection unchanged (already clamped upstream)',
            },
            {
                id: 'D5-mod-left',
                start: 'D5',
                key: `${MOD}+ArrowLeft`,
                expect: at('A5'),
                expectText: 'jumps to the data edge A5 (same as main)',
            },
            {
                id: 'A5-mod-left',
                start: 'A5',
                key: `${MOD}+ArrowLeft`,
                expect: stays,
                expectText: 'stays A5 (findNextGapRange already clamps)',
            },
            {
                id: 'multi-A5C5-activeC5-left',
                start: 'C5',
                preKeys: ['Shift+ArrowLeft', 'Shift+ArrowLeft'],
                expectStart: 'A5:C5',
                key: 'ArrowLeft',
                expect: () => true,
                expectText:
                    'OBSERVATION ONLY (operator undecided): Excel would move to B5; record what happens',
            },
            {
                id: 'hiddenA-B5-left',
                start: 'B5',
                key: 'ArrowLeft',
                setup: () =>
                    frame.evaluate(() =>
                        window.__notesheetUniverAPI
                            .getActiveWorkbook()
                            .getActiveSheet()
                            .hideColumns(0, 1),
                    ),
                teardown: () =>
                    frame.evaluate(() =>
                        window.__notesheetUniverAPI
                            .getActiveWorkbook()
                            .getActiveSheet()
                            .showColumns(0, 1),
                    ),
                expect: stays,
                expectText: 'stays B5 (column A hidden)',
            },
            // Corner: both edges at once (operator-requested corner-case run).
            {
                id: 'bottomright-right',
                start: refs.bottomRight,
                key: 'ArrowRight',
                expect: stays,
                expectText: `stays ${refs.bottomRight} (last row AND last column)`,
            },
            {
                id: 'bottomright-down',
                start: refs.bottomRight,
                key: 'ArrowDown',
                expect: stays,
                expectText: `stays ${refs.bottomRight} (last row AND last column)`,
            },
            // A sheet with a different column count: the clamp must follow the
            // sheet's real width, not assume 20. setup widens the sheet, teardown
            // restores the original width.
            ...[
                {
                    id: 'wide-lastcol-right',
                    start: `${colA1(WIDE_COLS - 1)}5`,
                    key: 'ArrowRight',
                    expect: stays,
                    expectText: `stays ${colA1(WIDE_COLS - 1)}5 (${WIDE_COLS}-column sheet)`,
                },
                {
                    id: 'wide-bottomright-down',
                    start: `${colA1(WIDE_COLS - 1)}${refs.rows}`,
                    key: 'ArrowDown',
                    expect: stays,
                    expectText: `stays ${colA1(WIDE_COLS - 1)}${refs.rows} (${WIDE_COLS}-column sheet)`,
                },
                {
                    id: 'wide-T5-right',
                    start: `${colA1(refs.cols - 1)}5`,
                    key: 'ArrowRight',
                    expect: at(`${colA1(refs.cols)}5`),
                    expectText: `moves to ${colA1(refs.cols)}5 (old last column is no longer the edge)`,
                },
            ].map((c) => ({
                ...c,
                setup: () =>
                    frame.evaluate(
                        (n) =>
                            window.__notesheetUniverAPI
                                .getActiveWorkbook()
                                .getActiveSheet()
                                .setColumnCount(n),
                        WIDE_COLS,
                    ),
                teardown: () =>
                    frame.evaluate(
                        (n) =>
                            window.__notesheetUniverAPI
                                .getActiveWorkbook()
                                .getActiveSheet()
                                .setColumnCount(n),
                        refs.cols,
                    ),
            })),
            {
                id: 'A5-type-x-then-left',
                start: 'A5',
                type: 'x',
                key: 'ArrowLeft',
                inspect: () =>
                    frame.evaluate(() => {
                        const ws = window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet();
                        return { a5Value: ws.getRange('A5').getValue() };
                    }),
                expect: (r) =>
                    r.a5Value === 'x' && r.after.range === 'A5' && r.after.current === 'A5',
                expectText: 'edit commits (A5 = "x") and active cell stays A5',
            },
        ];

        const results = [];
        for (const c of cases) {
            const r = await runKeyCase(page, frame, c);
            results.push(r);
            console.error(
                `${r.pass ? 'PASS' : 'FAIL'} ${r.id}: ${r.before.range}(${r.before.current}) --${r.key}--> ${r.after.range}(${r.after.current})` +
                    (r.a5Value !== undefined ? ` A5=${JSON.stringify(r.a5Value)}` : '') +
                    ` cmds=${JSON.stringify(r.commands.map((e) => `${e.phase}:${e.cancel ? 'CANCEL' : ''}`))}`,
            );
        }
        // Clean the typed value so re-runs start from the same state.
        await frame.evaluate(() =>
            window.__notesheetUniverAPI.getActiveWorkbook().getActiveSheet().getRange('A5').clear(),
        );
        const finalState = await state(frame);
        // The wide-sheet cases resize the note's sheet; prove teardown restored it.
        const colsAtEnd = await frame.evaluate(() =>
            window.__notesheetUniverAPI
                .getActiveWorkbook()
                .getActiveSheet()
                .getSheet()
                .getColumnCount(),
        );
        results.push({
            id: 'sheet-width-restored',
            pass: colsAtEnd === refs.cols,
            expected: `column count back to ${refs.cols}`,
            colsAtEnd,
        });
        console.error(
            `${colsAtEnd === refs.cols ? 'PASS' : 'FAIL'} sheet-width-restored: ${colsAtEnd}`,
        );

        if (DRAWING_NOTE_ID) {
            const d = await drawingCase(page);
            results.push(d);
            console.error(
                `${d.pass ? 'PASS' : 'FAIL'} drawing-nudge: left ${JSON.stringify(d.transformLeft)} ${d.error ?? ''}`,
            );
        }

        const relevantConsole = consoleMsgs.filter(
            (m) => /CanceledError|\[Notesheet\]/.test(m.text) || m.type === 'pageerror',
        );
        const scored = results.filter((r) => !r.id.startsWith('multi-'));
        const out = {
            label: LABEL,
            generatedAt: new Date().toISOString(),
            platform: process.platform,
            modifierKey: MOD,
            sheet: refs,
            noteId: NOTE_ID,
            drawingNoteId: DRAWING_NOTE_ID || null,
            allPass: scored.every((r) => r.pass),
            passCount: scored.filter((r) => r.pass).length,
            scoredCount: scored.length,
            commandStackDepthAtEnd: finalState.commandStackDepth,
            consoleRelevant: relevantConsole,
            consoleAllErrorsAndWarnings: consoleMsgs,
            results,
        };
        const utc = out.generatedAt.replace(/[:.]/g, '-');
        const file = path.join(OUT_DIR, `probe-${LABEL}-${utc}.json`);
        fs.writeFileSync(file, JSON.stringify(out, null, 2) + '\n');
        console.error(
            `arrow-edge-probe: ${out.passCount}/${out.scoredCount} pass; console relevant=${relevantConsole.length}`,
        );
        console.log(file);
        if (!out.allPass) process.exitCode = 5;
    } finally {
        await browser.close();
    }
}

main().catch((e) => {
    console.error(`arrow-edge-probe: ${e.stack || e.message}`);
    process.exit(4);
});
