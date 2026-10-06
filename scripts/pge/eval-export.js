// PGE export smoke test — exercises the plugin's REAL .xlsx export INSIDE
// the live Joplin webview (over CDP) and inspects the emitted BYTES.
//
// WHY this exists: for its entire history the harness only ever proved the
// RENDER — it screenshotted the Univer canvas / preview pane and sampled
// pixels. It never once opened the artifact the "Export .xlsx" button
// actually produces. That blind spot let a real bug ship undetected: images
// were silently dropped from every export because the export path called
// `Buffer.from(...)` and `Buffer` is UNDEFINED in the Joplin webview
// (it's a browser context, not Node). Jest passed the whole time because
// Node HAS Buffer — the divergence only manifested live. This script closes
// that gap: it drives the real button, captures the real Blob's bytes, and
// unzips them to assert the interesting parts (media / charts / drawings)
// are actually present.
//
// The webview-vs-Node Buffer nuance is the crux of the whole exercise, and
// it cuts BOTH ways here:
//   - Inside the frame (browser): there is NO Node `Buffer`. We read the
//     export Blob via `blob.arrayBuffer()` and base64-encode it with the
//     browser's `btoa` + chunked `String.fromCharCode` (spreading a large
//     Uint8Array into fromCharCode blows the call stack — we chunk at
//     0x8000, mirroring src/drawings/xlsxImage.ts:bytesToBase64).
//   - Back in Node: we decode that base64 with Node's `Buffer` (fine here —
//     this file only ever runs under Node), write the .xlsx, and unzip it
//     with the repo's `jszip` dependency.
//
// Lifecycle: same contract as eval-screenshot.js — attach over CDP to an
// already-running Joplin, do the work, `browser.close()` in a finally to
// DETACH. We do NOT quit Joplin; we don't own it.
//
// Where the export lives: the "Export .xlsx" button (textContent
// 'Export .xlsx') is in the editor toolbar INSIDE the UserWebviewIndex.html
// frame (created ~editorView.tsx:510). Its click handler calls
// handleExport(), which reads the active workbook snapshot, calls
// snapshotToXlsxBuffer(snapshot), wraps it in a Blob, and clicks a
// synthetic `<a download>` after `URL.createObjectURL(blob)`. We intercept
// at `URL.createObjectURL` so we never depend on a real file download
// landing on disk.
//
// Usage:
//   node eval-export.js <feature-id>
//   PGE_NOTE_ID=<id> node eval-export.js <feature-id>
//   PGE_OUT=/tmp/foo.xlsx node eval-export.js <feature-id>
//
// Env:
//   PGE_CDP_PORT   CDP port to attach (default 8315; must match
//                  launch-joplin.sh / Joplin's --remote-debugging-port).
//   PGE_NOTE_ID    Skip note lookup; open this id directly.
//   PGE_OUT        Override the exported .xlsx path (default
//                  screenshots/<feature-id>/export-<utc>.xlsx).
//   JOPLIN_TOKEN   Web Clipper Data API token (else read from
//                  .claude/joplin-token.local).
//   PGE_DEBUG_CONTEXTS  Print every CDP page's title/URL to stderr.
//
// Exit codes:
//   0  export captured AND the zip is a valid workbook (has
//      xl/workbook.xml). The exported .xlsx path is printed to stdout.
//   2  bad usage (no feature id).
//   3  playwright not installed.
//   4  HARNESS could not run: CDP unreachable, editor page / webview frame
//      missing, export never captured within the timeout. (The plugin's
//      export was never exercised — this is an environment problem.)
//   5  ARTIFACT is broken: bytes captured but the zip is invalid / empty /
//      missing xl/workbook.xml. (The export ran but produced garbage.)
// All diagnostics go to stderr; only the artifact path goes to stdout.

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

// Distinct exit codes so callers can tell "harness couldn't run" (4) apart
// from "export produced a broken/empty artifact" (5).
const EXIT_OK = 0;
const EXIT_USAGE = 2;
const EXIT_NO_PLAYWRIGHT = 3;
const EXIT_HARNESS = 4;
const EXIT_ARTIFACT = 5;

// A thrown HarnessError → exit 4; a thrown ArtifactError → exit 5; anything
// else → exit 4 (unknown failures are treated as "harness couldn't run").
class HarnessError extends Error {}
class ArtifactError extends Error {}

const FEATURE_ID = process.argv[2];
if (!FEATURE_ID) {
    console.error('usage: eval-export.js <feature-id>');
    process.exit(EXIT_USAGE);
}

const COLON = FEATURE_ID.indexOf(':');
const BASE_FEATURE_ID = COLON >= 0 ? FEATURE_ID.slice(0, COLON) : FEATURE_ID;

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const OUT_DIR = path.join(REPO_ROOT, 'screenshots', BASE_FEATURE_ID);
const CDP_PORT = process.env.PGE_CDP_PORT || '8315';
const CDP_URL = `http://localhost:${CDP_PORT}`;

// Per-feature note title prefix — kept in sync with eval-screenshot.js's
// TITLE_PREFIX_BY_FEATURE. The harness duplicates rather than shares (per
// its own comments), so this table is a local copy; features that only need
// export verification can pass PGE_NOTE_ID directly and skip the lookup.
const TITLE_PREFIX_BY_FEATURE = {
    'feature-1-smoke-red-cell': 'PGE smoke note ',
    'feature-1-m13-rotated-text-renders': 'PGE M13C eval ',
    'feature-1-m13-rich-text-renders': 'PGE M13D eval ',
    'feature-1-m13-theme-aware-banding:aptos': 'PGE M13E aptos eval ',
    'feature-1-m13-theme-aware-banding:classic': 'PGE M13E classic eval ',
    'feature-1-m15-conditional-formatting': 'PGE M15 CF eval ',
    'feature-1-m16-snapshot-to-html': 'PGE M16 HTML eval ',
    'feature-1-m17-chart-import-no-crash': 'PGE M17 chart eval ',
    'feature-3-m17-multisheet-import-editor-canvas': 'PGE M17 chart f3 eval ',
    'feature-7-m17-chart-svg-html-export-jest': 'PGE M17 chart f7 eval ',
    'feature-8-m17-chart-preview-pane-pge-smoke': 'PGE M17 chart f8 eval ',
};

function discoverApiPort() {
    const script = path.resolve(__dirname, 'discover-api-port.sh');
    try {
        return execFileSync('bash', [script], {
            stdio: ['ignore', 'pipe', 'pipe'],
            timeout: 5000,
        })
            .toString()
            .trim();
    } catch {
        return null;
    }
}

function discoverToken() {
    if (process.env.JOPLIN_TOKEN) return process.env.JOPLIN_TOKEN;
    // Same pattern joplin-api.js / eval-screenshot.js use; gitignored.
    const tokenFile = path.resolve(REPO_ROOT, '.claude', 'joplin-token.local');
    try {
        return fs.readFileSync(tokenFile, 'utf8').trim();
    } catch {
        return '';
    }
}

async function findLatestNoteByTitle(hostPort, token, prefix) {
    const q = new URLSearchParams({ fields: 'id,title,updated_time', limit: '100' });
    if (token) q.set('token', token);
    const res = await fetch(`http://${hostPort}/notes?${q}`);
    if (!res.ok) throw new Error(`Joplin API list-notes failed: ${res.status}`);
    const body = await res.json();
    const matches = (body.items ?? [])
        .filter((n) => (n.title || '').startsWith(prefix))
        .sort((a, b) => (b.updated_time ?? 0) - (a.updated_time ?? 0));
    if (matches.length === 0) {
        throw new Error(
            `no Joplin note with title prefix "${prefix}" — generator should have created one before invoking the evaluator`,
        );
    }
    return matches[0].id;
}

// Heuristic: pick the page that most likely is the Joplin main editor
// renderer. Same scoring as eval-screenshot.js:pickEditorPage.
async function pickEditorPage(context, debug) {
    const pages = context.pages();
    if (debug) {
        console.error(`eval-export: CDP exposes ${pages.length} page(s):`);
        for (const p of pages) {
            try {
                const url = p.url();
                let title = '';
                try {
                    title = await p.title();
                } catch {}
                console.error(`  - title="${title}" url="${url}"`);
            } catch (e) {
                console.error(`  - <unreadable page: ${e.message}>`);
            }
        }
    }
    if (pages.length === 0) {
        throw new HarnessError(
            'CDP attach succeeded but no pages are exposed. Joplin may still be booting.',
        );
    }
    const scored = await Promise.all(
        pages.map(async (p) => {
            let url = '';
            let title = '';
            try {
                url = p.url();
            } catch {}
            try {
                title = await p.title();
            } catch {}
            let score = 0;
            if (/index\.html/i.test(url)) score += 10;
            if (/Joplin/i.test(title)) score += 5;
            if (url && url !== 'about:blank') score += 1;
            return { page: p, url, title, score };
        }),
    );
    scored.sort((a, b) => b.score - a.score);
    const winner = scored[0];
    if (debug) {
        console.error(
            `eval-export: picked page title="${winner.title}" url="${winner.url}" score=${winner.score}`,
        );
    }
    return winner.page;
}

// Univer + the editor toolbar (with the Export button) mount inside the
// Joplin editor page's `UserWebviewIndex.html` frame. Grab it by URL match.
async function pickNotesheetWebview(page) {
    for (let attempt = 0; attempt < 30; attempt++) {
        const frame = page.frames().find((f) => /UserWebviewIndex\.html/.test(f.url()));
        if (frame) return frame;
        await page.waitForTimeout(200);
    }
    return null;
}

// Wait for Univer to actually render to the canvas — the workbook must be
// live before the Export button's handler can `save()` a snapshot.
async function waitForUniverRender(frame) {
    const sel = 'canvas[id^="univer-sheet-main-canvas"]';
    try {
        await frame.waitForSelector(sel, { timeout: 15_000, state: 'attached' });
    } catch {
        console.error(`eval-export: Univer canvas selector "${sel}" did not appear within 15s.`);
        return null;
    }
    try {
        await frame.waitForFunction(
            (s) => {
                const c = document.querySelector(s);
                return c && c.width > 0 && c.height > 0;
            },
            sel,
            { timeout: 5_000, polling: 100 },
        );
    } catch {
        console.error(
            'eval-export: Univer canvas attached but never resized; rendering likely incomplete.',
        );
    }
    await frame.waitForTimeout(250);
    return sel;
}

// Open a joplin:// URL via macOS `open`. Joplin registers the protocol
// handler; openNote is routed through the main process, so an OS-level
// `open` is the cleanest dispatch (same as eval-screenshot.js). We use
// execFileSync (argv array, no shell) so the id can't be shell-interpreted.
function openJoplinNote(id) {
    return new Promise((resolve, reject) => {
        try {
            execFileSync('open', [`joplin://x-callback-url/openNote?id=${id}`], {
                stdio: ['ignore', 'ignore', 'pipe'],
                timeout: 5000,
            });
            resolve();
        } catch (e) {
            reject(new Error(`failed to dispatch joplin:// URL: ${e.message}`));
        }
    });
}

// Install a one-shot monkeypatch on URL.createObjectURL inside the frame.
// handleExport() calls createObjectURL(blob) on the exported .xlsx Blob;
// we intercept the NEXT call, read the blob's bytes, base64-encode them in
// the browser (NO Node Buffer here — this runs in the webview), and stash
// the result on window.__notesheetExportB64. We still delegate to the real
// createObjectURL so the download <a> the handler builds keeps working.
//
// btoa can only take a binary string, and spreading a large Uint8Array into
// String.fromCharCode blows the call stack, so we chunk at 0x8000 exactly
// like src/drawings/xlsxImage.ts:bytesToBase64.
async function installExportInterceptor(frame) {
    await frame.evaluate(() => {
        window.__notesheetExportB64 = null;
        window.__notesheetExportErr = null;
        const realCreate = URL.createObjectURL.bind(URL);
        const patched = function (obj) {
            // Only intercept the first Blob-typed call; restore immediately
            // so nothing else in the app is affected beyond this one export.
            try {
                if (obj instanceof Blob && !window.__notesheetExportPending) {
                    window.__notesheetExportPending = true;
                    URL.createObjectURL = realCreate;
                    obj.arrayBuffer()
                        .then((ab) => {
                            const bytes = new Uint8Array(ab);
                            let binary = '';
                            const CHUNK = 0x8000;
                            for (let i = 0; i < bytes.length; i += CHUNK) {
                                binary += String.fromCharCode.apply(
                                    null,
                                    bytes.subarray(i, i + CHUNK),
                                );
                            }
                            window.__notesheetExportB64 = btoa(binary);
                        })
                        .catch((e) => {
                            window.__notesheetExportErr = String(e && e.message ? e.message : e);
                        });
                }
            } catch (e) {
                window.__notesheetExportErr = String(e && e.message ? e.message : e);
            }
            return realCreate(obj);
        };
        URL.createObjectURL = patched;
    });
}

// Click the "Export .xlsx" toolbar button inside the frame. Prefer a
// role-based locator (accessible name), fall back to a text locator that
// matches the codebase's literal button text.
async function clickExportButton(frame) {
    // getByRole first — button element with accessible name 'Export .xlsx'.
    try {
        const byRole = frame.getByRole('button', { name: 'Export .xlsx' });
        await byRole.waitFor({ state: 'visible', timeout: 5_000 });
        await byRole.click();
        return true;
    } catch {
        // Fall through to a text-based locator.
    }
    try {
        const byText = frame.locator('button', { hasText: 'Export .xlsx' }).first();
        await byText.waitFor({ state: 'visible', timeout: 5_000 });
        await byText.click();
        return true;
    } catch (e) {
        console.error(`eval-export: could not find/click the Export .xlsx button: ${e.message}`);
        return false;
    }
}

// Poll window.__notesheetExportB64 until populated or timeout. Returns the
// base64 string, or throws HarnessError on timeout / in-frame capture error.
async function waitForExportBytes(frame, timeoutMs = 20_000) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
        const state = await frame.evaluate(() => ({
            b64: window.__notesheetExportB64,
            err: window.__notesheetExportErr,
        }));
        if (state.err) {
            throw new HarnessError(`in-frame export capture errored: ${state.err}`);
        }
        if (state.b64) return state.b64;
        await frame.waitForTimeout(250);
    }
    throw new HarnessError(
        `export blob was never captured within ${timeoutMs}ms. The Export .xlsx handler may have thrown (check Joplin devtools console) or never called URL.createObjectURL.`,
    );
}

// Unzip the exported .xlsx (via the repo's jszip) and build a manifest of
// the interesting parts. Throws ArtifactError if the zip is invalid, empty,
// or missing xl/workbook.xml.
async function inspectXlsx(JSZip, buffer) {
    let zip;
    try {
        zip = await JSZip.loadAsync(buffer);
    } catch (e) {
        throw new ArtifactError(`exported bytes are not a valid zip: ${e.message}`);
    }
    const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
    if (names.length === 0) {
        throw new ArtifactError('exported zip is empty (no entries).');
    }
    if (!names.includes('xl/workbook.xml')) {
        throw new ArtifactError(
            `exported zip has no xl/workbook.xml — not a valid workbook. Entries: ${names.join(', ')}`,
        );
    }

    // Collect byte sizes for the interesting part families.
    async function sizeOf(name) {
        const buf = await zip.files[name].async('nodebuffer');
        return buf.length;
    }
    const media = [];
    const charts = [];
    const drawings = [];
    for (const name of names.sort()) {
        if (/^xl\/media\//.test(name)) media.push({ name, size: await sizeOf(name) });
        else if (/^xl\/charts\//.test(name)) charts.push({ name, size: await sizeOf(name) });
        else if (/^xl\/drawings\//.test(name)) drawings.push({ name, size: await sizeOf(name) });
    }
    return { entryCount: names.length, media, charts, drawings };
}

// Render + write the human/machine-readable sidecar. This is the file the
// evaluator Reads as evidence: it lists the media/chart/drawing parts and a
// one-line PASS/EMPTY summary per family.
function writeManifest(manifestPath, outXlsx, info) {
    function familyLine(label, parts) {
        return parts.length > 0 ? `${label}: ${parts.length} part(s)` : `${label}: EMPTY`;
    }
    const lines = [];
    lines.push(`# Notesheet export manifest`);
    lines.push(`feature: ${FEATURE_ID}`);
    lines.push(`artifact: ${outXlsx}`);
    lines.push(`generated: ${new Date().toISOString()}`);
    lines.push(`total-entries: ${info.entryCount}`);
    lines.push(`workbook: PRESENT (xl/workbook.xml)`);
    lines.push('');
    lines.push(familyLine('MEDIA', info.media));
    lines.push(familyLine('CHARTS', info.charts));
    lines.push(familyLine('DRAWINGS', info.drawings));
    lines.push('');
    lines.push(`## media parts`);
    if (info.media.length === 0) lines.push('  (none)');
    for (const p of info.media) lines.push(`  ${p.name}  ${p.size} bytes`);
    lines.push(`## chart parts`);
    if (info.charts.length === 0) lines.push('  (none)');
    for (const p of info.charts) lines.push(`  ${p.name}  ${p.size} bytes`);
    lines.push(`## drawing parts`);
    if (info.drawings.length === 0) lines.push('  (none)');
    for (const p of info.drawings) lines.push(`  ${p.name}  ${p.size} bytes`);
    lines.push('');
    fs.writeFileSync(manifestPath, lines.join('\n') + '\n');
}

async function main() {
    fs.mkdirSync(OUT_DIR, { recursive: true });
    const utc = new Date().toISOString().replace(/[:.]/g, '-');
    const outXlsx = process.env.PGE_OUT || path.join(OUT_DIR, `export-${utc}.xlsx`);
    // Make sure the parent of a caller-overridden PGE_OUT exists.
    fs.mkdirSync(path.dirname(outXlsx), { recursive: true });
    const manifestPath = `${outXlsx}.manifest.txt`;

    // jszip is a repo dependency (transitive via exceljs; used directly by
    // src/drawings/xlsxImage.ts). We only need it Node-side.
    let JSZip;
    try {
        JSZip = require('jszip');
    } catch (e) {
        throw new HarnessError(`jszip dependency did not resolve: ${e.message}`);
    }

    let playwright;
    try {
        playwright = require('playwright');
    } catch {
        console.error(
            [
                'Playwright not installed. Run from repo root:',
                '    npm install --save-dev playwright',
            ].join('\n'),
        );
        process.exit(EXIT_NO_PLAYWRIGHT);
    }
    const { chromium } = playwright;

    // 1. Sanity-check CDP is reachable (launch-joplin.sh should have already
    // verified; re-check so a stale/absent Joplin fails fast + clearly).
    try {
        const r = await fetch(`${CDP_URL}/json/version`, { signal: AbortSignal.timeout(3000) });
        if (!r.ok) throw new Error(`CDP /json/version → ${r.status}`);
    } catch (e) {
        throw new HarnessError(
            `CDP endpoint not reachable at ${CDP_URL}/json/version (${e.message}). Is Joplin running with --remote-debugging-port=${CDP_PORT}? Run scripts/pge/launch-joplin.sh.`,
        );
    }

    // 2. Attach.
    console.error(`eval-export: attaching CDP at ${CDP_URL}`);
    const browser = await chromium.connectOverCDP(CDP_URL);

    let capturedB64 = null;
    try {
        const contexts = browser.contexts();
        if (contexts.length === 0) {
            throw new HarnessError('CDP attach succeeded but no BrowserContexts present.');
        }
        const context = contexts[0];
        const debug =
            !!process.env.PGE_DEBUG_CONTEXTS || contexts.length > 1 || context.pages().length > 1;
        if (contexts.length > 1) {
            console.error(`eval-export: CDP exposes ${contexts.length} contexts; using the first.`);
        }

        const page = await pickEditorPage(context, debug);

        // Let the renderer settle before dispatching navigation.
        try {
            await page.waitForLoadState('networkidle', { timeout: 10_000 });
        } catch {
            console.error('eval-export: networkidle wait timed out at 10s; proceeding anyway.');
        }

        // 3. Open the test note (PGE_NOTE_ID wins; else lookup by prefix).
        const prefix =
            TITLE_PREFIX_BY_FEATURE[FEATURE_ID] || TITLE_PREFIX_BY_FEATURE[BASE_FEATURE_ID];
        const explicitNote = process.env.PGE_NOTE_ID;
        if (explicitNote) {
            console.error(`eval-export: opening note id ${explicitNote} (PGE_NOTE_ID)`);
            await openJoplinNote(explicitNote);
        } else if (prefix) {
            const hostPort = discoverApiPort();
            if (!hostPort)
                throw new HarnessError('Joplin Web Clipper API not found via discover-api-port.sh');
            const token = discoverToken();
            if (!token) {
                throw new HarnessError(
                    'No Joplin Web Clipper token found. Set JOPLIN_TOKEN or write the token to .claude/joplin-token.local. Get it from Joplin → Settings → Web Clipper → Authorization tokens.',
                );
            }
            const noteId = await findLatestNoteByTitle(hostPort, token, prefix);
            console.error(`eval-export: opening note id ${noteId} (prefix="${prefix}")`);
            await openJoplinNote(noteId);
        } else {
            throw new HarnessError(
                `no title-prefix mapping for "${FEATURE_ID}" and PGE_NOTE_ID unset. Set PGE_NOTE_ID to the test note id, or add a mapping to TITLE_PREFIX_BY_FEATURE.`,
            );
        }

        // 4. Drop into the UserWebviewIndex frame and wait for Univer paint.
        const frame = await pickNotesheetWebview(page);
        if (!frame) {
            throw new HarnessError(
                'UserWebviewIndex frame did not appear; the opened note may not be a Notesheet, or the plugin failed to load. (Note: install-plugin.sh against a LIVE Joplin removes the plugin from that session — quit → install → launch.)',
            );
        }
        const sel = await waitForUniverRender(frame);
        if (!sel) {
            throw new HarnessError(
                'Univer canvas never rendered; the workbook is not live, so Export .xlsx would have nothing to save.',
            );
        }
        console.error(`eval-export: Univer canvas rendered (selector "${sel}")`);

        // 5. Install the createObjectURL interceptor, then click Export.
        await installExportInterceptor(frame);
        console.error('eval-export: interceptor installed; clicking Export .xlsx');
        const clicked = await clickExportButton(frame);
        if (!clicked) {
            throw new HarnessError('Export .xlsx button not found/clickable in the webview frame.');
        }

        // 6. Poll for the captured base64 blob.
        capturedB64 = await waitForExportBytes(frame, 20_000);
        console.error(`eval-export: captured export blob (${capturedB64.length} base64 chars)`);
    } finally {
        // Detach but DO NOT close Joplin.
        await browser.close();
    }

    if (!capturedB64) {
        throw new HarnessError('export capture returned no bytes.');
    }

    // 7. Node-side: decode (Buffer is fine HERE — Node only), write, unzip.
    const buffer = Buffer.from(capturedB64, 'base64');
    fs.writeFileSync(outXlsx, buffer);
    const stat = fs.statSync(outXlsx);
    console.error(`eval-export: wrote ${outXlsx} (${stat.size} bytes)`);
    if (stat.size < 100) {
        throw new ArtifactError(`exported .xlsx is suspiciously small (${stat.size} bytes).`);
    }

    // 8. Inspect + manifest.
    const info = await inspectXlsx(JSZip, buffer);
    writeManifest(manifestPath, outXlsx, info);
    console.error(
        `eval-export: manifest → ${manifestPath} ` +
            `(MEDIA=${info.media.length} CHARTS=${info.charts.length} DRAWINGS=${info.drawings.length}, ${info.entryCount} entries)`,
    );

    // 9. Success: the artifact path to stdout (mirrors eval-screenshot.js).
    process.stdout.write(outXlsx + '\n');
}

main()
    .then(() => process.exit(EXIT_OK))
    .catch((e) => {
        if (e instanceof ArtifactError) {
            console.error('eval-export ARTIFACT-BROKEN:', e.message);
            process.exit(EXIT_ARTIFACT);
        }
        // HarnessError and everything unexpected → "couldn't run" (4).
        console.error('eval-export HARNESS-FAILURE:', e.stack || e.message);
        process.exit(EXIT_HARNESS);
    });
