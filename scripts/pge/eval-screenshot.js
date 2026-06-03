// Evaluator's authoritative screenshot capture via Playwright + CDP
// attach. Joplin must already be running (started by launch-joplin.sh)
// with `--remote-debugging-port=$PGE_CDP_PORT`. We attach to its
// renderer using `chromium.connectOverCDP`, find the editor page,
// open the feature's test note via the joplin:// URL scheme, wait for
// Univer to render, then call `page.screenshot()`.
//
// We do NOT use Playwright's `_electron.launch`: it injects
// `--inspect=0` into every Electron child, which Joplin rejects with
// a fatal "Unknown flag" modal. Connecting over CDP to an already-
// running Joplin sidesteps that entirely.
//
// Lifecycle: we attach, screenshot, and detach. We do NOT close
// Joplin; we don't own it. The operator (or run-cycle.sh) decides
// when to quit.
//
// Usage:
//   node eval-screenshot.js <feature-id>
//   PGE_OUT=/tmp/foo.png node eval-screenshot.js <feature-id>
//
// Env:
//   PGE_CDP_PORT      CDP port to attach (default 8315; must match
//                     launch-joplin.sh).
//   PGE_NOTE_ID       Skip note lookup; open this id directly.
//   PGE_OUT           Override output path (verification mode).
//   PGE_DEBUG_CONTEXTS Print every CDP page's title/URL to stderr.
//   JOPLIN_TOKEN      Web Clipper Data API token (else read from
//                     .claude/joplin-token.local).

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

const FEATURE_ID = process.argv[2];
if (!FEATURE_ID) {
    console.error('usage: eval-screenshot.js <feature-id>');
    process.exit(2);
}

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const SHOT_DIR = path.join(REPO_ROOT, 'screenshots', FEATURE_ID);
const CDP_PORT = process.env.PGE_CDP_PORT || '8315';
const CDP_URL = `http://localhost:${CDP_PORT}`;

// Per-feature note title prefix. Each feature spec in BUILD_PLAN.md
// dictates what the title prefix should be; keep this table in sync
// with eval-screenshot.sh / the planner.
const TITLE_PREFIX_BY_FEATURE = {
    'feature-1-smoke-red-cell': 'PGE smoke note ',
};

function discoverApiPort() {
    const script = path.resolve(__dirname, 'discover-api-port.sh');
    try {
        return execSync(`bash "${script}"`, {
            stdio: ['ignore', 'pipe', 'pipe'],
            timeout: 5000,
        }).toString().trim();
    } catch {
        return null;
    }
}

function discoverToken() {
    if (process.env.JOPLIN_TOKEN) return process.env.JOPLIN_TOKEN;
    // Same pattern joplin-api.js uses; gitignored.
    const tokenFile = path.resolve(REPO_ROOT, '.claude', 'joplin-token.local');
    try { return fs.readFileSync(tokenFile, 'utf8').trim(); }
    catch { return ''; }
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
        throw new Error(`no Joplin note with title prefix "${prefix}" — generator should have created one before invoking the evaluator`);
    }
    return matches[0].id;
}

// Heuristic: pick the page that most likely is the Joplin main editor
// renderer. Joplin's renderer URL is a `file://.../index.html` from
// the app bundle; helper subprocesses (printer, etc.) are usually
// `about:blank` or have no body. We prefer:
//   1. URL containing 'index.html' AND title 'Joplin'
//   2. URL containing 'index.html'
//   3. The page with the largest viewport (main window > popups)
//   4. Fall back to the first page.
async function pickEditorPage(context, debug) {
    const pages = context.pages();
    if (debug) {
        console.error(`eval-screenshot: CDP exposes ${pages.length} page(s):`);
        for (const p of pages) {
            try {
                const url = p.url();
                let title = '';
                try { title = await p.title(); } catch {}
                console.error(`  - title="${title}" url="${url}"`);
            } catch (e) {
                console.error(`  - <unreadable page: ${e.message}>`);
            }
        }
    }
    if (pages.length === 0) {
        throw new Error('CDP attach succeeded but no pages are exposed. Joplin may still be booting.');
    }

    // Score each page; higher is better.
    const scored = await Promise.all(pages.map(async (p) => {
        let url = '';
        let title = '';
        try { url = p.url(); } catch {}
        try { title = await p.title(); } catch {}
        let score = 0;
        if (/index\.html/i.test(url)) score += 10;
        if (/Joplin/i.test(title)) score += 5;
        if (url && url !== 'about:blank') score += 1;
        return { page: p, url, title, score };
    }));
    scored.sort((a, b) => b.score - a.score);
    const winner = scored[0];
    if (debug) {
        console.error(`eval-screenshot: picked page title="${winner.title}" url="${winner.url}" score=${winner.score}`);
    }
    return winner.page;
}

// Wait for Univer to render. Univer's spreadsheet UI mounts a canvas
// inside a container; the canvas is the load-bearing element. We
// also accept the broader Notesheet container as a fallback. If
// neither selector appears within the timeout, fall back to a fixed
// sleep — Univer's rendering can be quirky depending on plugin state,
// and we'd rather get a screenshot of "something rendered" than fail
// the evaluator on a selector mismatch.
async function waitForUniverRender(page) {
    const selectors = [
        // Univer spreadsheet's primary canvas; appears once cells are
        // laid out. (Univer uses a stacked-canvas renderer.)
        '.univer-render-canvas',
        'canvas.univer-render-canvas',
        // Notesheet plugin's editor wrapper (set by src/index.ts).
        '#joplin-plugin-content',
        // Generic Univer mount.
        '.univer-container',
    ];
    for (const sel of selectors) {
        try {
            await page.waitForSelector(sel, { timeout: 5_000, state: 'attached' });
            // Found it; give Univer a moment to actually paint cells.
            await page.waitForTimeout(2_000);
            return sel;
        } catch { /* try the next selector */ }
    }
    // Nothing matched. Document the gap and fall through to a fixed
    // delay so the screenshot still captures whatever's on screen.
    console.error('eval-screenshot: no Univer selector matched within 5s each; falling back to 5s sleep. Add a stable selector to waitForUniverRender() once you find one.');
    await page.waitForTimeout(5_000);
    return null;
}

async function main() {
    fs.mkdirSync(SHOT_DIR, { recursive: true });
    const utc = new Date().toISOString().replace(/[:.]/g, '-');
    const out = process.env.PGE_OUT || path.join(SHOT_DIR, `eval-${utc}.png`);
    // Make sure parent of PGE_OUT exists when caller overrode the path.
    fs.mkdirSync(path.dirname(out), { recursive: true });

    let playwright;
    try {
        playwright = require('playwright');
    } catch {
        console.error([
            'Playwright not installed. Run from repo root:',
            '    npm install --save-dev playwright',
        ].join('\n'));
        process.exit(3);
    }
    const { chromium } = playwright;

    // 1. Sanity-check CDP is reachable. (launch-joplin.sh should have
    // already verified this; we re-check so a stale Joplin is caught.)
    try {
        const r = await fetch(`${CDP_URL}/json/version`, { signal: AbortSignal.timeout(3000) });
        if (!r.ok) throw new Error(`CDP /json/version → ${r.status}`);
    } catch (e) {
        throw new Error(`CDP endpoint not reachable at ${CDP_URL}/json/version (${e.message}). Is Joplin running with --remote-debugging-port=${CDP_PORT}? Run scripts/pge/launch-joplin.sh.`);
    }

    // 2. Attach.
    console.error(`eval-screenshot: attaching CDP at ${CDP_URL}`);
    const browser = await chromium.connectOverCDP(CDP_URL);

    let savedPath = null;
    try {
        // Joplin exposes one BrowserContext (the default Electron
        // session); pages within it are the renderer windows.
        const contexts = browser.contexts();
        if (contexts.length === 0) {
            throw new Error('CDP attach succeeded but no BrowserContexts present.');
        }
        const context = contexts[0];
        const debug = !!process.env.PGE_DEBUG_CONTEXTS || contexts.length > 1 || context.pages().length > 1;
        if (contexts.length > 1) {
            console.error(`eval-screenshot: CDP exposes ${contexts.length} contexts; using the first.`);
        }

        const page = await pickEditorPage(context, debug);

        // 3. Open the test note (unless we're in PGE_OUT verification
        // mode without a feature-specific lookup — but we still take
        // whatever page is showing if PGE_NOTE_ID is unset and the
        // feature has no prefix, which is useful for verification).
        const prefix = TITLE_PREFIX_BY_FEATURE[FEATURE_ID];
        const explicitNote = process.env.PGE_NOTE_ID;
        if (explicitNote) {
            console.error(`eval-screenshot: opening note id ${explicitNote} (PGE_NOTE_ID)`);
            await openJoplinNote(explicitNote);
            await waitForUniverRender(page);
        } else if (prefix) {
            const hostPort = discoverApiPort();
            if (!hostPort) throw new Error('Joplin Web Clipper API not found via discover-api-port.sh');
            const token = discoverToken();
            if (!token) {
                throw new Error('No Joplin Web Clipper token found. Set JOPLIN_TOKEN or write the token to .claude/joplin-token.local. Get the token from Joplin → Settings → Web Clipper → Authorization tokens.');
            }
            const noteId = await findLatestNoteByTitle(hostPort, token, prefix);
            console.error(`eval-screenshot: opening note id ${noteId} (prefix="${prefix}")`);
            await openJoplinNote(noteId);
            await waitForUniverRender(page);
        } else {
            // No prefix table entry, no explicit note — verification
            // mode. Just screenshot whatever's on screen. Caller is
            // expected to set PGE_OUT to a /tmp path.
            console.error(`eval-screenshot: no title-prefix mapping for "${FEATURE_ID}" and PGE_NOTE_ID unset; capturing current page as-is (verification mode).`);
            await page.waitForTimeout(1_000);
        }

        // 4. Screenshot. We screenshot the whole page (not fullPage —
        // the renderer is a single viewport) so the eval includes
        // the editor chrome, sidebar, etc. — useful context for the
        // evaluator.
        await page.screenshot({ path: out, fullPage: false });
        savedPath = out;
        const stat = fs.statSync(out);
        console.error(`eval-screenshot: saved ${out} (${stat.size} bytes)`);
        if (stat.size < 1024) {
            throw new Error(`screenshot is suspiciously small (${stat.size} bytes); likely blank.`);
        }
    } finally {
        // Detach but DO NOT close Joplin.
        await browser.close();
    }

    if (savedPath) {
        process.stdout.write(savedPath + '\n');
    } else {
        process.exit(1);
    }
}

// Open a joplin:// URL via macOS `open`. Joplin registers the
// protocol handler and routes openNote to its internal navigation.
// We invoke it from outside the renderer (vs. window.evaluate) because
// our CDP-attached page is the renderer, and joplin:// is handled by
// the main process — the OS-level `open` is the cleanest dispatch.
function openJoplinNote(id) {
    return new Promise((resolve, reject) => {
        try {
            execSync(`open "joplin://x-callback-url/openNote?id=${id}"`, {
                stdio: ['ignore', 'ignore', 'pipe'],
                timeout: 5000,
            });
            resolve();
        } catch (e) {
            reject(new Error(`failed to dispatch joplin:// URL: ${e.message}`));
        }
    });
}

main().catch((e) => {
    console.error('eval-screenshot failed:', e.stack || e.message);
    process.exit(1);
});
