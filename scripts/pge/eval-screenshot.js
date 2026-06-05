// Evaluator's authoritative screenshot capture via Playwright + CDP
// attach. Joplin must already be running (started by launch-joplin.sh)
// with `--remote-debugging-port=$PGE_CDP_PORT`. We attach to its
// renderer using `chromium.connectOverCDP`, find the editor page,
// open the feature's test note via the joplin:// URL scheme, drop into
// the editor's UserWebviewIndex iframe (where Univer actually mounts),
// wait for the Univer canvas to attach AND be sized, sample row-0
// pixels for the machine-readable sidecar, then screenshot the whole
// editor page for human eyeball.
//
// Where Univer lives: NOT in the plugin sandbox CDP page (that page
// is `<body></body>` — plugin-process logic only, no UI). It lives
// inside `UserWebviewIndex.html`, which is a frame of the editor's
// main page. The plugin sandbox CDP page exists but doesn't host the
// view. This was the M13 false-blocker; the harness has no business
// trying to attach to the sandbox.
//
// Two outputs per run:
//   - <out>.png         — full editor screenshot (visual evidence)
//   - <out>.pixels.json — top-N non-background colours sampled from
//                         the Univer main canvas's row-0 slab
//                         (machine-checkable evidence). Future
//                         evaluators can assert "rgb(255,0,0)" appears
//                         in `top` without re-running the canvas pluck.
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
    console.error('usage: eval-screenshot.js <feature-id>[:variant]');
    process.exit(2);
}

// Variant suffix support — added for M13/E (first multi-screenshot
// cycle). Some features need two independent screenshots from one
// session (e.g. Aptos + Classic theme renderings). The convention is a
// `:variant` suffix on the feature id; the lookup tables below carry
// suffixed keys, and screenshots land under `screenshots/<base>/` with
// the variant baked into the filename.
const COLON = FEATURE_ID.indexOf(':');
const BASE_FEATURE_ID = COLON >= 0 ? FEATURE_ID.slice(0, COLON) : FEATURE_ID;
const VARIANT = COLON >= 0 ? FEATURE_ID.slice(COLON + 1) : null;

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const SHOT_DIR = path.join(REPO_ROOT, 'screenshots', BASE_FEATURE_ID);
const CDP_PORT = process.env.PGE_CDP_PORT || '8315';
const CDP_URL = `http://localhost:${CDP_PORT}`;

// Per-feature note title prefix. Each feature spec in BUILD_PLAN.md
// dictates what the title prefix should be; keep this table in sync
// with eval-screenshot.sh / the planner.
//
// Variant-suffixed keys (e.g. `feature-1-m13-theme-aware-banding:aptos`)
// map to per-variant title prefixes. The lookup falls back to the
// plain (un-suffixed) key when no `:variant` is provided. M13/E is the
// first cycle to exercise this — earlier cycles' single-key entries
// are intentionally preserved unchanged so the screenshot paths in
// PROGRESS.md `## Done` keep working.
const TITLE_PREFIX_BY_FEATURE = {
    'feature-1-smoke-red-cell': 'PGE smoke note ',
    'feature-1-m13-rotated-text-renders': 'PGE M13C eval ',
    'feature-1-m13-rich-text-renders': 'PGE M13D eval ',
    'feature-1-m13-theme-aware-banding:aptos': 'PGE M13E aptos eval ',
    'feature-1-m13-theme-aware-banding:classic': 'PGE M13E classic eval ',
    'feature-1-m15-conditional-formatting': 'PGE M15 CF eval ',
};

// Per-feature pixel-sampling region. Defaults to row-0 (the smoke
// pattern). Features whose visual evidence lives elsewhere on the
// canvas point at a different region helper. Variant-suffixed keys
// follow the same convention as TITLE_PREFIX_BY_FEATURE.
const REGION_BY_FEATURE = {
    'feature-1-smoke-red-cell': 'rowZero',
    'feature-1-m13-rotated-text-renders': 'rotatedRow',
    'feature-1-m13-rich-text-renders': 'richTextA1A2',
    'feature-1-m13-theme-aware-banding:aptos': 'tableHeaderRow',
    'feature-1-m13-theme-aware-banding:classic': 'tableHeaderRow',
    'feature-1-m15-conditional-formatting': 'cfAllColumns',
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

// Univer mounts inside the Joplin editor page's `UserWebviewIndex.html`
// frame, NOT in the plugin sandbox CDP page (the sandbox page is
// `<body></body>` — that's plugin process logic only, no UI). The
// editor view is a webview Joplin renders inside the main editor pane.
// We grab the frame by URL match.
//
// IMPORTANT: this can return null if the editor is showing a non-
// Notesheet note (or a markdown view of a Notesheet note). Callers
// must handle null and either retry after a navigation or fail loudly.
async function pickNotesheetWebview(page) {
    for (let attempt = 0; attempt < 30; attempt++) {
        const frame = page.frames().find((f) => /UserWebviewIndex\.html/.test(f.url()));
        if (frame) return frame;
        await page.waitForTimeout(200);
    }
    return null;
}

// Wait for Univer to actually render to the canvas. We probe selectors
// inside the UserWebviewIndex frame:
//   1. `canvas[id^="univer-sheet-main-canvas"]` — Univer 0.23's main
//      sheet canvas, id is `univer-sheet-main-canvas_<workbookId>`.
//      Load-bearing — this is what the user sees.
//   2. `[class*="univer-flex"]` — the Univer toolbar/header wrapper;
//      appears slightly before the canvas.
// We wait for the canvas to be attached AND have non-zero dimensions
// (Univer creates the element early, then sizes it asynchronously
// after layout).
async function waitForUniverRender(frame) {
    const sel = 'canvas[id^="univer-sheet-main-canvas"]';
    try {
        await frame.waitForSelector(sel, { timeout: 15_000, state: 'attached' });
    } catch {
        console.error(`eval-screenshot: Univer canvas selector "${sel}" did not appear within 15s.`);
        return null;
    }
    // Wait for the canvas to have non-zero size — Univer creates the
    // element first, then resizes it after the host frame settles.
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
        console.error('eval-screenshot: Univer canvas attached but never resized; rendering likely incomplete.');
    }
    // One frame to let any pending paint settle.
    await frame.waitForTimeout(250);
    return sel;
}

// Sample the rendered colour at the centre of cell A1 (row 0, col 0)
// from the Univer main canvas. Univer renders to canvas, so there is
// no per-cell DOM element — pixel sampling is the authoritative way
// to verify "did the cell visually render in the requested colour."
//
// Algorithm: find the most-common non-background, non-gridline pixel
// inside the bounding rect of A1 within the canvas. The "most common"
// rule is robust to anti-aliasing — fully-saturated red text on white
// background produces enough exact-#FF0000 pixels along stroke
// interiors to dominate any anti-aliased halo. Returns:
//   { dominant: 'rgb(R,G,B)', count, sampled, top: [['rgb(...)', n], ...] }
//
// We do NOT need the real cell rect from Univer's coordinate system —
// we sample a generous slab around row 0 (top ~60px of canvas).
// Future features that need precise per-cell rects can compute them
// from Univer's `defaultRowHeight` / `defaultColumnWidth` (snapshot
// fields).
// Sample five CF column bands (A, C, E, G, I) and return per-column
// pixel summaries plus an aggregated whole-CF-area summary. Used by
// the M15 cycle's `cfAllColumns` region kind — instead of one
// (regionFn) → one (summary) shape, this produces a `cfColumns`
// object keyed by column letter.
async function sampleCfColumns(frame, maxColors = 8) {
    return frame.evaluate(({ maxColors }) => {
        const canvas = document.querySelector('canvas[id^="univer-sheet-main-canvas"]');
        if (!canvas) return { error: 'no main canvas' };
        const dpr = canvas.clientWidth > 0 ? canvas.width / canvas.clientWidth : 1;
        const ROW_HEADER_W_CSS = 46;
        const COL_W_CSS = 73;
        const COL_HEADER_H_CSS = 18;
        const ROW_H_CSS = 19;
        const COLS = { A: 0, C: 2, E: 4, G: 6, I: 8 };

        function sampleRegion(rx, ry, rw, rh) {
            const off = document.createElement('canvas');
            off.width = rw;
            off.height = rh;
            const ctx = off.getContext('2d');
            ctx.drawImage(canvas, rx, ry, rw, rh, 0, 0, rw, rh);
            const data = ctx.getImageData(0, 0, rw, rh).data;
            const hist = new Map();
            const inkY = new Set();
            let sampled = 0;
            let redInk = 0, blueInk = 0, greenInk = 0, greyInk = 0;
            let pinkInk = 0, lightGreenInk = 0, yellowInk = 0;
            for (let y = 0; y < rh; y += 1) {
                for (let x = 0; x < rw; x += 1) {
                    const i = (y * rw + x) * 4;
                    const r = data[i], g = data[i + 1], b = data[i + 2], a = data[i + 3];
                    if (a < 200) continue;
                    if (r > 235 && g > 235 && b > 235) continue;
                    if (r < 30 && g < 30 && b < 30) continue;
                    sampled++;
                    inkY.add(y);
                    // M13/D originally required pure red (g/b ≤ 80); M15
                // broadens to g/b ≤ 140 so the colorScale's #F8696B
                // family (R=248, G=105, B=107) qualifies. Pure red text
                // (rgb(255,0,0)) still passes; Aptos accent3 dark green
                // (rgb(25,107,36)) doesn't (R=25 < 200). The threshold
                // change is monotonic (only widens), so no prior cycle's
                // gate that already passed can fail.
                if (r >= 200 && g <= 140 && b <= 140) redInk++;
                    // M15 broadens blueInk: dataBar colour #638EC6
                // (R=99, G=142, B=198) is the spec target. The
                // original M13/D thresholds (r<=80 AND g<=80 AND
                // b>=200) only matched pure blue text. We loosen to
                // "B clearly dominant by ≥30 over R AND G, AND B at
                // least 150" — catches every saturated blue from
                // pure rgb(0,0,255) down through Excel's dataBar
                // accent. Pure blue still passes; the change is
                // monotonic so no prior gate regresses.
                if (b > r + 30 && b > g + 30 && b >= 150) blueInk++;
                    if (g > r + 30 && g > b + 30 && g >= 80) greenInk++;
                    if (
                        r >= 140 && r <= 180
                        && g >= 140 && g <= 180
                        && b >= 140 && b <= 180
                        && Math.abs(r - g) <= 10
                        && Math.abs(g - b) <= 10
                    ) greyInk++;
                    if (r >= 220 && g >= 180 && g <= 220 && b >= 180 && b <= 220) pinkInk++;
                    if (r >= 180 && r <= 220 && g >= 220 && b >= 180 && b <= 220) lightGreenInk++;
                    if (r >= 200 && g >= 200 && b <= 120) yellowInk++;
                    const key = `rgb(${r},${g},${b})`;
                    hist.set(key, (hist.get(key) || 0) + 1);
                }
            }
            const top = Array.from(hist.entries()).sort((a, b) => b[1] - a[1]).slice(0, maxColors);
            const [dominant, count] = top[0] || [null, 0];
            return {
                dominant, count, sampled, top,
                regionX: rx, regionY: ry, regionWidth: rw, regionHeight: rh,
                inkRows: inkY.size,
                inkRowSpread: inkY.size / Math.max(1, rh),
                redInk, blueInk, greenInk, greyInk,
                pinkInk, lightGreenInk, yellowInk,
            };
        }

        const cfColumns = {};
        for (const col of Object.keys(COLS)) {
            const colIndex = COLS[col];
            const xCss = ROW_HEADER_W_CSS + colIndex * COL_W_CSS;
            const yCss = COL_HEADER_H_CSS + 1 * ROW_H_CSS;
            const wCss = COL_W_CSS;
            const hCss = 10 * ROW_H_CSS;
            const x = Math.round(xCss * dpr);
            const y = Math.round(yCss * dpr);
            const w = Math.round(wCss * dpr);
            const h = Math.round(hCss * dpr);
            const cw = Math.min(canvas.width - 1, x);
            const ch = Math.min(canvas.height - 1, y);
            const useW = Math.min(Math.max(canvas.width - x, 0), w);
            const useH = Math.min(Math.max(canvas.height - y, 0), h);
            cfColumns[col] = sampleRegion(cw, ch, useW, useH);
        }

        // Aggregate: a whole-band sample covering A2:I11 for the
        // top-level `dominant` and `top` histogram (legacy compat).
        const xCss = ROW_HEADER_W_CSS;
        const yCss = COL_HEADER_H_CSS + 1 * ROW_H_CSS;
        const wCss = 9 * COL_W_CSS;
        const hCss = 10 * ROW_H_CSS;
        const x = Math.round(xCss * dpr);
        const y = Math.round(yCss * dpr);
        const w = Math.round(wCss * dpr);
        const h = Math.round(hCss * dpr);
        const aggW = Math.min(Math.max(canvas.width - x, 0), w);
        const aggH = Math.min(Math.max(canvas.height - y, 0), h);
        const agg = sampleRegion(x, y, aggW, aggH);

        return { ...agg, cfColumns };
    }, { maxColors });
}

async function samplePixelsAt(frame, regionFn, maxColors = 8) {
    return frame.evaluate(({ regionFnSrc, maxColors }) => {
        const canvas = document.querySelector('canvas[id^="univer-sheet-main-canvas"]');
        if (!canvas) return { error: 'no main canvas' };
        const region = (new Function('return ' + regionFnSrc))()(canvas);
        const off = document.createElement('canvas');
        off.width = region.w;
        off.height = region.h;
        const ctx = off.getContext('2d');
        ctx.drawImage(canvas, region.x, region.y, region.w, region.h, 0, 0, region.w, region.h);
        const data = ctx.getImageData(0, 0, region.w, region.h).data;
        const hist = new Map();
        // rowsWithInk: how many distinct y-pixels carry text-coloured
        // ink. For horizontal text in a row, ink concentrates on a
        // narrow horizontal band (roughly the glyph height ~ 14-18px).
        // Rotated text spreads ink across many more y-pixels — that
        // spread is the rotation-positive signal independent of
        // colour.
        const inkY = new Set();
        let sampled = 0;
        // Aggregated colour bands (match the spec thresholds exactly,
        // independent of histogram bucketing). Anti-aliased glyph
        // edges spread saturated colour over many near-but-not-exact
        // RGB buckets, so the per-bucket histogram can under-report
        // the visual presence of a colour. These aggregates count
        // every pixel that satisfies the inequality, regardless of
        // exact RGB.
        // Red ink:   R >= 200 AND G <=  80 AND B <=  80
        // Blue ink:  R <=  80 AND G <=  80 AND B >= 200
        // Green ink: G > R+30 AND G > B+30 AND G >= 80
        //            (broader than M13/D's R<=80 AND G>=150 AND B<=80
        //            so the dark-but-saturated Aptos accent3 #196B24 =
        //            rgb(25,107,36) — used as the M13/E table-style
        //            HEADER fill — qualifies. The original tight
        //            threshold only matched pure greens like (0,255,0)
        //            and pastel #84E291 (132,226,145) failed too. We
        //            generalise to "green channel clearly dominant by
        //            ≥30 over red AND blue, AND green at least 80" —
        //            captures every saturated green in Excel's
        //            built-in TableStyle palette.)
        // Grey ink:  R, G, B all in [140, 180] AND
        //            abs(R-G) <= 10 AND abs(G-B) <= 10
        //            (all three channels mid-range AND mutually close
        //            — distinguishes grey from a tinted hue at similar
        //            luminance; M13/E uses this to gate the Classic
        //            fixture's grey header rendering.)
        let redInk = 0, blueInk = 0, greenInk = 0, greyInk = 0;
        // M15 CF aggregates. pinkInk targets #FFC7CE (cellIs > 50 fill);
        // lightGreenInk targets #C6EFCE (top-3 rank fill); yellowInk
        // targets the iconSet middle band's gold arrow (RGB roughly
        // 255,189,55 in Univer's ICON_MAP arrow["right-gold"]).
        // Loose-by-design — glyphs are small and anti-aliased.
        let pinkInk = 0, lightGreenInk = 0, yellowInk = 0;
        // Stride 1 (every pixel) — region is small and we want the
        // colour signal to cross the spec thresholds even on
        // narrow-glyph runs.
        for (let y = 0; y < region.h; y += 1) {
            for (let x = 0; x < region.w; x += 1) {
                const i = (y * region.w + x) * 4;
                const r = data[i], g = data[i + 1], b = data[i + 2], a = data[i + 3];
                if (a < 200) continue;
                if (r > 235 && g > 235 && b > 235) continue;       // background
                if (r < 30 && g < 30 && b < 30) continue;          // gridline
                sampled++;
                inkY.add(y);
                // M13/D originally required pure red (g/b ≤ 80); M15
                // broadens to g/b ≤ 140 so the colorScale's #F8696B
                // family (R=248, G=105, B=107) qualifies. Pure red text
                // (rgb(255,0,0)) still passes; Aptos accent3 dark green
                // (rgb(25,107,36)) doesn't (R=25 < 200). The threshold
                // change is monotonic (only widens), so no prior cycle's
                // gate that already passed can fail.
                if (r >= 200 && g <= 140 && b <= 140) redInk++;
                // M15 broadens blueInk: dataBar colour #638EC6
                // (R=99, G=142, B=198) is the spec target. The
                // original M13/D thresholds (r<=80 AND g<=80 AND
                // b>=200) only matched pure blue text. We loosen to
                // "B clearly dominant by ≥30 over R AND G, AND B at
                // least 150" — catches every saturated blue from
                // pure rgb(0,0,255) down through Excel's dataBar
                // accent. Pure blue still passes; the change is
                // monotonic so no prior gate regresses.
                if (b > r + 30 && b > g + 30 && b >= 150) blueInk++;
                if (g > r + 30 && g > b + 30 && g >= 80) greenInk++;
                if (
                    r >= 140 && r <= 180
                    && g >= 140 && g <= 180
                    && b >= 140 && b <= 180
                    && Math.abs(r - g) <= 10
                    && Math.abs(g - b) <= 10
                ) greyInk++;
                // Pink: #FFC7CE family — R high, G in [180,220], B in [180,220].
                if (r >= 220 && g >= 180 && g <= 220 && b >= 180 && b <= 220) pinkInk++;
                // Light green: #C6EFCE family — R in [180,220], G high, B in [180,220].
                if (r >= 180 && r <= 220 && g >= 220 && b >= 180 && b <= 220) lightGreenInk++;
                // Yellow: R high, G high, B low (gold arrow / yellow-flat icon).
                if (r >= 200 && g >= 200 && b <= 120) yellowInk++;
                const key = `rgb(${r},${g},${b})`;
                hist.set(key, (hist.get(key) || 0) + 1);
            }
        }
        const top = Array.from(hist.entries()).sort((a, b) => b[1] - a[1]).slice(0, maxColors);
        const [dominant, count] = top[0] || [null, 0];
        return {
            dominant,
            count,
            sampled,
            top,
            regionWidth: region.w,
            regionHeight: region.h,
            inkRows: inkY.size,
            inkRowSpread: inkY.size / Math.max(1, region.h),
            // Aggregated colour-band counts using spec thresholds.
            // Use these when the per-bucket `top` histogram is fragmented
            // by anti-aliasing (e.g. saturated colour text inside a
            // small region). For the M13/D rich-text gate the spec
            // requires redInk >= 30 and blueInk >= 30 within the A2
            // band.
            redInk,
            blueInk,
            greenInk,
            greyInk,
            pinkInk,
            lightGreenInk,
            yellowInk,
        };
    }, { regionFnSrc: regionFn.toString(), maxColors });
}

// Region of canvas covering row 0 (cell A1 column area). Univer renders
// row headers and column headers occupy ~25-30px each at default zoom;
// we use a generous slab to absorb that.
function rowZeroRegion(canvas) {
    return { x: 0, y: 0, w: Math.min(canvas.width, 400), h: Math.min(canvas.height, 80) };
}

// Region covering rows 5–6 of the canvas (zero-indexed: rows 4–6 in
// Univer's 0-based indexing, which lands on the rotated-text band of
// the MergedCellsAndAlignment fixture). Default row height is 19px and
// the column header takes ~25px; rotated text expands the row height
// further. We grab a generous slab from y=120 to y=320 covering
// roughly rows 4–6.
function rotatedRowRegion(canvas) {
    return {
        x: 0,
        y: Math.min(canvas.height - 1, 120),
        w: Math.min(canvas.width, 600),
        h: Math.min(Math.max(canvas.height - 120, 0), 200),
    };
}

// Region covering A2 only on the "RichText" sheet of the M13/D
// fixture — A2 carries the multi-colour pin-down (`Red` red + ` and `
// default + `Blue` blue + ` text` default) and is the cleanest
// signal source for the per-run colour gate. We deliberately exclude
// A3 (single-format hyperlink — ink is also blue, would alias the
// A2 `Blue` run) and A1 (whose active-cell selection border
// `rgb(44,83,241)` saturates the histogram with blue pixels even
// when the cell is plain).
//
// Empirical y-band on the running fixture: column header ~y=0–18,
// A1 ~y=22–37, A2 ~y=41–57, A3 ~y=58+. A2's text band (after
// excluding the cell border at y=38–40) is roughly y=41–57. We use
// y=41–58 with a 1px buffer to absorb anti-aliasing on the lower
// edge.
function richTextA1A2Region(canvas) {
    return {
        x: 0,
        y: Math.min(canvas.height - 1, 41),
        w: Math.min(canvas.width, 400),
        h: Math.min(Math.max(canvas.height - 41, 0), 17),
    };
}

// Region covering the header band of a built-in TableStyle table whose
// data range starts at A1. M13/E samples this band on both
// FormattingSmorgasboard fixtures: Aptos must show greenInk, Classic
// must show greyInk.
//
// y-band rationale:
//   - Univer's column header bar consumes ~y=0–18 at default zoom.
//   - The data area's row 1 (the table header row) is ~y=19–37 at
//     default row height (19px).
//   - Univer's active-cell selection paints `rgb(44,83,241)` blue
//     around A1 by default, contributing blue pixels at y≈19–20 and
//     y≈37–38 (top + bottom of the cell border).
// We sample y=22–35 (h=13) — sits inside the header-row text band but
// excludes the active-cell selection border on top/bottom.
//
// x-band rationale:
//   - The active-cell selection border also paints down the LEFT and
//     RIGHT edges of A1 (x≈0–1 and x≈73–74 at default col width).
//   - Sampling ALL columns of the header would mostly catch the
//     accent fill; the selection-border noise on A1 is small relative
//     to the band fill area, but to be safe we start the x sample at
//     col B (x≈74) so the active-cell A1 selection does NOT pollute
//     the colour gate.
//   - Default column width is 73px; col B starts at x≈74. We sample
//     x=80..min(width, 80+400)=480, covering cols B..G (or B..F for
//     the Classic fixture's 6-col table) — well past the active-cell
//     selection.
// Region covering one CF column band (rows 2-11 of the
// ConditionalFormatting-Variants fixture, after the header row at row
// 1). Five named columns: A, C, E, G, I. Univer's column header strip
// occupies y≈0–18 at default zoom; the header row at row 1 sits at
// y≈19–37; the data rows we care about (rows 2–11, i.e. snapshot rows
// 1–10) span ~y=37 to ~y=37+10*19=227 at default row height 19px CSS
// units.
//
// x positions: row-header column to the LEFT of A is ~46px wide; col A
// starts at x≈46, default col width 73px. Columns A, C, E, G, I are at
// snapshot indices 0, 2, 4, 6, 8 → x ≈ 46 + index*73.
//
// DPR-aware: scale by canvas.width / canvas.clientWidth per M13/E
// Phase 4 lesson; CSS px multiplied by DPR ratio at sample time.
function cfColumnRegion(canvas, col) {
    const dpr = canvas.clientWidth > 0 ? canvas.width / canvas.clientWidth : 1;
    const colIndex = ({ A: 0, C: 2, E: 4, G: 6, I: 8 })[col];
    const ROW_HEADER_W_CSS = 46;
    const COL_W_CSS = 73;
    const COL_HEADER_H_CSS = 18;
    const ROW_H_CSS = 19;
    const xCss = ROW_HEADER_W_CSS + colIndex * COL_W_CSS;
    // Skip the header row (row 1, y≈19..37) — start sampling at y≈37
    // so the colour signal is the CF fill, not the column-letter
    // string. Sample down through 10 data rows.
    const yCss = COL_HEADER_H_CSS + 1 * ROW_H_CSS;
    const wCss = COL_W_CSS;
    const hCss = 10 * ROW_H_CSS;
    const x = Math.round(xCss * dpr);
    const y = Math.round(yCss * dpr);
    const w = Math.round(wCss * dpr);
    const h = Math.round(hCss * dpr);
    return {
        x: Math.min(canvas.width - 1, x),
        y: Math.min(canvas.height - 1, y),
        w: Math.min(Math.max(canvas.width - x, 0), w),
        h: Math.min(Math.max(canvas.height - y, 0), h),
    };
}

function tableHeaderRowRegion(canvas) {
    // Univer's canvas has a backing store sized at devicePixelRatio
    // multiples of its CSS box. On a Retina display the canvas is 2x
    // wider/taller than at default DPR=1, so a hard-coded y=22 lands
    // INSIDE the column-letter strip instead of on the table header
    // row. Scale the region by the actual DPR ratio recovered from
    // `canvas.width / canvas.clientWidth`. clientWidth is the CSS
    // size; the ratio is 1 on standard displays and 2 on Retina.
    const dpr = canvas.clientWidth > 0
        ? canvas.width / canvas.clientWidth
        : 1;
    const x = Math.round(80 * dpr);
    const y = Math.round(22 * dpr);
    const w = Math.round(400 * dpr);
    const h = Math.round(13 * dpr);
    return {
        x: Math.min(canvas.width - 1, x),
        y: Math.min(canvas.height - 1, y),
        w: Math.min(Math.max(canvas.width - x, 0), w),
        h: Math.min(Math.max(canvas.height - y, 0), h),
    };
}

async function main() {
    fs.mkdirSync(SHOT_DIR, { recursive: true });
    const utc = new Date().toISOString().replace(/[:.]/g, '-');
    const namePrefix = VARIANT ? `eval-${VARIANT}` : 'eval';
    const out = process.env.PGE_OUT || path.join(SHOT_DIR, `${namePrefix}-${utc}.png`);
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

        // Belt-and-braces: wait for the renderer to settle before doing
        // anything else. A freshly-attached CDP session can land mid-
        // navigation (Joplin's startup loads a few async chunks); without
        // this gate, `joplin://` dispatch + screenshot can race the load
        // and capture a half-rendered editor. Bounded so a stuck network
        // doesn't hang the eval forever.
        try {
            await page.waitForLoadState('networkidle', { timeout: 10_000 });
        } catch {
            console.error('eval-screenshot: networkidle wait timed out at 10s; proceeding anyway.');
        }

        // 3. Open the test note (unless we're in PGE_OUT verification
        // mode without a feature-specific lookup — but we still take
        // whatever page is showing if PGE_NOTE_ID is unset and the
        // feature has no prefix, which is useful for verification).
        // Variant-suffixed key first, then the plain feature id. Plain-key
        // fallback preserves prior-cycle behaviour for un-suffixed
        // single-screenshot features.
        const prefix = TITLE_PREFIX_BY_FEATURE[FEATURE_ID] || TITLE_PREFIX_BY_FEATURE[BASE_FEATURE_ID];
        const explicitNote = process.env.PGE_NOTE_ID;
        let didOpenNote = false;
        if (explicitNote) {
            console.error(`eval-screenshot: opening note id ${explicitNote} (PGE_NOTE_ID)`);
            await openJoplinNote(explicitNote);
            didOpenNote = true;
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
            didOpenNote = true;
        } else {
            console.error(`eval-screenshot: no title-prefix mapping for "${FEATURE_ID}" and PGE_NOTE_ID unset; capturing current page as-is (verification mode).`);
            await page.waitForTimeout(1_000);
        }

        // 4. If a Notesheet note was opened, drop into the
        // UserWebviewIndex frame and wait for Univer to actually paint.
        // This is a robust stop-gap for the M13 failure mode: it
        // forces "the canvas exists AND has been sized" before we
        // screenshot, instead of a fixed sleep.
        let pixelSummary = null;
        let captureWebview = null;
        let captureCanvas = null;
        // Resolve the region kind: variant-suffixed key first, then the
        // plain feature id, then the row-0 default.
        const regionKind = REGION_BY_FEATURE[FEATURE_ID] || REGION_BY_FEATURE[BASE_FEATURE_ID] || 'rowZero';
        // For 'cfAllColumns' regionKind we use the dedicated multi-region
        // sampler `sampleCfColumns` below — no single-region regionFn.
        const regionFn = regionKind === 'rotatedRow'
            ? rotatedRowRegion
            : regionKind === 'richTextA1A2'
                ? richTextA1A2Region
                : regionKind === 'tableHeaderRow'
                    ? tableHeaderRowRegion
                    : rowZeroRegion;
        if (didOpenNote) {
            captureWebview = await pickNotesheetWebview(page);
            if (!captureWebview) {
                console.error('eval-screenshot: UserWebviewIndex frame did not appear; the opened note may not be a Notesheet, or the plugin failed to load.');
            } else {
                const sel = await waitForUniverRender(captureWebview);
                if (sel) {
                    console.error(`eval-screenshot: Univer canvas rendered (selector "${sel}")`);
                    captureCanvas = sel;
                    // Re-pick the frame just before sampling. When opening
                    // a different note via joplin://, the editor may swap
                    // out the UserWebviewIndex iframe; the original frame
                    // reference can become stale even though
                    // waitForSelector found a canvas. Fetching the frame
                    // again at sample time pins us to the live one.
                    let sampleFrame = captureWebview;
                    for (let attempt = 0; attempt < 6; attempt++) {
                        try {
                            if (regionKind === 'cfAllColumns') {
                                pixelSummary = await sampleCfColumns(sampleFrame);
                            } else {
                                pixelSummary = await samplePixelsAt(sampleFrame, regionFn);
                            }
                            if (pixelSummary && !pixelSummary.error) break;
                        } catch (e) {
                            console.error(`eval-screenshot: pixel sample attempt ${attempt} threw: ${e.message}`);
                        }
                        await page.waitForTimeout(300);
                        const refreshed = await pickNotesheetWebview(page);
                        if (refreshed) {
                            sampleFrame = refreshed;
                            captureWebview = refreshed;
                        }
                    }
                    if (pixelSummary && pixelSummary.error) {
                        console.error(`eval-screenshot: pixel sample returned error after retries: ${pixelSummary.error}`);
                    }
                }
            }
        }

        // 5. Screenshot. If a Notesheet canvas is in scope, screenshot
        // the canvas element directly — it's what the user sees and
        // it sidesteps Joplin's window-pane cropping (the editor pane
        // is often narrower than the canvas needs). Otherwise fall
        // back to the whole page (smoke / verification mode).
        if (captureWebview && captureCanvas) {
            await captureWebview.locator(captureCanvas).first().screenshot({ path: out });
        } else {
            await page.screenshot({ path: out, fullPage: false });
        }
        savedPath = out;
        const stat = fs.statSync(out);
        console.error(`eval-screenshot: saved ${out} (${stat.size} bytes)`);
        if (stat.size < 1024) {
            throw new Error(`screenshot is suspiciously small (${stat.size} bytes); likely blank.`);
        }

        // 6. Pixel-summary sidecar JSON. Future feature evaluators can
        // assert against this without doing their own canvas pluck.
        if (pixelSummary) {
            const sidecar = out.replace(/\.png$/, '.pixels.json');
            const regionLabel = regionKind === 'rotatedRow'
                ? 'rotated row band (y 120–320, slab covering rows 4–6 of the fixture)'
                : regionKind === 'richTextA1A2'
                    ? 'rich-text A2 band (y 41–58, multi-colour pin-down: Red+default+Blue+default)'
                    : regionKind === 'tableHeaderRow'
                        ? 'table header band (x 80–480, y 22–35; cols B+ of A1:_ table header row, excludes A1 active-cell selection border)'
                        : regionKind === 'cfAllColumns'
                            ? 'CF all columns (5 sub-regions A/C/E/G/I, rows 2-11; aggregate is the whole A2:I11 band)'
                            : 'row-0 (top 80px slab of main canvas)';
            fs.writeFileSync(sidecar, JSON.stringify({
                source: out,
                region: regionLabel,
                regionKind,
                ...pixelSummary,
            }, null, 2));
            console.error(`eval-screenshot: pixel summary → ${sidecar}`);
            console.error(`eval-screenshot:   dominant=${pixelSummary.dominant} count=${pixelSummary.count} sampled=${pixelSummary.sampled} inkRows=${pixelSummary.inkRows} inkRowSpread=${pixelSummary.inkRowSpread?.toFixed(3)} redInk=${pixelSummary.redInk} blueInk=${pixelSummary.blueInk} greenInk=${pixelSummary.greenInk} greyInk=${pixelSummary.greyInk}`);
            if (pixelSummary.cfColumns) {
                for (const col of Object.keys(pixelSummary.cfColumns)) {
                    const c = pixelSummary.cfColumns[col];
                    console.error(`eval-screenshot:   cfColumn[${col}]: dominant=${c.dominant} sampled=${c.sampled} redInk=${c.redInk} blueInk=${c.blueInk} greenInk=${c.greenInk} pinkInk=${c.pinkInk} lightGreenInk=${c.lightGreenInk} yellowInk=${c.yellowInk}`);
                }
            }
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
