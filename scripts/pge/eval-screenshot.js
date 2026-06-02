// Evaluator's authoritative screenshot capture. Launches Joplin via
// Playwright's Electron driver with the dev profile, waits for the
// Web Clipper API to come up (proves the app finished booting),
// creates or selects the feature's test note via the API + the
// joplin:// URL scheme, waits for Univer to render, and saves a
// screenshot under screenshots/<feature-id>/eval-<utc>.png.
//
// The evaluator agent invokes this via Bash, then opens the
// resulting PNG with the Read tool and judges its contents.
//
// Usage:
//   node eval-screenshot.js <feature-id>

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
const JOPLIN_PATH = process.env.JOPLIN_PATH ||
    '/Applications/Joplin.app/Contents/MacOS/Joplin';

// Per-feature note title prefix. The smoke creates notes with titles
// starting "PGE smoke note ". When we extend to real M13 features,
// each spec specifies a title prefix.
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
    const tokenFile = path.resolve(REPO_ROOT, '.claude', 'joplin-token.local');
    try { return fs.readFileSync(tokenFile, 'utf8').trim(); }
    catch { return ''; }
}

async function waitForApi(timeoutMs = 30000) {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
        const hostPort = discoverApiPort();
        if (hostPort) {
            try {
                const r = await fetch(`http://${hostPort}/ping`, { signal: AbortSignal.timeout(2000) });
                if (r.ok) return hostPort;
            } catch { /* not ready, retry */ }
        }
        await new Promise((r) => setTimeout(r, 1000));
    }
    throw new Error(`Joplin Web Clipper API did not respond within ${timeoutMs / 1000}s`);
}

async function findLatestNoteByTitle(hostPort, token, prefix) {
    const q = new URLSearchParams({ fields: 'id,title,updated_time' });
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

async function main() {
    fs.mkdirSync(SHOT_DIR, { recursive: true });
    const utc = new Date().toISOString().replace(/[:.]/g, '-');
    const out = path.join(SHOT_DIR, `eval-${utc}.png`);

    let playwright;
    try {
        playwright = require('playwright');
    } catch {
        console.error([
            'Playwright not installed. Run from repo root:',
            '    npm install --save-dev playwright',
            '',
            'Playwright bundles its Electron driver with the base',
            'install — no separate package needed.',
        ].join('\n'));
        process.exit(3);
    }

    // 1. Launch Joplin (must NOT be already running — Electron's
    // single-instance lock would error out).
    const { _electron: electron } = playwright;
    console.error('eval-screenshot: launching Joplin --env dev via Electron...');
    const app = await electron.launch({
        executablePath: JOPLIN_PATH,
        args: ['--env', 'dev'],
        timeout: 30_000,
    });

    let savedPath = null;
    try {
        const window = await app.firstWindow({ timeout: 30_000 });
        await window.waitForLoadState('domcontentloaded', { timeout: 30_000 });

        // 2. Wait for the Web Clipper Data API to come up inside the
        // launched Joplin. This is also our signal that the editor
        // service is alive.
        const hostPort = await waitForApi(45_000);
        console.error(`eval-screenshot: API up at http://${hostPort}`);

        const token = discoverToken();
        if (!token) {
            throw new Error(
                'No Joplin Web Clipper token found. Set JOPLIN_TOKEN env var, or write the token to .claude/joplin-token.local. Get the token from Joplin → Settings → Web Clipper → Authorization tokens.'
            );
        }

        // 3. Find the test note for this feature.
        const prefix = TITLE_PREFIX_BY_FEATURE[FEATURE_ID];
        if (!prefix) {
            throw new Error(
                `eval-screenshot.js does not know the title prefix for "${FEATURE_ID}". Add it to TITLE_PREFIX_BY_FEATURE.`
            );
        }
        const noteId = process.env.PGE_NOTE_ID || await findLatestNoteByTitle(hostPort, token, prefix);
        console.error(`eval-screenshot: opening note id ${noteId}`);

        // 4. Open the note via Joplin's URL scheme. Joplin's main
        // process registers `joplin://` and routes openNote to its
        // internal navigation. We dispatch via the Electron app's
        // first window.
        await window.evaluate((id) => {
            const a = document.createElement('a');
            a.href = `joplin://x-callback-url/openNote?id=${id}`;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            a.remove();
        }, noteId);

        // 5. Give Univer time to hydrate the editor.
        await window.waitForTimeout(5_000);

        // 6. Screenshot the active window.
        await window.screenshot({ path: out, fullPage: false });
        savedPath = out;
        console.error(`eval-screenshot: saved ${out}`);
    } finally {
        await app.close();
    }

    if (savedPath) {
        process.stdout.write(savedPath + '\n');
    } else {
        process.exit(1);
    }
}

main().catch((e) => {
    console.error('eval-screenshot failed:', e.stack || e.message);
    process.exit(1);
});
