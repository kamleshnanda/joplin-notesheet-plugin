// Evaluator's authoritative screenshot capture. Launches Joplin via
// Playwright's Electron driver, opens the test note for a feature,
// waits for Univer to render, screenshots the editor pane, saves
// under screenshots/<feature-id>/eval-<utc-timestamp>.png.
//
// The evaluator agent invokes this via Bash; the agent then opens
// the resulting PNG with the Read tool and judges its contents.
//
// Usage:
//   node eval-screenshot.js <feature-id>
//
// Configuration (env):
//   JOPLIN_API   — Joplin Web Clipper Data API base, default
//                  http://localhost:41184. Used to find the most
//                  recent test note matching a per-feature title
//                  prefix.
//   JOPLIN_PATH  — Path to Joplin's Electron binary (default
//                  /Applications/Joplin.app/Contents/MacOS/Joplin).
//   PGE_NOTE_ID  — If set, opens this note id directly instead of
//                  searching by title.
//
// Output:
//   - screenshots/<feature-id>/eval-<utc>.png on success.
//   - Non-zero exit + stderr on failure.

const fs = require('fs');
const path = require('path');

const FEATURE_ID = process.argv[2];
if (!FEATURE_ID) {
    console.error('usage: eval-screenshot.js <feature-id>');
    process.exit(2);
}

const REPO_ROOT = path.resolve(__dirname, '..', '..');
const SHOT_DIR = path.join(REPO_ROOT, 'screenshots', FEATURE_ID);
const JOPLIN_PATH = process.env.JOPLIN_PATH ||
    '/Applications/Joplin.app/Contents/MacOS/Joplin';

// Per-feature note title convention. The smoke creates notes with
// titles starting "PGE smoke note ". When we extend to real M13
// features, each feature spec specifies its own title prefix and
// this map gets that prefix.
const TITLE_PREFIX_BY_FEATURE = {
    'feature-1-smoke-red-cell': 'PGE smoke note ',
};

async function findLatestNoteByTitle(prefix) {
    const api = process.env.JOPLIN_API || 'http://localhost:41184';
    const tok = process.env.JOPLIN_TOKEN || '';
    const q = new URLSearchParams({ fields: 'id,title,updated_time' });
    if (tok) q.set('token', tok);
    const res = await fetch(`${api}/notes?${q}`);
    if (!res.ok) throw new Error(`Joplin API list-notes failed: ${res.status}`);
    const body = await res.json();
    const matches = (body.items ?? [])
        .filter((n) => (n.title || '').startsWith(prefix))
        .sort((a, b) => (b.updated_time ?? 0) - (a.updated_time ?? 0));
    if (matches.length === 0) {
        throw new Error(`no Joplin note with title prefix "${prefix}" — generator should have created one`);
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
            '    npx playwright install chromium',
            '',
            'Playwright bundles its Electron driver with the base',
            'install — no separate package needed.',
        ].join('\n'));
        process.exit(3);
    }

    let noteId = process.env.PGE_NOTE_ID;
    if (!noteId) {
        const prefix = TITLE_PREFIX_BY_FEATURE[FEATURE_ID];
        if (!prefix) {
            throw new Error(
                `eval-screenshot.js does not yet know the title prefix for ` +
                `feature "${FEATURE_ID}". Add it to TITLE_PREFIX_BY_FEATURE ` +
                `at the top of this file.`
            );
        }
        noteId = await findLatestNoteByTitle(prefix);
    }
    console.error(`eval-screenshot: using note id ${noteId}`);

    // Launch Joplin via Electron with the dev profile. Joplin must
    // NOT already be running — Playwright's _electron.launch spawns
    // a new instance, and a running Joplin will error out trying to
    // acquire the single-instance lock. The wrapper script handles
    // this by detecting a running Joplin and asking the operator to
    // quit it first.
    //
    // `--env dev` selects ~/.config/joplindev-desktop/, isolating PGE
    // cycles from the operator's main Joplin notes.
    const { _electron: electron } = playwright;
    const app = await electron.launch({
        executablePath: JOPLIN_PATH,
        args: ['--env', 'dev'],
        timeout: 30_000,
    });

    try {
        // Wait for the main window. Joplin's first window is the
        // editor; its title contains "Joplin".
        const window = await app.firstWindow({ timeout: 30_000 });
        await window.waitForLoadState('domcontentloaded', { timeout: 30_000 });

        // Joplin doesn't accept an "open this note" CLI arg; instead
        // we navigate via its in-app routing. Easiest path: simulate
        // Cmd+G "Goto note ID" doesn't exist either. So we use the
        // "search" sidebar to find the note by id (which Joplin's
        // search supports as `id:<id>`). If that fails, fall back to
        // clicking the most-recent note in the All Notes view.
        // For the smoke, we'll just wait and assume the operator's
        // "selected" note in Joplin's sidebar is what's open.
        //
        // TODO: make this deterministic. For the smoke we accept
        // that the generator must select the test note in Joplin
        // before the evaluator runs.
        await window.waitForTimeout(3_000); // give Univer time to hydrate

        await window.screenshot({ path: out, fullPage: false });
        console.error(`eval-screenshot: saved ${out}`);
    } finally {
        await app.close();
    }

    // Print the saved path on stdout for the evaluator agent to read.
    process.stdout.write(out + '\n');
}

main().catch((e) => {
    console.error('eval-screenshot failed:', e.stack || e.message);
    process.exit(1);
});
