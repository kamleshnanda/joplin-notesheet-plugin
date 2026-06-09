// Headless equivalent of the Notesheet plugin's "Import .xlsx as
// Notesheet" command. Runs the same `xlsxBufferToSnapshot` the plugin
// runs at runtime, wraps the result in the Notesheet markdown fence,
// and POSTs a new note via Joplin's Web Clipper Data API.
//
// Why TypeScript? `xlsxBufferToSnapshot` lives in src/xlsx.ts and
// resolves via TypeScript path resolution. Re-implementing the import
// in JS would diverge from the runtime — the harness must call the
// SAME code the plugin calls. We compile this file via tsc on demand
// (see import-fixture.sh) and execute the JS output with Node.
//
// Usage (via the .sh wrapper):
//   ./scripts/pge/import-fixture.sh <fixture-name> [--title <title>]
//
// Args (after compile):
//   $1  fixture-name (file under tests/ExcelBaseTestData/formatting-testdata/)
//   --title <title>  optional note title (default: fixture name + timestamp)

import * as fs from 'fs';
import * as path from 'path';
import { execSync } from 'child_process';

import { xlsxBufferToSnapshot } from '../../src/xlsx';
import { wrapSnapshot } from '../../src/snapshot';

async function main(): Promise<void> {
    const args = process.argv.slice(2);
    const fixtureName = args[0];
    if (!fixtureName) {
        console.error('usage: import-fixture <fixture-name> [--title <title>]');
        process.exit(2);
    }
    const titleIdx = args.indexOf('--title');
    const customTitle = titleIdx >= 0 ? args[titleIdx + 1] : null;

    // After tsc compile, __dirname lives inside the temp OUT_DIR. The
    // shell wrapper passes the repo root via PGE_REPO_ROOT so we always
    // resolve fixtures against the source tree, not the compile dir.
    const repoRoot = process.env.PGE_REPO_ROOT || path.resolve(__dirname, '..', '..');
    // Search the canonical fixture roots in order. The first one matching
    // the requested fixture name wins. M17 added tests/fixtures/charts/ —
    // chart fixtures live there, the M12-era formatting fixtures live
    // under tests/ExcelBaseTestData/formatting-testdata/.
    const fixtureRoots = [
        path.join(repoRoot, 'tests', 'ExcelBaseTestData', 'formatting-testdata'),
        path.join(repoRoot, 'tests', 'fixtures', 'charts'),
    ];
    let fixturePath: string | null = null;
    for (const root of fixtureRoots) {
        const candidate = path.join(root, fixtureName);
        if (fs.existsSync(candidate)) {
            fixturePath = candidate;
            break;
        }
    }
    if (!fixturePath) {
        console.error(`fixture not found: ${fixtureName} (searched ${fixtureRoots.join(', ')})`);
        process.exit(1);
    }

    const buffer = fs.readFileSync(fixturePath);
    const snapshot = await xlsxBufferToSnapshot(buffer);

    const title = customTitle
        || `PGE ${fixtureName} ${new Date().toISOString().replace(/[:.]/g, '-')}`;
    const body = wrapSnapshot(snapshot);

    const tmp = path.join(require('os').tmpdir(), `pge-fixture-${process.pid}.md`);
    fs.writeFileSync(tmp, body);
    try {
        const apiScript = path.join(repoRoot, 'scripts', 'pge', 'joplin-api.js');
        const out = execSync(
            `node "${apiScript}" create-note --title "${title.replace(/"/g, '\\"')}" --body-file "${tmp}"`,
            { stdio: ['ignore', 'pipe', 'pipe'] },
        ).toString();
        process.stdout.write(out);
    } finally {
        fs.unlinkSync(tmp);
    }
}

main().catch((e: Error) => { console.error(e.message); process.exit(1); });
