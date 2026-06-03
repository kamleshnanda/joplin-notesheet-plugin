#!/usr/bin/env node
// Create a fresh Notesheet note whose body is the EXACT output of
// `newSpreadsheet`'s execute handler — i.e. `wrapSnapshot(emptySnapshot())`
// — by importing the same source helpers via ts-node-style transpile.
//
// Why not `create-empty-notesheet.sh`? That script writes `{}` inside the
// fence. The plugin's `extractSnapshot()` succeeds on `{}` (it's valid
// JSON and an object), so the editor loads an empty Univer workbook —
// the seed in `emptySnapshot()` is bypassed. To actually exercise the
// seed, the note's body must be `wrapSnapshot(emptySnapshot())`, which
// is what the New Spreadsheet command does at runtime. This script is
// the headless equivalent: same output, no UI.
//
// Usage:
//   node scripts/pge/create-seeded-notesheet.js [title]
//
// Env:
//   JOPLIN_TOKEN  Web Clipper token (else read from .claude/joplin-token.local).

const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');

// ts-strip via esbuild-register would be heavier than required. Instead we
// re-implement `emptySnapshot()` here in lockstep with src/snapshot.ts.
// If src/snapshot.ts changes, this file must change too — that coupling
// is intentional: the harness is allowed to know what the runtime does.
function emptySnapshot() {
    return {
        id: 'workbook-' + Date.now(),
        sheetOrder: ['sheet-1'],
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        styles: {
            'pge-smoke-red': { cl: { rgb: '#FF0000' } },
        },
        sheets: {
            'sheet-1': {
                id: 'sheet-1',
                name: 'Sheet1',
                cellData: {
                    0: { 0: { v: 'harness-smoke-OK', s: 'pge-smoke-red' } },
                },
                rowCount: 100,
                columnCount: 26,
                defaultColumnWidth: 73,
                defaultRowHeight: 19,
            },
        },
    };
}

function wrapSnapshot(snap) {
    return '```notesheet v=1\n' + JSON.stringify(snap) + '\n```';
}

const title = process.argv[2] || `PGE seeded smoke ${new Date().toISOString().replace(/[:.]/g, '-')}`;
const body = wrapSnapshot(emptySnapshot());

const tmp = path.join(require('os').tmpdir(), `pge-seeded-${process.pid}.md`);
fs.writeFileSync(tmp, body);
try {
    const out = execSync(
        `node "${path.join(__dirname, 'joplin-api.js')}" create-note --title "${title.replace(/"/g, '\\"')}" --body-file "${tmp}"`,
        { stdio: ['ignore', 'pipe', 'pipe'] },
    ).toString();
    process.stdout.write(out);
} finally {
    fs.unlinkSync(tmp);
}
