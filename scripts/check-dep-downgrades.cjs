#!/usr/bin/env node
/**
 * Dependency-downgrade guard.
 *
 * Compares the working-tree package.json against the version committed at
 * git HEAD and FAILS (exit 1) if any dependency moved to a lower semver —
 * enforcing CONTRIBUTING.md's "downgrades are blocked by default" rule at
 * the earliest practical point (pre-commit) and again in CI.
 *
 * A downgrade that is genuinely intended can be allowed by listing it in
 * an `ALLOW_DOWNGRADE` env var (comma-separated `name` entries), which the
 * committer sets deliberately with a documented reason in the commit/PR —
 * never silently. Example:
 *   ALLOW_DOWNGRADE=some-pkg npm run check:deps
 *
 * Pure Node + git, no dependencies (it must run before installs). Uses
 * execFileSync with a fixed argv (no shell, no interpolation).
 */
const { execFileSync } = require('child_process');
const fs = require('fs');

function loadJson(text) {
    try {
        return JSON.parse(text);
    } catch {
        return null;
    }
}

// All dependency buckets we care about.
const BUCKETS = ['dependencies', 'devDependencies', 'optionalDependencies', 'peerDependencies'];

function collectDeps(pkg) {
    const out = {};
    if (!pkg) return out;
    for (const b of BUCKETS) {
        for (const [name, range] of Object.entries(pkg[b] || {})) {
            out[name] = range;
        }
    }
    return out;
}

// Strip a semver range down to a comparable [major, minor, patch]. Ranges
// like ^1.2.3, ~1.2.3, >=1.2.3, 1.2.3 all reduce to their base version.
// Returns null for non-comparable specs (git URLs, "latest", workspace:*).
function parseVersion(range) {
    if (typeof range !== 'string') return null;
    const m = range.match(/(\d+)\.(\d+)\.(\d+)/);
    if (!m) return null;
    return [Number(m[1]), Number(m[2]), Number(m[3])];
}

function cmp(a, b) {
    for (let i = 0; i < 3; i++) {
        if (a[i] !== b[i]) return a[i] - b[i];
    }
    return 0;
}

function main() {
    let headPkgText;
    try {
        // Fixed argv — no shell, no user input.
        headPkgText = execFileSync('git', ['show', 'HEAD:package.json'], {
            encoding: 'utf8',
            stdio: ['pipe', 'pipe', 'ignore'],
        });
    } catch {
        // No HEAD package.json (first commit, or not in git) — nothing to compare.
        console.log('check:deps — no committed package.json at HEAD; skipping.');
        return 0;
    }
    const headDeps = collectDeps(loadJson(headPkgText));
    const workDeps = collectDeps(loadJson(fs.readFileSync('package.json', 'utf8')));

    const allowed = new Set(
        (process.env.ALLOW_DOWNGRADE || '')
            .split(',')
            .map((s) => s.trim())
            .filter(Boolean),
    );

    const downgrades = [];
    for (const [name, workRange] of Object.entries(workDeps)) {
        const headRange = headDeps[name];
        if (!headRange) continue; // newly added — not a downgrade
        const a = parseVersion(headRange);
        const b = parseVersion(workRange);
        if (!a || !b) continue; // non-comparable spec
        if (cmp(b, a) < 0 && !allowed.has(name)) {
            downgrades.push({ name, from: headRange, to: workRange });
        }
    }

    if (downgrades.length === 0) {
        console.log('check:deps — no dependency downgrades. ✓');
        return 0;
    }

    console.error('\n✖ check:deps — BLOCKED: dependency downgrade(s) detected:\n');
    for (const d of downgrades) {
        console.error(`  ${d.name}: ${d.from}  →  ${d.to}`);
    }
    console.error(
        '\nDowngrades are blocked by default (see CONTRIBUTING.md → Dependency hygiene).\n' +
            'If a downgrade is genuinely intended, document the blocking reason in the\n' +
            'commit/PR and re-run with ALLOW_DOWNGRADE=<name>[,<name>...] set.\n',
    );
    return 1;
}

process.exit(main());
