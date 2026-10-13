// M20 guard: the two scoped npm `overrides` in package.json must actually
// take effect.
//
// They fix advisories that upstream can't ship on its own:
//   - exceljs 4.4.0 (latest, upstream inactive) asks for uuid ^8.3.0;
//     we force 11.1.1 (GHSA-w5hq-g745-h8pq).
//   - @univerjs/core 0.23.0 pins nanoid exactly at 5.1.11; we force 5.1.16
//     (GHSA-28wg-ghj8-5hjv) until the Univer 1.x milestone.
// Nothing else would fail if an override quietly stopped applying: someone
// drops the `overrides` block, a package moves in the tree, or a later npm
// resolves differently. The tests would stay green and the advisory would
// just come back in `npm audit`. This test resolves each dependency
// *from the consumer's own directory*, the way Node and webpack do, and
// checks the version the consumer really gets.
//
// The webview half of the uuid fix (bundling exceljs from source so the
// override reaches the editor) is a webpack concern that jest never sees.
// It is proven live by scripts/pge/eval-import.js + eval-export.js.

import * as fs from 'fs';
import * as path from 'path';

// Directory of an installed package. Uses the package's main entry and walks
// up to its package.json, because some packages (e.g. @univerjs/core) don't
// expose `./package.json` in their `exports` map.
function packageDir(name: string, fromDir: string = process.cwd()): string {
    let dir = path.dirname(require.resolve(name, { paths: [fromDir] }));
    while (!fs.existsSync(path.join(dir, 'package.json')) || !isPackageRoot(dir, name)) {
        const parent = path.dirname(dir);
        if (parent === dir) throw new Error(`no package root found for ${name}`);
        dir = parent;
    }
    return dir;
}

function isPackageRoot(dir: string, name: string): boolean {
    return JSON.parse(fs.readFileSync(path.join(dir, 'package.json'), 'utf8')).name === name;
}

// The version of `dep` that `consumer` gets at runtime.
function resolvedVersion(dep: string, consumer: string): string {
    const pkgJson = path.join(packageDir(dep, packageDir(consumer)), 'package.json');
    return JSON.parse(fs.readFileSync(pkgJson, 'utf8')).version;
}

// Numeric [major, minor, patch]; these packages publish plain x.y.z.
function parseSemver(v: string): [number, number, number] {
    const [major, minor, patch] = v.split('.').map((n) => parseInt(n, 10));
    return [major, minor, patch];
}

// Same major.minor as the patched release, and patch at or above it.
function expectPatchedLine(actual: string, floor: string): void {
    const [aMajor, aMinor, aPatch] = parseSemver(actual);
    const [fMajor, fMinor, fPatch] = parseSemver(floor);
    expect(`${aMajor}.${aMinor}`).toBe(`${fMajor}.${fMinor}`);
    expect(aPatch).toBeGreaterThanOrEqual(fPatch);
}

describe('M20 scoped npm overrides take effect', () => {
    test('exceljs resolves a patched uuid, and nanoid is patched for @univerjs/core', () => {
        const uuid = resolvedVersion('uuid', 'exceljs');
        const nanoid = resolvedVersion('nanoid', '@univerjs/core');
        expectPatchedLine(uuid, '11.1.1');
        expectPatchedLine(nanoid, '5.1.16');
    });
});
