// M14 spike — golden-snapshot regression baseline.
//
// For each fixture under `tests/ExcelBaseTestData/formatting-testdata/`,
// capture `src/xlsx.ts:xlsxBufferToSnapshot`'s output as a JSON file under
// `tests/golden-snapshots/<fixture-name>.json`. The test then asserts the
// current parser still produces the same snapshot — any change to the
// exceljs-driven output will fail this test, making the diff auditable.
//
// Phase 2 will run the same goldens against `src/xlsxSheetJS.ts` and merge
// only when every divergence is documented + accepted.
//
// **First run captures the goldens.** If the JSON file does not exist, the
// test writes it and passes. Subsequent runs compare. To intentionally
// refresh a golden, delete the JSON and re-run. This is identical to the
// pattern used by Jest's built-in `toMatchSnapshot`, but explicit so the
// diff is more readable in PRs.

import * as fs from 'fs';
import * as path from 'path';

import { xlsxBufferToSnapshot } from '../src/xlsx';

const FORMATTING_DIR = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');
const GOLDEN_DIR = path.join(__dirname, 'golden-snapshots');

function listFixtures(dir: string): string[] {
    if (!fs.existsSync(dir)) return [];
    return fs
        .readdirSync(dir)
        .filter((f) => f.toLowerCase().endsWith('.xlsx'))
        .sort();
}

// Strip volatile fields (timestamps + Math.random suffixes) so re-runs don't
// churn. src/xlsx.ts uses `<prefix>-<base36-now>-<rand36>` for table ids,
// table-column ids, and hyperlink range ids; the workbook id is
// `workbook-<unix-ms>`. We scrub all four families to a deterministic
// placeholder while preserving ordering.
//
// The scrub is applied recursively on the JSON-stringified snapshot, then
// parsed back. Because table data lives inside a JSON-string-typed `data`
// field on resources, the scrub re-stringifies that nested string before
// reparsing — otherwise nested ids inside the `data` payload would survive
// untouched.
const VOLATILE_RES = [
    { re: /workbook-\d+/g, replacement: 'workbook-STABLE' },
    { re: /tbl-([A-Za-z0-9_-]+)-[a-z0-9]+-[a-z0-9]+/g, replacement: 'tbl-$1-STABLE' },
    { re: /tblcol-(\d+)-[a-z0-9]+-[a-z0-9]+/g, replacement: 'tblcol-$1-STABLE' },
    { re: /lnk-[a-z0-9]+-[a-z0-9]+/g, replacement: 'lnk-STABLE' },
];

function scrubVolatile(s: string): string {
    let out = s;
    for (const { re, replacement } of VOLATILE_RES) {
        out = out.replace(re, replacement);
    }
    return out;
}

function stableSnapshot(snap: Record<string, unknown>): Record<string, unknown> {
    const json = JSON.stringify(snap);
    const scrubbed = scrubVolatile(json);
    return JSON.parse(scrubbed);
}

const fixtures = listFixtures(FORMATTING_DIR);

describe('golden snapshots: src/xlsx.ts (Phase-2 regression baseline)', () => {
    if (!fs.existsSync(GOLDEN_DIR)) {
        fs.mkdirSync(GOLDEN_DIR, { recursive: true });
    }

    test.each(fixtures)('%s', async (fixture) => {
        const buf = fs.readFileSync(path.join(FORMATTING_DIR, fixture));
        let captured: Record<string, unknown>;
        try {
            const snap = await xlsxBufferToSnapshot(buf);
            captured = stableSnapshot(snap as Record<string, unknown>);
        } catch (e) {
            // The current parser deliberately surfaces typed errors for
            // fixtures it can't handle (charts that exceljs's reconcile
            // can't resolve, multi-sheet+multi-table workbooks that hit
            // exceljs's table-reduce crash). The golden captures the
            // error code so Phase 2 must preserve the same surface.
            const err = e as { name?: string; code?: string; message?: string };
            captured = {
                __importError: {
                    name: err.name ?? 'Error',
                    code: err.code ?? null,
                    // Don't capture err.message — it's English prose that
                    // could be tweaked for clarity in a Phase-2 follow-up
                    // without representing a regression.
                },
            };
        }
        const goldenPath = path.join(GOLDEN_DIR, fixture.replace(/\.xlsx$/i, '') + '.json');
        if (!fs.existsSync(goldenPath)) {
            fs.writeFileSync(goldenPath, JSON.stringify(captured, null, 2) + '\n');
            expect(captured).toEqual(captured);
            return;
        }
        const golden = JSON.parse(fs.readFileSync(goldenPath, 'utf8'));
        expect(captured).toEqual(golden);
    });
});
