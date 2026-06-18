// M18 fix: workbooks whose relationship Targets are written ABSOLUTE
// (`Target="/xl/tables/table1.xml"`) must import. exceljs can't resolve an
// absolute rel target back to its loaded part, so its worksheet reconcile
// builds `tables[table.name]` over an `undefined` table and crashes with
// "Cannot read properties of undefined (reading 'name')". Excel itself and
// most tools write RELATIVE targets (`../tables/table1.xml`), which exceljs
// handles — so the same multi-sheet/multi-table layout imports fine when the
// targets are relative.
//
// FormulasAndStructuredRefs.xlsx uses absolute targets (openpyxl-authored)
// and used to fail; spreadsheet1.xlsx has the identical structure with
// relative targets and works. The fix normalizes absolute `/xl/...` (and
// `/_rels/...`) rel Targets to package-relative BEFORE exceljs loads.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import { xlsxBufferToSnapshot } from '../src/xlsx';

const FIX = path.join(
    __dirname,
    'fixtures',
    'formatting-testdata',
    'FormulasAndStructuredRefs.xlsx',
);

describe('M18 — absolute relationship Targets import cleanly', () => {
    test('FormulasAndStructuredRefs.xlsx (absolute /xl/... targets) imports without throwing', async () => {
        const buf = readFileSync(FIX);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        // Both sheets survive the import.
        const sheets = (snap as { sheets?: Record<string, unknown> }).sheets ?? {};
        expect(Object.keys(sheets).length).toBe(2);
    });

    test('both named tables round-trip into the snapshot table resource', async () => {
        const buf = readFileSync(FIX);
        const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
        const resources =
            (snap as { resources?: Array<{ name: string; data: string }> }).resources ?? [];
        const tableRes = resources.find((r) => r.name === 'SHEET_TABLE_PLUGIN');
        expect(tableRes).toBeDefined();
        // The two tables (one per sheet) are present in the table resource.
        const parsed = JSON.parse(tableRes!.data);
        const tableCount = Object.keys(parsed)
            .filter((k) => !k.startsWith('_'))
            .reduce((n, sheetId) => n + Object.keys(parsed[sheetId] ?? {}).length, 0);
        expect(tableCount).toBeGreaterThanOrEqual(2);
    });

    // REGRESSION GUARD (browser-renderer safety): the pre-load buffer
    // transforms run in the Joplin editor renderer, where Node's `Buffer` is
    // undefined. normalizeAbsoluteRelTargets() once emitted JSZip's
    // `type: 'nodebuffer'`, which threw "Buffer is not defined" on EVERY
    // import (surfaced as "Import failed: buffer not defined"). Node-only
    // jest can't faithfully simulate the browser (JSZip's own isBuffer()
    // touches the global), so we assert the SOURCE invariant instead: the
    // import-side zip transforms must produce 'arraybuffer', never
    // 'nodebuffer'. xlsx.ts:791 was the offender.
    test('import-path zip transforms never use nodebuffer (browser-safe)', () => {
        const src = readFileSync(path.join(__dirname, '..', 'src', 'xlsx.ts'), 'utf8');
        expect(src).not.toMatch(/generateAsync\(\{\s*type:\s*['"]nodebuffer['"]/);
    });
});
