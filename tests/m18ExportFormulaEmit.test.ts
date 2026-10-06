// Export-byte guard for FORMULA emit (GAP 1).
//
// Prior coverage only proved that IMPORTING a formula fixture "does not
// throw" — nothing asserted that the exported .xlsx actually carries the
// formula back out. This file unzips the buffer produced by
// snapshotToXlsxBuffer and asserts the RAW <f> elements in
// xl/worksheets/sheet1.xml, including a structured table reference
// (=SUM(Table1[Col1])) — the case most likely to be silently dropped.
//
// The source fixture FormulasAndStructuredRefs.xlsx carries (col F):
//   F2  =A2+B2                              (plain arithmetic)
//   F3  =SUM(A2:A11)                        (range formula)
//   F4  =SUM(Table1[Col1])                  (structured ref, whole column)
//   F5  =Table1[[#This Row],[Col1]]         (structured ref, #This Row)
//   F6  =SUM(Table1[[#All],[Col2]])         (structured ref, #All)
//   F13 =SUM(Table2[ColX])                  (cross-table structured ref)
//
// These assertions FAIL if the exporter blanked/dropped the <f> element
// (e.g. wrote only the cached <v>) or mangled the structured-ref text.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURE = path.join(
    __dirname,
    'fixtures',
    'formatting-testdata',
    'FormulasAndStructuredRefs.xlsx',
);

async function exportedSheet1Xml(): Promise<string> {
    const buf = readFileSync(FIXTURE);
    const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
    const exported = await snapshotToXlsxBuffer(snap);
    const zip = await JSZip.loadAsync(exported);
    const sheet1 = zip.files['xl/worksheets/sheet1.xml'];
    expect(sheet1).toBeDefined();
    return sheet1.async('string');
}

describe('GAP 1 — formula emit in exported OOXML', () => {
    test('plain and range formulas emit as <f> elements (not blanked to just <v>)', async () => {
        const xml = await exportedSheet1Xml();
        // Extract every formula element's text so we assert on real content,
        // not a truthy "some <f> exists".
        const formulas = (xml.match(/<f[^>]*>([^<]*)<\/f>/g) ?? []).map((m: string) =>
            m.replace(/<\/?f[^>]*>/g, ''),
        );
        expect(formulas).toContain('A2+B2');
        expect(formulas).toContain('SUM(A2:A11)');
    });

    test('structured table references survive export verbatim', async () => {
        const xml = await exportedSheet1Xml();
        const formulas = (xml.match(/<f[^>]*>([^<]*)<\/f>/g) ?? []).map((m: string) =>
            m.replace(/<\/?f[^>]*>/g, ''),
        );
        // Whole-column structured ref — the canonical GAP-1 case.
        expect(formulas).toContain('SUM(Table1[Col1])');
        // #This Row and #All variants must not be mangled.
        expect(formulas).toContain('Table1[[#This Row],[Col1]]');
        expect(formulas).toContain('SUM(Table1[[#All],[Col2]])');
        // Cross-table structured ref.
        expect(formulas).toContain('SUM(Table2[ColX])');
    });

    test('the structured-ref formula cell is emitted as a formula cell (has <f>), not a plain value cell', async () => {
        const xml = await exportedSheet1Xml();
        // Isolate the F4 cell markup and prove it carries an <f> child.
        // Guards against a regression that keeps only the cached result.
        const f4Start = xml.indexOf('r="F4"');
        expect(f4Start).toBeGreaterThan(-1);
        const f4Slice = xml.slice(f4Start, f4Start + 120);
        expect(f4Slice).toContain('<f>SUM(Table1[Col1])</f>');
    });
});
