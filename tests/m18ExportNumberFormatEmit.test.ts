// Export-byte guard for NUMBER-FORMAT emit breadth (GAP 3).
//
// Prior coverage verified only a single numFmt through export. This file
// imports the NumberFormats.xlsx fixture (16 representative formats), then
// unzips the exported buffer and asserts that xl/styles.xml's <numFmts>
// carries the expected custom formatCode strings AND that the cells in
// xl/worksheets/sheet1.xml reference a style whose xf points at the right
// numFmtId. We assert a representative spread: a date, a currency/accounting
// code, a percent, and a custom text code.
//
// The exporter writes numFmt via applyStyleToCell (src/xlsx.ts:
//   const numFmt = style.n?.pattern; if (numFmt) cell.numFmt = numFmt;)
// exceljs allocates numFmtId >= 164 for each distinct custom code and emits
// a <numFmt formatCode="..."/> entry. A regression that dropped the pattern
// leaves the <numFmts> block missing those codes.
//
// Built-in formats (0%, 0.00% -> ids 9/10) are referenced by their built-in
// id and are NOT written into <numFmts>; those are guarded via the cell's xf
// numFmtId instead (see the built-in-percent test below).

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import { readFileSync } from 'fs';
import path from 'path';
import JSZip from 'jszip';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURE = path.join(__dirname, 'fixtures', 'formatting-testdata', 'NumberFormats.xlsx');

interface ExportedParts {
    stylesXml: string;
    sheet1Xml: string;
}

async function exportFixture(): Promise<ExportedParts> {
    const buf = readFileSync(FIXTURE);
    const snap = await xlsxBufferToSnapshot(buf as unknown as Buffer);
    const exported = await snapshotToXlsxBuffer(snap);
    const zip = await JSZip.loadAsync(exported);
    const stylesXml = await zip.files['xl/styles.xml'].async('string');
    const sheet1Xml = await zip.files['xl/worksheets/sheet1.xml'].async('string');
    return { stylesXml, sheet1Xml };
}

// Parse <numFmt numFmtId="N" formatCode="CODE"/> entries into a formatCode
// -> numFmtId map. formatCode values may contain XML-escaped chars (&quot;
// etc.) — we compare against the escaped form as it appears in the file.
function parseNumFmts(stylesXml: string): Map<string, number> {
    const map = new Map<string, number>();
    const matches = stylesXml.match(/<numFmt\s+numFmtId="\d+"\s+formatCode="[^"]*"\s*\/>/g) ?? [];
    for (const entry of matches) {
        const id = /numFmtId="(\d+)"/.exec(entry)?.[1];
        const code = /formatCode="([^"]*)"/.exec(entry)?.[1];
        if (id !== undefined && code !== undefined) map.set(code, Number(id));
    }
    return map;
}

describe('GAP 3 — number-format emit breadth in exported OOXML', () => {
    test('exported styles.xml <numFmts> carries the custom currency/date/percent/text format codes', async () => {
        const { stylesXml } = await exportFixture();
        const codes = parseNumFmts(stylesXml);
        const allCodes = Array.from(codes.keys());

        // Currency (USD custom): $#,##0.00
        expect(allCodes).toContain('$#,##0.00');
        // Date: yyyy-mm-dd
        expect(allCodes).toContain('yyyy-mm-dd');
        // Date: dd-mmm-yy
        expect(allCodes).toContain('dd-mmm-yy');
        // Currency EUR: #,##0.00 "€" (euro sign preserved, quotes escaped).
        expect(allCodes).toContain('#,##0.00 &quot;€&quot;');
        // Custom text code: @ "suffix" (quotes are XML-escaped in the file).
        expect(allCodes).toContain('@ &quot;suffix&quot;');
        // Conditional-color code: [Red]#,##0.00;[Blue]#,##0.00
        expect(allCodes).toContain('[Red]#,##0.00;[Blue]#,##0.00');
    });

    test('built-in percent formats (0%, 0.00%) are emitted as built-in numFmtIds 9/10, not custom <numFmt> entries', async () => {
        // 0% and 0.00% are OOXML built-in formats (ids 9 and 10). Excel and
        // exceljs reference them by their built-in id and do NOT write them
        // into <numFmts> (that block is only for custom codes >= 164). So the
        // guard here is "the cell's xf carries numFmtId 9/10", NOT "the code
        // string appears in <numFmts>" — asserting the latter would be a false
        // alarm (they legitimately never appear there).
        const { stylesXml, sheet1Xml } = await exportFixture();
        const codes = parseNumFmts(stylesXml);
        // The built-in codes must NOT have been re-allocated as custom entries.
        expect(Array.from(codes.keys())).not.toContain('0%');
        expect(Array.from(codes.keys())).not.toContain('0.00%');

        const cellXfsBlock = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? '';
        const xfs = cellXfsBlock.match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) ?? [];
        const numFmtIdOfCell = (ref: string): number => {
            const cell = new RegExp(`<c\\s+r="${ref}"[^>]*>`).exec(sheet1Xml)?.[0];
            expect(cell).toBeDefined();
            const s = /\bs="(\d+)"/.exec(cell!)?.[1];
            expect(s).toBeDefined();
            const xf = xfs[Number(s)];
            const id = /numFmtId="(\d+)"/.exec(xf)?.[1];
            expect(id).toBeDefined();
            return Number(id);
        };
        // B10 is the 0% cell, B11 is the 0.00% cell in the fixture.
        expect(numFmtIdOfCell('B10')).toBe(9);
        expect(numFmtIdOfCell('B11')).toBe(10);
    });

    test('accounting code round-trips its four-section structure', async () => {
        const { stylesXml } = await exportFixture();
        const codes = parseNumFmts(stylesXml);
        const allCodes = Array.from(codes.keys());
        // Excel accounting format, four sections. exceljs normalizes away the
        // redundant backslash-escaping of the literal parens on write
        // ("\(" -> "(" ), so we assert the NORMALIZED form actually emitted —
        // the "$" literal, the "-"?? zero placeholder, and all four sections
        // must survive (quotes are XML-escaped).
        const accounting =
            '_(&quot;$&quot;* #,##0.00_);_(&quot;$&quot;* (#,##0.00);_(&quot;$&quot;* &quot;-&quot;??_);_(@_)';
        expect(allCodes).toContain(accounting);
    });

    test('a formatted cell references an xf whose numFmtId matches its custom formatCode', async () => {
        const { stylesXml, sheet1Xml } = await exportFixture();
        const codes = parseNumFmts(stylesXml);
        const dateId = codes.get('yyyy-mm-dd');
        expect(typeof dateId).toBe('number');

        // Find the <cellXfs> index (s="N") whose xf carries that numFmtId.
        // cellXfs is an ordered list of <xf .../>; the s attribute on a cell
        // is a 0-based index into it.
        const cellXfsBlock = /<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? '';
        const xfs = cellXfsBlock.match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) ?? [];
        const matchingXfIndex = xfs.findIndex((xf) => {
            const idMatch = /numFmtId="(\d+)"/.exec(xf);
            return idMatch !== null && Number(idMatch[1]) === dateId;
        });
        expect(matchingXfIndex).toBeGreaterThanOrEqual(0);

        // B13 holds the yyyy-mm-dd date value (serial 45000) in the fixture.
        // Its cell must reference that style index via s="matchingXfIndex".
        const b13 = /<c\s+r="B13"[^>]*>/.exec(sheet1Xml)?.[0];
        expect(b13).toBeDefined();
        const sAttr = /\bs="(\d+)"/.exec(b13!)?.[1];
        expect(sAttr).toBeDefined();
        expect(Number(sAttr)).toBe(matchingXfIndex);
    });
});
