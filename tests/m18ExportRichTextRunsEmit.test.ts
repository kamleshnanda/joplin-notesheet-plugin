// Export-byte guard for RICH-TEXT multi-run emit (GAP 2).
//
// Prior coverage (m13RichText.test.ts round-trip) reloaded the exported
// buffer with exceljs and asserted on exceljs's parsed { richText: [...] }
// — a circular write-then-read where BOTH legs could be wrong together.
// This file instead unzips the buffer and asserts on the RAW OOXML in
// xl/sharedStrings.xml: exactly two <r> run elements, with an <rPr><b/>
// on the bold run and NO <b/> on the plain run.
//
// The exporter routes multi-run cells (extractRichTextRunsFromCellP,
// src/xlsx.ts) through exceljs's { richText } cell value, which is
// serialized as a shared string with one <si> containing multiple <r>
// runs. A regression that flattened the runs (single <r>, or a plain
// <t> with no runs) fails these assertions.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() {
        /* sentinel */
    },
}));

import JSZip from 'jszip';

import { snapshotToXlsxBuffer } from '../src/xlsx';

// Minimal snapshot with ONE cell (A1) carrying two textRuns — mirrors the
// shape m13RichText.test.ts asserts on import: "Hello"(bold) + " world"(plain).
// documentStyle.pageSize is finite-shaped like buildHyperlinkCellP produces.
function richTextSnapshot() {
    return {
        id: 'wb-1',
        name: 'Spreadsheet',
        appVersion: '0.1.0',
        locale: 'enUS',
        sheetOrder: ['s1'],
        styles: {},
        sheets: {
            s1: {
                id: 's1',
                name: 'Sheet1',
                cellData: {
                    0: {
                        0: {
                            v: 'Hello world',
                            p: {
                                documentStyle: { pageSize: { width: 100, height: 100 } },
                                body: {
                                    dataStream: 'Hello world\r\n',
                                    paragraphs: [{ startIndex: 11, paragraphStyle: {} }],
                                    sectionBreaks: [{ startIndex: 12 }],
                                    textRuns: [
                                        { st: 0, ed: 5, ts: { bl: 1 } },
                                        { st: 5, ed: 11 },
                                    ],
                                },
                            },
                        },
                    },
                },
                rowCount: 10,
                columnCount: 5,
                defaultColumnWidth: 73,
                defaultRowHeight: 19,
                mergeData: [],
                rowData: {},
                columnData: {},
            },
        },
    };
}

async function exportedSharedStrings(): Promise<string> {
    const exported = await snapshotToXlsxBuffer(
        richTextSnapshot() as unknown as Parameters<typeof snapshotToXlsxBuffer>[0],
    );
    const zip = await JSZip.loadAsync(exported);
    const ss = zip.files['xl/sharedStrings.xml'];
    expect(ss).toBeDefined();
    return ss.async('string');
}

describe('GAP 2 — rich-text multi-run emit in exported OOXML', () => {
    test('sharedStrings carries exactly two <r> run elements for the bold+plain cell', async () => {
        const xml = await exportedSharedStrings();
        // The <si> for A1 must contain two run elements (not flattened to one).
        const runs = xml.match(/<r>/g) ?? [];
        expect(runs.length).toBe(2);
    });

    test('the first run carries <rPr> with <b/>; the second (plain) run has no bold', async () => {
        const xml = await exportedSharedStrings();
        // Grab the single <si> block and split it into runs to inspect each.
        const si = /<si>[\s\S]*?<\/si>/.exec(xml)?.[0];
        expect(si).toBeDefined();
        const runMatches = si!.match(/<r>[\s\S]*?<\/r>/g) ?? [];
        expect(runMatches.length).toBe(2);

        // Run 0 is "Hello", bold: <rPr> present with a <b/> child.
        expect(runMatches[0]).toContain('Hello');
        expect(/<rPr>[\s\S]*<b\/?>/.test(runMatches[0])).toBe(true);

        // Run 1 is " world", plain: no <b/> bold marker.
        expect(runMatches[1]).toContain('world');
        expect(/<b\/?>/.test(runMatches[1])).toBe(false);
    });

    test('the run texts concatenate back to the full cell string', async () => {
        const xml = await exportedSharedStrings();
        const texts = (xml.match(/<t[^>]*>([\s\S]*?)<\/t>/g) ?? []).map((m: string) =>
            m.replace(/<\/?t[^>]*>/g, ''),
        );
        expect(texts.join('')).toBe('Hello world');
    });
});
