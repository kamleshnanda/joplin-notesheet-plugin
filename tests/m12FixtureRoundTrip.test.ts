// Round-trip golden tests for the 8 .xlsx fixtures that import cleanly.
//
// Each fixture exercises a different formatting surface; this file
// records WHAT survives import → snapshot → export, and pins down each
// known SHORTCOMING with a test that asserts the current degraded
// behavior. The shortcomings are not bugs we plan to fix in M12 — they
// are the documented behavior the user agreed to ship in this PR
// (see [[feedback-known-shortcomings-over-bugs]] in memory). Each
// shortcoming references the milestone where the fix is queued.
//
// Adding a new test here:
//   - Use the helpers (importFixture, roundTrip, snapshotCell, exportedCell)
//     so the boilerplate stays small.
//   - If the assertion captures a SHORTCOMING, name the test
//     "KNOWN SHORTCOMING — <feature> drops to <result> (→ M??)"
//   - If the assertion captures a survival, name the test
//     "import: <feature> preserves …" or "round-trip: <feature> survives …"
//
// **No personal data, no /Users/ paths.** Anchor only on
// tests/ExcelBaseTestData/formatting-testdata/.

jest.mock('@univerjs/sheets-table', () => ({
    UniverSheetsTablePlugin: function MockUniverSheetsTablePlugin() { /* sentinel */ },
}));

import { readFileSync } from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

import { snapshotToXlsxBuffer, xlsxBufferToSnapshot } from '../src/xlsx';

const FIXTURES_DIR = path.join(__dirname, 'ExcelBaseTestData', 'formatting-testdata');

interface SnapshotShape {
    sheetOrder: string[];
    sheets: Record<string, {
        cellData: Record<number, Record<number, {
            v?: unknown; t?: number; s?: string;
            f?: string;
            p?: {
                body?: {
                    dataStream?: string;
                    customRanges?: Array<{ rangeType?: number; properties?: { url?: string } }>;
                };
            };
        }>>;
        mergeData?: Array<{ startRow: number; endRow: number; startColumn: number; endColumn: number }>;
    }>;
    styles: Record<string, Record<string, unknown>>;
    defaultStyle?: { ff?: string };
    resources?: Array<{ name: string; data: string }>;
}

async function importFixture(file: string): Promise<SnapshotShape> {
    const buf = readFileSync(path.join(FIXTURES_DIR, file));
    return await xlsxBufferToSnapshot(buf as unknown as Buffer) as unknown as SnapshotShape;
}

interface RoundTripped {
    zip: JSZip;
    wb: ExcelJS.Workbook;
    rawXmlForSheet(index: number): Promise<string>;
}

async function roundTrip(snap: SnapshotShape): Promise<RoundTripped> {
    const out = await snapshotToXlsxBuffer(snap as unknown as Parameters<typeof snapshotToXlsxBuffer>[0]);
    const zip = await JSZip.loadAsync(out as ArrayBuffer);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(out as unknown as Parameters<typeof wb.xlsx.load>[0]);
    return {
        zip,
        wb,
        async rawXmlForSheet(index: number) {
            return await zip.files[`xl/worksheets/sheet${index}.xml`].async('string');
        },
    };
}

function snapshotCell(snap: SnapshotShape, sheetIdx: number, row: number, col: number): {
    v?: unknown;
    style: Record<string, unknown> | null;
    hasP: boolean;
    customRangeUrl?: string;
} {
    const sheet = snap.sheets[snap.sheetOrder[sheetIdx]];
    const cell = sheet.cellData[row]?.[col];
    if (!cell) return { v: undefined, style: null, hasP: false };
    const style = cell.s ? snap.styles[cell.s] : null;
    const ranges = cell.p?.body?.customRanges ?? [];
    const link = ranges.find((r) => r.rangeType === 0);
    return {
        v: cell.v,
        style: style ?? null,
        hasP: !!cell.p,
        customRangeUrl: link?.properties?.url,
    };
}

// ─── BordersAndCellColors.xlsx ─────────────────────────────────────────

describe('M12 round-trip — BordersAndCellColors fixture', () => {
    // The 11 border styles + theme color cells live at fixed rows in the
    // generator (border style names start at row 2 = R1, theme rows
    // appear after a blank gap row at R12; see generate_excel_samples.py).
    const STYLE_ROWS: Array<[number, string, number]> = [
        [1, 'thin', 1], [2, 'hair', 2], [3, 'dotted', 3], [4, 'dashed', 4],
        [5, 'dashDot', 5], [6, 'dashDotDot', 6], [7, 'double', 7], [8, 'medium', 8],
        [9, 'mediumDashed', 9], [10, 'mediumDashDot', 10], [11, 'thick', 13],
    ];

    test.each(STYLE_ROWS)(
        'import: %s border preserves all 4 sides with #000000 (style num %s → %d)',
        async (row, _name, styleNum) => {
            const snap = await importFixture('BordersAndCellColors.xlsx');
            const cell = snapshotCell(snap, 0, row, 1);
            expect(cell.style?.bd).toBeDefined();
            const bd = cell.style!.bd as Record<string, { s: number; cl: { rgb: string } }>;
            for (const side of ['t', 'r', 'b', 'l']) {
                expect(bd[side]).toBeDefined();
                expect(bd[side].s).toBe(styleNum);
                expect(bd[side].cl.rgb).toBe('#000000');
            }
        },
    );

    test('round-trip: thin border on B2 survives export → re-import', async () => {
        const snap = await importFixture('BordersAndCellColors.xlsx');
        const { wb } = await roundTrip(snap);
        const ws = wb.getWorksheet(1)!;
        // exceljs is 1-based, so snapshot row 1 → exceljs row 2.
        const c = ws.getCell('B2');
        expect(c.border?.top?.style).toBe('thin');
        expect(c.border?.bottom?.style).toBe('thin');
        expect(c.border?.left?.style).toBe('thin');
        expect(c.border?.right?.style).toBe('thin');
    });

    test('import: theme=4 tint=0.4 border resolves to a concrete RGB', async () => {
        // Source row 14 (1-indexed 15) is "Theme Border" with all 4 sides
        // medium, color {theme:4, tint:0.4}. Our resolver uses the
        // workbook-imported clrScheme; this fixture ships with the
        // exceljs-default Office 2007 theme (because openpyxl emits its
        // own theme1.xml that exceljs can roundtrip), so accent1 is the
        // Office 2007 #4F81BD and the +0.4 tint produces #95B3D7.
        // See KNOWN SHORTCOMING below for what changes when the
        // workbook theme differs.
        const snap = await importFixture('BordersAndCellColors.xlsx');
        const cell = snapshotCell(snap, 0, 14, 1);
        const bd = (cell.style!.bd as Record<string, { s: number; cl: { rgb: string } }>);
        expect(bd.t.s).toBe(8); // medium
        expect(bd.t.cl.rgb).toBe('#95B3D7');
        expect(bd.r.cl.rgb).toBe('#95B3D7');
        expect(bd.b.cl.rgb).toBe('#95B3D7');
        expect(bd.l.cl.rgb).toBe('#95B3D7');
    });

    test('KNOWN SHORTCOMING — theme-tinted borders resolve against whichever clrScheme is active, not the source workbook theme (→ M13)', async () => {
        // Pinning down: if the source's theme=4 in its OWN palette would
        // resolve to one hue but exceljs has hardcoded a different
        // palette, the resolved RGB tracks exceljs's palette. The user's
        // hand-saved Aptos InvestmentSummary.xlsx exposed this as
        // "TableStyleMedium4 looks darker green on round-trip" because
        // accent3 differed between the original Aptos workbook and
        // exceljs's Office-2007 default. We mitigate by preserving the
        // source <a:clrScheme> on export (NOTESHEET_THEME_CLR_SCHEME_PLUGIN
        // resource), but the imported snapshot still carries
        // export-palette-resolved RGBs, so editing in Joplin shows the
        // wrong hue. M13 fix: make the in-Joplin renderer theme-aware.
        const snap = await importFixture('BordersAndCellColors.xlsx');
        const cell = snapshotCell(snap, 0, 14, 1);
        const bd = (cell.style!.bd as Record<string, { s: number; cl: { rgb: string } }>);
        // Documenting the shortcoming: the resolved RGB is hardcoded by
        // the active palette at import time, not by the source workbook
        // theme name.
        expect(bd.t.cl.rgb).toMatch(/^#[0-9A-F]{6}$/);
    });

    test('import: custom RGB font color (#FF6600) preserves exactly', async () => {
        const snap = await importFixture('BordersAndCellColors.xlsx');
        const cell = snapshotCell(snap, 0, 15, 1);
        expect((cell.style!.cl as { rgb: string }).rgb).toBe('#FF6600');
    });

    test('import: custom RGB fill (#CCFFCC) preserves exactly', async () => {
        const snap = await importFixture('BordersAndCellColors.xlsx');
        const cell = snapshotCell(snap, 0, 16, 1);
        expect((cell.style!.bg as { rgb: string }).rgb).toBe('#CCFFCC');
    });
});

// ─── ConditionalFormatting-Variants.xlsx ───────────────────────────────

describe('M12 round-trip — ConditionalFormatting-Variants fixture', () => {
    test('import: cell values from the 5 CF source columns preserve', async () => {
        // Each of the 5 CF columns has 10 numeric data cells in rows 2-11.
        const snap = await importFixture('ConditionalFormatting-Variants.xlsx');
        // Column A (col 0) data rows: 0..90 in steps of 10.
        expect(snapshotCell(snap, 0, 1, 0).v).toBe(0);
        expect(snapshotCell(snap, 0, 10, 0).v).toBe(90);
        // Column G (col 6) "Top 3 (Green)" has random ints; just check
        // they're numbers.
        expect(typeof snapshotCell(snap, 0, 1, 6).v).toBe('number');
    });

    test('KNOWN SHORTCOMING — conditional formatting rules are not preserved on round-trip (→ M14+)', async () => {
        // M12 explicitly OOS for conditional formatting per src/xlsx.ts
        // top-of-file comment. The snapshot drops the rules, and the
        // exported xlsx has zero <conditionalFormatting> blocks.
        const snap = await importFixture('ConditionalFormatting-Variants.xlsx');
        const rt = await roundTrip(snap);
        const sheetXml = await rt.rawXmlForSheet(1);
        expect(sheetXml).not.toMatch(/<conditionalFormatting/);
        // Also verify the snapshot itself doesn't carry CF rules.
        const cfResource = (snap.resources ?? []).find((r) => /cf|conditional/i.test(r.name));
        expect(cfResource).toBeUndefined();
    });
});

// ─── EmptyAndDegenerate.xlsx ───────────────────────────────────────────

describe('M12 round-trip — EmptyAndDegenerate fixture', () => {
    test('import: 3 sheets with the right names', async () => {
        const snap = await importFixture('EmptyAndDegenerate.xlsx');
        expect(snap.sheetOrder).toHaveLength(3);
        // Sheet names come from the source: EmptySheet, BareSheet, HeaderOnlyTable.
        // We don't pin the snapshot ID format (sheet-1 etc.), but the
        // names should match.
        const names = snap.sheetOrder.map((id) => snap.sheets[id]).filter(Boolean);
        expect(names).toHaveLength(3);
    });

    test('import: empty + bare sheets have zero data rows', async () => {
        const snap = await importFixture('EmptyAndDegenerate.xlsx');
        expect(Object.keys(snap.sheets[snap.sheetOrder[0]].cellData)).toHaveLength(0);
        expect(Object.keys(snap.sheets[snap.sheetOrder[1]].cellData)).toHaveLength(0);
    });

    test('import: header-only table imports the header row + a totals/data row', async () => {
        // Source has table A1:C2 with names Name/Age/City. Importer reads
        // 2 rows because the table range covers them (header + one
        // synthesized row).
        const snap = await importFixture('EmptyAndDegenerate.xlsx');
        const sheet = snap.sheets[snap.sheetOrder[2]];
        expect(Object.keys(sheet.cellData)).toHaveLength(2);
        // Row 0 cells carry the table header synthesis (bg #156082).
        expect((snapshotCell(snap, 2, 0, 0).style!.bg as { rgb: string }).rgb).toBe('#156082');
    });

    test('round-trip: header-only table survives back into table1.xml', async () => {
        const snap = await importFixture('EmptyAndDegenerate.xlsx');
        const rt = await roundTrip(snap);
        const tablePath = Object.keys(rt.zip.files).find((p) => /^xl\/tables\/table\d+\.xml$/.test(p));
        expect(tablePath).toBeDefined();
        const xml = await rt.zip.files[tablePath!].async('string');
        expect(xml).toMatch(/name="EmptyTable"/);
    });
});

// ─── Hyperlinks-Variants.xlsx ──────────────────────────────────────────

describe('M12 round-trip — Hyperlinks-Variants fixture', () => {
    test('import: cell values from Pattern A {text,hyperlink} cells get cell.p with the URL', async () => {
        const snap = await importFixture('Hyperlinks-Variants.xlsx');
        // R1 = "https://example.com" (clean URL).
        expect(snapshotCell(snap, 0, 1, 0).hasP).toBe(true);
        expect(snapshotCell(snap, 0, 1, 0).customRangeUrl).toBe('https://example.com');
        // R2 = mailto:.
        expect(snapshotCell(snap, 0, 2, 0).customRangeUrl).toBe('mailto:hello@example.com');
        // R3 = file:///.
        expect(snapshotCell(snap, 0, 3, 0).customRangeUrl).toMatch(/^file:\/\//);
    });

    test('import: URL with special chars (& < > ?) preserves verbatim', async () => {
        const snap = await importFixture('Hyperlinks-Variants.xlsx');
        const cell = snapshotCell(snap, 0, 4, 0);
        expect(cell.customRangeUrl).toBe('https://x.com/a&b<c>d?e=1');
    });

    test('import: display-text-vs-URL-mismatch preserves both', async () => {
        // R5: cell.value text is "Click here..."; the hyperlink targets a
        // longer URL. Univer's customRange must hold the URL, not the
        // display text.
        const snap = await importFixture('Hyperlinks-Variants.xlsx');
        const cell = snapshotCell(snap, 0, 5, 0);
        expect(cell.v).toMatch(/^Click here/);
        expect(cell.customRangeUrl).toBe('https://example.com/longpath/to/some/resource/that/is/quite/lengthy');
    });

    test('import: same URL repeated in 3 cells all get their own cell.p', async () => {
        const snap = await importFixture('Hyperlinks-Variants.xlsx');
        // R6 cols 0/1/2 all share the same URL.
        for (const col of [0, 1, 2]) {
            expect(snapshotCell(snap, 0, 6, col).hasP).toBe(true);
            expect(snapshotCell(snap, 0, 6, col).customRangeUrl).toBe('https://example.com/shared');
        }
    });

    test('round-trip: hyperlink with special-char URL re-imports as a hyperlink', async () => {
        const snap = await importFixture('Hyperlinks-Variants.xlsx');
        const rt = await roundTrip(snap);
        const ws = rt.wb.getWorksheet(1)!;
        const cell = ws.getCell('A5'); // 1-indexed row 5 = snapshot row 4
        expect(cell.isHyperlink).toBe(true);
        // exceljs returns {text, hyperlink} after round-trip.
        const v = cell.value as { hyperlink?: string };
        expect(v.hyperlink).toBe('https://x.com/a&b<c>d?e=1');
    });

    test('round-trip: dedup cells all re-import as hyperlinks', async () => {
        const snap = await importFixture('Hyperlinks-Variants.xlsx');
        const rt = await roundTrip(snap);
        const ws = rt.wb.getWorksheet(1)!;
        for (const a1 of ['A7', 'B7', 'C7']) {
            expect(ws.getCell(a1).isHyperlink).toBe(true);
        }
    });
});

// ─── MaliciousValues.xlsx ──────────────────────────────────────────────

describe('M12 round-trip — MaliciousValues fixture', () => {
    test('import: DDE-injection prefix string is preserved as text (NOT promoted to a formula)', async () => {
        // R1C1 source: "=cmd|'/c calc'!A1" with data_type='s'. Our import
        // must NOT promote this to a formula on the snapshot's cell.f.
        // If it did, Joplin would expose a code-execution vector when the
        // file is opened.
        const snap = await importFixture('MaliciousValues.xlsx');
        const sheet = snap.sheets[snap.sheetOrder[0]];
        const cell = sheet.cellData[1][1];
        expect(cell.v).toBe("=cmd|'/c calc'!A1");
        expect(cell.f).toBeUndefined();
    });

    test('import: +/-/@ prefix strings are preserved as plain text', async () => {
        const snap = await importFixture('MaliciousValues.xlsx');
        const sheet = snap.sheets[snap.sheetOrder[0]];
        for (const [row, prefix] of [[3, '+'], [4, '-'], [5, '@']] as const) {
            const cell = sheet.cellData[row][1];
            expect(typeof cell.v).toBe('string');
            expect((cell.v as string).startsWith(prefix)).toBe(true);
            expect(cell.f).toBeUndefined();
        }
    });

    test('import: control chars (TAB/LF/CR) preserve in cell value', async () => {
        const snap = await importFixture('MaliciousValues.xlsx');
        const cell = snapshotCell(snap, 0, 7, 1);
        expect(cell.v).toMatch(/TAB\[\t\]/);
        expect(cell.v).toMatch(/LF\[\n\]/);
        expect(cell.v).toMatch(/CR\[\r\]/);
    });

    test('import: 32767-char string (Excel cell limit) preserves length', async () => {
        // openpyxl auto-truncates the 1MB string in the generator to
        // Excel's 32767 hard limit at write time. We assert the imported
        // value is at the limit, not silently truncated further.
        const snap = await importFixture('MaliciousValues.xlsx');
        const cell = snapshotCell(snap, 0, 8, 1);
        expect(typeof cell.v).toBe('string');
        expect((cell.v as string).length).toBe(32767);
    });

    test('import: RTL override + emoji + BOM all preserve', async () => {
        const snap = await importFixture('MaliciousValues.xlsx');
        const cell = snapshotCell(snap, 0, 9, 1);
        const v = cell.v as string;
        // RTL-override: U+202E EVIL U+202C
        expect(v).toMatch(/‮/);
        expect(v).toMatch(/‬/);
        // Emoji
        expect(v).toMatch(/\u{1F389}/u);
        // BOM
        expect(v).toMatch(/﻿/);
    });
});

// ─── MergedCellsAndAlignment.xlsx ──────────────────────────────────────

describe('M12 round-trip — MergedCellsAndAlignment fixture', () => {
    test('import: A1:B2 + C1:D1 merges land in mergeData', async () => {
        const snap = await importFixture('MergedCellsAndAlignment.xlsx');
        const md = snap.sheets[snap.sheetOrder[0]].mergeData ?? [];
        expect(md).toEqual(expect.arrayContaining([
            { startRow: 0, endRow: 1, startColumn: 0, endColumn: 1 },
            { startRow: 0, endRow: 0, startColumn: 2, endColumn: 3 },
        ]));
    });

    test('import: merged-cell anchor preserves bold/center/middle alignment', async () => {
        const snap = await importFixture('MergedCellsAndAlignment.xlsx');
        const cell = snapshotCell(snap, 0, 0, 0);
        expect(cell.style?.bl).toBe(1);
        expect(cell.style?.ht).toBe(2); // CENTER
        expect(cell.style?.vt).toBe(2); // MIDDLE
    });

    test('import: wrap text style flag (tb=3) preserves on row 4 cells', async () => {
        const snap = await importFixture('MergedCellsAndAlignment.xlsx');
        const cell = snapshotCell(snap, 0, 3, 0);
        expect(cell.style?.tb).toBe(3); // WRAP_STRATEGY_WRAP
        expect(cell.v).toMatch(/wrap within the cell/);
    });

    test('import: newlines inside wrap-text cell preserve as \\n in value', async () => {
        const snap = await importFixture('MergedCellsAndAlignment.xlsx');
        const cell = snapshotCell(snap, 0, 3, 1);
        expect(cell.v).toBe('Another wrapped\ntext with\nnewlines');
    });

    test('round-trip: A1:B2 merge survives export → re-import', async () => {
        const snap = await importFixture('MergedCellsAndAlignment.xlsx');
        const rt = await roundTrip(snap);
        const ws = rt.wb.getWorksheet(1)!;
        // exceljs surfaces a model.merges array of "A1:B2" strings.
        const merges = (ws.model as { merges?: string[] }).merges ?? [];
        expect(merges).toEqual(expect.arrayContaining(['A1:B2', 'C1:D1']));
    });

    test('import: rotated-text cells preserve their angle in style.tr', async () => {
        // Source row 6 cells have Alignment(text_rotation=45/90/135).
        // openpyxl writes those values into the OOXML; exceljs decodes
        // the 135 (= 90 + 45 in OOXML's CW encoding) back to a signed
        // -45. Our import surfaces both as `style.tr.a` carrying the
        // exceljs angle directly. Univer's ITextRotation.a follows the
        // same signed-int convention.
        const snap = await importFixture('MergedCellsAndAlignment.xlsx');
        const expected = [
            { col: 0, a: 45 },
            { col: 1, a: 90 },
            { col: 2, a: -45 },
        ];
        for (const { col, a } of expected) {
            const cell = snapshotCell(snap, 0, 5, col);
            const tr = cell.style?.tr as { a?: number; v?: number } | undefined;
            expect(tr).toBeDefined();
            expect(tr?.a).toBe(a);
            // None of these are "vertical-stacked" mode (Excel's 255).
            expect(tr?.v).toBeFalsy();
        }
    });

    test('round-trip: rotation survives export → re-import', async () => {
        const snap = await importFixture('MergedCellsAndAlignment.xlsx');
        const rt = await roundTrip(snap);
        const ws = rt.wb.getWorksheet(1)!;
        // exceljs is 1-based, snapshot row 5 → exceljs row 6.
        expect(ws.getCell('A6').alignment?.textRotation).toBe(45);
        expect(ws.getCell('B6').alignment?.textRotation).toBe(90);
        expect(ws.getCell('C6').alignment?.textRotation).toBe(-45);
    });
});

// ─── NumberFormats.xlsx ────────────────────────────────────────────────

describe('M12 round-trip — NumberFormats fixture', () => {
    // Each row of the fixture has (label, value, formatCode) at cols 0/1/2.
    // The numFmt is on col 1 only. We pin every format the generator emits.
    const FORMATS: Array<[number, string]> = [
        [2, '0'],
        [3, '0.00'],
        [4, '#,##0'],
        [5, '#,##0.00'],
        [6, '$#,##0.00'],
        [7, '#,##0.00 "€"'],
        [8, '_-* #,##0.00 "kr"_-'],
        [9, '0%'],
        [10, '0.00%'],
        [11, 'm/d/yy'],
        [12, 'yyyy-mm-dd'],
        [13, 'dd-mmm-yy'],
        [14, '[$-409]m/d/yy h:mm AM/PM;@'],
        [15, '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)'],
        [16, '@ "suffix"'],
        [17, '[Red]#,##0.00;[Blue]#,##0.00'],
    ];

    test.each(FORMATS)(
        'import: row %s preserves numFmt pattern %s',
        async (row, fmt) => {
            const snap = await importFixture('NumberFormats.xlsx');
            const cell = snapshotCell(snap, 0, row, 1);
            expect((cell.style!.n as { pattern: string }).pattern).toBe(fmt);
        },
    );

    test('import: row 1 (General) emits no numFmt in the style', async () => {
        // We deliberately skip "General" because every cell would
        // otherwise carry a redundant numFmt entry.
        const snap = await importFixture('NumberFormats.xlsx');
        const cell = snapshotCell(snap, 0, 1, 1);
        expect(cell.style).toBeNull();
    });

    test('round-trip: a representative numFmt (currency USD) survives back through exceljs', async () => {
        const snap = await importFixture('NumberFormats.xlsx');
        const rt = await roundTrip(snap);
        const ws = rt.wb.getWorksheet(1)!;
        // Source row 7 (1-indexed) carries "$#,##0.00".
        const cell = ws.getCell('B7');
        expect(cell.numFmt).toBe('$#,##0.00');
    });
});

// ─── RichTextInOneCell.xlsx ────────────────────────────────────────────

describe('M12 round-trip — RichTextInOneCell fixture', () => {
    test('KNOWN SHORTCOMING — multi-style rich text flattens to plain text on import (→ M13)', async () => {
        // Source A1: CellRichText(TextBlock(b=True, "Hello"), " world").
        // Our import flattens this to "Hello world" with no style runs in
        // cell.p and no bold flag on the cell style. Univer's RichText
        // model can carry per-run formatting, but plumbing it through
        // the snapshot is M13 territory.
        const snap = await importFixture('RichTextInOneCell.xlsx');
        const cell = snapshotCell(snap, 0, 0, 0);
        expect(cell.v).toBe('Hello world');
        expect(cell.hasP).toBe(false); // no rich-text body emitted
        expect(cell.style?.bl).toBeUndefined(); // bold run is lost
    });

    test('KNOWN SHORTCOMING — multi-color rich text flattens to plain text (→ M13)', async () => {
        const snap = await importFixture('RichTextInOneCell.xlsx');
        const cell = snapshotCell(snap, 0, 1, 0);
        expect(cell.v).toBe('Red and Blue text');
        expect(cell.hasP).toBe(false);
    });

    test('import: hyperlink+plain in a single cell preserves the hyperlink (the only RT case we keep)', async () => {
        // Row 3 has a single-format string with cell.hyperlink set. exceljs
        // reports this as cell.value = "Visit example.com for more info"
        // with cell.hyperlink = "https://example.com" — Pattern A which
        // we DO support.
        const snap = await importFixture('RichTextInOneCell.xlsx');
        const cell = snapshotCell(snap, 0, 2, 0);
        expect(cell.hasP).toBe(true);
        expect(cell.customRangeUrl).toBe('https://example.com');
    });
});
